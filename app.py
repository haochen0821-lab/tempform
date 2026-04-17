import os
import json
from datetime import datetime
from functools import wraps

from flask import (
    Flask, request, jsonify, render_template,
    redirect, url_for, session, abort
)
from flask_sqlalchemy import SQLAlchemy

try:
    from anthropic import Anthropic
except ImportError:
    Anthropic = None

DB_PATH = os.environ.get("DB_PATH", os.path.join(os.path.dirname(__file__), "tempform.db"))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "tempform-dev-secret-change-me")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
anthropic_client = Anthropic(api_key=ANTHROPIC_API_KEY) if (Anthropic and ANTHROPIC_API_KEY) else None

MEMBERS = {
    "001": {"name": "劉浩晨", "role": "admin"},
    "002": {"name": "蔡崇正", "role": "member"},
    "003": {"name": "蕭淳勻", "role": "member"},
    "004": {"name": "洪瑞鶯", "role": "member"},
    "005": {"name": "蔡幸慧", "role": "member"},
}

GROUPS = [f"第{i}組" for i in range(1, 7)]
RUBRIC = ["內容深度", "實務連結", "批判反思", "整合脈絡", "表達呈現"]


class Submission(db.Model):
    __tablename__ = "submissions"
    id = db.Column(db.Integer, primary_key=True)
    member_code = db.Column(db.String(8), unique=True, nullable=False, index=True)
    member_name = db.Column(db.String(32), nullable=False)
    payload = db.Column(db.Text, nullable=False, default="{}")
    status = db.Column(db.String(16), nullable=False, default="draft")  # draft / submitted
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            "member_code": self.member_code,
            "member_name": self.member_name,
            "status": self.status,
            "payload": json.loads(self.payload or "{}"),
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else None,
        }


class AISummary(db.Model):
    __tablename__ = "ai_summaries"
    id = db.Column(db.Integer, primary_key=True)
    payload = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


with app.app_context():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    db.create_all()


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "code" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        code = session.get("code")
        if not code or MEMBERS.get(code, {}).get("role") != "admin":
            return abort(403)
        return f(*args, **kwargs)
    return wrapper


@app.route("/")
def index():
    code = session.get("code")
    if not code:
        return redirect(url_for("login"))
    if MEMBERS.get(code, {}).get("role") == "admin":
        return redirect(url_for("admin_page"))
    return redirect(url_for("form_page"))


def _is_admin():
    code = session.get("code")
    return bool(code) and MEMBERS.get(code, {}).get("role") == "admin"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        code = (request.form.get("code") or "").strip()
        if code in MEMBERS:
            session["code"] = code
            session["name"] = MEMBERS[code]["name"]
            return redirect(url_for("index"))
        return render_template("login.html", members=MEMBERS, error="代號錯誤")
    return render_template("login.html", members=MEMBERS, error=None)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/form")
@login_required
def form_page():
    me = session["code"]
    target = request.args.get("as", "").strip()
    if target and target != me:
        if not _is_admin():
            return abort(403)
        if target not in MEMBERS:
            return abort(404)
        edit_code = target
        admin_editing = True
    else:
        edit_code = me
        admin_editing = False

    sub = Submission.query.filter_by(member_code=edit_code).first()
    payload = json.loads(sub.payload) if sub else {}
    status = sub.status if sub else "new"
    return render_template(
        "form.html",
        code=edit_code,
        name=MEMBERS[edit_code]["name"],
        groups=GROUPS,
        rubric=RUBRIC,
        payload=payload,
        status=status,
        admin_editing=admin_editing,
        viewer_is_admin=_is_admin(),
    )


def _save_submission(code, payload, status):
    sub = Submission.query.filter_by(member_code=code).first()
    if not sub:
        sub = Submission(member_code=code, member_name=MEMBERS[code]["name"])
        db.session.add(sub)
    sub.payload = json.dumps(payload, ensure_ascii=False)
    sub.member_name = MEMBERS[code]["name"]
    sub.status = status
    sub.updated_at = datetime.utcnow()
    db.session.commit()
    return sub


@app.route("/api/save", methods=["POST"])
@login_required
def api_save():
    me = session["code"]
    data = request.get_json(silent=True) or {}
    payload = data.get("payload") or {}
    action = data.get("action", "draft")  # draft / submit
    target = (data.get("target_code") or me).strip()

    if target != me:
        if not _is_admin():
            return jsonify({"error": "forbidden"}), 403
        if target not in MEMBERS:
            return jsonify({"error": "not found"}), 404

    status = "submitted" if action == "submit" else "draft"
    sub = _save_submission(target, payload, status)
    return jsonify({"ok": True, "status": status, "updated_at": sub.updated_at.strftime("%Y-%m-%d %H:%M:%S")})


def _empty_payload():
    return {
        "groups": {g: {"q1": "", "q2": "", "q3": {r: 0 for r in RUBRIC}} for g in GROUPS},
        "overall": {"best_group": "", "best_reason": "", "worst_group": "", "worst_reason": "", "self_review": ""},
    }


@app.route("/api/admin/member/<code>", methods=["DELETE"])
@admin_required
def api_admin_delete_member(code):
    if code not in MEMBERS:
        return jsonify({"error": "not found"}), 404
    sub = Submission.query.filter_by(member_code=code).first()
    if sub:
        db.session.delete(sub)
        db.session.commit()
    return jsonify({"ok": True})


@app.route("/api/admin/member/<code>/clear", methods=["POST"])
@admin_required
def api_admin_clear(code):
    if code not in MEMBERS:
        return jsonify({"error": "not found"}), 404
    data = request.get_json(silent=True) or {}
    scope = data.get("scope", "")  # "group:第1組" / "overall" / "all"
    sub = Submission.query.filter_by(member_code=code).first()
    if not sub:
        return jsonify({"ok": True, "noop": True})
    payload = json.loads(sub.payload or "{}")
    base = _empty_payload()
    payload.setdefault("groups", base["groups"])
    payload.setdefault("overall", base["overall"])

    if scope == "all":
        payload = base
    elif scope == "overall":
        payload["overall"] = base["overall"]
    elif scope.startswith("group:"):
        g = scope.split(":", 1)[1]
        if g not in GROUPS:
            return jsonify({"error": "bad group"}), 400
        payload["groups"][g] = {"q1": "", "q2": "", "q3": {r: 0 for r in RUBRIC}}
    else:
        return jsonify({"error": "bad scope"}), 400

    _save_submission(code, payload, sub.status)
    return jsonify({"ok": True})


@app.route("/admin")
@admin_required
def admin_page():
    return render_template("admin.html", groups=GROUPS, rubric=RUBRIC, members=MEMBERS)


@app.route("/api/admin/overview")
@admin_required
def api_admin_overview():
    rows = []
    subs = {s.member_code: s for s in Submission.query.all()}
    for code, info in MEMBERS.items():
        sub = subs.get(code)
        rows.append({
            "code": code,
            "name": info["name"],
            "role": info["role"],
            "status": sub.status if sub else "none",
            "updated_at": sub.updated_at.strftime("%Y-%m-%d %H:%M:%S") if sub and sub.updated_at else None,
        })
    return jsonify({"rows": rows})


@app.route("/api/admin/member/<code>")
@admin_required
def api_admin_member(code):
    if code not in MEMBERS:
        return jsonify({"error": "not found"}), 404
    sub = Submission.query.filter_by(member_code=code).first()
    if not sub:
        return jsonify({
            "code": code,
            "name": MEMBERS[code]["name"],
            "status": "none",
            "payload": {},
            "updated_at": None,
        })
    return jsonify(sub.to_dict())


def _avg(nums):
    nums = [n for n in nums if isinstance(n, (int, float)) and n > 0]
    return round(sum(nums) / len(nums), 2) if nums else 0.0


@app.route("/api/admin/rankings")
@admin_required
def api_admin_rankings():
    subs = Submission.query.all()
    group_member_avgs = {g: [] for g in GROUPS}
    group_rubric_scores = {g: {r: [] for r in RUBRIC} for g in GROUPS}
    per_member = []
    for s in subs:
        payload = json.loads(s.payload or "{}")
        groups_data = payload.get("groups", {})
        member_entry = {"code": s.member_code, "name": s.member_name, "groups": {}, "rubric_detail": {}}
        for g in GROUPS:
            gd = groups_data.get(g, {})
            scores = gd.get("q3", {}) or {}
            vals = [scores.get(k, 0) for k in RUBRIC]
            avg = _avg(vals)
            if avg > 0:
                group_member_avgs[g].append(avg)
            member_entry["groups"][g] = avg
            rd = {}
            for r in RUBRIC:
                v = scores.get(r, 0)
                if isinstance(v, (int, float)) and v > 0:
                    group_rubric_scores[g][r].append(v)
                rd[r] = v
            member_entry["rubric_detail"][g] = rd
        per_member.append(member_entry)

    rankings = []
    for g in GROUPS:
        score = _avg(group_member_avgs[g])
        rubric = {}
        for r in RUBRIC:
            a = _avg(group_rubric_scores[g][r])
            rubric[r] = {"avg": a, "pct": round(a / 5 * 100, 1)}
        rankings.append({
            "group": g,
            "score": score,
            "pct": round(score / 5 * 100, 1),
            "voters": len(group_member_avgs[g]),
            "rubric": rubric,
        })
    rankings.sort(key=lambda x: x["score"], reverse=True)

    # 收集 Q5 / Q6 整體評價
    best_votes = {}
    worst_votes = {}
    self_reviews = []
    best_reasons = []
    worst_reasons = []
    for s in subs:
        payload = json.loads(s.payload or "{}")
        overall = payload.get("overall", {})
        b = overall.get("best_group")
        w = overall.get("worst_group")
        if b:
            best_votes[b] = best_votes.get(b, 0) + 1
            if overall.get("best_reason"):
                best_reasons.append({"by": s.member_name, "group": b, "reason": overall.get("best_reason")})
        if w:
            worst_votes[w] = worst_votes.get(w, 0) + 1
            if overall.get("worst_reason"):
                worst_reasons.append({"by": s.member_name, "group": w, "reason": overall.get("worst_reason")})
        if overall.get("self_review"):
            self_reviews.append({"by": s.member_name, "text": overall.get("self_review")})

    return jsonify({
        "rankings": rankings,
        "per_member": per_member,
        "best_votes": best_votes,
        "worst_votes": worst_votes,
        "best_reasons": best_reasons,
        "worst_reasons": worst_reasons,
        "self_reviews": self_reviews,
    })


AI_SYSTEM_PROMPT = """你是課堂同儕評分的整合助理。第2組5位成員針對第1～6組（共6組）進行課堂表現評分。
你的任務是把多位成員對同一題的意見整合成一段簡潔的綜合意見。

請以**純 JSON** 格式回覆，不要有任何 Markdown、code block 或其他文字，schema 必須完全符合：
{
  "groups": {
    "第1組": {"highlights": "…", "improvements": "…"},
    "第2組": {"highlights": "…", "improvements": "…"},
    "第3組": {"highlights": "…", "improvements": "…"},
    "第4組": {"highlights": "…", "improvements": "…"},
    "第5組": {"highlights": "…", "improvements": "…"},
    "第6組": {"highlights": "…", "improvements": "…"}
  },
  "best_overall": "…",
  "worst_overall": "…",
  "self_reflection": "…"
}

寫作要求：
- 用繁體中文，每段約 60～140 字
- 找出多位成員提到的共同點，用「多數成員認為…」「部分成員指出…」呈現
- 若意見少於 2 則，誠實寫出單一意見即可
- 若該題完全沒人填寫，寫「（無資料）」
- 客觀中立，避免攻擊性用詞
- 直接給結論，不要說「綜合來看」之類的引言"""


@app.route("/api/admin/ai_summary", methods=["GET"])
@admin_required
def api_get_ai_summary():
    s = AISummary.query.order_by(AISummary.id.desc()).first()
    if not s:
        return jsonify({"summary": None, "created_at": None})
    return jsonify({
        "summary": json.loads(s.payload),
        "created_at": s.created_at.strftime("%Y-%m-%d %H:%M:%S"),
    })


@app.route("/api/admin/ai_summary", methods=["POST"])
@admin_required
def api_generate_ai_summary():
    if not anthropic_client:
        return jsonify({"error": "伺服器未設定 ANTHROPIC_API_KEY"}), 500

    subs = Submission.query.all()
    by_group = {g: {"q1": [], "q2": []} for g in GROUPS}
    best_reasons, worst_reasons, self_reviews = [], [], []
    for s in subs:
        p = json.loads(s.payload or "{}")
        for g in GROUPS:
            gd = (p.get("groups") or {}).get(g, {})
            if (gd.get("q1") or "").strip():
                by_group[g]["q1"].append(f"{s.member_name}：{gd['q1'].strip()}")
            if (gd.get("q2") or "").strip():
                by_group[g]["q2"].append(f"{s.member_name}：{gd['q2'].strip()}")
        ov = p.get("overall") or {}
        if ov.get("best_group") and (ov.get("best_reason") or "").strip():
            best_reasons.append(f"{s.member_name} 推薦 {ov['best_group']}：{ov['best_reason'].strip()}")
        if ov.get("worst_group") and (ov.get("worst_reason") or "").strip():
            worst_reasons.append(f"{s.member_name} 認為 {ov['worst_group']} 需優化：{ov['worst_reason'].strip()}")
        if (ov.get("self_review") or "").strip():
            self_reviews.append(f"{s.member_name}：{ov['self_review'].strip()}")

    sections = []
    for g in GROUPS:
        h = "\n".join(f"- {x}" for x in by_group[g]["q1"]) or "（無）"
        i = "\n".join(f"- {x}" for x in by_group[g]["q2"]) or "（無）"
        sections.append(f"## {g}\n【亮點意見】\n{h}\n【優化意見】\n{i}")
    sections.append("## Q5a 最優秀組別投票理由\n" + ("\n".join(f"- {x}" for x in best_reasons) or "（無）"))
    sections.append("## Q5b 最需優化組別投票理由\n" + ("\n".join(f"- {x}" for x in worst_reasons) or "（無）"))
    sections.append("## Q6 第2組（自組）反思\n" + ("\n".join(f"- {x}" for x in self_reviews) or "（無）"))
    user_content = "\n\n".join(sections)

    try:
        msg = anthropic_client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4096,
            system=AI_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_content}],
        )
    except Exception as e:
        return jsonify({"error": f"Anthropic API 錯誤：{e}"}), 500

    text = msg.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
        if text.startswith("json"):
            text = text[4:].strip()
    try:
        result = json.loads(text)
    except Exception:
        return jsonify({"error": "AI 回傳非 JSON 格式", "raw": text}), 500

    rec = AISummary(payload=json.dumps(result, ensure_ascii=False))
    db.session.add(rec)
    db.session.commit()
    # 只保留最近 10 筆
    old = AISummary.query.order_by(AISummary.id.desc()).offset(10).all()
    for o in old:
        db.session.delete(o)
    db.session.commit()
    return jsonify({"summary": result, "created_at": rec.created_at.strftime("%Y-%m-%d %H:%M:%S")})


@app.route("/healthz")
def healthz():
    return {"ok": True}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5205)), debug=True)
