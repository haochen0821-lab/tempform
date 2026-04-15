import os
import json
from datetime import datetime
from functools import wraps

from flask import (
    Flask, request, jsonify, render_template,
    redirect, url_for, session, abort
)
from flask_sqlalchemy import SQLAlchemy

DB_PATH = os.environ.get("DB_PATH", os.path.join(os.path.dirname(__file__), "tempform.db"))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "tempform-dev-secret-change-me")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

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
    code = session["code"]
    sub = Submission.query.filter_by(member_code=code).first()
    payload = json.loads(sub.payload) if sub else {}
    status = sub.status if sub else "new"
    return render_template(
        "form.html",
        code=code,
        name=MEMBERS[code]["name"],
        groups=GROUPS,
        rubric=RUBRIC,
        payload=payload,
        status=status,
    )


@app.route("/api/save", methods=["POST"])
@login_required
def api_save():
    code = session["code"]
    data = request.get_json(silent=True) or {}
    payload = data.get("payload") or {}
    action = data.get("action", "draft")  # draft / submit
    status = "submitted" if action == "submit" else "draft"

    sub = Submission.query.filter_by(member_code=code).first()
    if not sub:
        sub = Submission(member_code=code, member_name=MEMBERS[code]["name"])
        db.session.add(sub)
    sub.payload = json.dumps(payload, ensure_ascii=False)
    sub.member_name = MEMBERS[code]["name"]
    sub.status = status
    sub.updated_at = datetime.utcnow()
    db.session.commit()
    return jsonify({"ok": True, "status": status, "updated_at": sub.updated_at.strftime("%Y-%m-%d %H:%M:%S")})


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
    # 對每個 group 收集每個成員給的 Q3 5項平均分
    group_member_avgs = {g: [] for g in GROUPS}
    per_member = []  # 每位成員對每組的平均
    for s in subs:
        payload = json.loads(s.payload or "{}")
        groups_data = payload.get("groups", {})
        member_entry = {"code": s.member_code, "name": s.member_name, "groups": {}}
        for g in GROUPS:
            gd = groups_data.get(g, {})
            scores = gd.get("q3", {}) or {}
            vals = [scores.get(k, 0) for k in RUBRIC]
            avg = _avg(vals)
            if avg > 0:
                group_member_avgs[g].append(avg)
            member_entry["groups"][g] = avg
        per_member.append(member_entry)

    rankings = []
    for g in GROUPS:
        rankings.append({"group": g, "score": _avg(group_member_avgs[g]), "voters": len(group_member_avgs[g])})
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


@app.route("/healthz")
def healthz():
    return {"ok": True}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5205)), debug=True)
