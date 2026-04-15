# Tempform 部署說明

- VPS 路徑：`/opt/tempform`
- Container：`tempform`
- Port：`5205`（原規劃 5204 已被 daypass 占用）
- DB：`/opt/tempform/data/tempform.db`
- URL：https://peer.dayspringmatsu.com

## 常用指令

```bash
# 查看狀態
docker ps | grep tempform
docker logs -f tempform

# 重新部署
cd /opt/tempform && git pull && docker build -t tempform . \
  && docker stop tempform && docker rm tempform \
  && docker run -d --name tempform --restart unless-stopped \
       -p 5205:5205 -v /opt/tempform/data:/data tempform

# 備份 DB
cp /opt/tempform/data/tempform.db ~/tempform_backup_$(date +%F).db
```

## GitHub Actions Secrets
- `VPS_HOST` = 152.42.205.215
- `VPS_USER` = root
- `VPS_SSH_KEY` = 私鑰內容
