# Tempform

第2組同儕評分系統。Flask + SQLite + Docker，port 5205，部署於 https://peer.dayspringmatsu.com 。

Always respond in English only. Do not use any Chinese characters in terminal output.

## 成員
- 001 劉浩晨（組長/admin）
- 002 蔡崇正
- 003 蕭淳勻
- 004 洪瑞鶯
- 005 蔡幸慧

## 路由
- /login 選擇代號登入
- /form 評分表單（001–005）
- /admin 管理後台（限 001）

## 部署
容器名 tempform，DB 掛載於 /opt/tempform/data/tempform.db。
