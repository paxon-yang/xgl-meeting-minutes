# XGL Meeting Minutes DOCX Service

Flask microservice — receives meeting minutes JSON → generates branded DOCX → returns binary file.

## 部署到 Render.com（免费）

1. 把 `docx_service/` 整个文件夹上传到 GitHub 新仓库
2. **注意：把公司模板 `模板.docx` 重命名为 `template.docx` 放入仓库根目录**
3. 打开 https://render.com → New → Web Service
4. 连接 GitHub 仓库
5. 配置如下：
   - **Name**: xgl-meeting-minutes
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
6. 点击 Deploy

部署完成后会得到一个 URL，如：`https://xgl-meeting-minutes.onrender.com`

## 在 n8n 中更新节点

将工作流中节点 **"11. Generate DOCX"** 的 URL 改为：
```
https://xgl-meeting-minutes.onrender.com/generate-minutes
```

## 测试

```bash
curl -X POST https://your-service.onrender.com/generate-minutes \
  -H "Content-Type: application/json" \
  -d '{
    "meeting_info": {"title_en": "Test Meeting", "title_zh": "测试会议", "date": "2026-03-24", "location": "Mine Site"},
    "minutes": {"overview": {"en": "Test", "zh": "测试"}, "key_decisions": [], "action_items": [], "pending_matters": []},
    "full_transcript": []
  }' --output test.docx
```

## API

**POST** `/generate-minutes`

Request body:
```json
{
  "meeting_info": {
    "title_en": "...", "title_zh": "...",
    "date": "YYYY-MM-DD", "location": "...",
    "participants": ["Name1", "Name2"]
  },
  "minutes": {
    "overview": {"en": "...", "zh": "..."},
    "key_decisions": [{"en": "...", "zh": "...", "speaker": "...", "evidence_time_range": "HH:MM:SS–HH:MM:SS"}],
    "action_items": [{"task_en": "...", "task_zh": "...", "owner": "...", "deadline": "YYYY-MM-DD"}],
    "pending_matters": [{"en": "...", "zh": "...", "speaker": "..."}]
  },
  "full_transcript": [
    {"speaker": "A", "start_time": "00:01:00", "end_time": "00:02:00", "en_text": "...", "zh_text": "..."}
  ]
}
```

Response: DOCX binary file

**GET** `/health` → `{"status": "ok", "template_found": true}`
