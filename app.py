"""Web 客户端：上传 Excel → 选方向 → 流式翻译 → 下载。

启动：python app.py，浏览器开 http://127.0.0.1:8000
文件上传走原始请求体（避开 python-multipart 依赖）；进度走 SSE。
"""
import json
import os
import queue
import threading
import uuid
import importlib.util

import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse

# translate-auto.py 含连字符，无法直接 import，按文件路径加载
_spec = importlib.util.spec_from_file_location("translate_auto", os.path.join(os.path.dirname(__file__), "translate-auto.py"))
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
run_pipeline = _mod.run_pipeline

app = FastAPI()

EXCEL_DIR = os.path.join(os.path.dirname(__file__), "excel")
os.makedirs(EXCEL_DIR, exist_ok=True)

# 单任务守护：同一时刻只跑一个翻译
running = threading.Event()
# jobId -> {"q": Queue}
jobs: dict[str, dict] = {}

HTML = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>Excel 翻译</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 900px; margin: 24px auto; padding: 0 16px; }
  .row { margin: 12px 0; }
  pre { background: #f5f5f5; padding: 12px; border-radius: 6px; white-space: pre-wrap;
        word-break: break-word; max-height: 420px; overflow: auto; font-size: 13px; }
  .reasoning { color: #6a737d; }
  .content { color: #111; }
  button { padding: 8px 18px; }
  .err { color: #c00; }
  #status { margin: 8px 0; color: #0366d6; }
</style>
</head>
<body>
  <h2>Excel 翻译</h2>
  <div class="row">
    <input type="file" id="file" accept=".xlsx">
  </div>
  <div class="row">
    翻译方向：
    <select id="translate">
      <option>中文翻译成英文</option>
      <option>英文翻译成中文</option>
      <!--
      <option>中文翻译成日文</option>
      <option>中文翻译成繁体中文</option>
      -->
    </select>
  </div>
  <div class="row">
    <button id="start">开始翻译</button>
    <span id="status"></span>
  </div>
  <h4>思考过程</h4>
  <pre id="reasoning" class="reasoning"></pre>
  <h4>完整回复</h4>
  <pre id="content" class="content"></pre>

<script>
const $ = id => document.getElementById(id);
let es = null;

$("start").onclick = async () => {
  const file = $("file").files[0];
  if (!file) { alert("请先选择 xlsx 文件"); return; }
  const translate = $("translate").value;
  $("reasoning").textContent = "";
  $("content").textContent = "";
  $("status").textContent = "上传中...";
  $("start").disabled = true;

  let jobId;
  try {
    const url = `/upload?filename=${encodeURIComponent(file.name)}&translate=${encodeURIComponent(translate)}`;
    const res = await fetch(url, { method: "POST", body: file });
    if (res.status === 409) { $("status").textContent = "已有任务在跑，请稍后"; $("start").disabled = false; return; }
    if (!res.ok) { $("status").textContent = "上传失败 " + res.status; $("start").disabled = false; return; }
    jobId = (await res.json()).jobId;
  } catch (e) { $("status").textContent = "上传出错 " + e; $("start").disabled = false; return; }

  $("status").textContent = "翻译中...";
  es = new EventSource(`/stream/${jobId}`);
  es.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.type === "status") $("status").textContent = "阶段：" + m.stage + (m.total != null ? `（${m.total} 条）` : "");
    else if (m.type === "reasoning") { const el = $("reasoning"); el.append(m.text); el.scrollTop = el.scrollHeight; }
    else if (m.type === "content" || m.type === "chunk") { const el = $("content"); el.append(m.text); el.scrollTop = el.scrollHeight; }
    else if (m.type === "done") {
      es.close(); es = null;
      $("status").textContent = "完成，开始下载…";
      window.location.href = `/download/${encodeURIComponent(m.file)}`;
      $("start").disabled = false;
    } else if (m.type === "error") {
      es.close(); es = null;
      $("status").innerHTML = `<span class="err">错误：${m.message}</span>`;
      $("start").disabled = false;
    }
  };
  es.onerror = () => { es = null; $("status").textContent = "连接中断"; $("start").disabled = false; };
};
</script>
</body>
</html>
"""


@app.get("/")
def index() -> HTMLResponse:
    return HTMLResponse(HTML)


@app.post("/upload")
async def upload(request: Request):
    filename = request.query_params.get("filename") or "upload.xlsx"
    translate = request.query_params.get("translate") or "中文翻译成英文"

    if running.is_set():
        return JSONResponse({"error": "已有任务在跑"}, status_code=409)

    safe = os.path.basename(filename)
    data = await request.body()
    path = os.path.join(EXCEL_DIR, safe)
    with open(path, "wb") as f:
        f.write(data)

    job_id = uuid.uuid4().hex
    q: queue.Queue = queue.Queue()
    jobs[job_id] = {"q": q}

    def worker():
        running.set()
        try:
            run_pipeline(path, translate,
                         json_path=os.path.join(EXCEL_DIR, "translate.json"),
                         on_event=lambda ev: q.put(ev))
        except Exception:
            # run_pipeline 已 emit error；兜底再放一条，防 SSE 卡住
            q.put({"type": "error", "message": "pipeline 异常退出"})
        finally:
            running.clear()
            q.put({"type": "end"})

    threading.Thread(target=worker, daemon=True).start()
    return {"jobId": job_id}


@app.get("/stream/{job_id}")
def stream(job_id: str):
    job = jobs.get(job_id)
    if job is None:
        return JSONResponse({"error": "未知任务"}, status_code=404)
    q: queue.Queue = job["q"]

    def gen():
        while True:
            ev = q.get()
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"
            if ev.get("type") in ("done", "error", "end"):
                break

    return StreamingResponse(gen(), media_type="text/event-stream")


@app.get("/download/{filename}")
def download(filename: str):
    safe = os.path.basename(filename)
    path = os.path.join(EXCEL_DIR, safe)
    if not os.path.exists(path):
        return JSONResponse({"error": "文件不存在"}, status_code=404)
    return FileResponse(path, filename=safe)


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
