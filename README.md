# Excel 内容翻译工具

## 如何变成 skill
```
 /document-skills:skill-creator 将该项目变成 skill，通过提示词完成翻译。比如：将文件 ./excel/translate.xlsx 由中文翻译成英文，新的文件保存到 ./translate-translated.xlsx 中。skill 用中文书写
```

## 通过 skill 翻译如何使用

创建文件夹 `excel` ，将要翻译的文件放入文件夹中 `translate.xlsx`
```
将文件 ./excel/translate.xlsx 由中文翻译成英文，新的文件保存到 ./translate-translated.xlsx 中
```

## 手动翻译如何使用

创建文件夹 `excel` ，将要翻译的文件放入文件夹中 `translate.xlsx`
### 1. 获取需要翻译的信息
```python
    path = "./excel/translate.xlsx"
    # 1. 获取翻译的 excel 文本到 1.txt 中
    print_translate(path)
```
### 2. 大模型翻译，提示词：将 @translate.json 中的 key 由中文翻译成英文，翻译结果写入对应的 value
### 3. Excel 翻译替换
```python
    # 3. 替换 Excel 中的文本
    excel_cell_replace("./1.txt", "./2.txt", path)
```

## Web 应用（translate-auto.py + app.py）

`translate-auto.py` 把「提取 → 大模型翻译 → 替换」封装为 `run_pipeline(path, translate)`，`app.py` 提供一个 Web 端：浏览器上传 xlsx、选择翻译方向、实时看到 LLM 流式进度、完成后自动下载 `*-translated.xlsx`。

### 本地运行

```bash
pip install fastapi uvicorn openai openpyxl python-dotenv
python app.py
# 浏览器打开 http://127.0.0.1:8000
```

`.env` 需配置：

```
OPENAI_API_KEY=...
BASE_URL=https://llm-4m24ghuzgkr9cr4e.cn-beijing.maas.aliyuncs.com/compatible-mode/v1
MODEL=deepseek-v3.2
```

### 流程说明

- 上传文件走原始请求体（不依赖 `python-multipart`），进度走 SSE。
- 首遍翻译 → 写 `translate.json`；二次翻译（value 仍含中文的条目，按位置对齐的数组回填，避免 LLM 改写 key 导致对不回）→ 重写 `translate.json` → 最后才生成 `*-translated.xlsx`。
- 同一时刻只允许一个翻译任务（进程内锁），正在跑时上传返回 409。

## 部署到服务器

### 本应用的部署约束（关键）

1. **必须单 worker**。任务锁和任务表是进程内内存状态，多 worker（`--workers 2` / gunicorn 多进程）会导致锁失效、状态错乱。只能 `--workers 1`。
2. **SSE 流式**。前置 nginx 必须关缓冲，否则进度要等整段才一次性吐出：`proxy_buffering off;` + `proxy_cache off;`，并放宽超时 `proxy_read_timeout 600s;`（LLM 翻译要几分钟）。
3. **写文件到工作目录**。上传的 xlsx 进 `excel/`、产物 `*-translated.xlsx` 写在进程工作目录，需设稳定的 `WorkingDirectory` 并定期清理。
4. **无鉴权**。当前 `/upload` 谁都能调，等于公费让别人刷 LLM 额度。公网暴露必须加鉴权（nginx basic auth / IP 白名单 / token）。
5. `.env` 要带上 `OPENAI_API_KEY`、`BASE_URL`、`MODEL`。

### 方式一：systemd + nginx（裸服务器推荐）

**1. 装依赖**
```bash
sudo apt install python3-venv nginx
python3 -m venv ~/venv && source ~/venv/bin/activate
pip install -r requirements.txt
# pip install fastapi uvicorn openai openpyxl python-dotenv
# pip freeze > requirements.txt
```

**2. 同步代码 + `.env` 到服务器**（`.env` 不要进 git）。

**3. systemd 服务** `/etc/systemd/system/excel-translate.service`：
```ini
[Unit]
Description=Excel Translate
After=network.target

[Service]
User=rick
WorkingDirectory=/home/rick/excel-translate
EnvironmentFile=/home/rick/excel-translate/.env
ExecStart=/home/rick/venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000 --workers 1
Restart=on-failure

[Install]
WantedBy=multi-user.target
```
> 仍绑 `127.0.0.1`，由前面 nginx 对外；`--workers 1` 必须保留。
```bash
sudo systemctl daemon-reload && sudo systemctl enable --now excel-translate
```

**4. nginx 反代** `/etc/nginx/sites-available/excel-translate`：
```nginx
server {
    listen 80;
    server_name your.domain.or.ip;

    auth_basic "Restricted";
    auth_basic_user_file /etc/nginx/.htpasswd;

    client_max_body_size 200m;        # xlsx 可能很大
    proxy_read_timeout 600s;          # LLM 翻译耗时

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_buffering off;          # SSE 必须
        proxy_cache off;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```
```bash
sudo htpasswd -c /etc/nginx/.htpasswd rick
sudo ln -s /etc/nginx/sites-available/excel-translate /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```
HTTPS：`sudo certbot --nginx -d your.domain`。

**5. 清理产物**（crontab）：
```cron
0 3 * * * find /home/rick/excel-translate -maxdepth 1 -name '*-translated.xlsx' -mtime +7 -delete
```

### 方式二：Docker

`Dockerfile`：
```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
```

`docker-compose.yml`：
```yaml
services:
  web:
    build: .
    ports:
      - "127.0.0.1:8000:8000"
    env_file: .env
    volumes:
      - ./excel:/app/excel
    restart: unless-stopped
```
> 仍只对本地暴露 8000，外网走同机 nginx 反代（配置同上）；`--workers 1` 同样保留。

### 上线前清单

- [ ] `.env` 已上传、含 `OPENAI_API_KEY/BASE_URL/MODEL`
- [ ] 服务起来后 `curl http://127.0.0.1:8000/` 返回 200
- [ ] nginx `proxy_buffering off` 已生效（进度能实时流）
- [ ] 已加鉴权或限定内网/IP 白名单
- [ ] `client_max_body_size` 够大（150MB 级 xlsx → 200m+）
- [ ] 产物清理 cron 已加
