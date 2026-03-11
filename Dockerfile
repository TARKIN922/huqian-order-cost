# 使用轻量 Python 镜像
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 先复制依赖文件（利用 Docker 层缓存，代码改动不重装包）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制项目文件
COPY app.py .
COPY templates/ templates/

# 创建工作区目录
RUN mkdir -p workspace

# 暴露端口
EXPOSE 5000

# 用 gunicorn 生产模式启动，支持多并发
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--threads", "4", "--timeout", "300", "app:app"]
