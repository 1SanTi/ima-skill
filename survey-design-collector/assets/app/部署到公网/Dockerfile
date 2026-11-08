# 问卷系统 · Docker 镜像
# 构建（在「在线问卷系统」目录下执行）：
#   docker build -f 部署到公网/Dockerfile -t survey:latest .
# 运行：
#   docker run -d --name survey -p 80:8000 -v survey_data:/app/data --restart unless-stopped survey:latest
FROM python:3.12-slim
WORKDIR /app
COPY . /app/
ENV PORT=8000
EXPOSE 8000
CMD ["python3", "server.py"]
