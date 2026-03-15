# 腾讯云轻量服务器上传包（991x）

此目录用于上传到腾讯云轻量服务器部署：

- `991x/`：应用代码目录（已整理可上传）
- `install-on-lighthouse.sh`：服务器端一键安装/启动脚本
- `nginx-991x.conf`：Nginx 反向代理配置模板

## 1. 上传方式

将本目录中的 `991x` 上传到服务器，例如目标路径：

- `/opt/991x`

## 2. 服务器执行

```bash
cd /opt/991x
chmod +x ../install-on-lighthouse.sh
sudo ../install-on-lighthouse.sh
```

## 3. 访问

- 应用默认监听：`127.0.0.1:9910`
- 通过 Nginx 对外提供：`80/443`

## 4. 可选环境变量

- `PORT`：应用端口（默认 `9910`）
- `HOST`：监听地址（默认 `127.0.0.1`）
