# run_app.py —— 生产模式：禁用 dev 模式 + 指定静态资源 + 端口就绪后再开浏览器
import os, sys, socket, threading, time, webbrowser
from pathlib import Path

def base_dir() -> Path:
    # PyInstaller 解包目录优先
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).parent

def port_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.25)
        return s.connect_ex((host, port)) == 0

def find_free_port(host: str, start: int, tries: int = 20) -> int:
    p = start
    for _ in range(tries):
        if not port_in_use(host, p):
            return p
        p += 1
    return start

def open_browser_when_ready(url: str, timeout: float = 30.0):
    deadline = time.time() + timeout
    host = url.split("//", 1)[1].split(":")[0]
    port = int(url.rsplit(":", 1)[-1])
    while time.time() < deadline:
        if port_in_use(host, port):
            try: webbrowser.open(url)
            except Exception: pass
            return
        time.sleep(0.4)

def set_streamlit_static_path():
    """为打包产物绑定 Streamlit 的前端静态资源目录。"""
    import streamlit
    candidates = [
        Path(streamlit.__file__).parent / "static",
        base_dir() / "streamlit" / "static",
        Path(sys.executable).parent / "streamlit" / "static",
    ]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / "streamlit" / "static")

    for p in candidates:
        if p.exists():
            os.environ["STREAMLIT_STATIC_PATH"] = str(p)
            return p
    return None

def ensure_config_file(host: str, port: int):
    """在工作目录创建 .streamlit/config.toml，强制 production 配置。"""
    cfg_dir = base_dir() / ".streamlit"
    cfg_dir.mkdir(exist_ok=True)
    cfg = f"""
[global]
developmentMode = false

[server]
address = "{host}"
port = {port}
headless = true

[browser]
gatherUsageStats = false
"""
    (cfg_dir / "config.toml").write_text(cfg.strip(), encoding="utf-8")

def run():
    # 工作目录切到程序根，便于读写相对路径
    os.chdir(base_dir())

    # 让自定义包可被 import（兼容 _internal 情况）
    sys.path.insert(0, str(base_dir()))
    sys.path.insert(0, str(base_dir() / "_internal"))

    logs = base_dir() / "logs"
    logs.mkdir(exist_ok=True)

    # 入口脚本：兼容根目录或 _internal 目录
    candidates = [base_dir() / "streamlit_app.py", base_dir() / "_internal" / "streamlit_app.py"]
    app_path = next((p for p in candidates if p.exists()), None)
    if not app_path:
        (logs / "fatal.log").write_text("[FATAL] 未找到 streamlit_app.py\n", encoding="utf-8")
        raise FileNotFoundError("streamlit_app.py")

    host = os.environ.get("APP_ADDR", "127.0.0.1")
    port = find_free_port(host, int(os.environ.get("APP_PORT", "8501")), tries=30)
    url  = f"http://{host}:{port}"

    # 强制关闭开发模式（双保险：环境变量 + 配置文件）
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    ensure_config_file(host, port)

    # 绑定静态资源目录，避免回退到 Node dev server(3000)
    static_path = set_streamlit_static_path()

    # 记录日志
    with open(logs / "app.log", "a", encoding="utf-8") as f:
        f.write(f"[INFO] Starting at {url}\n[INFO] BaseDir={base_dir()}\n[INFO] StaticPath={static_path}\n")

    # 端口就绪再开一次浏览器
    threading.Thread(target=open_browser_when_ready, args=(url,), daemon=True).start()

    # 直接启动（不要走 CLI）
    try:
        from streamlit.web import bootstrap as st_bootstrap
    except Exception:
        import streamlit.bootstrap as st_bootstrap  # 极旧版兜底

    st_bootstrap.run(
        str(app_path),
        "",
        [],
        flag_options={
            "server.headless": True,
            "server.address": host,
            "server.port": port,
            "browser.gatherUsageStats": False,
            "server.fileWatcherType": "auto",
        },
    )

if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        try:
            with open(base_dir() / "logs" / "fatal.log", "a", encoding="utf-8") as f:
                f.write(f"[EXCEPTION] {e}\n")
        except Exception:
            pass
        raise
