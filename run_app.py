import os
import sys
import subprocess
import time
import socket
import webbrowser
import atexit

def find_free_port():
    """找一个空闲端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def cleanup():
    """退出时清理"""
    print("程序已退出")

def main():
    atexit.register(cleanup)
    
    # 获取 app.py 路径
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(base_dir, "app.py")
    
    if not os.path.exists(app_path):
        print(f"错误：找不到 app.py，路径：{app_path}")
        input("按回车键退出...")
        return
    
    # 找个空闲端口
    port = find_free_port()
    
    print("=" * 50)
    print("  测试报告自动生成工具 正在启动...")
    print("=" * 50)
    print(f"  服务地址：http://localhost:{port}")
    print("  浏览器将自动打开，请勿关闭此窗口")
    print("=" * 50)
    
    # 启动 Streamlit 服务
    streamlit_cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
    ]
    
    try:
        # 启动 Streamlit 进程
        process = subprocess.Popen(
            streamlit_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True
        )
        
        # 等待服务启动，最多等 30 秒
        print("  等待服务启动", end="")
        for i in range(30):
            time.sleep(1)
            print(".", end="", flush=True)
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(0.5)
                result = sock.connect_ex(('localhost', port))
                sock.close()
                if result == 0:
                    print("\n")
                    print("  ✅ 服务已启动！正在打开浏览器...")
                    webbrowser.open(f"http://localhost:{port}")
                    break
            except:
                pass
        else:
            print("\n  ⚠️ 服务启动超时，请手动打开浏览器访问：")
            print(f"  http://localhost:{port}")
        
        # 持续打印 Streamlit 输出（方便调试）
        print("\n" + "=" * 50)
        print("  工具运行中...")
        print("  关闭此窗口即可退出程序")
        print("=" * 50 + "\n")
        
        for line in process.stdout:
            print(line, end="")
            
    except Exception as e:
        print(f"\n❌ 启动失败：{e}")
        input("\n按回车键退出...")

if __name__ == "__main__":
    main()
