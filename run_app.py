import os
import sys

def main():
    # 获取 app.py 路径
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(base_dir, "app.py")
    
    # 直接运行 app.py
    with open(app_path, "r", encoding="utf-8") as f:
        code = f.read()
    exec(code)

if __name__ == "__main__":
    main()
