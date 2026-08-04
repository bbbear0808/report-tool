import os
import sys
import importlib.util

def main():
    # 获取 app.py 路径
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(base_dir, "app.py")
    
    # 动态加载 app 模块
    spec = importlib.util.spec_from_file_location("app", app_path)
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    
    # 调用 app 的 main 函数
    app_module.main()

if __name__ == "__main__":
    main()
