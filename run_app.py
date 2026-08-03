import os
import sys
import streamlit.web.bootstrap as bootstrap

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(base_dir, "app.py")
    sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false", "--server.headless=true", "--browser.serverAddress=localhost"]
    bootstrap.run()

if __name__ == "__main__":
    main()
