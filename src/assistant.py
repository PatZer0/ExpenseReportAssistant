import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from ui import MainWindow

def resource_path(relative_path):
    """获取资源的绝对路径，兼容开发环境和PyInstaller打包后的环境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后的路径
        base_path = sys._MEIPASS
    else:
        # 开发环境下的路径
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    return os.path.join(base_path, relative_path)

def main():
    """程序入口函数"""
    app = QApplication(sys.argv)
    
    # 设置应用程序图标，使用兼容打包环境的路径
    icon_path = resource_path(os.path.join('icon', 'icon.ico'))
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    else:
        print(f"图标文件未找到: {icon_path}")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
