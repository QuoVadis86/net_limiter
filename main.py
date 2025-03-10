import sys
import ctypes
import psutil
import subprocess
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QInputDialog,
                            QMessageBox, QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton)
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtCore import QObject, pyqtSignal, Qt

class SpeedInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("限速设置")
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("上传速度 (KB/s):"))
        self.upload_input = QLineEdit("100")
        layout.addWidget(self.upload_input)
        
        layout.addWidget(QLabel("下载速度 (KB/s):"))
        self.download_input = QLineEdit("100")
        layout.addWidget(self.download_input)
        
        self.confirm_btn = QPushButton("确认")
        self.confirm_btn.clicked.connect(self.accept)
        layout.addWidget(self.confirm_btn)
        
        self.setLayout(layout)

class TrayApp(QSystemTrayIcon):
    toggle_speed = pyqtSignal(bool)
    
    def __init__(self, icon, parent=None):
        super().__init__(icon, parent)
        self.setToolTip("网络限速控制器")
        self.active_limit = False
        
        # 创建上下文菜单
        self.menu = QMenu()
        
        self.toggle_action = self.menu.addAction("启用限速")
        self.toggle_action.triggered.connect(self.toggle_limit)
        
        self.settings_action = self.menu.addAction("参数设置")
        self.settings_action.triggered.connect(self.show_settings)
        
        self.menu.addSeparator()
        exit_action = self.menu.addAction("退出")
        exit_action.triggered.connect(self.exit_app)
        
        self.setContextMenu(self.menu)
        
        # 初始化配置
        self.process_name = "game.exe"
        self.upload_limit = 100  # KB/s
        self.download_limit = 100  # KB/s
        
        # 快捷键设置
        self.shortcut = QApplication.instance().shortcut = QKeySequence(
            Qt.CTRL + Qt.SHIFT + Qt.Key_L, self.toggle_limit)
        
    def toggle_limit(self):
        """切换限速状态"""
        self.active_limit = not self.active_limit
        self.toggle_speed.emit(self.active_limit)
        self.update_menu_status()
        
    def update_menu_status(self):
        self.toggle_action.setText("禁用限速" if self.active_limit else "启用限速")
        self.setIcon(QIcon("active.ico" if self.active_limit else "normal.ico"))
        
    def show_settings(self):
        """显示参数设置对话框"""
        dialog = QInputDialog()
        process_name, ok = QInputDialog.getText(
            None, "进程设置", "输入目标进程名:", text=self.process_name)
        if ok:
            self.process_name = process_name
            
        speed_dialog = SpeedInputDialog()
        if speed_dialog.exec_():
            try:
                self.upload_limit = int(speed_dialog.upload_input.text())
                self.download_limit = int(speed_dialog.download_input.text())
            except ValueError:
                QMessageBox.warning(None, "错误", "请输入有效的数字")

    def exit_app(self):
        """清理退出"""
        NetLimiter.remove_policy()
        QApplication.quit()

class NetLimiter(QObject):
    @staticmethod
    def set_limit(process_name, upload_kbps, download_kbps):
        """设置网络限制"""
        exe_path = NetLimiter.find_process_exe(process_name)
        if not exe_path:
            return False, "进程未找到"
        
        # 转换为bps（1KBps = 1024 Bytes/s = 8192 bits/s）
        upload_bps = upload_kbps * 8192
        download_bps = download_kbps * 8192
        
        # 创建上传策略
        ps_cmd = f'''
        New-NetQosPolicy -Name "UL_GameLimiter" -AppPathName "{exe_path}" `
        -ThrottleRateActionBitsPerSecond {upload_bps} -PolicyStore ActiveStore -Confirm:$false;
        New-NetQosPolicy -Name "DL_GameLimiter" -AppPathName "{exe_path}" `
        -ThrottleRateActionBitsPerSecond {download_bps} -PolicyStore ActiveStore -Confirm:$false
        '''
        return NetLimiter.execute_powershell(ps_cmd)
    
    @staticmethod
    def remove_policy():
        """移除所有策略"""
        ps_cmd = '''
        Remove-NetQosPolicy -Name "UL_GameLimiter" -Confirm:$false;
        Remove-NetQosPolicy -Name "DL_GameLimiter" -Confirm:$false
        '''
        return NetLimiter.execute_powershell(ps_cmd)
    
    @staticmethod
    def find_process_exe(process_name):
        for proc in psutil.process_iter(['name', 'exe']):
            if proc.info['name'] == process_name:
                return proc.info['exe']
        return None
    
    @staticmethod
    def execute_powershell(command):
        result = subprocess.run(
            ["powershell", "-Command", command],
            capture_output=True,
            text=True
        )
        return result.returncode == 0, result.stderr

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()
        
    app = QApplication(sys.argv)
    app.setApplicationName("GameNetLimiter Pro")
    app.setQuitOnLastWindowClosed(False)
    
    # 初始化托盘图标
    tray_icon = TrayApp(QIcon("normal.ico"))
    tray_icon.show()
    
    # 连接信号
    def handle_toggle(enable):
        if enable:
            success, msg = NetLimiter.set_limit(
                tray_icon.process_name,
                tray_icon.upload_limit,
                tray_icon.download_limit
            )
            if not success:
                QMessageBox.critical(None, "错误", f"启用限制失败: {msg}")
                tray_icon.active_limit = False
                tray_icon.update_menu_status()
        else:
            NetLimiter.remove_policy()
    
    tray_icon.toggle_speed.connect(handle_toggle)
    
    sys.exit(app.exec_())