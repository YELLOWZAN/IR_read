import sys
import cv2
import win32com.client
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QComboBox, QMessageBox)
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import QTimer, Qt

class CameraWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Windows Hello 画面读取工具")
        self.setGeometry(100, 100, 800, 600)
        
        # 初始化变量
        self.cap = None
        self.timer = QTimer()
        self.camera_list = self.get_camera_list()
        
        # 中心Widget和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 1. 功能选择区
        func_layout = QHBoxLayout()
        self.func_label = QLabel("功能选择：")
        self.func_combo = QComboBox()
        self.func_combo.addItems(["无红外补光(普通摄像头)", "有红外补光(红外摄像头)", "实时画面(普通摄像头)"])
        func_layout.addWidget(self.func_label)
        func_layout.addWidget(self.func_combo)
        main_layout.addLayout(func_layout)
        
        # 2. 摄像头选择区
        cam_layout = QHBoxLayout()
        self.cam_label = QLabel("摄像头选择：")
        self.cam_combo = QComboBox()
        cam_layout.addWidget(self.cam_label)
        cam_layout.addWidget(self.cam_combo)
        main_layout.addLayout(cam_layout)
        
        # 3. 控制按钮区
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始预览")
        self.stop_btn = QPushButton("停止预览")
        self.stop_btn.setEnabled(False)  # 初始禁用停止按钮
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        main_layout.addLayout(btn_layout)
        
        # 加载摄像头列表（确保在创建所有UI组件后调用）
        self.update_cam_combo()
        
        # 4. 画面显示区
        self.video_label = QLabel()
        self.video_label.setAlignment(Qt.AlignCenter)
        self.video_label.setText("请选择功能和摄像头，点击开始预览")
        main_layout.addWidget(self.video_label)
        
        # 信号与槽连接
        self.func_combo.currentTextChanged.connect(self.update_cam_combo)
        self.start_btn.clicked.connect(self.start_preview)
        self.stop_btn.clicked.connect(self.stop_preview)
        self.timer.timeout.connect(self.update_frame)

    def get_camera_list(self):
        """枚举摄像头，区分普通/红外，特别优化Windows Hello摄像头检测"""
        camera_dict = {"普通摄像头": [], "红外摄像头": []}
        
        # 添加备用检测方法：直接尝试连接不同的摄像头索引
        direct_cams = []
        max_cams_to_check = 10  # 尝试检查前10个摄像头索引
        
        # 1. 尝试直接使用OpenCV检测可用摄像头
        for i in range(max_cams_to_check):
            cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
            if cap.isOpened():
                direct_cams.append(f"摄像头 {i}")
                cap.release()
        
        # 2. 使用WMI查询所有视频相关设备，扩大查询范围
        try:
            wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            service = wmi.ConnectServer(".", "root\CIMV2")
            
            # 扩大查询范围，不仅限于Image类，还包括可能包含Windows Hello设备的其他类
            queries = [
                "SELECT * FROM Win32_PnPEntity WHERE PNPClass='Image'",
                "SELECT * FROM Win32_PnPEntity WHERE PNPClass='Camera'",
                "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%camera%'",
                "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%webcam%'",
                "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%hello%'",
                "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%红外%'"
            ]
            
            wmi_cams = []
            seen_names = set()  # 用于去重
            
            for query in queries:
                try:
                    results = service.ExecQuery(query)
                    for item in results:
                        name = str(item.Name)
                        if name not in seen_names:
                            seen_names.add(name)
                            wmi_cams.append(name)
                except Exception:
                    # 忽略单个查询失败
                    pass
            
            # 合并WMI检测到的摄像头并分类
            for cam_name in wmi_cams:
                name_lower = cam_name.lower()
                # 红外摄像头识别关键词扩展
                is_infrared = any(kw in name_lower for kw in 
                                 ["infrared", "ir", "红外", "hello", "windows hello", 
                                  "ir camera", "ir cam", "face", "人脸识别"])
                if is_infrared:
                    camera_dict["红外摄像头"].append(cam_name)
                else:
                    camera_dict["普通摄像头"].append(cam_name)
            
            # 3. 如果WMI没有找到足够的摄像头，添加OpenCV直接检测的结果
            if len(camera_dict["普通摄像头"]) == 0 and len(camera_dict["红外摄像头"]) == 0:
                # 如果没有检测到任何摄像头，使用直接检测的结果作为普通摄像头
                for cam in direct_cams:
                    camera_dict["普通摄像头"].append(cam)
        
        except Exception as e:
            # 如果WMI查询失败，使用OpenCV直接检测的结果
            QMessageBox.warning(self, "警告", f"设备枚举部分失败，尝试直接检测：{str(e)}")
            for cam in direct_cams:
                camera_dict["普通摄像头"].append(cam)
        
        # 4. 添加备选方法：手动编号的摄像头选项
        if len(camera_dict["红外摄像头"]) == 0:
            # 如果没有识别到红外摄像头，添加一些可能的编号选项
            for i in range(max_cams_to_check):
                camera_dict["红外摄像头"].append(f"尝试红外摄像头 {i}")
        
        return camera_dict

    def update_cam_combo(self):
        """根据所选功能更新摄像头下拉列表"""
        self.cam_combo.clear()
        selected_func = self.func_combo.currentText()
        if "无红外" in selected_func or "实时" in selected_func:
            cams = self.camera_list["普通摄像头"]
        else:
            cams = self.camera_list["红外摄像头"]
        
        if not cams:
            self.cam_combo.addItem("无可用摄像头")
            self.start_btn.setEnabled(False)
        else:
            self.cam_combo.addItems(cams)
            self.start_btn.setEnabled(True)

    def start_preview(self):
        """开始画面预览"""
        selected_func = self.func_combo.currentText()
        cam_name = self.cam_combo.currentText()
        if "无可用" in cam_name:
            QMessageBox.warning(self, "提示", "请先连接可用摄像头")
            return
        
        # 确定摄像头索引
        cam_idx = None
        
        # 检查是否为直接索引格式的摄像头名称
        if "摄像头 " in cam_name or "尝试红外摄像头 " in cam_name:
            # 尝试从摄像头名称中提取数字索引
            try:
                # 提取名称中的数字部分
                import re
                numbers = re.findall(r'\d+', cam_name)
                if numbers:
                    cam_idx = int(numbers[0])
            except (ValueError, IndexError):
                pass
        
        # 如果无法从名称中提取索引，则使用原始方法（通过列表索引）
        if cam_idx is None:
            try:
                if "无红外" in selected_func or "实时" in selected_func:
                    cam_idx = self.camera_list["普通摄像头"].index(cam_name)
                else:
                    cam_idx = self.camera_list["红外摄像头"].index(cam_name)
            except ValueError:
                # 如果仍然找不到，默认尝试索引0
                QMessageBox.warning(self, "提示", f"无法确定摄像头 '{cam_name}' 的索引，默认使用索引0")
                cam_idx = 0
        
        # 初始化摄像头捕获，尝试不同的API选项提高兼容性
        try:
            # 首先尝试DSHOW API（Windows推荐）
            self.cap = cv2.VideoCapture(cam_idx, cv2.CAP_DSHOW)
            
            # 如果DSHOW失败，尝试默认API
            if not self.cap.isOpened():
                self.cap.release()
                self.cap = cv2.VideoCapture(cam_idx)
        except:
            if not self.cap.isOpened():
                QMessageBox.critical(self, "错误", "无法打开所选摄像头")
                return
        
        # 红外摄像头设置分辨率（可根据设备调整）
        if "有红外" in selected_func:
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        
        # 启动定时器刷新画面
        self.timer.start(30)  # 30ms刷新一次（约33帧/秒）
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.func_combo.setEnabled(False)
        self.cam_combo.setEnabled(False)

    def update_frame(self):
        """更新画面到GUI"""
        ret, frame = self.cap.read()
        if not ret:
            self.stop_preview()
            QMessageBox.warning(self, "提示", "画面读取失败")
            return
        
        # 转换OpenCV帧为PyQt可用格式
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = frame_rgb.shape
        bytes_per_line = ch * w
        q_image = QImage(frame_rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(q_image).scaled(
            self.video_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        self.video_label.setPixmap(pixmap)

    def stop_preview(self):
        """停止预览"""
        if self.timer.isActive():
            self.timer.stop()
        if self.cap is not None and self.cap.isOpened():
            self.cap.release()
        self.video_label.clear()
        self.video_label.setText("预览已停止")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.func_combo.setEnabled(True)
        self.cam_combo.setEnabled(True)

    def closeEvent(self, event):
        """窗口关闭时释放资源"""
        self.stop_preview()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CameraWindow()
    window.show()
    sys.exit(app.exec_())