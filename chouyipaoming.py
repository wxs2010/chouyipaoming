import sys
import random
import openpyxl
import time

from BlurWindow.blurWindow import GlobalBlur
from PySide2.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QWidget, QVBoxLayout, QHBoxLayout, QGraphicsBlurEffect
from PySide2.QtCore import Qt, QTimer, QPoint, QEasingCurve, QPropertyAnimation
from PySide2.QtGui import QColor

class NamePickerApp(QMainWindow):
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.names = self.read_excel(file_path)
        self.last_name = None
        self.last_activity_time = 0
        self.moving = False
        self.drag_start_position = QPoint()

        # 获取屏幕宽度和高度
        screen_geometry = QApplication.primaryScreen().geometry()
        self.screen_width = screen_geometry.width()
        self.screen_height = screen_geometry.height()

        # 设置窗口大小和位置
        self.window_width = 150
        self.window_height = 100
        x_cord = (self.screen_width // 2) - (self.window_width // 2)
        y_cord = (self.screen_height // 2) - (self.window_height // 2) - 75
        self.setGeometry(x_cord, y_cord, self.window_width, self.window_height)

        # 创建中心部件和布局
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        # 设置窗口背景色和圆角
        GlobalBlur(self.winId(), Dark=True, QWidget=self)
        self.setStyleSheet("border-radius: 10px;background-color: rgba(252, 220, 249, 64)")

        # 创建按钮和标签
        self.result_label = QLabel("", self)
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setStyleSheet("color: #100006;font-size: 30px;font-family: 楷体;")
        layout.addWidget(self.result_label)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)

        self.button = QPushButton("点名", self)
        self.button.setFixedSize(60, 30)  # 调整按钮大小
        self.button.setStyleSheet("""
            background-color: #ffa6e1;
            color: #e956a5;
            font-size: 20px;
            border-radius: 5px;  /* 给按钮加上圆角 */
            font-family: 楷体; /* 修改为你想要的字体 */
        """)
        self.button.clicked.connect(self.show_random_name_with_animation)
        button_layout.addWidget(self.button)
        button_layout.addStretch(1)

        layout.addLayout(button_layout)

        self.setCentralWidget(central_widget)

        # 启动检查无操作的定时器
        self.idle_timer = QTimer(self)
        self.idle_timer.timeout.connect(self.check_idle)
        self.idle_timer.start(100)

        # 确保窗口始终在最顶层并且不在任务栏显示
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        # 定时器用于刷新窗口的置顶状态
        self.refresh_topmost_timer = QTimer(self)
        self.refresh_topmost_timer.timeout.connect(self.refresh_topmost)
        self.refresh_topmost_timer.start(1000)  # 每秒刷新一次
    def read_excel(self, file_path):
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            names = [cell.value for cell in sheet['A'] if cell.value is not None and cell.value.strip()]
            workbook.close()
            return names
        except Exception as e:
            raise IOError(f"读取 Excel 文件时发生错误: {e}")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
            self.last_activity_time = time.time()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton:
            new_pos = event.globalPos() - self.drag_start_position
            new_x = new_pos.x()
            new_y = new_pos.y()

            # 检查新位置是否超出屏幕范围
            if new_x < 0:
                new_x = 0
            elif new_x + self.window_width > self.screen_width:
                new_x = self.screen_width - self.window_width

            if new_y < -85:
                new_y = -85
            elif new_y + self.window_height > self.screen_height:
                new_y = self.screen_height - self.window_height

            self.move(new_x, new_y)
            self.last_activity_time = time.time()

    def show_random_name_with_animation(self):
        while True:
            new_name = random.choice(self.names)
            if new_name != self.last_name:
                break

        self.result_label.setText(new_name)
        self.last_name = new_name
        self.last_activity_time = time.time()

        # 开始按钮动画
        self.animate_button()

    def animate_button(self):
        animation = QPropertyAnimation(self.button, b"backgroundColor")
        animation.setDuration(200)
        animation.setStartValue(QColor("#ffa6e1"))
        animation.setEndValue(QColor("#ff6b6b"))
        animation.setEasingCurve(QEasingCurve.InOutQuad)
        animation.finished.connect(lambda: self.reset_button_color())
        animation.start()

    def reset_button_color(self):
        self.button.setStyleSheet("""
            background-color: #ffa6e1;
            color: #e956a5;
            font-size: 20px;
            border-radius: 5px;  
            font-family: 楷体;
        """)

    def clear_result(self):
        self.result_label.setText("？？？")
        self.last_activity_time = time.time()

    def animate_move_to_edge(self):
        if not self.moving:
            return

        current_y = self.y()
        target_y = -85
        delta = -30

        # 计算新位置
        new_y = current_y + delta
        new_y = min(new_y, target_y) if delta > 0 else max(new_y, target_y)

        # 更新窗口位置
        self.move(self.x(), new_y)

        # 递归调用实现动画效果
        if new_y != target_y:
            QTimer.singleShot(10, self.animate_move_to_edge)
        else:
            self.moving = False
            self.clear_result()

    def check_idle(self):
        current_time = time.time()
        if current_time - self.last_activity_time > 15:
            self.moving = True
            self.animate_move_to_edge()

    def refresh_topmost(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWin = NamePickerApp('data\\名单.xlsx')
    mainWin.show()
    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        print("程序被用户中断")



