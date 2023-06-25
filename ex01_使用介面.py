import subprocess
import sys
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, \
    QMessageBox, QProgressDialog, QTextEdit
from PyQt5.QtCore import Qt, QDate
from PyQt5 import QtWidgets, QtCore
from ex02_文件處理 import 文件讀取與輸出


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.檔案選擇 = False

    def initUI(self):
        # 創建一個 QVBoxLayout 佈局
        layout = QVBoxLayout()

        # 添加一個 QLabel 顯示標題
        title_label = QLabel("Output 自動輸出", self)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        layout.addWidget(title_label, alignment=Qt.AlignCenter)
        layout.addSpacing(120)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)

        # 將容器設置為中央控件
        self.setCentralWidget(container)

        self.setStyleSheet("""
                    QLabel {
                        color: #000000;
                        font-size: 12px;
                    }
                """)

        # 介面標題與大小
        self.setWindowTitle("Output自動輸出")
        self.setGeometry(600, 300, 1000, 350)

        # 建立 QLabel 用於顯示背景圖片
        background_label = QLabel(self)

        # 加載背景圖片
        pixmap = QPixmap('../公司Output文件/背景-1.jpg')

        # 設置 QLabel 的尺寸和背景圖片
        # self.width() 和 self.height() 分別返回視窗的寬度和高度，這樣可以確保背景圖片的大小與視窗相符
        # (x, y) 設置為 (0, 0)，它將位於視窗的左上角
        background_label.setGeometry(0, 0, self.width(), self.height())
        background_label.setPixmap(pixmap)
        background_label.setScaledContents(True)

        # 將背景圖片置於最底層
        background_label.lower()

        # 按鈕設置與大小
        # 在這個上下文中，self 是一個特殊的參數，它指向正在創建的類的實例。
        # 在 MainWindow 類的方法中，self 用於引用類的實例本身。
        button_排程 = QPushButton("Output檔輸出", self)
        button_排程.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        button_排程.setGeometry(400, 150, 200, 30)
        button_排程.clicked.connect(self.AutoOutput)

        button_檔案選擇 = QPushButton("排程檔選擇", self)
        button_檔案選擇.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        button_檔案選擇.setGeometry(400, 100, 200, 30)
        button_檔案選擇.clicked.connect(self.selectFile)

        # 添加一個 QLabel 顯示選擇的文件路徑
        self.file_label = QLabel(self)
        self.file_label.setStyleSheet(
            "font-size: 18px;border: 3px groove black; background-color: white; padding: 5px;")
        self.file_label.setMinimumSize(300, 10)
        layout.addWidget(self.file_label, alignment=Qt.AlignCenter)
        # 自動換行設定
        # self.file_label.setWordWrap(True)

        # 創建日期選擇的 DateRangePicker
        # 並將其添加到主視窗的布局中
        date_range_picker = DateRangePicker()
        layout.addWidget(date_range_picker)

    def AutoOutput(self):
        if not self.檔案選擇:
            QMessageBox.warning(self, "警告", "請先選擇文件！")
            return

        confirm = QtWidgets.QMessageBox.question(self, "確認", "即將開始產生Output文件，請確認日期與文件設定正確!",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if confirm == QtWidgets.QMessageBox.Yes:
            try:
                dialog = QProgressDialog("文件處理中，請稍後...", None, 0, 0, self)
                dialog.setWindowModality(Qt.WindowModal)
                dialog.setWindowTitle("Loading")
                dialog.setCancelButton(None)
                dialog.show()
                文件讀取與輸出(起始日期, 結束日期, self.檔案選擇)
                dialog.close()
                QMessageBox.information(self, '結果', '文件處理完成!')
            except Exception as e:
                # 異常處理
                QMessageBox.critical(self, "錯誤", f"發生錯誤：{str(e)}")
                print(f"錯誤：{str(e)}")

    def selectFile(self):
        global 文件路徑

        # 創建一個 QFileDialog 的實例，用於顯示文件對話框。
        文件選擇視窗 = QFileDialog()

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        self.檔案選擇, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.file_label.setText(f"選擇的文件：{self.檔案選擇}")


class DateRangePicker(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(20, 0, 270, 0)  # 設置元件之間的間距

        # 創建一個空的佈局元素作為間距
        spacer_item = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        layout.addItem(spacer_item)

        # 創建起始日期的 QDateEdit
        # QtWidgets.QDateEdit() 函數內可設置預設日期
        self.start_date_edit = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setFixedWidth(110)
        self.start_date_edit.setStyleSheet("font-size: 18px;font-Family: Times New Roman")
        layout.addWidget(self.start_date_edit)

        # 計算預設的結束日期（當天日期 + 6天）
        default_end_date = QtCore.QDate.currentDate().addDays(1)

        # 創建結束日期的 QDateEdit
        self.end_date_edit = QtWidgets.QDateEdit(default_end_date)
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setFixedWidth(110)
        self.end_date_edit.setStyleSheet("font-size: 18px;font-Family: Times New Roman")
        layout.addWidget(self.end_date_edit)

        self.date_label = QtWidgets.QLabel()
        self.date_label.setStyleSheet(
            "font-size: 16px;border: 3px double black; background-color: white; padding: 5px;")
        self.date_label.setMinimumSize(170, 40)  # 設置方框的最小大小
        layout.addWidget(self.date_label, alignment=Qt.AlignCenter)

        # 連接按鈕的點擊事件到槽函數
        self.start_date_edit.dateChanged.connect(self.updateDateRange)
        self.end_date_edit.dateChanged.connect(self.updateDateRange)

        self.初始化日期()

    def 初始化日期(self):
        global 起始日期, 結束日期
        # 在此自動取得預設日期
        初始化日期_起始日期 = QtCore.QDate.currentDate()
        初始化日期_結束日期 = 初始化日期_起始日期.addDays(1)

        # 將預設日期轉換為字串
        初始化日期_起始日期 = 初始化日期_起始日期.toString("yyyy/MM/dd")
        初始化日期_結束日期 = 初始化日期_結束日期.toString("yyyy/MM/dd")

        起始日期 = 初始化日期_起始日期
        結束日期 = 初始化日期_結束日期
        # 在方框中顯示預設日期
        self.date_label.setText(f'起始日期: {起始日期}\n結束日期: {結束日期}')

    def updateDateRange(self):
        global 起始日期, 結束日期
        # 獲取選擇的起始日期和結束日期
        起始日期 = self.start_date_edit.date().toString("yyyy/MM/dd")
        結束日期 = self.end_date_edit.date().toString("yyyy/MM/dd")

        self.date_label.setText(f'起始日期: {起始日期}\n結束日期: {結束日期}')


# 這是 Python 中的慣用語法，表示如果這個程式碼是直接被執行而不是被當作模組引入，則執行下面的程式碼塊
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
