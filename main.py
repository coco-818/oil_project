import os
import sys

# 动态获取当前运行程序（或exe）所在的绝对目录
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 拼接出 JSON 文件的绝对路径
JSON_PATH = os.path.join(BASE_DIR, 'chemical_devices.json')

import json
import time

# 强制定位 Qt 插件路径（防止平台初始化报错） ----
import PyQt5
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins', 'platforms')
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QDialog, QLabel, QTextEdit,
                             QAbstractItemView, QHeaderView, QComboBox, QLineEdit, QMessageBox)
from PyQt5.QtCore import Qt

# 尝试引入 Excel 导出库
try:
    import openpyxl
except ImportError:
    pass

DATA_FILE = "JSON_PATH"
DEFAULT_DATA = {
    "储存设备": [
        {"id": "ST-01", "name": "1号原油常压储罐", "status": "运行中", "param": "液位: 12.5m | 温度: 25℃",
         "log": "2026-06-30 常规巡检：罐体无腐蚀，呼吸阀动作正常。 (巡检员：张师傅)"},
        {"id": "ST-02", "name": "2号精制汽油内浮顶罐", "status": "检修中", "param": "液位: 2.1m | 温度: 22℃",
         "log": "2026-07-01 异常上报：阻火器处发现微量渗漏，已开具工单转入紧急检修。 (填报人：李工)"}
    ],
    "反应设备": [
        {"id": "RE-01", "name": "加氢裂化一段反应器", "status": "运行中", "param": "压力: 15.2MPa | 温度: 380℃",
         "log": "2026-06-28 运行分析：催化剂床层温升正常。"}
    ],
    "输送设备": [
        {"id": "TR-01", "name": "常减压主进料高压离心泵", "status": "运行中", "param": "转速: 2950rpm | 轴承温度: 65℃",
         "log": "2026-06-27 机泵测振正常，润滑油位处于标准区间。"}
    ]
}

def init_data():
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_DATA, f, ensure_ascii=False, indent=4)


# 子窗口一：设备检修与状态调整子窗口 ---
class InspectionDialog(QDialog):
    def __init__(self, device_info, parent=None):
        super().__init__(parent)
        self.device_info = device_info
        self.initUI()

    def initUI(self):
        self.setWindowTitle(f"设备维保检修档案 - {self.device_info['name']}")
        self.setFixedSize(450, 420)
        layout = QVBoxLayout()

        layout.addWidget(QLabel(f"<b>设备位号(ID):</b> {self.device_info['id']}"))
        layout.addWidget(QLabel(f"<b>设备名称:</b> {self.device_info['name']}"))
        layout.addWidget(QLabel(f"<b>当前工艺参数:</b> {self.device_info['param']}"))

        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("<b>设备运行状态调整:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(["运行中", "停机", "检修中"])
        self.status_combo.setCurrentText(self.device_info['status'])
        status_layout.addWidget(self.status_combo)
        layout.addLayout(status_layout)

        layout.addWidget(QLabel("<b>厂区维保/巡检历史日志录入:</b>"))
        self.log_text = QTextEdit()
        self.log_text.setText(self.device_info['log'])
        layout.addWidget(self.log_text)

        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 保存修改并同步台账")
        self.btn_close = QPushButton("🔙 取消并返回")

        self.btn_save.clicked.connect(self.save_data)
        self.btn_close.clicked.connect(self.close)

        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_close)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def save_data(self):
        self.device_info['status'] = self.status_combo.currentText()
        self.device_info['log'] = self.log_text.toPlainText()
        self.accept()


# 子窗口二：新设备入库新增子窗口 ---
class AddDeviceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.new_device = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("⚙️ 新增厂区设备入库台账")
        self.setFixedSize(350, 250)
        layout = QVBoxLayout()

        self.txt_id = QLineEdit()
        self.txt_id.setPlaceholderText("请输入设备位号 (如: ST-04)")
        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("请输入设备名称 (如: 4号轻油罐)")
        self.txt_param = QLineEdit()
        self.txt_param.setPlaceholderText("请输入初始设计/工艺参数")

        layout.addWidget(QLabel("<b>设备编号(位号):</b>"))
        layout.addWidget(self.txt_id)
        layout.addWidget(QLabel("<b>设备名称:</b>"))
        layout.addWidget(self.txt_name)
        layout.addWidget(QLabel("<b>设计参数/技术指标:</b>"))
        layout.addWidget(self.txt_param)

        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定入库")
        btn_cancel = QPushButton("取消")
        btn_ok.clicked.connect(self.on_submit)
        btn_cancel.clicked.connect(self.close)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def on_submit(self):
        if not self.txt_id.text().strip() or not self.txt_name.text().strip():
            QMessageBox.warning(self, "警告", "设备编号和名称不能为空！")
            return
        self.new_device = {
            "id": self.txt_id.text().strip(),
            "name": self.txt_name.text().strip(),
            "status": "停机",
            "param": self.txt_param.text().strip() if self.txt_param.text().strip() else "暂无参数描述",
            "log": f"{time.strftime('%Y-%m-%d')} 新设备登记入库台账。"
        }
        self.accept()

# 主窗口：化工设备分级台账主系统 ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_category = "储存设备"
        self.load_all_data()
        self.initUI()

    def load_all_data(self):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            self.all_data = json.load(f)

    def save_all_data(self):
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.all_data, f, ensure_ascii=False, indent=4)

    def initUI(self):
        self.setWindowTitle("石化厂区化工设备分级台账管理系统 v2.0")
        self.resize(900, 550)

        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # 1. 顶部厂区设备分级导航栏
        nav_layout = QHBoxLayout()
        self.btn_store = QPushButton("🛢️ 储存设备名录")
        self.btn_react = QPushButton("⚗️ 反应设备名录")
        self.btn_trans = QPushButton(" Pump 输送设备名录")
        nav_layout.addWidget(self.btn_store)
        nav_layout.addWidget(self.btn_react)
        nav_layout.addWidget(self.btn_trans)

        self.btn_store.clicked.connect(lambda: self.switch_category("储存设备"))
        self.btn_react.clicked.connect(lambda: self.switch_category("反应设备"))
        self.btn_trans.clicked.connect(lambda: self.switch_category("输送设备"))
        main_layout.addLayout(nav_layout)

        # 2. 工具栏：搜索过滤 + 新增 + 报表导出
        tool_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 按设备编号/名称实时过滤过滤搜索...")
        self.search_input.textChanged.connect(self.refresh_table)  # 实时过滤
        tool_layout.addWidget(self.search_input, 2)

        self.btn_add = QPushButton("➕ 新增设备")
        self.btn_add.clicked.connect(self.add_device)
        self.btn_delete = QPushButton("❌ 设备报废")
        self.btn_delete.clicked.connect(self.delete_device)
        self.btn_export = QPushButton("📊 导出Excel盘点表")
        self.btn_export.setStyleSheet("background-color: #2E7D32; color: white;")
        self.btn_export.clicked.connect(self.export_to_excel)

        tool_layout.addWidget(self.btn_add)
        tool_layout.addWidget(self.btn_delete)
        tool_layout.addWidget(self.btn_export)
        main_layout.addLayout(tool_layout)

        # 3. 中间台账数据表格
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["设备位号(ID)", "设备名称", "生产运行状态", "实时工艺参数明细"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(self.open_inspection)  # 双击直接跳转
        main_layout.addWidget(self.table)

        # 4. 底部查看调阅按钮
        self.btn_inspect = QPushButton("🔍 调取设备详细维保子窗口")
        self.btn_inspect.setStyleSheet("background-color: #1E88E5; color: white; font-weight: bold; height: 35px;")
        self.btn_inspect.clicked.connect(self.open_inspection)
        main_layout.addWidget(self.btn_inspect)

        # 5. 审计日志栏
        self.status_label = QLabel("📋 系统审计日志: 台账系统初始化完成，等待厂区调度指令。")
        self.status_label.setStyleSheet("color: #555; background-color: #EEE; padding: 5px;")
        main_layout.addWidget(self.status_label)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        self.refresh_table()

    def switch_category(self, category_name):
        self.current_category = category_name
        self.search_input.clear()  # 切换分类时清空搜索栏
        self.refresh_table()
        self.log_status(f"用户切换了分类视图 -> 调取【{category_name}】分级台账名录")

    def refresh_table(self):
        devices = self.all_data.get(self.current_category, [])
        search_keyword = self.search_input.text().strip().lower()

        # 过滤筛选符合搜索条件的设备
        filtered_devices = []
        for dev in devices:
            if not search_keyword or search_keyword in dev["id"].lower() or search_keyword in dev["name"].lower():
                filtered_devices.append(dev)

        self.table.setRowCount(len(filtered_devices))

        # 渲染表格（缓存被过滤的数据索引，以便双击或点击时准确定位）
        self.displayed_devices = filtered_devices

        for row, dev in enumerate(filtered_devices):
            self.table.setItem(row, 0, QTableWidgetItem(dev["id"]))
            self.table.setItem(row, 1, QTableWidgetItem(dev["name"]))

            status_item = QTableWidgetItem(dev["status"])
            status_item.setTextAlignment(Qt.AlignCenter)
            if dev["status"] == "运行中":
                status_item.setBackground(PyQt5.QtGui.QColor(230, 245, 230))
                status_item.setForeground(PyQt5.QtGui.QColor(0, 128, 0))
            elif dev["status"] == "检修中":
                status_item.setBackground(PyQt5.QtGui.QColor(255, 230, 230))
                status_item.setForeground(PyQt5.QtGui.QColor(255, 0, 0))
            else:
                status_item.setBackground(PyQt5.QtGui.QColor(240, 240, 240))
                status_item.setForeground(PyQt5.QtGui.QColor(128, 128, 128))

            self.table.setItem(row, 2, status_item)
            self.table.setItem(row, 3, QTableWidgetItem(dev["param"]))

    def open_inspection(self):
        selected_row = self.table.currentRow()
        if selected_row < 0:
            self.log_status("❌ 操作驳回！请先在表格中选择一行化工设备。")
            return

        device_info = self.displayed_devices[selected_row]

        dialog = InspectionDialog(device_info, self)
        if dialog.exec_() == QDialog.Accepted:
            self.save_all_data()
            self.refresh_table()
            self.log_status(f"💾 数据留存：设备 [{device_info['id']}] 维保档案已同步写入本地。")
        else:
            self.log_status(f"🔙 审计事件：用户关闭了 [{device_info['id']}] 档案，未做修改。")

    def add_device(self):
        dialog = AddDeviceDialog(self)
        if dialog.exec_() == QDialog.Accepted and dialog.new_device:
            new_dev = dialog.new_device
            # 防止编号重复
            if any(d['id'] == new_dev['id'] for category in self.all_data.values() for d in category):
                QMessageBox.critical(self, "错误", f"设备位号 {new_dev['id']} 在全厂台账中已存在，无法重复录入！")
                return
            self.all_data[self.current_category].append(new_dev)
            self.save_all_data()
            self.refresh_table()
            self.log_status(
                f"➕ 资产变动：新设备 [{new_dev['id']} - {new_dev['name']}] 登记入库，归属于【{self.current_category}】")

    def delete_device(self):
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "提示", "请先选择需要申请报废的设备行！")
            return
        device_info = self.displayed_devices[selected_row]

        reply = QMessageBox.question(self, "报废确认",
                                     f"确定要将设备 [{device_info['name']}] 从当前设备台账中注销报废吗？此操作不可逆！",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.all_data[self.current_category].remove(device_info)
            self.save_all_data()
            self.refresh_table()
            self.log_status(f"❌ 资产变动：申请批准！设备 [{device_info['id']}] 已执行报废下线并清退台账。")

    def export_to_excel(self):
        try:
            wb = openpyxl.Workbook()
            # 移出默认工作表
            default_sheet = wb.active
            wb.remove(default_sheet)

            # 分品类导出成不同的工作表(Sheet)
            for category, devices in self.all_data.items():
                ws = wb.create_sheet(title=category)
                ws.append(["设备编号(位号)", "设备名称", "生产运行状态", "工艺参数明细", "维保/检修历史日志"])
                for dev in devices:
                    ws.append([dev["id"], dev["name"], dev["status"], dev["param"], dev["log"]])

            excel_name = f"全厂化工设备分级台账_{int(time.time())}.xlsx"
            wb.save(excel_name)
            QMessageBox.information(self, "成功", f"📊 盘点表导出成功！\n文件已保存至：\n{os.path.abspath(excel_name)}")
            self.log_status(f"📊 报表审计：成功导出全厂资产盘点表 -> 文件名: {excel_name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败，错误原因：{str(e)}")

    def log_status(self, message):
        timestamp = time.strftime('%H:%M:%S', time.localtime())
        self.status_label.setText(f"📋 系统审计日志: [{timestamp}] {message}")


if __name__ == "__main__":
    init_data()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())