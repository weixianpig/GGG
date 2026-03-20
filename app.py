import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QFileDialog, QLabel, QMessageBox
)
from openpyxl import load_workbook

# ====== 你的核心逻辑 ======
SLOT_COLS = [4, 7, 8, 9, 11, 13, 14, 17, 21, 22]

ROW_MAP = {
    "Full SKU": 11,
    "Partner SKU": 12,
    "Product Page(s)": 13,
    "Item Name": 14,
    "Category": 15,
    "SubCategory": 16,
    "Current Inventory": 17,
    "QTY Sold": 18,
    "Average First Cost": 19,
    "First Cost Sales": 20,
}

TEMPLATE_FILE = "fwtop10reportleslieneighbor/Jeffan - Top 10 - 1.8.26. KB v1.xlsx"


def safe_float(v):
    try:
        return float(v)
    except:
        return float("-inf")


def load_data(raw_file):
    wb = load_workbook(raw_file, data_only=True)
    ws = wb["Partner Central"]

    cols = []
    for c in range(4, ws.max_column + 1):
        if ws.cell(10, c).value:
            cols.append(c)

    items = []
    for c in cols:
        items.append({
            "Full SKU": ws.cell(10, c).value,
            "Partner SKU": ws.cell(11, c).value,
            "Product Page(s)": ws.cell(13, c).value,
            "Item Name": ws.cell(14, c).value,
            "Category": ws.cell(15, c).value,
            "SubCategory": ws.cell(17, c).value,
            "Current Inventory": ws.cell(18, c).value,
            "QTY Sold": ws.cell(19, c).value,
            "Average First Cost": ws.cell(20, c).value,
            "First Cost Sales": ws.cell(21, c).value,
        })

    items.sort(key=lambda x: safe_float(x["First Cost Sales"]), reverse=True)
    return items[:10]


def generate_report(raw_file, output_file):
    items = load_data(raw_file)

    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for i, item in enumerate(items):
        col = SLOT_COLS[i]
        for key, row in ROW_MAP.items():
            ws.cell(row=row, column=col).value = item.get(key)

    wb.save(output_file)


# ====== GUI ======
class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("gg")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        self.label = QLabel("No file selected")
        layout.addWidget(self.label)

        self.select_btn = QPushButton("Select Partner Central File")
        self.select_btn.clicked.connect(self.select_file)
        layout.addWidget(self.select_btn)

        self.generate_btn = QPushButton("Generate Report")
        self.generate_btn.clicked.connect(self.generate)
        layout.addWidget(self.generate_btn)

        self.setLayout(layout)

        self.file_path = None

    def select_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xlsx)")
        if file:
            self.file_path = file
            self.label.setText(file)

    def generate(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file first")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "Output.xlsx", "Excel Files (*.xlsx)")
        if not save_path:
            return

        try:
            generate_report(self.file_path, save_path)
            QMessageBox.information(self, "Success", "Report generated!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
