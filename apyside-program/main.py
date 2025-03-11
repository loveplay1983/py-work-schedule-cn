import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QComboBox, QSpinBox, QPushButton, QFileDialog, QLabel)
from PySide6.QtCore import Qt
from datetime import datetime
import calendar
from common import generate_schedule, coworkers, MONTH_NAMES

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("生成排班表")
        self.setGeometry(100, 100, 450, 250)  # Adjusted size for better layout

        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(20)

        # Month selection
        self.month_label = QLabel("选择月份:")
        self.month_label.setStyleSheet("font-weight: bold; font-size: 16px; color: #2c3e50;")
        layout.addWidget(self.month_label)
        
        self.month_combo = QComboBox()
        for m in range(1, 13):
            self.month_combo.addItem(MONTH_NAMES[m-1], m)
        self.month_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                font-size: 14px;
                border: 2px solid #dfe6e9;
                border-radius: 5px;
            }
            QComboBox:hover {
                border-color: #3498db;
            }
        """)
        layout.addWidget(self.month_combo)

        # Year selection
        self.year_label = QLabel("选择年份:")
        self.year_label.setStyleSheet("font-weight: bold; font-size: 16px; color: #2c3e50;")
        layout.addWidget(self.year_label)
        
        self.year_spin = QSpinBox()
        self.year_spin.setRange(1900, 2100)
        self.year_spin.setStyleSheet("""
            QSpinBox {
                padding: 8px;
                font-size: 14px;
                border: 2px solid #dfe6e9;
                border-radius: 5px;
            }
            QSpinBox:hover {
                border-color: #3498db;
            }
        """)
        layout.addWidget(self.year_spin)

        # Generate button
        self.generate_button = QPushButton("生成排班表")
        self.generate_button.setStyleSheet("""
            QPushButton {
                padding: 12px;
                background-color: #3498db;
                color: white;
                border-radius: 6px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
                box-shadow: 0 4px 15px rgba(52, 152, 219, 0.4);
            }
        """)
        self.generate_button.clicked.connect(self.generate_schedule)
        layout.addWidget(self.generate_button)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #7f8c8d; font-size: 14px;")
        layout.addWidget(self.status_label)

        # Set defaults to current date
        current_date = datetime.now()
        self.month_combo.setCurrentIndex(current_date.month - 1)
        self.year_spin.setValue(current_date.year)

    def generate_schedule(self):
        month = self.month_combo.currentData()
        year = self.year_spin.value()
        self.status_label.setText("正在生成排班表...")
        QApplication.processEvents()  # Update GUI

        wb = generate_schedule(year, month, coworkers)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存排班表", f"schedule_{year}_{month}.xlsx", "Excel Files (*.xlsx)"
        )
        
        if file_path:
            wb.save(file_path)
            self.status_label.setText(f"排班表已保存至 {file_path}")
        else:
            self.status_label.setText("生成已取消")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())