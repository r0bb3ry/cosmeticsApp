import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QComboBox, QTableWidget, QTableWidgetItem, QPushButton, QFileDialog
from PyQt5 import uic
import mysql.connector
from datetime import datetime
# from docx import Document # убираем docx
# from docx.shared import Inches # убираем docx
import os
import csv

class TeamStatsWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('team_stats.ui', self)

        self.mydb = None
        self.connect_to_db()

        self.comboBox_year = self.findChild(QComboBox, "comboBox_year")
        self.tableWidget = self.findChild(QTableWidget, "tableWidget")
       # self.pushButton_report = self.findChild(QPushButton, "pushButton_report") #убираем кнопку
        self.pushButton_load_csv = self.findChild(QPushButton, "pushButton_load_csv")  # Добавили кнопку

        self.load_years()
        #self.pushButton_report.clicked.connect(self.create_report) #убираем сигнал
        self.comboBox_year.currentIndexChanged.connect(self.load_table_data)
        self.pushButton_load_csv.clicked.connect(self.load_data_from_csv) # добавили сигнал для csv

        self.load_table_data()

    def connect_to_db(self):
        try:
           self.mydb = mysql.connector.connect(
               host="localhost",
               user="root",
               password="2810300719",
               database="dianadb"
           )
           if self.mydb.is_connected():
               print("Connected to MySQL Database")
        except mysql.connector.Error as e:
           print(f"Error connecting to MySQL: {e}")
           sys.exit(1)

    def load_years(self):
       years = [str(year) for year in range(2020, datetime.now().year + 1)]
       self.comboBox_year.addItems(years)


    def load_table_data(self):
        selected_year = self.comboBox_year.currentText()

        self.tableWidget.clearContents() #очищаем данные
        self.tableWidget.setRowCount(0)

        try:
           cursor = self.mydb.cursor()
           sql = """
                SELECT e.first_name, e.last_name, SUM(s.quantity * p.price) as revenue, SUM(s.quantity) as total_quantity
                FROM employees e
                LEFT JOIN sales s ON e.employee_id = s.employee_id
                LEFT JOIN products p ON s.product_id = p.product_id
                WHERE YEAR(s.sale_date) = %s OR s.sale_date IS NULL
                GROUP BY e.employee_id
                """

           cursor.execute(sql, (selected_year,))
           results = cursor.fetchall()
           self.tableWidget.setColumnCount(4)
           self.tableWidget.setHorizontalHeaderLabels(["Имя", "Фамилия", "Выручка", "Продано"])
           for row_num, row_data in enumerate(results):
              self.tableWidget.insertRow(row_num)
              for col_num, cell_data in enumerate(row_data):
                 if cell_data is None:
                    item = QTableWidgetItem("0")
                 else:
                     item = QTableWidgetItem(str(cell_data))
                 self.tableWidget.setItem(row_num, col_num, item)


           self.tableWidget.resizeColumnsToContents()
        except mysql.connector.Error as e:
           print(f"Error getting sales data: {e}")
        finally:
           cursor.close()

    def load_data_from_csv(self):
       filename, _ = QFileDialog.getOpenFileName(self, "Выберите CSV-файл", "", "CSV Files (*.csv)")

       if filename:
           self.tableWidget.clearContents()
           self.tableWidget.setRowCount(0)
           try:
                with open(filename, 'r', encoding="utf-8") as file:
                   reader = csv.reader(file)
                   header = next(reader)  # Считываем заголовок

                   self.tableWidget.setColumnCount(len(header))
                   self.tableWidget.setHorizontalHeaderLabels(header)
                   for row_num, row_data in enumerate(reader):
                        self.tableWidget.insertRow(row_num)
                        for col_num, cell_data in enumerate(row_data):
                           item = QTableWidgetItem(str(cell_data))
                           self.tableWidget.setItem(row_num, col_num, item)

                   self.tableWidget.resizeColumnsToContents()
           except Exception as e:
                print(f"Error loading CSV: {e}")

    def closeEvent(self, event):
        if self.mydb:
            self.mydb.close()
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TeamStatsWindow()
    window.show()
    sys.exit(app.exec_())