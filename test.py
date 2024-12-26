# login_window.py
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QPushButton, QLabel, QMessageBox
from PyQt5 import uic
from PyQt5.QtCore import Qt
import mysql.connector
# from buttons_window import ButtonsWindow

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('login.ui', self)

        self.lineEdit_login = self.findChild(QLineEdit, "lineEdit_login")
        self.lineEdit_password = self.findChild(QLineEdit, "lineEdit_password")
        self.pushButton_login = self.findChild(QPushButton, "pushButton_login")
        self.label_message = self.findChild(QLabel, "label_message")

        self.lineEdit_password.setEchoMode(QLineEdit.Password)
        self.pushButton_login.clicked.connect(self.login)

        self.mydb = None
        self.connect_to_db()

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

    def login(self):
        login = self.lineEdit_login.text()
        password = self.lineEdit_password.text()

        if not login or not password:
           self.label_message.setText("Пожалуйста, заполните все поля")
           return

        if self.mydb and self.mydb.is_connected():
          cursor = self.mydb.cursor()
          try:
            sql = "SELECT employee_id, first_name, last_name, username, role, email, phone_number  FROM employees WHERE username = %s AND password_hash = %s"
            cursor.execute(sql, (login, password))
            result = cursor.fetchone()

            if result:
                self.label_message.clear()
                self.open_buttons_window(result)
                self.close()
            else:
               self.label_message.setText("Неверный логин или пароль")

          except mysql.connector.Error as e:
            self.label_message.setText(f"Ошибка при запросе: {e}")
            print(f"Error when loging in: {e}")
          finally:
            cursor.close()
        else:
          self.label_message.setText("Ошибка подключения к БД")
          print("Error connecting to DB")


    def open_buttons_window(self, user_data):
      from MainWindow import ButtonsWindow
      self.buttons_window = ButtonsWindow(user_data)
      self.buttons_window.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LoginWindow()
    window.show()
    sys.exit(app.exec_())