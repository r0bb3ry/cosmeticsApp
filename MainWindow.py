import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QTabWidget, QLabel
from PyQt5 import uic
from StatsWindow import StatsWindow
from TeamStatsWindow import TeamStatsWindow  # импортируем окно статистики команды

class ButtonsWindow(QWidget):
    def __init__(self, user_data=None):
        super().__init__()
        try:
            uic.loadUi('ButtonsWindow.ui', self)
            self.user_data = user_data

            # Получаем ссылки на вкладки и кнопки
            self.tabWidget = self.findChild(QTabWidget, "tabWidget")
            self.pushButton_LK = self.findChild(QPushButton, "pushButton_LK")
            self.pushButton_stats = self.findChild(QPushButton, "pushButton_stats")
            self.label_user_info = self.findChild(QLabel, "label_user_info")
            self.pushButton_team_stats = self.findChild(QPushButton, "pushButton_team_stats") #кнопка статистики команды


            # Подключаем сигналы к слотам
            self.pushButton_stats.clicked.connect(self.open_stats_window)
            self.pushButton_team_stats.clicked.connect(self.open_team_stats_window)  # сигнал для кнопки статистики команды

            # Отображение информации о пользователе
            if user_data:
                self.show_user_info(user_data)

        except Exception as e:
            print(f"Error in ButtonsWindow init: {e}")

    def open_stats_window(self):
        try:
           if self.user_data:
              self.stats_window = StatsWindow(self.user_data)
              self.stats_window.show()
           else:
             print("Нет данных о пользователе")
        except Exception as e:
             print(f"Error opening stats window: {e}")

    def open_team_stats_window(self):
        try:
            self.team_stats_window = TeamStatsWindow()
            self.team_stats_window.show()
        except Exception as e:
            print(f"Error opening team stats window: {e}")


    def show_user_info(self, user_data):
       # self.user_data = user_data #убираем
       if user_data and len(user_data) >= 7:
          first_name = user_data[1] if user_data[1] else "Неизвестно"
          last_name = user_data[2] if user_data[2] else "Неизвестно"
          email = user_data[5] if user_data[5] else "Неизвестно"
          phone_number = user_data[6] if user_data[6] else "Неизвестно"
          role = user_data[4] if user_data[4] else "Неизвестно"  # рол
          user_info = f"""
             <b>{first_name} {last_name}</b><br><br>
             <b>Должность:</b> {role}<br>
             <b>Email:</b> {email}<br>
             <b>Номер телефона:</b> {phone_number}<br>
             """

          self.label_user_info.setText(user_info)
       else:
           self.label_user_info.setText("Нет данных о пользователе")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        window = ButtonsWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Global exception: {e}")