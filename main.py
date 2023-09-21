import logs
import macros
import func

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys
import os




class MyTread(QtCore.QThread):
    mysignal = QtCore.pyqtSignal(str)
    params = {}

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        error = None
        try:
            worker = macros.Macros(self.mysignal)
            worker.set_options(self.params)
            worker.run()
            self.mysignal.emit("Программа завершила работу")
        except Exception as err_all:
            self.mysignal.emit(f"Непредвиденная ошибка: {err_all}")




class Window(QMainWindow):
    options = {'file_in': '', 'file_in_name': '', 'type_out': '', 'list_manuf': [], 'num_header': 0, 'col_name': 0, 'folder_out': '', 'col_fn': 0}

    def __init__(self):

        #описываем окно
        super(Window, self).__init__()
        self.setWindowTitle("Обработчик")
        self.setGeometry(300, 100, 820, 620)
        #Добавлям кнопку выбора файла
        self.btn_upload = QtWidgets.QPushButton(self)  # кнопка загрузить из файла
        self.btn_upload.move(10, 10)
        self.btn_upload.setText("Выбрать файла")
        self.btn_upload.setFixedWidth(150)
        self.btn_upload.clicked.connect(self.click_btn_upload)
        self.l_file_in = QtWidgets.QLabel(self)
        self.l_file_in.setGeometry(QtCore.QRect(10, 40, 420, 30))

        #радиокнопка выбора типа результата
        self.gr_type_out = QtWidgets.QGroupBox(self)
        self.gr_type_out.setGeometry(QtCore.QRect(10, 80, 141, 80))
        self.gr_type_out.setObjectName("groupBox")
        self.rb_one_file = QtWidgets.QRadioButton(self.gr_type_out)
        self.rb_one_file.setGeometry(QtCore.QRect(20, 20, 82, 17))
        self.rb_one_file.setObjectName("rb_one_file")
        self.rb_many_file = QtWidgets.QRadioButton(self.gr_type_out)
        self.rb_many_file.setGeometry(QtCore.QRect(20, 40, 150, 17))
        self.rb_many_file.setObjectName("rb_many_file")
        self.rb_many_file.setChecked(True)
        self.gr_type_out.setTitle("Сохранять в ")
        self.rb_one_file.setText("один файл")
        self.rb_many_file.setText("несколько файлов")
        self.o_type_out = QtWidgets.QButtonGroup()
        self.o_type_out.addButton(self.rb_one_file)
        self.o_type_out.addButton(self.rb_many_file)
        #добавляем элемент для выбора производителей
        try:
            self.all_manuf = func.get_all_manuf()
        except Exception as err_all:
            logs.write_log(f'ошибка доступа к справочнику {err_all}')
            sys.exit()
        self.o_manuf = QtWidgets.QListWidget(self)
        self.o_manuf.setGeometry(QtCore.QRect(450, 80, 220, 520))
        self.o_manuf.setObjectName('Выберите производителей')
        for i, option in enumerate(self.all_manuf):
            item_ch = QtWidgets.QListWidgetItem()
            item_ch.setText(option)
            item_ch.setCheckState(QtCore.Qt.Unchecked)
            self.o_manuf.addItem(item_ch)
        self.btn_check_all = QtWidgets.QPushButton(self)
        self.btn_check_all.setGeometry(QtCore.QRect(450, 10, 80, 25))
        self.btn_check_all.setText('Выбрать все')
        self.btn_check_all.clicked.connect(self.click_btn_check_all)
        self.btn_uncheck_all = QtWidgets.QPushButton(self)
        self.btn_uncheck_all.setGeometry(QtCore.QRect(535, 10, 70, 25))
        self.btn_uncheck_all.setText('Снять все')
        self.btn_uncheck_all.clicked.connect(self.click_btn_uncheck_all)
        self.txt_manuf_search = QtWidgets.QLineEdit(self)
        self.txt_manuf_search.setGeometry(QtCore.QRect(450, 50, 170, 25))
        self.btn_manuf_search = QtWidgets.QPushButton(self)
        self.btn_manuf_search.setGeometry(QtCore.QRect(630, 50, 50, 25))
        self.btn_manuf_search.setText('найти')
        self.btn_manuf_search.clicked.connect(self.search_manuf)
        self.btn_plus =QtWidgets.QPushButton(self)
        self.btn_plus.setGeometry(QtCore.QRect(675, 175, 25, 25))
        self.btn_plus.setText('+')
        self.btn_plus.clicked.connect(self.btn_plus_clicked)
        self.btn_minus =QtWidgets.QPushButton(self)
        self.btn_minus.setGeometry(QtCore.QRect(675, 200, 25, 25))
        self.btn_minus.setText('-')
        self.btn_minus.clicked.connect(self.btn_minus_clicked)
        #колонка с ФН
        self.lbl_num_fn = QtWidgets.QLabel(self)
        self.lbl_num_fn.setGeometry(QtCore.QRect(10, 190, 200, 30))
        self.lbl_num_fn.setText('Номер столбца с ФН (для ОФД)')
        self.o_num_fn = QtWidgets.QTextEdit(self)
        self.o_num_fn.setGeometry(220, 190, 50, 30)
        #номер строки заголовка
        self.lbl_num_header = QtWidgets.QLabel(self)
        self.lbl_num_header.setGeometry(QtCore.QRect(10, 230, 200, 30))
        self.lbl_num_header.setText('Введите номер строки заголовка')
        self.o_num_header = QtWidgets.QTextEdit(self)
        self.o_num_header.setGeometry(220, 230, 50, 30)
        self.o_num_header.setText('1')
        #номер колонки с наименованием
        self.l_col_name = QtWidgets.QLabel(self)
        self.l_col_name.setGeometry(QtCore.QRect(10, 270, 200, 30))
        self.l_col_name.setText('Номер колонки с названием')
        self.o_col_name = QtWidgets.QTextEdit(self)
        self.o_col_name.setGeometry(QtCore.QRect(220, 270, 50, 30))
        self.o_col_name.setText('8')
        #кнопка запуска
        self.btn_process = QtWidgets.QPushButton(self)
        self.btn_process.setGeometry(QtCore.QRect(10, 320, 200, 50))
        self.btn_process.setText("Запустить")
        self.btn_process.clicked.connect(self.click_btn_process)
        #добавляем элемент для логов
        self.lbl_log = QtWidgets.QLabel(self)
        self.lbl_log.setText('вывод информации о работе программы')
        self.lbl_log.setGeometry(QtCore.QRect(10, 380, 300, 20))
        self.txt_logs = QtWidgets.QTextEdit(self)
        self.txt_logs.setGeometry(QtCore.QRect(10, 400, 400, 200))
        self.txt_logs.setReadOnly(True)
        self.txt_logs.setBackgroundRole(QtGui.QPalette.Base)
        p = self.txt_logs.palette()
        p.setColor(self.txt_logs.backgroundRole(), QtGui.QColor(225, 230, 229))
        self.txt_logs.setPalette(p)
        #поток для работы
        self.mythread = MyTread()
        self.mythread.finished.connect(self.mythread_finish)
        self.mythread.mysignal.connect(self.mythread_change, QtCore.Qt.QueuedConnection)



    def mythread_change(self, s):
        self.txt_logs.append(s)


    def click_btn_upload(self):
        tmp = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel files (*.xlsx)")
        self.options['file_in'] = tmp[0]
        self.options['file_in_name'] = os.path.basename(self.options['file_in']).replace('.xlsx', '').replace('.xls', '')
        self.options['folder_out'] = os.path.dirname(self.options['file_in'])
        self.l_file_in.setText(self.options['file_in_name'])


    def save_options(self):
        fl_run = True
        #1
        if self.rb_many_file.isChecked() is True:
            self.options['type_out'] = 'many'
        else:
            tmp = QtWidgets.QFileDialog.getSaveFileName(self, "Выберите файл", "", "Excel files (*.xlsx)")
            self.options['type_out'] = tmp[0]
        #2
        if self.options['file_in'] == '':
            self.txt_logs.append('Выберите файл')
            fl_run = False
        #3
        self.options['col_name'] = self.o_col_name.toPlainText().strip()
        if self.options['col_name'] == '':
            self.txt_logs.append('не указан номер колонки с наименованием')
            fl_run = False
        else:
            self.options['col_name'] = int(self.options['col_name'])
        #4 список производителей
        self.options['list_manuf'].clear()
        for i in range(self.o_manuf.count()):
            if self.o_manuf.item(i).checkState() == QtCore.Qt.Checked:
                self.options['list_manuf'].append(self.o_manuf.item(i).text().strip())
        if len(self.options['list_manuf']) == 0:
            self.txt_logs.append('Не выбран ни один производитель')
            fl_run = False
        #5
        self.options['num_header'] = self.o_num_header.toPlainText().strip()
        if self.options['num_header'] == '':
            self.txt_logs.append('Не указан номер строки с заголовком')
            fl_run = False
        else:
            self.options['num_header'] = int(self.options['num_header'])
        #6
        self.options['col_fn'] = self.o_num_fn.toPlainText().strip()
        if self.options['col_fn'] == '':
            self.options['col_fn'] = 0
        else:
            self.options['col_fn'] = int(self.o_num_fn.toPlainText().strip())
        return fl_run



    def click_btn_process(self):
        if self.save_options() is True:
            self.btn_process.setEnabled(False)
            self.txt_logs.append('Запускаем процесс')
            self.mythread.params = self.options.copy()
            self.mythread.start()
        else:
            self.txt_logs.append('макрос не запущен')

    def mythread_finish(self):
        self.btn_process.setEnabled(True)

    def click_btn_check_all(self):
        for i in range(self.o_manuf.count()):
            self.o_manuf.item(i).setCheckState(QtCore.Qt.Checked)


    def click_btn_uncheck_all(self):
        for i in range(self.o_manuf.count()):
            self.o_manuf.item(i).setCheckState(QtCore.Qt.Unchecked)

    def search_manuf(self):
        txt_search = self.txt_manuf_search.text().lower().strip()
        if txt_search is None or txt_search == '':
            return False
        for i in range(self.o_manuf.count()):
            manufacturer = self.o_manuf.item(i).text().lower()
            if manufacturer.find(txt_search) > -1:
                self.o_manuf.setCurrentRow(i)
                break
        return True

    def btn_plus_clicked(self):
        w_h = self.geometry().size().height() - 100
        self.o_manuf.setFixedHeight(w_h)

    def btn_minus_clicked(self):
        self.o_manuf.setFixedHeight(520)



def application():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    application()
