import csv
import sys
import docx
import pandas as pd
from PySide6.QtCore import QUrl
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QPushButton, QFrame, QFileDialog, QTableView, QTextEdit

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Настройка основных параметров окна
        self.setWindowTitle("Многофреймовое приложение")
        self.setGeometry(100, 100, 600, 400)

        # Создание и установка центрального виджета (фрейма) для главного окна
        self.centralFrame = QFrame()
        self.setCentralWidget(self.centralFrame)

        # Инициализация текущего макета и установка его для центрального фрейма
        self.currentLayout = QVBoxLayout()
        self.centralFrame.setLayout(self.currentLayout)

        # Показать начальный макет с кнопкой выбора файла
        self.showChooseFileLayout()

    def clearLayout(self):
        # Метод для очистки текущего макета.
        while self.currentLayout.count():
            child = self.currentLayout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def showChooseFileLayout(self):
        # Метод для отображения вкладки выбора файла.

        # Очистка текущего макета перед добавлением новых элементов.
        self.clearLayout()

        # Создание кнопки "Выбрать файл".
        btnChooseFile = QPushButton("Выбрать файл")
        btnChooseFile.setStyleSheet("font-size: 16px; background-color: #4CAF50; color: white;")

        # Привязка обработчика события нажатия на кнопку "Выбрать файл".
        btnChooseFile.clicked.connect(self.chooseFileAction)

        # Добавление кнопки "Выбрать файл" на текущий макет.
        self.currentLayout.addWidget(btnChooseFile)

    def showTextLayout(self, filePath: str, fileType: str):
        # Метод для отображения вкладки с текстовой информацией.

        # Очистка текущего макета перед добавлением новых элементов.
        self.clearLayout()

        # Создание кнопки "Назад".
        btnBack = QPushButton("Назад")
        btnBack.setStyleSheet("font-size: 16px; background-color: #f44336; color: white;")

        # Привязка обработчика события нажатия на кнопку "Назад".
        btnBack.clicked.connect(self.showChooseFileLayout)

        # Добавление кнопки "Назад" на текущий макет.
        self.currentLayout.addWidget(btnBack)

        # Создание виджета для отображения текста.
        textEdit = QTextEdit()
        textEdit.setStyleSheet("font-size: 14px; background-color: #ecf0f1; color: #2c3e50;")

        # Установка текста в виджете с использованием метода getText.
        textEdit.setPlainText(self.getText(filePath, fileType))

        # Добавление виджета текста на текущий макет.
        self.currentLayout.addWidget(textEdit)

    def showTableLayout(self, filePath: str, fileType: str):
        # Метод для отображения вкладки с таблицей данных.

        # Очистка текущего макета перед добавлением новых элементов.
        self.clearLayout()

        # Создание кнопки "Назад".
        btnBack = QPushButton("Назад")
        btnBack.setStyleSheet("font-size: 16px; background-color: #f44336; color: white;")
        # Привязка обработчика события нажатия на кнопку "Назад".
        btnBack.clicked.connect(self.showChooseFileLayout)
        # Добавление кнопки "Назад" на текущий макет.
        self.currentLayout.addWidget(btnBack)

        # Создание виджета для отображения таблицы.
        tableView = QTableView()
        tableView.setStyleSheet("font-size: 14px; background-color: #ecf0f1; color: #2c3e50;")

        # Получение данных таблицы с использованием метода getTable.
        data = self.getTable(filePath, fileType)

        # Создание модели данных для таблицы и заполнение ее значениями из данных.
        model = QStandardItemModel(len(data), len(data[0]))
        for row in range(len(data)):
            for column in range(len(data[0])):
                item = data[row][column]
                model.setItem(row, column, QStandardItem(str(item)))

        # Установка модели данных для таблицы.
        tableView.setModel(model)

        # Добавление виджета таблицы на текущий макет.
        self.currentLayout.addWidget(tableView)

    def showPdfLayout(self, filePath: str):
        # Метод для отображения вкладки с просмотром PDF файла.

        # Очистка текущего макета перед добавлением новых элементов.
        self.clearLayout()

        # Создание кнопки "Назад".
        btnBack = QPushButton("Назад")
        btnBack.setStyleSheet("font-size: 16px; background-color: #f44336; color: white;")

        # Привязка обработчика события нажатия на кнопку "Назад".
        btnBack.clicked.connect(self.showChooseFileLayout)

        # Добавление кнопки "Назад" на текущий макет.
        self.currentLayout.addWidget(btnBack)

        # Создание виджета для отображения PDF-файла.
        webView = QWebEngineView()
        webView.settings().setAttribute(webView.settings().WebAttribute.PluginsEnabled, True)
        webView.settings().setAttribute(webView.settings().WebAttribute.PdfViewerEnabled, True)
        webView.setStyleSheet("background-color: #ecf0f1;")

        # Добавление виджета PDF-просмотра на текущий макет.
        self.currentLayout.addWidget(webView)

        # Установка URL для открытия PDF-файла.
        webView.setUrl(QUrl(f"file://{filePath}"))

    def chooseFileAction(self):
        # Метод для обработки события выбора файла из диалогового окна.
        fileDialog = QFileDialog(self)

        # Установка фильтра для выбора файлов с определенными расширениями.
        fileDialog.setNameFilter("Допустимые файлы (*.csv *.xlsx *.pdf *.txt *.docx);;Все файлы (*)")

        # Показ диалогового окна выбора файла.
        if fileDialog.exec():
            # Получение пути к выбранному файлу.
            filePath = fileDialog.selectedFiles()[0]

            # Получение расширения файла.
            fileType = filePath.split('.')[-1]

            # Проверка типа файла и вызов соответствующего метода для отображения содержимого.
            if fileType:
                if fileType == 'csv' or fileType == 'xlsx':
                    self.showTableLayout(filePath, fileType)
                elif fileType == 'txt' or fileType == 'docx':
                    self.showTextLayout(filePath, fileType)
                elif fileType == 'pdf':
                    self.showPdfLayout(filePath)
                else:
                    print('Неизвестный файл! ОШИБКА!')

    def getText(self, filePath: str, fileType: str) -> str:
        # Метод для получения текста из файла в зависимости от его расширения.
        if fileType == 'txt':
            with open(filePath, 'r', encoding='utf-8') as file:
                return file.read()
        elif fileType == 'docx':
            # Использование библиотеки python-docx для чтения текста из файла .docx.
            doc = docx.Document(filePath)
            fullText = [para.text for para in doc.paragraphs]
            return '\n'.join(fullText)
        return ""

    def getTable(self, filePath: str, fileType: str) -> list:
        # Метод для получения таблицы данных в зависимости от расширения файла.
        if fileType == 'xlsx':
            # Использование библиотеки pandas для чтения данных из файла Excel (.xlsx).
            df = pd.read_excel(filePath, 0, header=None)
            return df.values.tolist()
        elif fileType == 'csv':
            data = []
            try:
                # Использование модуля csv для чтения данных из файла CSV.
                with open(filePath, newline='') as csvfile:
                    csvReader = csv.reader(csvfile, delimiter=',')
                    for row in csvReader:
                        data.append(row)
            except Exception as e:
                print(f"Ошибка при чтении CSV файла: {e}")
            return data
        return []

# Точка входа в приложение
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
