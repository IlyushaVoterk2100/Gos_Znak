import sys
import os
import cv2
import pytesseract
import openpyxl
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication


# вызов GUI
Form, Window = uic.loadUiType('pyqt1.ui')

app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

def Start():
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        cv2.imshow('frame', frame)
        cv2.imwrite('cam.png', frame)
        img = cv2.imread('cam.png')

        # Блок работы с цветом изображения
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
        result = []
        pytesseract.pytesseract.tesseract_cmd = 'Tesseract-ORC\\tesseract.exe'
        result = pytesseract.image_to_string(thresh, lang='rus+en', config='-c page_separator=" "')
        ras = []
        ras = result.split()

        # Считывание Excel
        wb = openpyxl.reader.excel.load_workbook(filename='BAZA.xlsx')
        wb.active = 0
        sheet = wb.active
        NUMBase = []
        for i in range(1, 100):
            NUMBase += [(sheet['A' + str(i)].value)]

            # Проверка Базы
        ans = ''
        l = len(ras)
        if l < 1:
            ans = 'Въезд запрещён!'
        else:
            for i in range(l):
                if ras[i] in NUMBase:
                    ans = 'Въезд разрешен!'
                else:
                    ans = 'Въезд запрещён!'

        form.label_2.setText(ans)

        if ans == 'Въезд разрешен!':
            cv2.waitKey(5000)

        if cv2.waitKey(1) == 27:
            break

def Excel():

    # открытие Excel
    a = os.system('BAZA.xlsx')
    return a

def NSD():
    form.label_2.setText("Шлагбаум открыт!")
    cv2.waitKey(5000)


form.Start.clicked.connect(Start) # Кнопка Старт
form.NSD.clicked.connect(NSD) # Кнопка Открыть шлагбаум
form.Excel.clicked.connect(Excel) # Кнопка Excel
form.Exit.clicked.connect(sys.exit) # Кнопка Выход

app.exec()


