import sys  # sys нужен для передачи argv в QApplication
from PySide2 import QtCore,QtGui,QtWidgets 
from design import Ui_Dialog
from Try_to_find_reference import *
from PyQt5.QtWidgets import QPushButton

#Create app
app = QtWidgets.QApplication(sys.argv)

#init
Dialog = QtWidgets .QDialog()  
ui = Ui_Dialog()
ui.setupUi(Dialog)
Dialog.show()

#hook logic

def get_data():
    filename1 = ui.rozkladName.text()
    filename2 = ui.subjName.text()
    firstData = ui.beginData.text()
    secondData = ui.endData.text()
    sheetsString = ui.numberSheets.text()
    filenameRings = ui.rings.text()
    firstClas = ui.firstLineRing.text() #раздитель по линиям звонков
    schoolName = (ui.schoolName_2.text()).strip()
    rozklad = fewSheetsResult(filename1.strip(),filename2.strip(),sheetsString.strip(),filenameRings.strip(),firstData.strip(),secondData.strip(),firstClas,schoolName.strip())
    stringRozklad = "<table border = '1'><tr><td>Аудиторія</td><td>Предмет</td><td>Вчитель</td><td>Клас</td><td>Початок сем.</td><td>Кінець сем.</td><td>День тижня</td><td>Періодичність</td><td>Час уроку</td></tr>"
    del rozklad[0]
    for line in rozklad:
        newstring = "<tr><td>"+line[0]+"</td>"+"<td>"+line[1]+"</td>"+"<td>"+line[2]+"</td>"+"<td>"+line[3]+"</td>"+"<td>"+line[4]+"</td>"+"<td>"+line[5]+"</td>"+"<td>"+line[6]+"</td>"+"<td>"+line[7]+"</td>"+"<td>"+line[8]+"</td></tr>"
        stringRozklad = stringRozklad + newstring      
    stringRozklad = stringRozklad + "</table>"
    ui.textEditRozklad.textCursor().insertHtml(stringRozklad)

    #-----------fix------------------#
    nullMass = changeSpace(uniqNullMass(findNull(rozklad)))
    nullFixString = "<table border = '1'><tr><td>Предмет</td><td>Клас</td><td>Вчитель</td></tr>"
    #del nullMass[0]
    for line in nullMass:
        newstring = "<tr><td>"+line[0]+"</td><td>"+line[1]+"</td><td>"+""+"</td></tr>"
        nullFixString = nullFixString + newstring
    nullFixString = nullFixString + "</table>"
    ui.textEditFix.textCursor().insertHtml(nullFixString)
    #recordFile(rozklad,schoolName)
             
    
ui.pushButton.clicked.connect(get_data)

def get_lists_data():
    filename1 = ui.studentsFile.text()
    filename2 = ui.teachersFile.text()
    schoolName = ui.schoolName.text()
    stList = studentsList(filename1.strip(),schoolName.strip())
    tchList = teachersList(filename2.strip(),schoolName.strip())
    stringStudent = "<table border = '1'><tr><th>Ім'я</th><th>По батькові</th><th>Прізвище</th><th>Клас</th><th>Школа</th></tr>"
    for line in  stList:
        newstring = "<tr><td>"+line[0]+"</td>"+"<td>"+line[1]+"</td>"+"<td>"+line[2]+"</td>"+"<td>"+line[3]+"</td>"+"<td>"+line[4]+"</td></tr>"
        stringStudent = stringStudent + newstring
    stringStudent = stringStudent + "</table><p> <br> <p>"
    ui.textEdit_2.textCursor().insertHtml(stringStudent)
    stringTeaher = "<table border = '1'><tr><th>Ім'я</th><th>По батькові</th><th>Прізвище</th><th>Школа</th></tr>"
    for line in  tchList:
        newstring = "<tr><td>"+line[0]+"</td>"+"<td>"+line[1]+"</td>"+"<td>"+line[2]+"</td>"+"<td>"+line[3]+"</td></tr>"
        stringTeaher = stringTeaher + newstring
    stringTeaher = stringTeaher + "</table>"
    ui.textEdit_2.textCursor().insertHtml(stringTeaher)
    
ui.pushButton_2.clicked.connect(get_lists_data)

def Fix():
    filename1 = ui.rozkladName.text()
    filename2 = ui.subjName.text()
    firstData = ui.beginData.text()
    secondData = ui.endData.text()
    sheetsString = ui.numberSheets.text()
    filenameRings = ui.rings.text()
    firstClas = ui.firstLineRing.text() #раздитель по линиям звонков
    schoolName = (ui.schoolName_2.text()).strip()
    rozklad = fewSheetsResult(filename1.strip(),filename2.strip(),sheetsString.strip(),filenameRings.strip(),firstData.strip(),secondData.strip(),firstClas,schoolName.strip())
    takeFromEdit = changeSlash(rebornTextEditList((ui.textEditFix.toPlainText()).split()))
    #print(takeFromEdit)
    result = ResultFix(takeFromEdit,rozklad)
    stringRozklad = "<table border = '1'><tr><td>Аудиторія</td><td>Предмет</td><td>Вчитель</td><td>Клас</td><td>Початок сем.</td><td>Кінець сем.</td><td>День тижня</td><td>Періодичність</td><td>Час уроку</td></tr>"
    del rozklad[0]
    for line in result:
        newstring = "<tr><td>"+line[0]+"</td>"+"<td>"+line[1]+"</td>"+"<td>"+line[2]+"</td>"+"<td>"+line[3]+"</td>"+"<td>"+line[4]+"</td>"+"<td>"+line[5]+"</td>"+"<td>"+line[6]+"</td>"+"<td>"+line[7]+"</td>"+"<td>"+line[8]+"</td></tr>"
        stringRozklad = stringRozklad + newstring      
    stringRozklad = stringRozklad + "</table>"
    ui.textEditRozklad.clear()
    ui.textEditRozklad.textCursor().insertHtml(stringRozklad)
    
ui.fixNull.clicked.connect(Fix)
        
#main loop
sys.exit(app.exec_())





