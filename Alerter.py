# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Alerter.ui'
#
# Created: ZK Aug 13 23:58:36 2017
#      by: PyQt4 UI code generator 4.11.3
#
# WARNING! All changes made in this file will be lost!

import sys
reload(sys)
sys.setdefaultencoding("utf-8")

# Dependency
import os
import time
import xlrd
import shutil
import webbrowser

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_Dialog(object):
    # Cache dialog QWidget
    _dialog = None;
    # Time template
    ISOTIMEFORMAT = '%Y%m%d%H%M%S'
    LOGISOTIMEFORMAT='%Y-%m-%d %X'

    # Index 4 excel
    KEY_INDEX = 0
    ALERT_GROUP_INDEX = 1
    ALERT_KEY_INDEX = 2
    SEVERITY_INDEX = 3
    SUMMARY_INDEX = 4

    def setupUi(self, Dialog):
        _dialog = Dialog
        Dialog.setObjectName(_fromUtf8("Dialog"))
        Dialog.resize(641, 467)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Dialog.sizePolicy().hasHeightForWidth())
        Dialog.setSizePolicy(sizePolicy)
        Dialog.setMinimumSize(QtCore.QSize(641, 467))
        Dialog.setMaximumSize(QtCore.QSize(641, 467))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Consolas"))
        font.setPointSize(11)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        Dialog.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/google_alerts.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        icon.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/google_alerts.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        Dialog.setWindowIcon(icon)
        Dialog.setAutoFillBackground(True)
        self.groupBox_config = QtGui.QGroupBox(Dialog)
        self.groupBox_config.setEnabled(True)
        self.groupBox_config.setGeometry(QtCore.QRect(30, 20, 581, 191))
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_config.sizePolicy().hasHeightForWidth())
        self.groupBox_config.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.groupBox_config.setFont(font)
        self.groupBox_config.setFlat(False)
        self.groupBox_config.setCheckable(False)
        self.groupBox_config.setObjectName(_fromUtf8("groupBox_config"))
        self.label_input = QtGui.QLabel(self.groupBox_config)
        self.label_input.setGeometry(QtCore.QRect(10, 30, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label_input.setFont(font)
        self.label_input.setObjectName(_fromUtf8("label_input"))
        self.radioButton_outputDir = QtGui.QRadioButton(self.groupBox_config)
        self.radioButton_outputDir.setGeometry(QtCore.QRect(290, 30, 211, 16))
        self.radioButton_outputDir.setObjectName(_fromUtf8("radioButton_outputDir"))
        self.radioButton_outputFile = QtGui.QRadioButton(self.groupBox_config)
        self.radioButton_outputFile.setGeometry(QtCore.QRect(290, 100, 211, 16))
        self.radioButton_outputFile.setObjectName(_fromUtf8("radioButton_outputFile"))
        self.lineEdit_outputDir = QtGui.QLineEdit(self.groupBox_config)
        self.lineEdit_outputDir.setGeometry(QtCore.QRect(310, 60, 231, 20))
        self.lineEdit_outputDir.setObjectName(_fromUtf8("lineEdit_outputDir"))
        self.lineEdit_outputFile = QtGui.QLineEdit(self.groupBox_config)
        self.lineEdit_outputFile.setGeometry(QtCore.QRect(310, 130, 231, 20))
        self.lineEdit_outputFile.setObjectName(_fromUtf8("lineEdit_outputFile"))
        self.lineEdit_input = QtGui.QLineEdit(self.groupBox_config)
        self.lineEdit_input.setGeometry(QtCore.QRect(10, 60, 221, 20))
        self.lineEdit_input.setObjectName(_fromUtf8("lineEdit_input"))
        self.pushButton_input = QtGui.QPushButton(self.groupBox_config)
        self.pushButton_input.setGeometry(QtCore.QRect(240, 60, 21, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_input.setFont(font)
        self.pushButton_input.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton_input.setText(_fromUtf8(""))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/up122.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.pushButton_input.setIcon(icon1)
        self.pushButton_input.setIconSize(QtCore.QSize(19, 19))
        self.pushButton_input.setObjectName(_fromUtf8("pushButton_input"))
        self.line = QtGui.QFrame(self.groupBox_config)
        self.line.setGeometry(QtCore.QRect(270, 30, 20, 141))
        self.line.setFrameShape(QtGui.QFrame.VLine)
        self.line.setFrameShadow(QtGui.QFrame.Sunken)
        self.line.setObjectName(_fromUtf8("line"))
        self.label_note = QtGui.QLabel(self.groupBox_config)
        self.label_note.setGeometry(QtCore.QRect(10, 100, 261, 51))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_note.setFont(font)
        self.label_note.setWordWrap(True)
        self.label_note.setObjectName(_fromUtf8("label_note"))
        self.pushButton_outputDir = QtGui.QPushButton(self.groupBox_config)
        self.pushButton_outputDir.setEnabled(False)
        self.pushButton_outputDir.setGeometry(QtCore.QRect(550, 60, 21, 21))
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_outputDir.sizePolicy().hasHeightForWidth())
        self.pushButton_outputDir.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_outputDir.setFont(font)
        self.pushButton_outputDir.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton_outputDir.setText(_fromUtf8(""))
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/doc_lines_stright.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.pushButton_outputDir.setIcon(icon2)
        self.pushButton_outputDir.setIconSize(QtCore.QSize(21, 21))
        self.pushButton_outputDir.setObjectName(_fromUtf8("pushButton_outputDir"))
        self.pushButton_outputFile = QtGui.QPushButton(self.groupBox_config)
        self.pushButton_outputFile.setEnabled(False)
        self.pushButton_outputFile.setGeometry(QtCore.QRect(550, 130, 21, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_outputFile.setFont(font)
        self.pushButton_outputFile.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton_outputFile.setText(_fromUtf8(""))
        self.pushButton_outputFile.setIcon(icon2)
        self.pushButton_outputFile.setIconSize(QtCore.QSize(21, 21))
        self.pushButton_outputFile.setObjectName(_fromUtf8("pushButton_outputFile"))
        self.scrollArea_log = QtGui.QScrollArea(Dialog)
        self.scrollArea_log.setGeometry(QtCore.QRect(30, 240, 581, 121))
        self.scrollArea_log.setAutoFillBackground(False)
        self.scrollArea_log.setFrameShape(QtGui.QFrame.Panel)
        self.scrollArea_log.setWidgetResizable(True)
        self.scrollArea_log.setObjectName(_fromUtf8("scrollArea_log"))
        self.scrollAreaWidgetContents = QtGui.QLabel()
        self.scrollAreaWidgetContents.setWordWrap(True)
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 579, 119))
        self.scrollAreaWidgetContents.setObjectName(_fromUtf8("scrollAreaWidgetContents"))
        self.scrollArea_log.setWidget(self.scrollAreaWidgetContents)
        self.progressBar = QtGui.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(30, 370, 491, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName(_fromUtf8("progressBar"))
        self.pushButton_start = QtGui.QPushButton(Dialog)
        self.pushButton_start.setGeometry(QtCore.QRect(540, 430, 75, 23))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_start.setFont(font)
        self.pushButton_start.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/play86.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.pushButton_start.setIcon(icon3)
        self.pushButton_start.setObjectName(_fromUtf8("pushButton_start"))
        self.pushButton_download = QtGui.QPushButton(Dialog)
        self.pushButton_download.setGeometry(QtCore.QRect(290, 430, 241, 23))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.pushButton_download.setFont(font)
        self.pushButton_download.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/download14.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.pushButton_download.setIcon(icon4)
        self.pushButton_download.setIconSize(QtCore.QSize(21, 21))
        self.pushButton_download.setObjectName(_fromUtf8("pushButton_download"))
        self.pushButton_clearLog = QtGui.QPushButton(Dialog)
        self.pushButton_clearLog.setGeometry(QtCore.QRect(530, 370, 81, 23))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_clearLog.setFont(font)
        self.pushButton_clearLog.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap(_fromUtf8("./icons/025.png")), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.pushButton_clearLog.setIcon(icon5)
        self.pushButton_clearLog.setIconSize(QtCore.QSize(18, 18))
        self.pushButton_clearLog.setObjectName(_fromUtf8("pushButton_clearLog"))
        self.line_2 = QtGui.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(30, 400, 581, 16))
        self.line_2.setFrameShape(QtGui.QFrame.HLine)
        self.line_2.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_2.setObjectName(_fromUtf8("line_2"))
        self.line_3 = QtGui.QFrame(Dialog)
        self.line_3.setGeometry(QtCore.QRect(260, 420, 20, 41))
        self.line_3.setFrameShape(QtGui.QFrame.VLine)
        self.line_3.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_3.setObjectName(_fromUtf8("line_3"))
        self.commandLinkButton = QtGui.QCommandLinkButton(Dialog)
        self.commandLinkButton.setGeometry(QtCore.QRect(50, 420, 188, 41))
        self.commandLinkButton.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.commandLinkButton.setObjectName(_fromUtf8("commandLinkButton"))
        self.label_3 = QtGui.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(40, 220, 61, 21))
        self.label_3.setFrameShadow(QtGui.QFrame.Plain)
        self.label_3.setTextFormat(QtCore.Qt.LogText)
        self.label_3.setObjectName(_fromUtf8("label_3"))

        # Binding pushButton_input click event
        self.pushButton_input.clicked.connect(self.uploadExcel)
        # Binding pushButton_outputDir click event
        self.pushButton_outputDir.clicked.connect(self.selectDir)
        # Binding pushButton_outputFile click event
        self.pushButton_outputFile.clicked.connect(self.selectFile)
        # Binding pushButton_download click event
        self.pushButton_download.clicked.connect(self.download)
        # Binding pushButton_clearLog click event
        self.pushButton_clearLog.clicked.connect(self.clear)
        # Binding pushButton_start click event
        self.pushButton_start.clicked.connect(self.execute)
        # Binding commandLinkButton click event
        self.commandLinkButton.clicked.connect(self.readUserManual)

        # Binding radioButton_outputDir select event
        self.radioButton_outputDir.clicked.connect(self.enableDirOutput)
        # Binding radioButton_outputFile select event
        self.radioButton_outputFile.clicked.connect(self.enableFileOutput)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(_translate("Dialog", "Alerter", None))
        self.groupBox_config.setTitle(_translate("Dialog", "Config", None))
        self.label_input.setText(_translate("Dialog", "Select Input Excel", None))
        self.radioButton_outputDir.setText(_translate("Dialog", "Select Output Directory", None))
        self.radioButton_outputFile.setText(_translate("Dialog", "Select File To Append", None))
        self.pushButton_input.setToolTip(_translate("Dialog", "<html><head/><body><p>上传输入表格</p></body></html>", None))
        self.label_note.setText(_translate("Dialog", "* Please use our template excel. It can be got by clicking the right-bottom button!", None))
        self.pushButton_outputDir.setToolTip(_translate("Dialog", "<html><head/><body><p>选择输出路径</p></body></html>", None))
        self.pushButton_outputFile.setToolTip(_translate("Dialog", "<html><head/><body><p>选择输出目标文件</p></body></html>", None))
        self.pushButton_start.setToolTip(_translate("Dialog", "<html><head/><body><p>开始</p></body></html>", None))
        self.pushButton_start.setText(_translate("Dialog", "Start", None))
        self.pushButton_download.setToolTip(_translate("Dialog", "<html><head/><body><p>下载表格模板</p></body></html>", None))
        self.pushButton_download.setText(_translate("Dialog", "Download Template Excel", None))
        self.pushButton_clearLog.setToolTip(_translate("Dialog", "<html><head/><body><p>清空控制台输出</p></body></html>", None))
        self.pushButton_clearLog.setText(_translate("Dialog", "Clear", None))
        self.commandLinkButton.setToolTip(_translate("Dialog", "<html><head/><body><p>阅读用户手册</p></body></html>", None))
        self.commandLinkButton.setText(_translate("Dialog", "Read User Manual", None))
        self.label_3.setText(_translate("Dialog", "Console", None))

    # Read user mannual
    def readUserManual(self):
        #TODO:
        webbrowser.open("./static/index.html")

    # Upload the input excel
    def uploadExcel(self):
        # Open file Dialog
        filePath = QtGui.QFileDialog.getOpenFileName(self._dialog, "Open Document", '/', "Excels(*.xls *.xlsx)")
        if filePath.length() != 0 :
            # Display the selected excel path
            self.lineEdit_input.setText(filePath)

    # Select the output directory path
    def selectDir(self):
        options = QtGui.QFileDialog.DontResolveSymlinks | QtGui.QFileDialog.ShowDirsOnly;
        dirPath = QtGui.QFileDialog.getExistingDirectory(self._dialog, "Open Folder", '/', options)
        if dirPath.length() != 0:
            # Display the selected output dir path
            self.lineEdit_outputDir.setText(dirPath)

    # Select the output rules path
    def selectFile(self):
        # Open file Dialog
        filePath = QtGui.QFileDialog.getOpenFileName(self._dialog, "Open Document", '/', "Rules(*.rules)")
        if filePath.length() != 0 :
            # Display the selected output file path
            self.lineEdit_outputFile.setText(filePath)

    # Download the template excel
    def download(self):
        dirPath = QtGui.QFileDialog.getExistingDirectory(self._dialog, "Open Folder", '/')
        if dirPath.length != 0:
            try:
                srcFilePath = "./template/template.xlsx"
                dstFilePath = dirPath + "/template.xlsx"
                shutil.copyfile(srcFilePath, dstFilePath)
                # TODO: Line 310 doesn't support the path contains chinese
                os.system("explorer.exe %s" % dirPath)
            except IOError , errMsg:
                self.logError("./log/download_log_" + time.strftime(self.ISOTIMEFORMAT, time.localtime()) + ".log", "IOException: " + errMsg)
                print e

    # Clear the content of scroll area & reset the progress
    def clear(self):
        # Clear content
        self.scrollAreaWidgetContents.clear()
        # Reset progress
        self.progressBar.setProperty("value", 0)

    # Call back after dir radiobox selected
    def enableDirOutput(self):
        self.lineEdit_outputDir.setText("")
        self.pushButton_outputDir.setEnabled(True)
        self.pushButton_outputFile.setEnabled(False)

    # Call back after dir radiobox selected
    def enableFileOutput(self):
        self.lineEdit_outputFile.setText("")
        self.pushButton_outputDir.setEnabled(False)
        self.pushButton_outputFile.setEnabled(True)

    # Lock start button to avoid exception
    def lockUserOperation(self, isLock):
        # isLock == True  -> lock
        # isLock == False -> unlock
        self.pushButton_start.setEnabled(not isLock)

    # Check excel path
    def isExcel(self, filePath):
        suffix = os.path.splitext(filePath)[-1]
        return (suffix == ".xls" or suffix == ".xlsx") and os.path.isfile(filePath)

    # Check rules path
    def isRules(self, filePath):
        suffix = os.path.splitext(filePath)[-1]
        return suffix == ".rules" and os.path.isfile(filePath)

    # Check can execute
    # Param outputMode: directory or target file
    def canExecute(self, inputPath, outputMode, outputPath):
        if outputMode == "Dir":
            return self.isExcel(inputPath) and os.path.isdir(outputPath)
        elif outputMode == "File":
            return self.isExcel(inputPath) and self.isRules(outputPath)
        else:
            return False

    # Log error
    def logError(self, logFilePath, msg):
        msg = "[" + time.strftime(self.LOGISOTIMEFORMAT, time.localtime()) + "]" + "[ERROR]" + msg + "\n"
        # Write log file
        with open(logFilePath, "a") as file:
            file.write(msg)
        # Write console
        self.scrollAreaWidgetContents.setText(self.scrollAreaWidgetContents.text() + msg)

    # Log warning
    def logWarning(self, logFilePath, msg):
        msg = "[" + time.strftime(self.LOGISOTIMEFORMAT, time.localtime()) + "]" + "[WARNING]" + msg + "\n"
        # Write log file
        with open(logFilePath, "a") as file:
            file.write(msg)
        # Write console
        self.scrollAreaWidgetContents.setText(self.scrollAreaWidgetContents.text() + msg)

    # Log info
    def logInfo(self, logFilePath, msg):
        msg = "[" + time.strftime(self.LOGISOTIMEFORMAT, time.localtime()) + "]" + "[INFO]" + msg + "\n"
        # Write log file
        with open(logFilePath, "a") as file:
            file.write(msg)
        # Write console
        self.scrollAreaWidgetContents.setText(self.scrollAreaWidgetContents.text() + msg)

    # Log debug
    def logDebug(self, logFilePath, msg):
        msg = "[" + time.strftime(self.LOGISOTIMEFORMAT, time.localtime()) + "]" + "[DEBUG]" + msg + "\n"
        # Write log file
        with open(logFilePath, "a") as file:
            file.write(msg)

    # Write output
    def writeResult(self, outputPath, key, alertGroup, alertKey, severity, summary):
        with open("./ason/output.ason", "r") as rFile:
            content = rFile.read()
        with open(outputPath, "a") as wFile:
            content = content.replace("@1", time.strftime(self.LOGISOTIMEFORMAT, time.localtime()))
            content = content.replace("@2", key)
            content = content.replace("@3", alertGroup)
            content = content.replace("@4", alertKey)
            content = content.replace("@5", str(int(severity)))
            summaryItems = summary.split('+')
            for summaryItem in summaryItems:
                summaryItem = summaryItem.strip(" ")
                if summaryItem.startswith('@') or summaryItem.startswith('$'):
                    continue
                # TODO: line 408 doesn't work
                summaryItem = "\"" + summaryItem + "\""
            seperator = " + "
            summary = seperator.join(summaryItems)
            content = content.replace("@6", summary)
            wFile.write(content)

    # Deal fail
    def exitByFail(self, logFilePath, msg):
        self.logError(logFilePath, msg)
        self.progressBar.setProperty("value", 100)

    # Deal success
    def exitBySuccess(self, logFilePath):
        self.logInfo(logFilePath, "Success 100%")
        self.progressBar.setProperty("value", 100)

    # Parse excel & write(append) .rules
    def execute(self):
        # 0.Lock user operateion
        self.lockUserOperation(True)

        # 1.Init log file name
        logFilePath = "./log/Alerter_log_" + time.strftime(self.ISOTIMEFORMAT, time.localtime()) + ".log"

        self.logInfo(logFilePath, "Start")

        # 2. Get input params (use abspath to avoid dir attack)
        inputPath = os.path.abspath(self.lineEdit_input.text()).decode("utf-8")
        outputMode = ""
        outputPath = ""
        if self.radioButton_outputDir.isChecked():
            outputMode = "Dir"
            outputPath = os.path.abspath(self.lineEdit_outputDir.text()).decode("utf-8")
        elif self.radioButton_outputFile.isChecked:
            outputMode = "File"
            outputPath = os.path.abspath(self.lineEdit_outputFile.text()).decode("utf-8")

        self.logInfo(logFilePath, "Velidate params start")
        self.logDebug(logFilePath, "inputPath:" + inputPath + " outputMode:" + outputMode + " outputPath:" + outputPath)

        # 3. Check input params
        if self.canExecute(inputPath, outputMode, outputPath) == False:
            self.exitByFail(logFilePath, "Input file or output params error!")
            return

        self.logInfo(logFilePath, "Velidate params end")
        self.logInfo(logFilePath, "Execute start")
        # If user select dir output mode, we generate the out file
        if outputMode == "Dir":
            outputPath = outputPath + "\Alerter_result_" + time.strftime(self.ISOTIMEFORMAT, time.localtime()) + ".rules"

        self.logInfo(logFilePath, "OutputPath is " + outputPath)

        # 4. Read input excel
        workbook = xlrd.open_workbook(inputPath)
        table = workbook.sheet_by_index(0)
        nrows = table.nrows;

        # 5. Generate result
        for index in range(1, nrows):
            rowData = table.row(index)
            self.writeResult(outputPath, rowData[self.KEY_INDEX].value, rowData[self.ALERT_GROUP_INDEX].value,
                    rowData[self.ALERT_KEY_INDEX].value, rowData[self.SEVERITY_INDEX].value, rowData[self.SUMMARY_INDEX].value)
            self.progressBar.setProperty("value", index / nrows)

        self.logInfo(logFilePath, "Execute end")
        self.exitBySuccess(logFilePath)

        # 6. unlock user operateion
        self.lockUserOperation(False)

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    Form = QtGui.QWidget()
    ui = Ui_Dialog()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
