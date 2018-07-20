# -*- coding: UTF-8 -*-
"""
1.加载一个指定路径文件夹内的所有txt内容
2.把解析出来的指定内容写入Excel表格
"""

import xlrd
import xlwt
from xlutils.copy import copy
import os
import re
import threading
import time
import shutil
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout, qApp, \
    QDesktopWidget, QFileDialog, QPlainTextEdit

__author__ = "yooongchun"
__email__ = "yooongchun@foxmail.com"


# 加载某文件夹下的所有TXT文件，返回其绝对路径
def loadTXT(file_path):
    txt_files = []  # 保存文件地址和名称
    files = os.listdir(file_path)
    for _file in files:
        if not os.path.splitext(_file)[1] == '.txt':  # 判断是否为txt文件
            continue
        abso_path = os.path.join(file_path, _file)
        txt_files.append(abso_path)
    return txt_files


# 提取TXT文件内容
def extractor(txt_path):
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as fp:
        text = fp.read()
    broker_name = ""
    CRD = ""
    headers = text.split("\n")
    for index, one in enumerate(headers):
        if re.sub(r"\s+", "", one) == "BrokerCheckReport":
            broker_name = headers[index + 1]
            break
    for index, one in enumerate(headers):
        if re.match(r"^CRD#\s+|d+$", one):
            CRD = re.split(r"#\s+", one)[-1]
            break
    text = re.sub(
        r"©\d{4}|FINRA|All rights reserved|Report about|www\.finra\.org/brokercheck",
        "", text)
    disclosure = re.split(r"Disclosure\s+\d+\s+of\s+\d+", text)
    IDs = re.findall(r"Disclosure\s+\d+\s+of\s+\d+", text)
    new_IDs = []
    INFO = []
    counter = -1
    for one, id_ in zip(disclosure[1:], IDs):
        if "Reporting Source" not in one:
            continue
        new_IDs.append(id_)
        counter += 1
        text = re.split(r"\n", one)
        info = {}
        flag = [True for i in range(26)]
        for index, item in enumerate(text):
            cc = -1
            # 1
            cc += 1
            key = "Reporting Source"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 2
            cc += 1
            key = "Current Status"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 3
            cc += 1
            key = "Allegations"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                cnt = index + 1
                while cnt < len(
                        text
                ) and "Initiated By" not in text[cnt] and "Arbitration Forum" not in text[cnt] and not re.sub(
                        r"\s+", "", text[cnt]) == "\n" and not re.sub(
                            r"\s+", "", text[cnt]) == "":
                    v += " " + text[cnt]
                    cnt += 1
                v = re.sub(r"\s+", " ", v)
                key_name = re.split(r",|\.|\s+&", broker_name)
                for k in key_name:
                    v = re.sub(k, "", v)
                info[key] = v
                continue
            # 4
            cc += 1
            key = "Initiated By"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 5
            cc += 1
            key = "Date Initiated"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 6
            cc += 1
            key = "Docket/Case Number"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 7
            cc += 1
            key = "Principal Product Type"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 8
            cc += 1
            key = "Other Product Type(s)"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 9
            cc += 1
            key = "Principal Sanction(s)/Relief"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split("Relief")[-1]
                info[key + " Sought"] = v
            # 10
            cc += 1
            key = "Other Sanction(s)/Relief"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split("Relief")[-1]
                info[key + " Sought"] = v
            # 11
            cc += 1
            key = "Resolution"
            if key in item and "Date" not in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 12
            cc += 1
            key = "Resolution Date"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 13
            cc += 1
            key = "Does the order constitute a"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split("a")[-1]
                key = "Does the order constitute a final order based on violations of any laws or regulations that prohibit fraudulent, manipulative, or deceptive conduct?"
                info[key] = v
                continue
            # 14
            cc += 1
            key = "Sanctions Ordered"
            if key in item and "Other Sanctions Ordered" not in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 15
            cc += 1
            key = "Other Sanctions Ordered"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 16
            cc += 1
            key = "Sanction Details"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                cnt = index + 1
                while cnt < len(text) and not re.sub(
                        r"\s+", "", text[cnt]
                ) == "i" and "Firm Statement" not in text[cnt] and "Regulator Statement" not in text[cnt] and not re.sub(
                        r"\s+", "", text[cnt]) == "\n" and not re.sub(
                            r"\s+", "", text[cnt]) == "":
                    v += " " + text[cnt]
                    cnt += 1
                v = re.sub(r"\s+", " ", v)
                key_name = re.split(r",|\.|\s+&", broker_name)
                for k in key_name:
                    v = re.sub(k, "", v)
                info[key] = v
                continue
            # 17
            cc += 1
            key = "Monetary/Fine"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", "", item)
                v = v.split("Fine")[-1]
                info[key] = v
                continue
            # 18
            cc += 1
            key = "Type of Event"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 19
            cc += 1
            key = "Arbitration Forum"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 20
            cc += 1
            key = "Case Initiated"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 21
            cc += 1
            key = "Case Number"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 22
            cc += 1
            key = "Disputed Product Type"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 23
            cc += 1
            key = "Sum of All Relief Requested"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 24
            cc += 1
            key = "Disposition"
            if key in item and "Disposition Date" not in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 25
            cc += 1
            key = "Disposition Date"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
            # 26
            cc += 1
            key = "Sum of All Relief Awarded"
            if key in item and flag[cc]:
                flag[cc] = False
                v = re.sub(r"\s+", " ", item)
                v = v.split(":")[-1]
                info[key] = v
                continue
        info["Broker name"] = broker_name
        info["CRD#"] = CRD
        info["ID of disclosure event"] = new_IDs[counter]
        INFO.append(info)
    return INFO


# 为每个PDF文件添加sheet
def add_sheet(Excel_path, names):
    book = xlrd.open_workbook(Excel_path)  # 打开一个wordbook
    sheet = book.sheet_by_name("output")
    key_words = sheet.row_values(1, 0, sheet.ncols)
    book = xlwt.Workbook()
    for name in names:
        sheet = book.add_sheet(name, cell_overwrite_ok=True)
        for i in range(len(key_words)):
            sheet.write(0, i, key_words[i])
    new_path = os.path.splitext(Excel_path)[0] + "_Result.xls"
    book.save(new_path)
    return new_path


# 保存到Excel中
def save2Excel(INFO, Excel_path, sheet_name):
    book = xlrd.open_workbook(Excel_path)  # 打开一个wordbook
    sheet = book.sheet_by_name(sheet_name)
    key_words = sheet.row_values(0, 0, sheet.ncols)
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet(sheet_name)

    for index, info in enumerate(INFO):
        for key, value in info.items():
            for ind, key_out in enumerate(key_words):
                if key_out == key:
                    col = ind
                    row = index + 1
                    sheet_copy.write(row, col, value)
                    break
    copy_book.save(Excel_path)


# GUI界面代码
class MYGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.exit_flag = False
        self.try_time = 1532052770.2892158 + 24 * 60 * 60  # 试用时间

        self.initUI()

    def initUI(self):
        self.pdf_label = QLabel("TXT文件夹路径: ")
        self.pdf_btn = QPushButton("选择")
        self.pdf_btn.clicked.connect(self.open_pdf)
        self.pdf_path = QLineEdit("TXT文件夹路径...")
        self.pdf_path.setEnabled(False)
        self.excel_label = QLabel("Excel Demo 路径: ")
        self.excel_btn = QPushButton("选择")
        self.excel_btn.clicked.connect(self.open_excel)
        self.excel_path = QLineEdit("Excel Demo路径...")
        self.excel_path.setEnabled(False)
        self.output_label = QLabel("保存路径: ")
        self.output_path = QLineEdit("保存文件路径...")
        self.output_path.setEnabled(False)
        self.output_btn = QPushButton("选择")
        self.output_btn.clicked.connect(self.open_output)
        self.info = QPlainTextEdit()

        h1 = QHBoxLayout()
        h1.addWidget(self.pdf_label)
        h1.addWidget(self.pdf_path)
        h1.addWidget(self.pdf_btn)

        h2 = QHBoxLayout()
        h2.addWidget(self.excel_label)
        h2.addWidget(self.excel_path)
        h2.addWidget(self.excel_btn)

        h3 = QHBoxLayout()
        h3.addWidget(self.output_label)
        h3.addWidget(self.output_path)
        h3.addWidget(self.output_btn)

        self.run_btn = QPushButton("运行")
        self.run_btn.clicked.connect(self.run)

        self.auth_label = QLabel("密码")
        self.auth_ed = QLineEdit("test_mode")

        exit_btn = QPushButton("退出")
        exit_btn.clicked.connect(self.Exit)
        h4 = QHBoxLayout()
        h4.addWidget(self.auth_label)
        h4.addWidget(self.auth_ed)
        h4.addStretch(1)
        h4.addWidget(self.run_btn)
        h4.addWidget(exit_btn)

        v = QVBoxLayout()
        v.addLayout(h1)
        v.addLayout(h2)
        v.addLayout(h3)
        v.addWidget(self.info)
        v.addLayout(h4)
        self.setLayout(v)
        width = int(QDesktopWidget().screenGeometry().width() / 3)
        height = int(QDesktopWidget().screenGeometry().height() / 3)
        self.setGeometry(100, 100, width, height)
        self.setWindowTitle('TXT to Excel')
        self.show()

    def Exit(self):
        self.exit_flag = True
        qApp.quit()

    def open_pdf(self):
        fname = QFileDialog.getExistingDirectory(self, "Open folder", "/home")
        if fname:
            self.pdf_path.setText(fname)

    def open_excel(self):
        fname = QFileDialog.getOpenFileName(self, "Open Excel", "/home")
        if fname[0]:
            self.excel_path.setText(fname[0])

    def open_output(self):
        fname = QFileDialog.getExistingDirectory(self, "Open output folder",
                                                 "/home")
        if fname:
            self.output_path.setText(fname)

    def run(self):
        self.info.setPlainText("")
        threading.Thread(target=self.scb, args=()).start()
        if self.auth_ed.text() == "a3s7wt29yn1m48zj":
            self.info.insertPlainText("密码正确，开始运行程序!\n")
            threading.Thread(target=self.main_fcn, args=()).start()
        elif self.auth_ed.text() == "test_mode":
            if time.time() < self.try_time:
                self.info.insertPlainText("试用模式，截止时间：2018-07-22 10:00\n")
                threading.Thread(target=self.main_fcn, args=()).start()
            else:
                self.info.insertPlainText(
                    "试用已结束，继续使用请联系yooongchun获取密码，微信：18217235290\n")

        else:
            self.info.insertPlainText(
                "密码错误，请联系yooongchun(微信：18217235290)获取正确密码!\n")

    def scb(self):
        flag = True
        cnt = self.info.document().lineCount()
        while not self.exit_flag:
            if flag:
                self.info.verticalScrollBar().setSliderPosition(
                    self.info.verticalScrollBar().maximum())
            time.sleep(0.01)
            if cnt < self.info.document().lineCount():
                flag = True
                cnt = self.info.document().lineCount()
            else:
                flag = False
            time.sleep(0.01)

    def main_fcn(self):
        # 加载TXT文件夹
        if os.path.isdir(self.pdf_path.text()):
            try:
                txt_path = self.pdf_path.text()
            except Exception:
                self.info.insertPlainText("加载TXT文件夹出错，请重试！\n")
                return
        else:
            self.info.insertPlainText("TXT路径错误，请重试！\n")
            return
        # 加载Excel路径
        if os.path.isfile(self.excel_path.text()):
            demo_path = self.excel_path.text()
        else:
            self.info.insertPlainText("Excel路径错误，请重试！\n")
            return
        # 加载保存路径
        if os.path.isdir(self.output_path.text()):
            name = os.path.basename(demo_path)
            out_path = os.path.join(self.output_path.text(),
                                    name.replace(".xlsx", ".xls"))
        else:
            self.info.insertPlainText("输出路径错误，请重试！\n")
            return
        try:
            shutil.copyfile(demo_path, out_path)
        except Exception:
            self.info.insertPlainText("拷贝临时文件出错，请确保程序有足够运行权限再重试！\n")
            return
        try:
            self.info.insertPlainText("加载TXT文件...\n")
            txt_paths = loadTXT(txt_path)
        except Exception:
            self.info.insertPlainText("加载TXT文件出错！\n")
            return
        names = [os.path.basename(name).split(".")[0] for name in txt_paths]
        try:
            new_path = add_sheet(out_path, names)
        except Exception:
            self.info.insertPlainText("生成sheet失败！\n")
            return
        counter = 0
        for name, path in zip(names, txt_paths):
            counter += 1
            self.info.insertPlainText("正在处理文件: %s %d/%d" %
                                      (name, counter, len(names)) + "\n")
            try:
                INFO = extractor(path)
                save2Excel(INFO, new_path, name)
            except Exception:
                self.info.insertPlainText("文件：%s 出错，跳过...\n" % name)
                continue
        self.info.insertPlainText("运行完成！\n")


# 程序入口
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MYGUI()
    sys.exit(app.exec_())
