from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import time

class MyFrame(QWidget):

    def __init__(self):
        super().__init__()
        self.table = QTableWidget(5, 5, self)  # row, column
        self.table.setHorizontalHeaderLabels(["", "종목명", "현재가(문자)",  "현재가(숫자)", "거래량"])

        data = [
            ("삼성전자", "200000", 200000, "25000"),
            ("셀트리온", "2100", 2100, "1500"),
            ("현대차", "190000", 190000, "300000"),
            ("기아차", "150000", 150000, "240000")
                ]


        for idx, (hname, price_str, price, vol) in enumerate(data):
            # 사용자정의 item 과 checkbox widget 을, 동일한 cell 에 넣어서 , 추후 정렬 가능하게 한다.
            item = MyQTableWidgetItemCheckBox()
            self.table.setItem(idx, 0, item)
            chbox = MyCheckBox(item)
            self.table.setCellWidget(idx, 0, chbox)
            chbox.stateChanged.connect(self.__checkbox_change)  # sender() 확인용 예..
            self.table.setItem(idx, 1, QTableWidgetItem(hname))
            self.table.setItem(idx, 2, QTableWidgetItem(price_str))

            # 숫자를 기준으로 정렬하기 위함. -- default 는 '문자'임.
            item = QTableWidgetItem()
            item.setData(Qt.DisplayRole, price)
            self.table.setItem(idx, 3, item)
            self.table.setItem(idx, 4, QTableWidgetItem(vol))

        self.table.setSortingEnabled(False)  # 정렬기능
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()  # 이것만으로는 checkbox 컬럼은 잘 조절안됨.
        self.table.setColumnWidth(0, 15)  # checkbox 컬럼 폭 강제 조절.
        self.table.cellClicked.connect(self._cellclicked)

        # 컬럼 헤더를 click 시에만 정렬하기.
        hheader = self.table.horizontalHeader()  # qtablewidget --> qtableview --> horizontalHeader() --> QHeaderView
        hheader.sectionClicked.connect(self._horizontal_header_clicked)
        vbox = QVBoxLayout(self)
        vbox.addWidget(self.table)
        self.setLayout(vbox)


    def __checkbox_change(self, checkvalue):
        chbox = self.sender()  # signal을 보낸 MyCheckBox instance
        print("checkbox sender row = ", chbox.get_row())

    def _cellclicked(self, row, col):
        print("_cellclicked... ", row, col)

    def _horizontal_header_clicked(self, idx):
        self.table.setSortingEnabled(True)  # 정렬기능 on
        self.table.setSortingEnabled(False)  # 정렬기능 off

class MyCheckBox(QCheckBox):

    def __init__(self, item):
        super().__init__()
        self.item = item
        self.mycheckvalue = 0   # 0 --> unchecked, 2 --> checked
        self.stateChanged.connect(self.__checkbox_change)
        self.stateChanged.connect(self.item.my_setdata)  # checked 여부로 정렬을 하기위한 data 저장

    def __checkbox_change(self, checkvalue):
        self.mycheckvalue = checkvalue
        print("checkbox row= ", self.get_row())

    def get_row(self):
        return self.item.row()

class MyQTableWidgetItemCheckBox(QTableWidgetItem):

    def __init__(self):
        super().__init__()
        self.setData(Qt.UserRole, 0)

    def __lt__(self, other):
        return self.data(Qt.UserRole) < other.data(Qt.UserRole)

    def my_setdata(self, value):
        self.setData(Qt.UserRole, value)


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    frame = MyFrame()
    frame.setWindowTitle("정렬하기")
    frame.resize(600, 400)  # width, height
    frame.show()
    app.exec_()