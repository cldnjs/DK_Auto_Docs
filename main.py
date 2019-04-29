import os
import sys

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from ExcelFunction import *


class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)

        # 기본 저장 경로 존재여부 확인 후 설정
        self.check_default_path()

        # 화면 및 레이아웃 선언
        self.setWindowTitle('DK Auto Docs')
        self.setWindowIcon(QIcon('./images/DK_logo.png'))
        layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.setLayout(layout)

        # 헤더 선언 및 적용
        self.table = self.create_table()
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignCenter)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        layout.addWidget(self.table)

        # 적용 버튼 선언
        self.apply_btn = QPushButton('OK')
        self.apply_btn.clicked.connect(self.apply_btn_clicked)
        layout.addWidget(self.apply_btn)

    def check_default_path(self):
        """
        기본 저장 경로가 설정 여부 판단 후 없으면 설정
        :return:
        """
        with open('./settings/path.txt', encoding='utf-8') as f:
            path = f.readline()

            # 경로가 존재하지 않으면 다시 설정
            if os.path.exists(path) is False:
                print('저장된 경로가 실제로 존재하지 않음')
                QMessageBox.about(self, 'Message', '저장 경로가 올바르지 않습니다.\n다시 설정해주세요.')
                self.set_default_path()
            else:
                print('설정된 저장 경로: {0}'.format(path))

    def set_default_path(self):
        """
        기본 저장 경로를 설정
        :return:
        """
        path = QFileDialog.getExistingDirectory(self, '저장 경로 선택')

        # 저장 경로를 제대로 선택하지 않으면 다시 선택
        if len(path) <= 0:
            QMessageBox.about(self, 'Message', '저장 경로를 설정해주세요.')
            self.set_default_path()
        else:
            with open('./settings/path.txt', 'w', encoding='utf-8') as f:
                f.flush()
                f.writelines(path)

    @staticmethod
    def get_save_path():
        """
        기본 저장 경로 반환
        :return: String
        """
        with open('./settings/path.txt', 'r', encoding='utf-8') as f:
            path = f.readline()
            if os.path.exists(path) is False:
                    print('폴더 존재하지 않음')
            else:
                return path

    @staticmethod
    def get_product_info():
        """
        엑셀 파일을 읽어 제품 정보를 리스트로 가공
        :return: List
        """
        # 엑셀 파일 로드
        wb, sheet = load_excel(
            filename='제품목록.xlsx',
            sheet_name='Sheet1',
            read_only=False,
            data_only=False
        )
        product_list = load_column_data(sheet, 2, 4, 320)
        standard_list = load_column_data(sheet, 3, 4, 320)
        buy_price = load_column_data(sheet, 5, 4, 320)
        correction_price = load_column_data(sheet, 6, 4, 320)

        product_data = []
        for i in range(len(product_list)):
            data = {
                'product_name': product_list[i],
                'standard': standard_list[i],
                'buy_price': buy_price[i],
                'correction_price': correction_price[i]
            }
            product_data.append(data)

        return product_data

    def create_table(self):
        """
        엑셀에서 데이터를 읽서와서 QTableWidget 생성
        :return: QTableWidget
        """
        # 엑셀에서 받은 데이터로 테이블 생성
        product_data = self.get_product_info()

        table = QTableWidget()
        table.setColumnCount(5)
        table.setRowCount(len(product_data))
        table.setHorizontalHeaderLabels(['선택', '설비명', '규격', '구입', '교정'])

        # 셀에 데이터 체우기 및 체크박스 추가
        for idx, data in enumerate(product_data):
            product_name = QTableWidgetItem(data['product_name'])
            standard = QTableWidgetItem(data['standard'])

            if data['buy_price'] is None:
                buy_price = QTableWidgetItem('')
            else:
                buy_price = QTableWidgetItem(str(data['buy_price']))

            if data['correction_price'] is None:
                correction_price = QTableWidgetItem('')
            else:
                correction_price = QTableWidgetItem(str(data['correction_price']))

            product_name.setTextAlignment(Qt.AlignCenter)
            standard.setTextAlignment(Qt.AlignCenter)
            buy_price.setTextAlignment(Qt.AlignCenter)
            correction_price.setTextAlignment(Qt.AlignCenter)

            table.setItem(idx, 1, product_name)
            table.setItem(idx, 2, standard)
            table.setItem(idx, 3, buy_price)
            table.setItem(idx, 4, correction_price)

        for i in range(table.rowCount()):
            ch = QCheckBox(parent=table)
            ch.clicked.connect(self.product_checked)
            table.setCellWidget(i, 0, ch)

        return table

    @pyqtSlot()
    def apply_btn_clicked(self):
        print('apply button clicked')

    @pyqtSlot()
    def product_checked(self):
        ch = self.sender()
        ix = self.table.indexAt(ch.pos())
        print(ix.row(), ix.column(), ch.isChecked())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = Form()
    form.show()
    exit(app.exec_())
