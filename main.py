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

        # 위젯 선언 및 적용
        tw = self.create_table()
        layout.addWidget(tw)

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
    def create_table():
        """
        엑셀에서 데이터를 읽서와서 QTableWidget 생성
        :return: QTableWidget
        """
        # 엑셀 파일 로드
        wb, sheet = load_excel(
            filename='제품목록.xlsx',
            sheet_name='Sheet1',
            read_only=False,
            data_only=False
        )

        # 제품명, 헤더 얻기
        product_name = load_column_data(sheet, 2, 4)

        # 엑셀에서 읽은 데이터로 테이블위젯 생성
        table = QTableWidget()
        table.setColumnCount(2)
        table.setRowCount(len(product_name))
        table.setHorizontalHeaderLabels(['제품명', 'header2'])

        for i in range(0, len(product_name)):
            table.setItem(i, 0, QTableWidgetItem(product_name[i]))

        # 헤더 넓이 조정
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)

        return table


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = Form()
    form.show()
    exit(app.exec_())
