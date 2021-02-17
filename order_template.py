from order_template_gui import *
from openpyxl import *
import sys
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem

big_dict = {}
stock_dict = {}
stock_path = ''
stock_start_row = 6
stock_ysku_col = 2
stock_value_col = 6
assr_dict = {}
assortment_path = ''
assr_ysku_col = 29
assr_sku_col = 2
assr_art_col = 10
stat_dict = {}
stat_sales_path = ''
stat_start_row = 3
stat_sku_col = 4
stat_value_col = 6
route_path = ''


class MyWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # кнопки
        self.ui.stock_Button.clicked.connect(lambda: self.ui.stock_path.setText(QFileDialog.getOpenFileName()[0]))
        self.ui.assortment_Button.clicked.connect(
            lambda: self.ui.assortment_path.setText(QFileDialog.getOpenFileName()[0]))
        self.ui.stat_sales_Button.clicked.connect(
            lambda: self.ui.stat_sales_path.setText(QFileDialog.getOpenFileName()[0]))
        self.ui.route_Button.clicked.connect(
            lambda: self.ui.route_kotelniki_path.setText(QFileDialog.getOpenFileName()[0]))
        self.ui.create_Button.clicked.connect(self.get_paths)

    def get_paths(self):
        global stock_path, assortment_path, stat_sales_path, route_path
        stock_path = self.ui.stock_path.text()
        assortment_path = self.ui.assortment_path.text()
        stat_sales_path = self.ui.stat_sales_path.text()
        route_path = self.ui.route_kotelniki_path.text()
        self.create_assortment_dict()
        self.create_stock_dict()
        self.create_stat_dict()
        self.create_big_dict()
        #TODO отчет заказы обработать способом дубль плюс
        #TODO фильтруя через остатки в биг создать словарь, остатик вытягивать из бига
        #TODO удалить из бига не пустые остатки
        #TODO апдэйт буги сделать


    def create_big_dict(self):
        print('start big')
        for big_key in big_dict:
            if big_key in stock_dict:
                if big_dict[big_key]['Stock'] > stat_dict[big_key]['Order'] / 2:
                    big_dict.pop(big_key)
                else:
                    big_dict[big_key]['Order'] = stat_dict[big_key]['Order']
                    print('del', big_key)
            elif big_dict[big_key]['Stock'] > 0:
                big_dict.pop(big_key)
                print('no_orders_del', big_key)
        for key in big_dict:
            print(key, big_dict[key])


    def create_stat_dict(self):
        print('start stat')
        global stat_dict
        stat = load_workbook(stat_sales_path)
        stat_sheet = stat.worksheets[0]
        self.del_rep(stat_sheet, stat_start_row, stat_sku_col, stat_value_col)
        stat_dict = {x[3].value: {'YSKU': '', 'ART': 'УТ', 'Stock': 0, 'Order': int(x[5].value)}
                      for i, x in enumerate(stat_sheet) if x[3].value in assr_dict and i >= stat_start_row - 1}

    def create_assortment_dict(self):
        print('start assr')
        global assr_dict
        assr = load_workbook(assortment_path)
        assr_sheet = assr.worksheets[2]
        assr_dict = {x[1].value: {'YSKU': x[28].value, 'ART': x[9].value, 'Stock': 0, 'Order': 5} \
                      for i, x in enumerate(assr_sheet) if i > 3}

    def create_stock_dict(self):
        print('start stock')
        global stock_dict, big_dict
        stock = load_workbook(stock_path)
        stock_sheet = stock.worksheets[0]
        self.del_rep(stock_sheet, stock_start_row, stock_ysku_col, stock_value_col)
        stock_dict = {x[0].value: {'YSKU': x[1].value, 'ART': 'УТ', 'Stock': int(x[5].value), 'Order': 5} \
                      for i, x in enumerate(stock_sheet) if i >= stock_start_row - 1 and x[0].value in assr_dict}
        big_dict = assr_dict.copy()
        big_dict.update(stock_dict)


    def del_rep(self, sheet, start_row, ysku_col, value_col):
        sheet_mr = sheet.max_row
        for i in range(start_row, sheet_mr + 1):
            ysku = str(sheet.cell(row=i, column=ysku_col).value).strip()
            if ysku == None or ysku == '' or ysku == 'None':
                continue
            for x in range(sheet_mr, i, - 1):
                if ysku == str(sheet.cell(row=x, column=ysku_col).value).strip():
                    founded_val = int(sheet.cell(row=i, column=value_col).value)
                    sheet.cell(row=i, column=value_col).value = \
                        sheet.cell(row=x, column=value_col).value + founded_val
                    sheet.delete_rows(x)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())
