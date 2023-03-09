from PyPDF2 import PdfReader
import os
import sys
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QMessageBox)
import datetime
import csv

class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        self.title = "PDF TO EXCEL CONVERTER"
        self.top = 100
        self.left = 100
        self.width = 480
        self.height = 300

        self.pushButton = QPushButton("Click me to start the process", self)
        self.pushButton.move(110, 70)
        self.pushButton.resize(260,120)
        self.pushButton.setFont(QFont('Arial', 13))

        self.pushButton.clicked.connect(lambda: self.process())

        self.main_window()

    def main_window(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.top, self.left, self.width, self.height)
        self.show()

    def process(self):
        dlg = QMessageBox()
        dlg.setIcon(QMessageBox.Information)
        dlg.setWindowTitle("PDF TO EXCEL CONVERTER")
        dlg.setFont(QFont('Calibri', 13))

        check1 = 0

        # Reading all the pdf files from the directory and store them in all_files list
        all_files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]

        if len(all_files) == 0:
            dlg.setText("PDF File Error")
            dlg.exec()
            check1 = 1


        if check1 == 0:
            d = datetime.datetime.now()
            dt = str(str(d.day) + '-' + str(d.month) + '-' + str(d.year))
            check = 0
            m = 0
            for file in all_files:
                m += 1
                reader = PdfReader(file)
                number_of_pages = len(reader.pages)
                n = len(file)
                output_file = file[0:n-4] + ' _ ' + dt
                output_filename = output_file + '.csv'

                if (' '.join(reader.pages[0].extract_text().split("\n"))).find("SKU") == -1:
                    continue

                fields = ['ship-address-1', 'Street', 'Town', 'Country', 'ship-postal-code', 'first_name', 'last_name',
                          'delivery_instructions', 'contents', 'labelreference', 'reference2', 'Order Quantity',
                          'Weight', 'Length', 'Width', 'Height', 'item-price', 'quantity-purchased', 'serviceid','sequence']
                rows = []

                for i in range(0, number_of_pages):
                    try:
                        page = reader.pages[i]
                        text = (page.extract_text()).split("\n")
                        check_seller = 0
                        index_of_orderId = 0
                        quantity = 1
                        address1 = street = town = postal = first_name = last_name = order_id = contents = sku = refrence2 = ""
                        for j in text:
                            if j.find("Seller Name") != -1:
                                check_seller = 1
                                print(i+1)
                                check = 1
                                index_of_current = text.index(j)
                                name = text[index_of_current + 4]
                                last_name = ""
                                if name.find('-') != -1:
                                    first_name = name.split("-")[0]
                                    last_name = name.split("-")[1]
                                elif len(name.split()) > 1:
                                    first_name = name.split()[0]
                                    last_name = ' '.join(name.split()[1:])
                                else:
                                    first_name = name
                                address1 = text[index_of_current + 5]
                                street = text[index_of_current + 6]
                                postal = text[index_of_current + 8]
                                country = text[index_of_current + 9]
                                address6 = text[index_of_current + 10]
                                if country.find("Order ID") != -1:  # if 5 lines
                                    town = ""
                                    postal = text[index_of_current + 7]
                                    index_of_orderId = index_of_current + 9
                                elif address6.find("Order ID") != -1:  # if 6 lines
                                    town = text[index_of_current + 7]
                                    index_of_orderId = index_of_current + 10
                                else:  # if 7 lines
                                    town = text[index_of_current + 7]
                                    index_of_current += 8
                                    while(text[index_of_current + 1].find("United")) == -1:
                                        town += text[index_of_current]
                                        index_of_current += 1
                                    postal = text[index_of_current]
                                    index_of_orderId = index_of_current + 1
                        if check_seller == 0:
                            index_of_orderId = 5
                        while (text[index_of_orderId].find('Order ID')) == -1:
                            index_of_orderId += 1
                        order_id = text[index_of_orderId].split(": ")[1]
                        list1 = text[index_of_orderId + 3].split(" ")
                        quantity = int(list1[0])
                        del list1[0]  # deleting quantity from list
                        content = " "
                        contents = content.join(list1)  # converting list to str
                        index = index_of_orderId + 4
                        while text[index].find("SKU") == -1:
                            contents += text[index]
                            index += 1
                        sku = text[index].split(": ")[1]
                        reference2 = ""
                        if len(sku) > 20:
                            reference2 = sku[20:]
                            sku = sku[0:20]
                    except Exception as x:
                        print(x)
                        continue
                    finally:
                        row = [address1, street, town, 'United Kingdom', postal, first_name, last_name, order_id,
                                   contents, sku, reference2, quantity, 5, 60, 45, 10, 20, 1, 'Standard Parcel',i+1]
                        rows.append(row)
                # writing to csv file
                with open(output_filename, 'w', newline='', encoding="utf-8") as csvfile:
                    # creating a csv writer object
                    csvwriter = csv.writer(csvfile)
                    # writing the fields
                    csvwriter.writerow(fields)
                    for r in rows:
                        csvwriter.writerow(r)

            if check == 0:
                dlg.setText("Wrong pdf files given")
                dlg.exec()
            else:
                dlg.setText("Success")
                dlg.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec())