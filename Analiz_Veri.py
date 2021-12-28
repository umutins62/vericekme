import requests
from bs4 import BeautifulSoup
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import *
import sys
import matplotlib
matplotlib.use('Qt5Agg')
from PyQt5 import  QtWidgets
import os
from selenium.webdriver.chrome.options import Options
import openpyxl
from selenium import webdriver
import time
import pandas as pd


class analiz_cek(QWidget):


    def __init__(self):
        super().__init__()
        self.setUI()
        # self.hesap()


    def setUI(self):
        self.setWindowTitle("Poz Analizi Öğrenme")
        self.setWindowIcon(QIcon(":hesap.ico"))
        self.setGeometry(700, 300, 840, 600)
        self.setFixedSize(self.size())

        self.alan=QLineEdit()
        self.alan.setPlaceholderText("Poz nosu veya tanımı giriniz!")


        self.yil=QComboBox()

        self.sinif=QComboBox()


        self.yapimaliyet=QLabel("-----")
        self.yapiyaklasikmaliyet=QLabel("-----")
        self.bos1=QLabel("")
        self.bos2=QLabel("")
        self.bos3=QLabel("Bu Program Umut Çelik tarafından yapılmıştır.")

        self.hesp=QPushButton("Getir")
        self.hesp.clicked.connect(self.hesap)

        self.acik=QPushButton("Excel Aç")
        self.acik.clicked.connect(self.exceleaktar)

        self.ara=QPushButton("Ara")
        self.ara.clicked.connect(self.sayfalar)

        self.acik.setEnabled(False)

        self.yil=QComboBox()

        self.tableWidget = QTableWidget()
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setColumnCount(7)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(False)
        self.tableWidget.horizontalHeader().setSortIndicatorShown(False)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.verticalHeader().setCascadingSectionResizes(False)
        self.tableWidget.verticalHeader().setStretchLastSection(False)
        self.tableWidget.setHorizontalHeaderLabels(("Kitap","Poz No","Açıklama", "Birim","Miktar", "Birim Fiyatı","Tutar(TL)"))
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        style = "::section {""background-color: lightblue; }"
        self.tableWidget.horizontalHeader().setStyleSheet(style)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)


        self.listWidget = QListWidget()


        hbox1=QHBoxLayout()


        hbox1.addStretch()
        hbox1.addWidget(self.ara)
        hbox1.addWidget(self.hesp)
        hbox1.addWidget(self.acik)


        h_box = QtWidgets.QVBoxLayout()
        self.groupbox = QGroupBox("Girdiler")
        self.groupbox1 = QGroupBox("")
        h_box.addWidget(self.groupbox)
        h_box.addWidget(self.groupbox1)

        form=QFormLayout()
        form.addRow("Aranacak Poz veya Tanım :",self.alan)
        form.addRow(self.bos2,hbox1)
        hbox2=QVBoxLayout()
        self.poz_adi=QLabel("Poz Adı :")
        self.poz_tanim=QLabel("Poz Açıklama :")

        hbox2.addWidget(self.listWidget)
        hbox2.addWidget(self.poz_adi)
        hbox2.addWidget(self.poz_tanim)
        hbox2.addWidget(self.tableWidget)


        self.groupbox.setLayout(form)
        self.groupbox1.setLayout(hbox2)



        self.setLayout(h_box)

        self.show()

    def exceleaktar(self):
        os.system(str(self.deger1)+".xlsx")

    def sayfalar(self):
        chromeOptions = Options()
        chromeOptions.headless = True


        driver=webdriver.Chrome(executable_path="C:\webdrivers\chromedriver.exe")

        driver.get("https://www.birimfiyat.net/")
        driver.maximize_window()
        self.deger=self.alan.text()
        time.sleep(2)

        driver.find_element_by_id("poz-ara-text").send_keys(self.deger)

        driver.find_element_by_xpath("//*[@id='home']/div/div/form/div/div/div/div/div[1]/div/div[2]/div[2]/div/div[2]").click()


        time.sleep(1)
        driver.find_element_by_xpath("//*[@id='aramaSecenekleri ']/div/div[1]/div/div[1]/div/label/span").click()
        time.sleep(1)

        driver.find_element_by_xpath("//*[@id='home']/div/div/form/div/div/div/div/div[1]/div/div[1]/div/button").click()



        liste1=driver.find_element_by_xpath("//*[@id='sonuclaraKonumlan']/div/div/div[2]/span").text


        url=driver.current_url


        self.pozara(url,liste1)



    def pozara(self,url,liste1):


        r=requests.get(url)
        print(r.status_code)
        soup=BeautifulSoup(r.content,"lxml")




        liste0=soup.find("tbody",attrs={"class":"arama_sonuc_tbody_on_yuz"}).select("tr>td:nth-of-type(1)")
        liste=soup.find("tbody",attrs={"class":"arama_sonuc_tbody_on_yuz"}).select("tr>td:nth-of-type(2)")
        liste2=soup.find("tbody",attrs={"class":"arama_sonuc_tbody_on_yuz"}).select("tr>td:nth-of-type(4)")





        self.ara.setText("Ara("+str(liste1)+")")





        for i in range(1,10):
            poz=liste[i].find("a").text
            aciklama=liste2[i].find("a").text
            kurum=liste0[i].find("a").text
            self.listWidget.addItem(kurum+"-"+poz+"-"+aciklama)





    def hesap(self):
        if self.alan.text()!="":
            chromeOptions = Options()
            chromeOptions.headless = True

            driver=webdriver.Chrome(executable_path="C:\webdrivers\chromedriver.exe")

            driver.get("https://www.birimfiyat.net/")
            driver.maximize_window()
            self.deger=self.alan.text()
            time.sleep(4)
            driver.find_element_by_id("poz-ara-text").send_keys(self.deger)
            try:
                driver.find_element_by_xpath("//*[@id='home']/div/div/form/div/div/div/div/div[1]/div/div[2]/div[2]/div/div[2]").click()
                time.sleep(2)
                driver.find_element_by_xpath("//*[@id='aramaSecenekleri ']/div/div[1]/div/div[1]/div/label/span").click()
                time.sleep(2)
                driver.find_element_by_xpath("//*[@id='home']/div/div/form/div/div/div/div/div[1]/div/div[1]/div/button").click()
                time.sleep(2)
                driver.find_element_by_xpath("//*[@id='veritablosu']/tbody/tr[1]/td[2]/a").click()
                poz_adi=driver.find_element_by_xpath("/html/body/section[2]/div/div/div/div[2]/div/div[1]/div[1]/div/table/tbody/tr[2]/td[2]").text
                poz_tanim=driver.find_element_by_xpath("//*[@id='poz-tarifi']/div").text
                self.poz_adi.setText("Poz Adı :"+poz_adi)
                self.poz_tanim.setText("Poz Tarifi :"+poz_tanim)
                url2=driver.current_url
                d=pd.read_html(url2,encoding = 'utf-8', decimal=",", thousands='.', converters={'Account': str})

                try:
                    self.df=d[3]

                    print(self.df)


                    n=len(self.df)
                    print(n)

                    self.df.drop(self.df.tail(1).index,inplace=True)
                    self.df.drop(self.df.head(1).index,inplace=True) # drop first n rows


                    self.df[str(self.deger)+' Pozu Analizi.4']=self.df[str(self.deger)+' Pozu Analizi.4'].astype('float')
                    self.df[str(self.deger)+' Pozu Analizi.5']=self.df[str(self.deger)+' Pozu Analizi.5'].astype('float')
                    self.df[str(self.deger)+' Pozu Analizi.6']=self.df[str(self.deger)+' Pozu Analizi.6'].astype('float')

                    print(self.df)
                    self.deger1=self.deger.replace("/","-")

                    self.df.to_excel(str(self.deger1)+".xlsx")

                    filename = str(self.deger1)+".xlsx"
                    wb = openpyxl.load_workbook(filename,data_only=True)
                    sheet = wb['Sheet1']





                    sheet.delete_cols(1)
                    row1_ = int(sheet.max_row)
                    sum=row1_+1

                    sheet.column_dimensions['a'].width = 15
                    sheet.column_dimensions['b'].width = 15
                    sheet.column_dimensions['C'].width = 100
                    sheet.column_dimensions['e'].width = 15
                    sheet.column_dimensions['f'].width = 20
                    sheet.column_dimensions['g'].width = 20

                    sheet['A1']="Kitap"
                    sheet['B1']="Poz No"
                    sheet['C1']="Açıklama"
                    sheet['D1']="Birim"
                    sheet['E1']="Miktar"
                    sheet['F1']="Birim Fiyat(TL)"
                    sheet['G1']="Tutar(TL)"


                    _cell = sheet["G"+str(sum)]
                    _cell.number_format = '0.00'

                    top= self.df[str(self.deger)+' Pozu Analizi.6'].sum()
                    sheet["G"+str(sum)] = top
                    sheet["F"+str(sum)] = "Analiz Toplam"

                    print(top)

                    sheet["F"+str(sum+1)] = "%25 Kar"
                    sheet["G"+str(sum+1)] =top*0.25

                    sheet["F"+str(sum+2)] = "Analiz Karlı Toplam"
                    sheet["G"+str(sum+2)] =top*0.25+top




                    wb.save(filename)
                    self.aciklama()
                    self.acik.setEnabled(True)
                except:
                    QMessageBox.information(self, "Dikkat",
                                            "Aranan pozun analizi bulunmamaktadır.")
                    driver.quit()
            except:
                QMessageBox.information(self, "Dikkat",
                                        "Aranan poz bulunmamaktadır.")
                driver.quit()

        
        else:
            QMessageBox.information(self, "Dikkat",
                                    "Arama kutucuğunu boş bırakmayınız.")


    def aciklama(self):
        df1 = pd.read_excel(str(self.deger1)+".xlsx", header=0)
        df = df1.where(pd.notnull(df1), "")

        self.tableWidget.setColumnCount(len(df.columns))
        self.tableWidget.setRowCount(len(df.index))

        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()




if __name__ == "__main__":
    app = QApplication(sys.argv)
    pencere = analiz_cek()
    sys.exit(app.exec())