from PyQt5 import QtWidgets , QtGui , QtCore 
from PyQt5.QtWidgets import QApplication, QMainWindow , QAbstractItemView , QMessageBox
import sys
import pandas as pd 
import zipfile , requests , io , openpyxl , os , numpy
from pathlib import Path

parent = Path(__file__).resolve().parent
parent = str(parent)

class MyWindow(QMainWindow):

	def __init__(self):
		super(MyWindow, self).__init__()
		self.setWindowTitle("MWPL Difference Calculator")
		self.setFixedSize(640, 480)
		self.initUI()
		self.btn_2.setEnabled(False)

	def initUI(self):
		self.label = QtWidgets.QLabel(self)  #MWPL Title
		self.label.setText("MWPL")
		self.label.setGeometry(260,25,125,30)
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI Light")
		font.setPointSize(28)
		self.label.setFont(font)

		self.label_2 = QtWidgets.QLabel(self) #Previous Date
		self.label_2.setText("Previous Date")
		self.label_2.setGeometry(75,110,105,12)
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI Light")
		font.setPointSize(12)
		self.label_2.setFont(font)

		self.label_3 = QtWidgets.QLabel(self)  #Current Date
		self.label_3.setText("Current Date")
		self.label_3.setGeometry(465,110,105,12)
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI Light")
		font.setPointSize(12)
		self.label_3.setFont(font)

		self.label_4 = QtWidgets.QLabel(self)  #Percentage
		self.label_4.setText("Percentage")
		self.label_4.setGeometry(270,100,105,20)
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI Light")
		font.setPointSize(12)
		self.label_4.setFont(font)


		self.dateEdit = QtWidgets.QDateEdit(self)  #Previous Date dateEdit
		self.dateEdit.setGeometry(65,140,150,30)
		self.dateEdit.setMinimumDate(QtCore.QDate(2020,1,1))
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI")
		font.setPointSize(15)
		self.dateEdit.setFont(font)
		self.dateEdit.setDisplayFormat("dd-MM-yyyy")
		self.date = self.dateEdit.date()
	

		self.dateEdit_2 = QtWidgets.QDateEdit(self)  #Current Date dateEdit
		self.dateEdit_2.setGeometry(455,140,150,30)
		self.dateEdit_2.setMinimumDate(QtCore.QDate(2020,1,1))
		font = QtGui.QFont()
		font.setFamily("Microsoft JhengHei UI")
		font.setPointSize(15)
		self.dateEdit_2.setFont(font)
		self.dateEdit_2.setDisplayFormat("dd-MM-yyyy")
		self.date_2 = self.dateEdit_2.date()
		

		self.spinBox = QtWidgets.QSpinBox(self)     #Percentage Spin Box
		self.spinBox.setGeometry(272,140,75,30)

		self.btn = QtWidgets.QPushButton(self)      #Download Button
		self.btn.setText("Download files")
		self.btn.setGeometry(200,200,100,35)
		self.btn.clicked.connect(self.click)

		self.btn_2 = QtWidgets.QPushButton(self)      #Check Button
		self.btn_2.setText("Check Percentage")
		self.btn_2.setGeometry(330,200,100,35)
		self.btn_2.clicked.connect(self.check)

		self.listView = QtWidgets.QListView(self)
		self.listView.setEditTriggers(QAbstractItemView.NoEditTriggers)
		self.listView.setGeometry(170,250,300,220)
		


	def process(self,date,url):
	
		#To download the Zip file
		r = requests.get(url)

		#To extract the downloaded Zip file
		try:
			z = zipfile.ZipFile(io.BytesIO(r.content))
			z.extractall()
		except Exception:
			msg = QMessageBox()
			msg.setWindowTitle("Error")
			msg.setText("No file found in the given date")
			msg.setIcon(QMessageBox.Critical)
			x = msg.exec_()
			os.startfile(__file__)
			sys.exit()

		
		#To Convert CSV to Excel 
		csv_directory = parent + "\combineoi_"+(date)+".csv"
		csv = pd.read_csv(csv_directory)
		csv.to_excel( parent +"\combineoi_"+(date)+".xlsx",index = None, header = True)

		#To delete the CSV file and XML file
		os.remove(parent + "\combineoi_"+(date)+".csv")
		try:
			os.remove(parent + "\combineoi_"+(date)+".xml")
		except Exception as e:
			pass
		
		#To open excel file as DataFrame and sort it.
		df = pd.read_excel(parent + "\combineoi_"+(date)+".xlsx")
		sorted_excel = df.sort_values(' Scrip Name')
		sorted_excel.to_excel(parent + "\combineoi_"+(date)+".xlsx", index = False, header=True)
		excel_directory = parent + "\combineoi_"+(date)+".xlsx"
		workbook = openpyxl.load_workbook(excel_directory)
		sheet = workbook.active

		#To modify the excel sheet
		sheet['H1'] = "Percentage"
		no_of_rows = sheet.max_row

		#To Calculate Percentage
		i = 2
		while(i<=no_of_rows):
			i_inString = str(i)
			limit_for_next_day_object = sheet.cell(row = i , column = 7)
			limit_for_next_day_value = limit_for_next_day_object.value
			if(limit_for_next_day_value == 'No Fresh Positions'):
				i = i + 1
				continue
			mwpl_object = sheet.cell(row = i , column = 5)
			mwpl_value = mwpl_object.value
			sheet['H'+i_inString]= (int(limit_for_next_day_value)/int(mwpl_value))*100
			i = i + 1 
		#To save the file
		workbook.save("combineoi_"+(date)+".xlsx")
		self.btn_2.setEnabled(True)


	


	def click(self):
		self.value = self.spinBox.value()
		percentage = self.value

		self.date = self.dateEdit.date()
		self.pydate = self.date.toPyDate()
		date1 = self.pydate
		date1 = date1.strftime("%d%m%Y")
		url = "https://www1.nseindia.com/archives/nsccl/mwpl/combineoi_"+(date1)+".zip"
		self.process(date1,url)

		self.date_2 = self.dateEdit_2.date()
		self.pydate_2 = self.date_2.toPyDate()
		date2 = self.pydate_2
		date2 = date2.strftime("%d%m%Y")
		url = "https://www1.nseindia.com/archives/nsccl/mwpl/combineoi_"+(date2)+".zip"
		self.process(date2,url)


	


		#My_MWPL excel
		self.mwpl_workbook = openpyxl.Workbook()
		self.mwpl_sheet = self.mwpl_workbook.active
		i =2
		no_of_rows = 150

		while(i<=no_of_rows):
			#Adding data from Date1 excel file into MWPL excel
			excel_directory = parent + "\combineoi_"+(date1)+".xlsx"
			workbook = openpyxl.load_workbook(excel_directory)
			sheet = workbook.active
			no_of_rows_in_1 = sheet.max_row
			
			#Column : 1 Stock Name_1 
			stockname_object = sheet.cell(row = i,column = 3)
			stockname_value = stockname_object.value
			i_inString = str(i)
			a = "A" + i_inString
			mwpl_stockname = self.mwpl_sheet[a]
			mwpl_stockname.value = stockname_value
			
			#Column : 2 Percentage_1 
			percentage_object = sheet.cell(row = i, column = 8)
			percentage_value = percentage_object.value
			b = "B" + i_inString
			mwpl_percentage = self.mwpl_sheet[b]
			mwpl_percentage.value = percentage_value
			
			#Adding data from Date2 excel file into MWPL excel
			excel_directory2 = parent + "\combineoi_"+(date2)+".xlsx"
			workbook_2 = openpyxl.load_workbook(excel_directory2)
			sheet_2 = workbook_2.active
			no_of_rows_in_2 = sheet_2.max_row

			#Column : 3 Stock Name_2 
			stockname_object_2 = sheet_2.cell(row = i,column = 3)
			stockname_value_2 = stockname_object_2.value
			i_inString = str(i)
			c = "C" + i_inString
			mwpl_stockname_2 = self.mwpl_sheet[c]
			mwpl_stockname_2.value = stockname_value_2
			
			#Column : 4 Percentage_2 
			percentage_2_object = sheet_2.cell(row = i , column = 8)
			percentage_2_value = percentage_2_object.value
			d = "D" + i_inString
			mwpl_percentage_2 = self.mwpl_sheet[d]
			mwpl_percentage_2.value = percentage_2_value
			
			i = i+ 1

		self.rows=self.mwpl_sheet.max_row
		i=1
		j=1

		for i in range(1,self.rows+1):
			i_inString = str(i)
			a = "A" + i_inString
			mwpl_stockname = self.mwpl_sheet[a]
			for j in range(1,self.rows+1):
				j_inString = str(j)
				c = "C" + j_inString
				mwpl_stockname_2 = self.mwpl_sheet[c]
				
				if(mwpl_stockname.value == mwpl_stockname_2.value):
					b = "B"+i_inString
					mwpl_percentage = self.mwpl_sheet[b]
					d = "D"+j_inString
					mwpl_percentage_2 = self.mwpl_sheet[d]
					e = "E"+i_inString
					difference_mwpl = self.mwpl_sheet[e]
					if(mwpl_percentage.value == None or mwpl_percentage_2.value == None):
						continue
					difference_mwpl.value = mwpl_percentage.value - mwpl_percentage_2.value


		#self.mwpl_workbook.save(parent + "\mwpl.xlsx")
		os.remove(parent + "\combineoi_"+(date1)+".xlsx")
		os.remove(parent + "\combineoi_"+(date2)+".xlsx")


	def func(self,percentage):
			i = 2
			entries = []
			while(i<=self.rows):
				difference_object = self.mwpl_sheet.cell(row = i, column = 5)
				difference_object_value = difference_object.value
				if(difference_object_value == None ):
					i = i + 1 
					continue
				if(difference_object_value >= percentage):
					mwpl_stockname = self.mwpl_sheet.cell(row = i, column = 1)
					mwpl_stockname_value = mwpl_stockname.value
					entries.append(mwpl_stockname_value)
				i = i + 1

			model = QtGui.QStandardItemModel()
			self.listView.setModel(model)
			for i in entries:
				item = QtGui.QStandardItem(i)
				model.appendRow(item)


	def check(self):
		self.value = self.spinBox.value()
		percentage = self.value 
		self.func(percentage)

		
def window():
	app = QApplication(sys.argv)
	win = MyWindow()
	win.show()
	sys.exit(app.exec_())
window()
