# -*- coding: gb2312 -*-
# written by kebiao, 2010/08/20

from win32com.client import Dispatch
import os
import sys

class ExcelTool:
	"""
	�򵥵ķ�װexcel���ֲ���
	ϵͳҪ�� windowsϵͳ�� ��װpython2.6�Լ�pywin32-214.win32-py2.6.exe, �Լ�ms office
	"""
	def __init__(self, fileName):
		#try:
		#	self.close()
		#except:
		#	pass

		self.__xapp = Dispatch("Excel.Application")

		self.__xlsx = None

		self.fileName = os.path.abspath(fileName)

	def getWorkbook(self, forcedClose = False):
		"""
		���Workbook�Ѿ�����Ҫ�ȹرպ��
		forcedClose���Ƿ�ǿ�ƹرգ���򿪸�Workbook
		"""
		try:
			wn  = len(self.__xapp.Workbooks)
		except:
			print('�����쳣�˳������������򿪱༭��"ĳ�ļ�"��û�б�����ļ���ɵģ��뱣����ļ�')
			sys.exit(1)

		for x in range(0, wn):
			Workbook = self.__xapp.Workbooks[x]

			if self.fileName == os.path.join(Workbook.Path, Workbook.name):
				if forcedClose:
					Workbook.Close(SaveChanges = False)
				return False

		self.__xlsx = self.__xapp.Workbooks.Open(self.fileName)			#���ļ�
		return True

	def getXApp(self):
		return self.__xapp

	def getXLSX(self):
		return self.__xlsx

	def close(self, saveChanges = False):
		"""
		�ر�excelӦ��
		"""
		if self.__xapp:
			self.__xlsx.Close(SaveChanges = saveChanges)
			if len(self.__xapp.Workbooks) ==0:
				self.__xapp.Quit()
		else:
			return False

	def getSheetCount(self):
		"""
		��ù��������
		"""
		return self.__xlsx.Sheets.Count

	def getSheetNameByIndex(self, index):
		"""
		���excel��ָ������λ���ϵı�����
		"""
		return self.getSheetByIndex(index).Name

	def getSheetByIndex(self, index):
		"""
		���excel��ָ������λ���ϵı�
		"""
		if index in range(1, len(self.__xlsx.Sheets)+1):
			return self.__xlsx.Sheets(index)

		else:
			return None

	def getRowCount(self, sheetIndex):
		"""
		���һ���ж���Ԫ��
		"""
		return self.getSheetByIndex(sheetIndex).Cells(1).CurrentRegion.Columns.Count

	def getColCount(self, sheetIndex):
		"""
		���һ���ж���Ԫ��
		"""
		return self.getSheetByIndex(sheetIndex).Cells(1).CurrentRegion.Rows.Count

	def getValue(self, sheet, row, col):
		"""
		���ĳ���������ĳ��λ���ϵ�ֵ
		"""
		return sheet.Cells(row, col).Value

	def getText(self, sheet, row, col):
		"""
		���ĳ���������ĳ��λ���ϵ�ֵ
		"""
		return sheet.Cells(row, col).Text

	def getRowValues(self, sheet, row):
		"""
		����
		"""
		return sheet.Cells(1).CurrentRegion.Rows[row].Value[0]

	def getSheetRowIters(self, sheet, row):
		"""
		�е�����
		"""
		return sheet.Cells(1).CurrentRegion.Rows

	def getSheetColIters(self, sheet, col):
		"""
		�е�����
		"""
		return sheet.Cells(1).CurrentRegion.Columns

	def getColValues(self, sheet, col):
		"""
		����
		"""
		return sheet.Cells(1).CurrentRegion.Columns[col].Value

#---------------------------------------------------------------------
#   ʹ������
#---------------------------------------------------------------------
def main():
	xbook = ExcelTool("d:\\test1.xlsx")

	print("sheetCount=%i" % xbook.getSheetCount())

	for x in range(1, xbook.getSheetCount() +1 ):
	   print( "      ", xbook.getSheetNameByIndex(x))

	print( "sheet1:rowCount=%i, colCount=%i" % (xbook.getRowCount(1), xbook.getColCount(1)))

	for r in range(1, xbook.getRowCount(1) + 1):
		for c in range(1, xbook.getColCount(1) + 1):
			val = xbook.getValue(xbook.getSheetByIndex(2), r, c)
			if val:
				print( "DATA:", val)

if __name__ == "__main__":
	main()




