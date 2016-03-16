import codecs
from bs4 import BeautifulSoup
import sys
import os
import re
import xlsxwriter


def OpenHtml(fileName):
	os.popen("pdf2txt.py -o " + fileName + ".html " + fileName + ".pdf" )
	html = codecs.open(fileName+".html", 'r')
	return html

def ParseHtml(html,District):
	soup = BeautifulSoup(html,'lxml')
	style = soup.find_all('span', {'style':'font-family: AAAAAC+Arial; font-size:10px'})
	pattern = re.compile('^[ANB|UP|GLT]')
	
	workbook = xlsxwriter.Workbook(District + '.xlsx')
	worksheet = workbook.add_worksheet()
	worksheet.set_column('A:A', 20)
	worksheet.set_column('B:B', 20)
	worksheet.set_column('C:C', 20)
	worksheet.set_column('D:D', 22)
	worksheet.set_column('E:E', 27)

	worksheet.write('A1', 'EPIC NO')
	worksheet.write('B1', 'Name')
	worksheet.write('C1', 'Name (Hindi)')
	worksheet.write('D1', "Father/Husband's Name")
	worksheet.write('E1', "Father/Husband's Name(Hindi)")
	worksheet.write('F1', 'Age')
	worksheet.write('G1', 'Gender')
	i = 1
	for epic in style:
		if pattern.match(str(epic.text)):
			i = i + 1
			worksheet.write('A'+ str(i), str(epic.text).rsplit(' ')[0]) 
	workbook.close()


if __name__ == "__main__":
	fileName = sys.argv[1]
	District = sys.argv[2]
	html = OpenHtml(fileName)
	ParseHtml(html,District)
	os.remove(fileName+".html")

