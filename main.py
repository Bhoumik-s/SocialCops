import sys
import os

from getEpic import ConvertToHtml, ParseEpicFromHtml
from getData import getData

if __name__ == "__main__":
	fileName = sys.argv[1]
	print("Converting PDF to html")
	html = ConvertToHtml(fileName)
	print("Parsing Epic No from html")
	ParseEpicFromHtml(html)
	os.remove("output.html")
	print("Getting Data from CEO site")
	getData("data.xlsx")