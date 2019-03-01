import re
import xlsxwriter

from os import listdir
from os.path import isfile, join

def main():
	foodfiles = [f for f in listdir(r".\Mods\AutoGen\Food") if isfile(join(r".\Mods\AutoGen\Food", f))]
	seedfiles = [f for f in listdir(r".\Mods\AutoGen\Seed") if isfile(join(r".\Mods\AutoGen\Seed", f))]
	foodandseedfilecontent = []

	for file in foodfiles:
		file_object = open(r".\Mods\AutoGen\Food\{f}".format(f=file),"r")
		foodandseedfilecontent.append(file_object.read())
	for file in seedfiles:
		file_object = open(r".\Mods\AutoGen\Seed\{f}".format(f=file),"r")
		foodandseedfilecontent.append(file_object.read())

	workbook = xlsxwriter.Workbook('fooddata.xlsx')
	worksheet = workbook.add_worksheet()

	for idx, text in enumerate(foodandseedfilecontent):
		try:
			namemach = re.search(r"DisplayName .*\"(.*)\"",text)
			worksheet.write(idx,0, namemach[1])

			caloriematch = re.search(r"Calories.*n (-?\d{1,4})\;",text)
			worksheet.write(idx,1, caloriematch[1])

			carbsmatch = re.search(r"Carbs = (\d{1,4})",text)
			worksheet.write(idx,2, carbsmatch[1])

			fatmatch = re.search(r"Fat = (\d{1,4})",text)
			worksheet.write(idx,3, fatmatch[1])

			protmatch = re.search(r"Protein = (\d{1,4})",text)
			worksheet.write(idx,4, protmatch[1])

			vitaminsmatch = re.search(r"Vitamins = (\d{1,4})",text)
			worksheet.write(idx,5, vitaminsmatch[1])
		except Exception as e:
			print(text)
			print(e)
			raise

	workbook.close()

if __name__ == '__main__':
    main()
