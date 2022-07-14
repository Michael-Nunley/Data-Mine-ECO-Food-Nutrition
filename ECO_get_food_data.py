import re
import xlsxwriter

from os import listdir
from os.path import isfile, join

def main():
	foodfiles = [f for f in listdir(r".\Mods\__core__\AutoGen\Food") if isfile(join(r".\Mods\__core__\AutoGen\Food", f))]
	seedfiles = [f for f in listdir(r".\Mods\__core__\AutoGen\Seed") if isfile(join(r".\Mods\__core__\AutoGen\Seed", f))]
	foodandseedfilecontent = []

	for file in foodfiles:
		file_object = open(r".\Mods\__core__\AutoGen\Food\{f}".format(f=file),"r")
		foodandseedfilecontent.append(file_object.read())
	for file in seedfiles:
		file_object = open(r".\Mods\__core__\AutoGen\Seed\{f}".format(f=file),"r")
		foodandseedfilecontent.append(file_object.read())

	workbook = xlsxwriter.Workbook('fooddata.xlsx')
	worksheet = workbook.add_worksheet()

	numformat = workbook.add_format({'num_format': '###.##0'})

	for idx, text in enumerate(foodandseedfilecontent):
		try:
			namemach = re.search(r"LocDisplayName.*\"(.*)\"",text)
			worksheet.write(idx+1,0, namemach[1])

			caloriematch = re.search(r"Calories.*(?:=>)? (-?\d{1,4})\;",text)
			worksheet.write(idx+1,1, int(caloriematch[1]))

			carbsmatch = re.search(r"Carbs = (\d{1,4})",text)
			worksheet.write(idx+1,2, int(carbsmatch[1]) if not int(carbsmatch[1]) == None else 0)

			fatmatch = re.search(r"Fat = (\d{1,4})",text)
			worksheet.write(idx+1,3, int(fatmatch[1]) if not int(fatmatch[1]) == None else 0)

			protmatch = re.search(r"Protein = (\d{1,4})",text)
			worksheet.write(idx+1,4, int(protmatch[1]) if not int(protmatch[1]) == None else 0)

			vitaminsmatch = re.search(r"Vitamins = (\d{1,4})",text)
			worksheet.write(idx+1,5, int(vitaminsmatch[1]) if not int(vitaminsmatch[1]) == None else 0)

			worksheet.write_formula(idx+1,6, '=SUM(C{}:F{})'.format(idx+2,idx+2),numformat)

			worksheet.write_formula(idx+1,7, '=G{}/(MAX(C{}:F{})*4)*2'.format(idx+2,idx+2,idx+2),numformat)

			worksheet.write_formula(idx+1,8, '=((G{}*B{})/B{})*H{}+12'.format(idx+2,idx+2,idx+2,idx+2),numformat)
		except Exception as e:
			print(text)
			print(e)
			raise

	worksheet.add_table('A1:I{}'.format(len(foodandseedfilecontent)),{'columns': [  {'header': 'Name'},
                                          						{'header': 'Calories'},
											{'header': 'Carbs'},
                                          						{'header': 'Fat'},
                                          						{'header': 'Protein'},
                                          						{'header': 'Vitamins'},
											{'header': 'Nutrition'},
											{'header': 'Balance Mult.'},
											{'header': 'Skill points/day'},]})
	workbook.close()

if __name__ == '__main__':
    main()
