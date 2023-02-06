import re
import xlsxwriter
from os import listdir
from os.path import isfile, join

def parse_file(file_path):
    with open(file_path, "r") as file_object:
        text = file_object.read()
        namematch = re.search(r"LocDisplayName.*\"(.*)\"", text)
        caloriematch = re.search(r"Calories.*(?:=>)? (-?\d{1,4})\;", text)
        carbsmatch = re.search(r"Carbs = (\d{1,4})", text)
        fatmatch = re.search(r"Fat = (\d{1,4})", text)
        protmatch = re.search(r"Protein = (\d{1,4})", text)
        vitaminsmatch = re.search(r"Vitamins = (\d{1,4})", text)
        return {
            "name": namematch[1] if namematch else None,
            "calories": int(caloriematch[1]) if caloriematch else None,
            "carbs": int(carbsmatch[1]) if carbsmatch else None,
            "fat": int(fatmatch[1]) if fatmatch else None,
            "protein": int(protmatch[1]) if protmatch else None,
            "vitamins": int(vitaminsmatch[1]) if vitaminsmatch else None,
        }

def main():
    food_dir = r"./Mods/__core__/AutoGen/Food"
    seed_dir = r"./Mods/__core__/AutoGen/Seed"
    food_files = [f for f in listdir(food_dir) if isfile(join(food_dir, f))]
    seed_files = [f for f in listdir(seed_dir) if isfile(join(seed_dir, f))]

    food_and_seed_data = [parse_file(join(food_dir, file)) for file in food_files] + [parse_file(join(seed_dir, file)) for file in seed_files]

    workbook = xlsxwriter.Workbook("fooddata.xlsx")
    worksheet = workbook.add_worksheet()
    num_format = workbook.add_format({"num_format": "###.##0"})

    worksheet.add_table('A1:I{}'.format(len(food_and_seed_data)),{'columns': [  {'header': 'Name'},{'header': 'Calories'},{'header': 'Carbs'},{'header': 'Fat'},{'header': 'Protein'},{'header': 'Vitamins'},{'header': 'Nutrition'},{'header': 'Balance Mult.'},{'header': 'Skill points/day'}]})

    for row, data in enumerate(food_and_seed_data, start=1):
        col = 0
        worksheet.write(row, col, data.get("name"))
        worksheet.write(row, col + 1, data.get("calories"))
        worksheet.write(row, col + 2, data.get("carbs"))
        worksheet.write(row, col + 3, data.get("fat"))
        worksheet.write(row, col + 4, data.get("protein"))
        worksheet.write(row, col + 5, data.get("vitamins"))
        worksheet.write_formula(row, col + 6, f"=SUM(C{row + 1}:F{row + 1})", num_format)
        worksheet.write_formula(row, col + 7, f"=G{row + 1}/(MAX(C{row + 1}:F{row + 1})*4)*2",num_format)
        worksheet.write_formula(row, col + 8, f"=((G{row + 1}*B{row + 1})/B{row + 1})*H{row + 1}+12",num_format)

    workbook.close()

if __name__ == '__main__':
    main()
