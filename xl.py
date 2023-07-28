import xlsxwriter
workbook = xlsxwriter.Workbook("allaboutpyxl.xlsx")
worksheet = workbook.add_worksheet("firstsheet")
data = [
    {
        'name': "chiru",
        'class': 2023,
        'age':22
    },
    {
        'name':"ram",
        'class':2023,
        'age':23
    }
]

worksheet.write(0,0,"#")
worksheet.write(0,1,"Name")
worksheet.write(0,2,"Class")
worksheet.write(0,3,"age")


for index, entry in enumerate(data):
    worksheet.write(index+1,0,str(index))
    worksheet.write(index+1,1,entry["name"])
    worksheet.write(index+1,2,entry["class"])
    worksheet.write(index+1,3,entry["age"])
    
workbook.close()
'''
import xlsxwriter
workbook = xlsxwriter.Workbook("allaboutpyxl.xlsx")
worksheet = workbook.add_worksheet("firstsheet")
data = [{},{},{}]

def generate_xl(workbook_name: str, worksheet_name: str, header_list: list, data: list):
    #creating  workbook
    workbook = xlsxwriter.Workbook("workbook_name")

    #creatin worksheet
    worksheet = workbook.add_worksheet("worksheet_name")

    #adding headers
    for index, header in enumerate(header_list):
        worksheet.write(0, index, str(header).capitalize())

    #adding data
    for index1, entry in enumerate(data):
        for index2, header in enumerate(header_list):
            worksheet.write(index1+1,index2,entry[header])

    #close workbook
    workbook.close()
        
generate_excel("testxl.xlsx", "secondsheet", ["name',phone","email","address","country"].data)
'''