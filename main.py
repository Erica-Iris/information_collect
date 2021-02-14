import json
import chardet
# test_dict = {'bigberg': [7600, {1: [['iPhone', 6300], ['Bike', 800], ['shirt', 300]]}]}
# # print(test_dict)
# # print(type(test_dict))
# json_str=json.dumps(test_dict)
# # print(json_str)
# # print(type(json_str))

# new_dict=json.loads(json_str)


# with open("E:\Work_Place\information_collect/record.json","rb") as f:
#     json.dump(new_dict,f)
#     print("successed onload files")

import xlrd

xlsfile = r"E:\\QQ Temp\\2427940916\\FileRecv\\collect\\a.xls"

book = xlrd.open_workbook(xlsfile)

sheet0 = book.sheet_by_index(0)
# sheet_name=book.sheet_names()[0]

json_text = {}

# for sheet_name in book.sheet_names():
#     # print(book.sheet_by_name(sheet_name))
#     sheet1=book.sheet_by_name(sheet_name)
#     nrows=sheet1.nrows
#     for i in range(nrows):
#         print(i,sheet1.row_values(i)[0])
#     ncols=sheet1.ncols
#     # print("4.",ncols)

for i in range(1, sheet0.nrows):
    # ,'学号':sheet0.(i)2],'院系':sheet0.(i)3],'年级':sheet0.(i)4],'专业':sheet0.(i)5],'班级':sheet0.(i)6],'层次':sheet0.(i)7]
    # print(sheet0.row_values(i)[0],sheet0.row_values(i)[1])
    appended = {
        '性别': sheet0.row_values(i)[1],
        '学号': sheet0.row_values(i)[2],
        '院系': sheet0.row_values(i)[3],
        '年级': int(sheet0.row_values(i)[4]),
        '专业': sheet0.row_values(i)[5],
        '班级': sheet0.row_values(i)[6],
        '层次': sheet0.row_values(i)[7]
        }
    json_text[sheet0.row_values(i)[0]]=appended

print(type(json_text))

json_data = json.dumps(json_text,ensure_ascii=False)

json_file=r"‪E:\\Work_Place\\information_collect\\record.json"
# print(json_data)
# print(json_data.encode('utf-8'))
# print(type(json_data))
f=open('record.json','w',encoding='utf-8')
f.write(json_data)
f.close()

