import requests
import xlrd
import json

# workbook = xlrd.open_workbook("/Users/hupc/Desktop/case.xlsx")  # 文件路径
# # # 通过sheet名获得sheet对象
# worksheet = workbook.sheet_by_name("case_list")
# # param = case(worksheet)
# for i in range(worksheet.nrows):
#     if i == 0:  # 跳过第一行
#         continue
#     nrows = worksheet.row_values(i)  # 循环打印每一行
#     print(nrows)
#     case_id = nrows[0]
#     print("case_id ：", case_id)
#     Method = nrows[1]
#     url = nrows[2]
#     headers = json.loads(nrows[3])
#     print(headers)
#     payload = json.loads(nrows[4])
#     print(payload)
#     if Method == "GET":
#         response = requests.get(url = url, headers=headers, data = payload)
#         print("返回：",response.text)
#     else:
#         response = requests.post(url = url, headers=headers, data = payload)
#         print("返回：", response.text)



def xlrd_excel():
    data = xlrd.open_workbook('/Users/hupc/Desktop/case.xlsx')  # 打开xls文件
    table = data.sheet_by_name("case_list") #通过名称获取
    nrows=table.nrows  # 获取表的行数
    li=[]
    for i in range(nrows):     # 循环逐行打印
        if i!=0:        #跳过第一行
            li.append(str(table.row_values(i)))#将取回的值放入列表中
    return li
list = xlrd_excel()
print(type(list))
print(list)

def is_Method(method,url,headers,payload):
    if method == "GET":
        response = requests.request("GET", url, headers = headers, params = payload)
        print("返回：",response.text)
    else:
        response = requests.request("POST", url, headers = headers, params = payload)
        print("返回：", response.text)

for dict in xlrd_excel():
    print(dict)
    method = dict[1]
    # print(method)
    url = dict[2]
    print(url)
    headers = dict[3]
    print(headers)
    payload = dict[4]
    print(payload)
    # if method == "GET":
    #     payload = json.loads(dict[4])
    #     response = requests.request(method, url, headers=headers, params=payload)
    #     print("返回：", json.loads(response.text))
    # else:
    #     response = requests.request(method, url, headers=headers, data=payload)
    #     print("返回：", response.text)



