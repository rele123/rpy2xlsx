import glob
import os
import openpyxl
import json


#初始化全局变量
suffix = ".r2e"
input_path = ""
output_path = ""
fail_line = []    #错误行结尾print，手动修改【文件名，第几行】
inject_font = ""

def setup():
    global files_dir, files_n, input_path, output_path, inject_font
    #用户输入两个需要的目录
    if os.path.exists("./tran_files"):
        input_path = os.path.abspath("./tran_files").replace('\\', '/')
    else:
        input_path = input("请输入翻译后的excel文件目录::").replace('\\', '/')
    if os.path.exists("./rpy_files"):
        output_path = os.path.abspath("./rpy_files").replace('\\', '/')
    else:
        output_path = input("请输入要写入的rpy目录").replace('\\', '/')
    inject_font = input("请输入要注入的字体路径(以/game/为根目录)").replace('\\', '/')
    if inject_font.startswith('/'):
        inject_font = inject_font.replace('/', '', 1)
    if not os.path.exists(input_path):
        os.makedirs(input_path)
        print("目录" + input_path + "已创建。")
        if not os.path.exists(output_path):
            os.makedirs(output_path)
            print("目录" + input_path + "已创建。")
    else:
        print("目录检测完成，一切正常")
    

def load_excel(excel_dir):
    workbook = openpyxl.load_workbook(excel_dir)  # 替换为你的Excel文件路径
    sheet = workbook.active  # 获取活动工作表
    # 初始化存储有效数据的列表
    excel_data = []
    
    # 遍历 Excel 表的除第一行的没一行的1，2列
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=2):
        row_data = []
        if row[0].value is None:
            row[0].value = ""
        if row[1].value is None:
            row_data = [row[0].value, row[0].value]
        else:
            if isinstance(row[1].value, int):
                row_data = [row[0].value, str(row[1].value)]
            elif isinstance(row[1].value, str):
                row_data = [row[0].value, row[1].value]
            else:
                row_data = [row[0].value, row[0].value]
        excel_data.append(row_data)  # 将处理后的行数据加入总数据中
    
    rpy_name = os.path.basename(excel_dir)
    rpy_name = rpy_name.replace(".xlsx", "")
    
    return excel_data

def load_json(json_dir):
    if not os.path.isfile(json_dir):
        raise FileNotFoundError(f"File {json_dir} not found.")
    
    with open(json_dir, 'r', encoding='utf-8') as json_file:
        json_file_data = json.load(json_file)
    
    json_data = [
        [item["line"], item["left"], item["text"], item["right"]] 
        for item in json_file_data
    ]
    
    return json_data


def read_file(rpy_dir):
    tran_data = []
    excel_dir = os.path.join(input_path + "/EXCEL_FILE", rpy_dir[1] + ".xlsx")
    json_dir = os.path.join(input_path + "/JSON_FILE", rpy_dir[1] + suffix)
    excel_data = load_excel(excel_dir)
    json_data = load_json(json_dir)
    for i in range(len(json_data)):
        row_e = excel_data[i]
        row_j = json_data[i]
        untranslate = row_j[1] + row_j[2] + row_j[3]
        if inject_font == "":
            translate = row_j[1] + row_e[1] + row_j[3]
        else:
            translate = row_j[1] + "{font=" + inject_font + "}" + row_e[1] + "{/font}" + row_j[3]
#        translate = row_j[1] + row_e[1] + row_j[3]
        if row_e[0] == row_j[2]:
            tran_data.append([row_j[0], untranslate, translate, 1])
        else:
            print("\033[31m错误：" + excel_dir + "与" + json_dir + "在rpy文件第" + str(row_j[0]) + "行的原文不匹配！！！\033[0m")
            if isinstance(row_j[0], int):
                tran_data.append([row_j[0], "", "", 0])
    
    return tran_data

def write_rpy(rpy_dir, tran_data):
    with open(rpy_dir[0], 'r', encoding='utf-8') as rpy_file_R:
        rpy_file_R_line = rpy_file_R.readlines()
    for row in tran_data:
        if row[3] == 1:
            rpy_file_R_line[row[0] - 1] = row[2]
        elif row[3] == 0:    #读取失败，不予写入
            fail_line.append(rpy_name, row[0])
        else:
            pass
    with open(rpy_dir[0], 'w', encoding='utf-8') as rpy_file_R:
        rpy_file_R.writelines(rpy_file_R_line)


def change_files():
    rpy_dirs = []
    for root, dirs, files in os.walk(output_path):
        for file in files:
            if file.endswith('.rpy'):
                # 获取文件的绝对路径
                full_path = os.path.abspath(os.path.join(root, file))
                relative_path = os.path.relpath(full_path, output_path)
                rpy_dirs.append([full_path.replace('\\', '/'), relative_path.replace('\\', '/')])
    for rpy_dir in rpy_dirs:
        tran_data = read_file(rpy_dir)
        write_rpy(rpy_dir, tran_data)


def E2R():
    pass


if __name__ == '__main__':
    setup()
    print("正在进行任务.....请稍后")
    change_files()

