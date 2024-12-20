import glob
import os
import openpyxl
import json

#可调选项：
konghang = 1    #为生成的excel表格顶部空行，默认空一行,以匹配T++(default=1)

#以下为初始化全局变量-------------------------------------------------------------------
input_path = ""
output_path = ""
rpy_dirs = []    #所有文件目录




def setup():
    global rpy_dirs, files_n, input_path, output_path
    #用户输入两个需要的目录
    if os.path.exists("./rpy_files"):
        input_path = os.path.abspath("./rpy_files")
    else:
        input_path = input("请要转化的rpy目录:").replace('\\', '/')
    if os.path.exists("./rpy_files"):
        output_path = os.path.abspath("./tran_files")
    else:
        output_path = input("请输入导出的excel文件目录:").replace('\\', '/')
    if not os.path.exists(input_path):
        os.makedirs(input_path)
        print("目录" + input_path + "已创建。")
        if not os.path.exists(output_path):
            os.makedirs(output_path)
            print("目录" + input_path + "已创建。")
    else:
        print("目录检测完成，一切正常")
    #开始初始化文件目录变量
    for root, dirs, files in os.walk(input_path):
        for file in files:
            if file.endswith('.rpy'):
                # 获取文件的绝对路径
                full_path = os.path.abspath(os.path.join(root, file))
                relative_path = os.path.relpath(full_path, input_path)
                rpy_dirs.append([full_path.replace('\\', '/'), relative_path.replace('\\', '/')])


def old_get_lines(rpy_dir):
    with open(rpy_dir[0], 'r', encoding='utf-8') as file:
        lines = []
        for line in file:
            #line = line.rstrip('\n')
            line = line
            lines.append(line)
    return lines

def get_line(rpy_dir):
    file_data = []
    find_count = 0
    with open(rpy_dir[0], 'r', encoding='utf-8') as file:
        lines = file.readlines()
    row_count = 0
    tran_text = ""
    for line in lines:
        row_count = row_count + 1
        if line.count('"') >= 2:
            first_yinhao = line.find('"')
            last_yinhao = line.rfind('"')
            left = line[:first_yinhao + 1]
            t_data = line[first_yinhao + 1:last_yinhao]
            right = line[last_yinhao:]
            find_count = find_count + 1
            if find_count % 2 != 0:
                tran_text = t_data
            else:
                file_data.append([row_count, left, tran_text, right])
    return file_data

def old_cut_line_1to3(line):
    # 判断字符串中是否至少有两个引号
    if line.count('"') < 2:
        return False
    
    # 找到第一个引号和最后一个引号的位置
    first_yinhao = line.find('"')
    last_yinhao = line.rfind('"')
    
    # 提取左部分、引号部分和右部分
    left = line[:first_yinhao + 1]
    t_data = line[first_yinhao + 1:last_yinhao]
    right = line[last_yinhao:]
    return left, t_data, right

def old_process_line(line, row):
    global line_odd_even
    result = cut_line_1to3(line)
    if result is False:
        return False    # 排除非引号行
    else:
        left, t_data, right = result
        line_data = [row, left, t_data, right]
        line_odd_even = line_odd_even + 1
        if line_odd_even % 2 == 0:
            return False    #排除偶数行
        else:
            return line_data    #给生成json回传偶数行数据

def create_excel(file_data, rpy_dir):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in range(len(file_data)):
        sheet.cell(row=row + 1 + konghang, column=1, value=file_data[row][2])  # 空一行，把t_data写入第一列
    out_excel = os.path.join(output_path + "/EXCEL_FILE", rpy_dir[1] + ".xlsx")
    if not os.path.exists(os.path.dirname(out_excel)):
        os.makedirs(os.path.dirname(out_excel), exist_ok=True)
    workbook.save(out_excel)


def create_json(file_data, rpy_dir):
    suffix = ".r2e"    #自定义识别后缀
    # 创建一个列表来存储最终数据
    json_data = []
    headers = ["line", "left", "text", "right"]
    # 将每一行数据与表头结合起来形成字典
    for row in file_data:
        # 创建字典，将表头作为键，行数据作为值
        json_data.append(dict(zip(headers, row)))
    # 构造完整的文件路径
    json_dir = os.path.join(output_path + "/JSON_FILE", rpy_dir[1] + suffix)
    # 将生成的列表保存为 JSON 文件
    if not os.path.exists(os.path.dirname(json_dir)):
        os.makedirs(os.path.dirname(json_dir), exist_ok=True)
    with open(json_dir, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, ensure_ascii=False, indent=4)









#处理单个文件的方法
def process_file(rpy_dir):
    file_data = get_line(rpy_dir)
    #file_data写入表格和json(配套)
    create_excel(file_data, rpy_dir)
    create_json(file_data, rpy_dir)
    
#处理多个文件的总进程
def process_files():
    for rpy_dir in rpy_dirs:
        process_file(rpy_dir)

#rpy2excel的主脚本进程
def R2E():
    pass


if __name__ == '__main__':
    setup()
    print("正在从 " + input_path + " 运行导出脚本")
    process_files()
    print("已保存至目录: " + output_path)