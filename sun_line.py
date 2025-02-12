import openpyxl
from scipy.stats import norm

# 向excel填写最左侧三列
def write_outline(outline):
    file_path = "属性表.xlsx"
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    header_row = 3
    for i, row_data in enumerate(outline):
        sheet.cell(row=i+header_row, column=1, value=row_data[0])
        sheet.cell(row=i+header_row, column=2, value=row_data[1])
        sheet.cell(row=i+header_row, column=3, value=row_data[2])
    wb.save(file_path)

# 获取转职加成
def get_fixed_attr(job):
    wb = openpyxl.load_workbook("转职加成.xlsx", data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if row[0] == job:
            fixed_attr = [
                [row[29], row[30], row[31], row[32], row[33], row[34], row[35]],
                [row[1],  row[2],  row[3],  row[4],  row[5],  row[6],  row[7]],
                [row[8],  row[9],  row[10], row[11], row[12], row[13], row[14]],
                [row[15], row[16], row[17], row[18], row[19], row[20], row[21]],
                [row[22], row[23], row[24], row[25], row[26], row[27], row[28]]
            ]
            return fixed_attr

# 获取模型倾向
def get_variable_attr(model):
    start = [0, 0, 0, 0, 0, 0, 0]
    growth_m = [0, 0, 0, 0, 0, 0, 0]
    growth_v = [0, 0, 0, 0, 0, 0, 0]
    wb = openpyxl.load_workbook("统计表.xlsx", data_only=True)
    ws = wb.active
    jobs = []
    cnt = 0
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if row[0] == model:
            if row[1] not in jobs:
                jobs.append(row[1])
            cnt += 1
            start[0] += row[2]
            start[1] += row[3]
            start[2] += row[4]
            start[3] += row[5]
            start[4] += row[6]
            start[5] += row[7]
            start[6] += row[8]
    for i in range(0, 7):
        start[i] = round((1.0 / cnt)*start[i], 1)
    for job in jobs:
        fixed_attr = get_fixed_attr(job)[1]
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if row[0] == model and row[1] == job:
                growth_m[0] += (row[18] - fixed_attr[0])
                growth_m[1] += (row[19] - fixed_attr[1])
                growth_m[2] += (row[20] - fixed_attr[2])
                growth_m[3] += (row[21] - fixed_attr[3])
                growth_m[4] += (row[22] - fixed_attr[4])
                growth_m[5] += (row[23] - fixed_attr[5])
                growth_m[6] += (row[24] - fixed_attr[6])
                growth_v[0] += (row[18] - fixed_attr[0]) * (row[18] - fixed_attr[0])
                growth_v[1] += (row[19] - fixed_attr[1]) * (row[19] - fixed_attr[1])
                growth_v[2] += (row[20] - fixed_attr[2]) * (row[20] - fixed_attr[2])
                growth_v[3] += (row[21] - fixed_attr[3]) * (row[21] - fixed_attr[3])
                growth_v[4] += (row[22] - fixed_attr[4]) * (row[22] - fixed_attr[4])
                growth_v[5] += (row[23] - fixed_attr[5]) * (row[23] - fixed_attr[5])
                growth_v[6] += (row[24] - fixed_attr[6]) * (row[24] - fixed_attr[6])
    for i in range(0, 7):
        growth_m[i] = (1.0 / cnt)*growth_m[i]
        growth_v[i] = (1.0 / cnt)*growth_v[i] - growth_m[i] * growth_m[i]
        growth_m[i] = (1.0 / 126)*growth_m[i]
        growth_v[i] = (1.0 / 126)*growth_v[i]
    return [start, growth_m, growth_v]

# 向excel填写模型倾向和转职加成
def write_data():
    file_path = "属性表.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True), start=3):
        print("正在统计: " + row[2] + " -> " + row[0])
        attr_index = 0 if row[1] == "力" else 1 if row[1] == "魔" else 2 if row[1] == "技" else 3 if row[1] == "速" else 4 if row[1] == "体" else 5
        variable_attr = get_variable_attr(row[2])  
        fixed_attr = get_fixed_attr(row[0]) 
        ws.cell(row=i, column=4, value=variable_attr[1][attr_index])
        ws.cell(row=i, column=5, value=variable_attr[2][attr_index])
        ws.cell(row=i, column=6, value=variable_attr[0][attr_index])
        ws.cell(row=i, column=7, value=fixed_attr[1][attr_index])
        ws.cell(row=i, column=8, value=fixed_attr[2][attr_index])
        ws.cell(row=i, column=9, value=fixed_attr[3][attr_index])
        ws.cell(row=i, column=10, value=fixed_attr[4][attr_index])
        ws.cell(row=i, column=11, value=fixed_attr[0][attr_index])
    wb.save(file_path)

# 向excel填写日10线
def write_sun_line():
    file_path = "属性表.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True), start=3):
        print("计算日10线: " + row[2] + " -> " + row[0])
        s = row[5]       # 初始属性
        m = row[3] * 293 # 总体期望
        v = row[4] * 293 # 总体方差
        f = row[10]      # 转职加成
        ws.cell(row=i, column=19, value=round(s + f + norm.ppf(1-0.1,      loc=m, scale=v**0.5)))
        ws.cell(row=i, column=23, value=round(s + f + norm.ppf(1-0.05,     loc=m, scale=v**0.5)))
        ws.cell(row=i, column=27, value=round(s + f + norm.ppf(1-0.01,     loc=m, scale=v**0.5)))
        ws.cell(row=i, column=31, value=round(s + f + norm.ppf(1-0.0001,   loc=m, scale=v**0.5)))
        ws.cell(row=i, column=35, value=round(s + f + norm.ppf(1-0.000001, loc=m, scale=v**0.5)))
    wb.save(file_path)

def count_model():
    models = []
    wb = openpyxl.load_workbook("统计表.xlsx", data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if row[0] not in models:
            models.append(row[0])
    for model in models:
        cnt = 0
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if row[0] == model:
                cnt += 1
        print(model, " ", cnt)