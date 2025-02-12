
import numpy as np
import openpyxl
from scipy.optimize import root

cost = [59657, 21556, 340000, 1306000]

def gen_p(mean, variance):
    """
    计算在支持集{0,1,2,3,4}上，满足总概率为1、期望为mean、方差为variance的最大熵分布。
    最大熵分布具有形式：
        p(i) = exp(B*i + C*i^2) / Z,  i=0,1,2,3,4,
    其中 Z = sum(exp(B*j + C*j^2), j=0..4)。
    参数：
        mean: 目标均值（浮点数）
        variance: 目标方差（浮点数）
    返回：
        长度为5的 numpy 数组，对应+0到+4的概率。
    注意：此函数要求输入的 mean 和 variance 在该离散支持下是可行的。
    """
    xs = np.arange(5)
    # 定义方程组，待求解参数为 [B, C]
    # 设 p(i) = exp(B*i + C*i^2) / Z，其中 Z = sum(exp(B*j + C*j^2))
    # 要求：sum(i*p(i)) = mean,  sum(i^2*p(i)) = mean^2 + variance
    def equations(params):
        B, C = params
        weights = np.exp(B * xs + C * xs**2)
        Z = np.sum(weights)
        p = weights / Z
        m_computed = np.sum(xs * p)
        m2_computed = np.sum(xs**2 * p)
        return [m_computed - mean, m2_computed - (mean**2 + variance)]
    # 初始猜测：如果mean=2, variance=2, 则均匀分布或指数族中的B,C都可能接近0。
    initial_guess = [0.0, 0.0]
    sol = root(equations, initial_guess)
    # if not sol.success:
    #     raise ValueError("无法找到满足条件的分布: " + sol.message)
    B, C = sol.x
    weights = np.exp(B * xs + C * xs**2)
    Z = np.sum(weights)
    p = weights / Z
    return p

# 返回给定4星5星月百日10线、模型倾向，计算资源的期望消耗量
# 这里的线把转职加成和初始属性先抠了，最后一起加回来
def sum_cost(四星线, 五星线, 月百线, 日十线, 期望, 方差):
    resource_sum = 0

    # 100的位置是0属性对应的位置，这里错开的目的是统一迭代的表达式，加速计算
    dist = np.zeros(905)
    dist[100] = 1
    p = np.array(gen_p(期望, 方差)) 
    四星线 = 四星线 + 100
    五星线 = 五星线 + 100
    月百线 = 月百线 + 100
    日十线 = 日十线 + 100

    # 计算到了4星50的分布情况
    for _ in range(126):
        dist[100:905] = (
            p[0] * dist[100:905] +
            p[1] * dist[99:904] +
            p[2] * dist[98:903] +
            p[3] * dist[97:902] +
            p[4] * dist[96:901]
        )
    dist[900] += np.sum(dist[901:905])  #800属性已经镇堡的不能再镇堡了，后面加几已经没所谓了，不影响概率分布的分析。
    dist[901:905] = 0  

    #至此，100%的英雄被升级到四星，资源消耗量为100% * cost[0]
    resource_sum += cost[0]

    #分布完成之后截断达不到四星线的部分
    rest_part = np.sum(dist[四星线:901])
    dist[100:四星线] = 0

    # 计算到了5星60的分布情况
    for _ in range(59):
        dist[100:905] = (
            p[0] * dist[100:905] +
            p[1] * dist[99:904] +
            p[2] * dist[98:903] +
            p[3] * dist[97:902] +
            p[4] * dist[96:901]
        )
    dist[900] += np.sum(dist[901:905])
    dist[901:905] = 0  

    #至此，rest_part的英雄被升级到五星，资源消耗量为rest_part * cost[1]
    resource_sum += rest_part * cost[1]

    #分布完成之后截断达不到五星线的部分
    rest_part = np.sum(dist[五星线:901])
    dist[四星线:五星线] = 0

    # 计算到了月百的分布情况
    for _ in range(99):
        dist[100:905] = (
            p[0] * dist[100:905] +
            p[1] * dist[99:904] +
            p[2] * dist[98:903] +
            p[3] * dist[97:902] +
            p[4] * dist[96:901]
        )
    dist[900] += np.sum(dist[901:905])
    dist[901:905] = 0  

    #至此，rest_part的英雄被升级到月百，资源消耗量为rest_part * cost[2]
    resource_sum += rest_part * cost[2]

    #分布完成之后截断达不到月百线的部分
    rest_part = np.sum(dist[月百线:901])
    dist[五星线:月百线] = 0

    # 计算到了日十的分布情况
    for _ in range(9):
        dist[100:905] = (
            p[0] * dist[100:905] +
            p[1] * dist[99:904] +
            p[2] * dist[98:903] +
            p[3] * dist[97:902] +
            p[4] * dist[96:901]
        )
    dist[900] += np.sum(dist[901:905])
    dist[901:905] = 0  

    #至此，rest_part的英雄被升级到日十，资源消耗量为rest_part * cost[3]
    resource_sum += rest_part * cost[3]

    #分布完成之后截断达不到日十线的部分，也就是说，计算超过了日10线的占比 -> 出货率
    rest_part = np.sum(dist[日十线:901])
    
    #升出一个过线冒险者的资源消耗期望，等于刚才这一轮分布的资源消耗，除以出货率。
    resource_sum /= rest_part
    return round(resource_sum)

def optimize_lines(t, m, v):
    """
    在 0 <= l1 <= l2 <= l3 <= t 范围内，寻找使 sum_cost(l1, l2, l3, t, m, v) 最小的整数组合。
    这里先根据比例（126/293, 185/293, 284/293）给出初始估计，然后采用多步长局部搜索优化。
    """
    # 初始估计
    l1 = round(126 * t / 293)
    l2 = round(185 * t / 293)
    l3 = round(284 * t / 293)
    # 保证满足约束
    l1 = max(0, min(l1, t))
    l2 = max(l1, min(l2, t))
    l3 = max(l2, min(l3, t))
    
    best_cost = sum_cost(l1, l2, l3, t, m, v)
    step = max(1, t // 10)  # 初始步长，可根据 t 调整
    
    while step > 0:
        improved = True
        while improved:
            improved = False
            # 尝试调整 l1
            for delta in [-step, step]:
                new_l1 = l1 + delta
                if 0 <= new_l1 <= l2:
                    c = sum_cost(new_l1, l2, l3, t, m, v)
                    if c < best_cost:
                        best_cost = c
                        l1 = new_l1
                        improved = True
                        break
            if improved:
                continue
            # 尝试调整 l2
            for delta in [-step, step]:
                new_l2 = l2 + delta
                if l1 <= new_l2 <= l3:
                    c = sum_cost(l1, new_l2, l3, t, m, v)
                    if c < best_cost:
                        best_cost = c
                        l2 = new_l2
                        improved = True
                        break
            if improved:
                continue
            # 尝试调整 l3
            for delta in [-step, step]:
                new_l3 = l3 + delta
                if l2 <= new_l3 <= t:
                    c = sum_cost(l1, l2, new_l3, t, m, v)
                    if c < best_cost:
                        best_cost = c
                        l3 = new_l3
                        improved = True
                        break
        step //= 2  # 缩小步长
    return l1, l2, l3

# 向excel填写月百五星四星线
def write_moon_line():
    file_path = "属性表.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True), start=3):
        print("计算月百五星四星线: " + row[2] + " -> " + row[0])
        s = row[5]       # 初始属性
        m = row[3]       # 每级期望
        v = row[4]       # 每级方差
        f = row[10]      # 转职加成
        for j in range(18, 38, 4):
            print("    日十" + str(row[j]), end="")
            # 日十减去初始和转职部分
            t = row[j] - round(s) - f

            # 计算最合适的line取值
            # ========================================================
            line1, line2, line3 = optimize_lines(t, m, v)
            cost = sum_cost(line1, line2, line3, t, m, v)
            # ========================================================

            line1 += round(s) + row[6]
            line2 += round(s) + row[6] + row[7]
            line3 += round(s) + row[6] + row[7] + int(row[8])
            print("    月百" + str(line3), end="")
            print("    五星" + str(line2), end="")
            print("    四星" + str(line1), end="")
            print("    资源消耗期望" + str(round(cost/10000)) + "w")
            ws.cell(row=i, column=j, value=line3)
            ws.cell(row=i, column=j-1, value=line2)
            ws.cell(row=i, column=j-2, value=line1)

    wb.save(file_path)