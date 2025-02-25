

data = [162 , 145 , 180 , 166 , 134 , 152 , 139 , 147 , 183 , 157 , 156 , 129 , 173 , 158 , 164 , 154 , 131 , 137 , 147 , 142 , 170 , 158 , 172 , 131 , 155 , 165 , 164 , 149 , 156 , 176 , 175 , 147 , 176 , 141 , 145 , 146 , 161 , 173 , 167 , 164 , 154 , 157 , 171 , 137 , 165 , 158 , 157 , 148 , 167 , 175 , 171 , 164 , 138 , 153 , 156 , 173 , 139 , 139 , 137 , 120 , 146 , 146 , 177 , 153 , 184 , 163 , 146 , 159 , 153 , 157 , 136 , 161 , 169 , 160 , 144 , 144 , 138 , 174 , 162 , 194 , 203 , 166 , 172 , 169 , 165 , 135 , 170 , 163 , 160 , 134 , 163 , 153 , 166 , 150 , 149 , 118 , 178 , 143 , 154 , 152 , 159 , 140 , 163 , 162 , 165 , 158 , 131 , 163 , 157 , 154 , 163 , 170 , 140 , 148 , 176 , 152 , 148 , 148 , 129 , 155 , 142 , 164 , 160 , 161 , 155 , 125 , 150 , 169 , 175 , 150 , 178 , 138 , 161 , 143 , 146 , 127 , 160 , 165 , 183 , 166 , 157 , 149 , 182 , 172 , 200 , 149 , 168 , 168 , 162 , 123 , 168 , 139 , 137 , 131 , 156 , 130 , 145 , 156 , 139]


import numpy as np

#data = [162 , 145 , 180 , 166 , 134, ...]

b = [0,0,0,0]
for d in data:
    b[0] += d ** 1
    b[1] += d ** 2
    b[2] += d ** 3
    b[3] += d ** 4
b[0] /= len(data)
b[1] /= len(data)
b[2] /= len(data)
b[3] /= len(data)

def solve_a(b):
    '''
    Todo:
    b1 = 126 * a1
    b2 = 126 * a2 + 126 * 125 * a1**2
    b3 = 126 * a3 + 3 * 126 * 125 * a1 * a2 + 126 * 125 * 124 * a1**3
    b4 = 126 * a4 + 4 * 126 * 125 * a1 * a3 + 3 * 126 * 125 * a2**2 + 6 * 126 * 125 * 124 * a1**2 * a2 + 126 * 125 * 124 * 123 * a1**4
    return [a1,a2,a3,a4]
    '''
    a1 = b[0] / 126
    a2 = (b[1] - 126 * 125 * a1**2) / 126
    a3 = (b[2] - 3 * 126 * 125 * a1 * a2 - 126 * 125 * 124 * a1**3) / 126
    a4 = (b[3] - 4 * 126 * 125 * a1 * a3 - 3 * 126 * 125 * a2**2 - 6 * 126 * 125 * 124 * a1**2 * a2 - 126 * 125 * 124 * 123 * a1**4) / 126
    return [a1, a2, a3, a4]

def solve_p(a):
    '''
    Todo:
    a[0] = p1 * 1 + p2 * 2 + p3 * 3 + p4 * 4
    a[1] = p1 * 1 + p2 * 4 + p3 * 9 + p4 * 16
    a[2] = p1 * 1 + p2 * 8 + p3 * 27+ p4 * 64
    a[3] = p1 * 1 + p2 * 16+ p3 * 81+ p4 * 256
    1    = p0 + p1 + p2 + p3 + p4
    return [p0,p1,p2,p3,p4]
    '''
    A = np.array([
        [1, 2, 3, 4],
        [1, 4, 9, 16],
        [1, 8, 27, 64],
        [1, 16, 81, 256]
    ])
    b = np.array(a)
    p = list(np.linalg.solve(A, b))
    p0 = 1 - p[0] - p[1] - p[2] - p[3] 
    return [p0, p[0], p[1], p[2], p[3]]

def solve_p2(a):
    from scipy.optimize import minimize
    # 目标函数：最小化残差平方和
    def objective(x):
        p1, p2, p3, p4 = x
        p0 = 1 - p1 - p2 - p3 - p4
        # 确保所有概率非负
        if p0 < 0 or p1 <0 or p2 <0 or p3 <0 or p4 <0:
            return float('inf')
        e1 = p1 + 2*p2 +3*p3 +4*p4
        e2 = p1 +4*p2 +9*p3 +16*p4
        e3 = p1 +8*p2 +27*p3 +64*p4
        e4 = p1 +16*p2 +81*p3 +256*p4
        return (e1 - a[0])**2 + (e2 - a[1])**2 + (e3 - a[2])**2 + (e4 - a[3])**2
    
    # 初始猜测
    x0 = [0.25, 0.25, 0.25, 0.25]
    # 约束：p1+p2+p3+p4 <=1，且每个p>=0
    bounds = [(0,1) for _ in range(4)]
    res = minimize(objective, x0, bounds=bounds, method='SLSQP')
    p1, p2, p3, p4 = res.x
    p0 = 1 - p1 - p2 - p3 - p4
    # 确保p0非负
    p0 = max(0, p0)
    # 重新标准化（防止浮点误差）
    total = p0 + p1 + p2 + p3 + p4
    p0 /= total
    p1 /= total
    p2 /= total
    p3 /= total
    p4 /= total
    return [p0, p1, p2, p3, p4]

a = solve_a(b)
p = solve_p2(a)

print(b)
print(a)
print(p)