from ortools.sat.python import cp_model

# 模型初始化
model = cp_model.CpModel()

# 定義參數
num_people = 5  # 假設有5個人
num_days = 25  # 計劃的天數

# 每天最多可休息人數列表，例如: [2, 2, 3, 1, 2, ...]，這個列表可以自由設定
max_rest_per_day = [
    2,3,1,3,1,3,1,1,1,2,1,1,3,2,2,3,3,3,2,3,1,2,2,3,1
]

# 月份
month = 9

# 專員姓名 與 mandatory_off 對應
people_list = [
    '小明',
    '小花',
    '小白',
    '小城',
    '小胖'
]

# 固定休假日 與 people_list 對應
mandatory_off = [
    [1, 6, 14],  # 第一個人的固定休假日
    [4, 16, 20], # 第二個人的固定休假日
    [2, 10, 18], # 第三個人的固定休假日
    [5, 15, 21], # 第四個人的固定休假日
    [3, 12, 19]  # 第五個人的固定休假日
]

# 定義變量：work[p][d] 表示第 p 個人在第 d 天是否上班 (1 表示上班，0 表示休息)
work = []
for p in range(num_people):
    person_schedule = []
    for d in range(num_days):
        person_schedule.append(model.NewBoolVar(f'work_{p}_{d}'))
    work.append(person_schedule)

# 約束條件 1：每個人休息 10 天（包括固定休假）
for p in range(num_people):
    model.Add(sum(1 - work[p][d] for d in range(num_days)) == 10)

# 約束條件 2：連續上班不能超過 4 天
for p in range(num_people):
    for d in range(num_days - 4):
        model.Add(sum(work[p][d + i] for i in range(5)) <= 4)

# 約束條件 3：不可單獨上班一天，第一天和最後一天例外
for p in range(num_people):
    for d in range(1, num_days - 1):
        model.AddBoolOr([
            work[p][d - 1],
            work[p][d + 1]
        ]).OnlyEnforceIf(work[p][d])

# 約束條件 4：固定休假日
for p in range(num_people):
    for day in mandatory_off[p]:
        model.Add(work[p][day - 1] == 0)

# 新增的約束條件 5：每天最多可休息的人數，使用 max_rest_per_day 列表
for d in range(num_days):
    model.Add(sum(1 - work[p][d] for p in range(num_people)) <= max_rest_per_day[d])

# 創建求解器
solver = cp_model.CpSolver()

# 設定求解器參數
solver.parameters.max_time_in_seconds = 60

# 求解
status = solver.Solve(model)

# 輸出結果
if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
    # 第一行顯示日期
    print('Day:    ', end='')
    for d in range(1, num_days + 1):
        print(f'{d:>5}', end='')
    print()

    for p in range(num_people):
        print(f'Person {p + 1}:', end=' ')
        for d in range(num_days):
            if d + 1 in mandatory_off[p]:
                print(f'{ "指休":>5}', end=' ')
            elif solver.Value(work[p][d]) == 0:
                print(f'{ "休假":>5}', end=' ')
            else:
                print(f'{ "上班":>5}', end=' ')
        print()
else:
    print('No solution found within the time limit.')
