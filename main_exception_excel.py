import openpyxl
from openpyxl.styles import PatternFill
from ortools.sat.python import cp_model

# 模型初始化
model = cp_model.CpModel()

# 定義參數
month = 9
people_list = ['小明', '小花', '小白', '小城', '小胖']
num_people = len(people_list)
num_days = 25  # 計劃的天數

# 每天最多可休息人數列表
max_rest_per_day = [
    2,3,1,3,1,3,1,1,1,2,1,1,3,2,2,3,3,3,2,3,1,2,2,3,1
]

# 固定休假日
mandatory_off = [
    [1, 6, 14],  # 小明的固定休假日
    [4, 16, 20], # 小花的固定休假日
    [2, 10, 18], # 小白的固定休假日
    [5, 15, 21], # 小城的固定休假日
    [3, 12, 19]  # 小胖的固定休假日
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

# 約束條件 2：連續上班不能超過 4 天 (若無解可以寬限到 5 天一次)
# 為每個人定義一個 "破例" 的布林變量，表示是否允許一次連續 5 天上班
exception_used = []
for p in range(num_people):
    exception_used.append(model.NewBoolVar(f'exception_used_{p}'))

for p in range(num_people):  # 對每個專員進行處理
    for d in range(num_days - 4):  # 遍歷可以開始的工作天數，檢查 5 天內的狀況
        # 生成約束條件，表示在這 5 天內不能有超過 4 天上班
        five_day_sum = sum(work[p][d + i] for i in range(5))

        # 使用布林變量來控制是否允許破例
        model.Add(five_day_sum <= 4).OnlyEnforceIf(exception_used[p].Not())  # 如果還沒使用破例，遵守最多 4 天的限制
        model.Add(five_day_sum == 5).OnlyEnforceIf(exception_used[p])  # 如果使用了破例，允許這 5 天內剛好 5 天上班

# 每個人只能使用一次破例，即最多允許一次連續 5 天上班
model.Add(sum(exception_used[p] for p in range(num_people)) <= 1)

# 確保沒有跨區間的 6 天連續上班情況
for p in range(num_people):
    for d in range(num_days - 5):  # 這次檢查的是跨越 6 天的連續上班情況
        six_day_sum = sum(work[p][d + i] for i in range(6))
        model.Add(six_day_sum <= 5)  # 確保 6 天內最多有 5 天上班


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

# 構建 Excel 數據
if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
    # 創建一個新的 Workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "排班表"

    # 定義填充樣式
    fill_mandatory_off = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黃色底色（指休）
    fill_vacation = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 綠色底色（休假）

    # 第一行顯示日期
    ws.append(['日期'] + [f'{month}/{d + 1}' for d in range(num_days)])

    # 添加當日可休人數
    ws.append(['當日可休人力'] + max_rest_per_day)

    # 輸出每個人的排班狀態
    for p in range(num_people):
        row = [people_list[p]]
        for d in range(num_days):
            if d + 1 in mandatory_off[p]:
                row.append('指休')
            elif solver.Value(work[p][d]) == 0:
                row.append('休假')
            else:
                row.append('上班')
        ws.append(row)

    # 設置單元格的背景顏色
    for p in range(num_people):
        for d in range(num_days):
            cell = ws.cell(row=p + 3, column=d + 2)  # +3 是因為要跳過標題行和當日可休人數行
            if d + 1 in mandatory_off[p]:
                cell.fill = fill_mandatory_off  # 指休
            elif solver.Value(work[p][d]) == 0:
                cell.fill = fill_vacation  # 休假

    # 保存 Excel 文件
    excel_file = './schedule_output.xlsx'
    wb.save(excel_file)

    # 文件路徑返回
    excel_file
else:
    "No solution found within the time limit."
