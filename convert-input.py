# 讀取 input.txt 並解析
with open('input.txt', 'r', encoding='utf-8') as file:
    lines = file.readlines()

# 分割行為兩部分：第一行為 max_rest_per_day，後面的行為是排班表
max_rest_per_day = list(map(int, lines[0].strip().split('\t')))
schedules = [line.strip().split('\t') for line in lines[2:]]

# 定義 people_list 使用 A-Z 人
people_list = [f"'{chr(65 + i)}人'" for i in range(len(schedules))]

# 初始化 mandatory_off 列表來存儲每個人的休息日
mandatory_off = []

# 處理排班表，找到"休"的位置
for schedule in schedules:
    off_days = [index + 1 for index, value in enumerate(schedule) if value == '休']
    mandatory_off.append(off_days)

# 輸出結果
print("max_rest_per_day = [")
print(f"    {', '.join(map(str, max_rest_per_day))}")
print("]")
print("people_list = [")
print(f"    {', '.join(map(str, people_list))}")
print("]")
print("mandatory_off = [")
for off in mandatory_off:
    print(f"    {off},")
print("]")
