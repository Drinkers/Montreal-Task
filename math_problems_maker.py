import random
import xlsxwriter

problem_level1 = []  # 储存第二简单水平的数学题
for i in range(0, 200):
    case1 = random.choice([2, 3])  # 算术题里有2或3个数字
    if case1 == 2:
        problem = '10'  # 将算术题初始值定为10，如果算术题的解不是0-9或者已经在题库里则持续生成
        while eval(problem) >= 10 or eval(problem) < 0 or (problem in problem_level1):
            first_number = random.randint(1, 99)
            second_number = random.randint(1, 99)
            problem = str(first_number) + random.choice(['+', '-']) + str(second_number)
            # print(problem)

        problem_level1.append(problem)

    else:
        problem = '10'
        while eval(problem) >= 10 or eval(problem) < 0 or (problem in problem_level1):
            first_number = random.randint(1, 99)
            second_number = random.randint(1, 99)
            third_number = random.randint(1, 99)
            problem = str(first_number) + random.choice(['+', '-']) + str(second_number) + random.choice(['+', '-']) + str(third_number)
            # print(problem)

        problem_level1.append(problem)

problem_level2 = []  # 储存第三简单
for i in range(0, 200):
    case2 = random.choice([3, 4])  # 算术题里有3或4个数字
    if case2 == 3:
        problem = '10'  # 将算术题初始值定为10，如果算术题的解不是0-9或者已经在题库里，则持续生成
        while eval(problem) >= 10 or eval(problem) < 0 or (problem in problem_level1):
            first_number = random.randint(1, 99)
            second_number = random.randint(1, 99)
            third_number = random.randint(1, 99)
            problem = str(first_number) + random.choice(['+', '-', '*']) + str(second_number) + \
                random.choice(['+', '-', '*']) + str(third_number)
            # print(problem)

        problem_level2.append(problem)

    else:
        problem = '10'
        while eval(problem) >= 10 or eval(problem) < 0 or (problem in problem_level1) or len(problem) >= 9:
            first_number = random.randint(1, 99)
            second_number = random.randint(1, 99)
            third_number = random.randint(1, 99)
            fourth_number = random.randint(1, 99)
            problem = str(first_number) + random.choice(['+', '-', '*']) + str(second_number) + random.choice\
                (['+', '-', '*']) + str(third_number) + random.choice(['+', '-', '*']) +str(fourth_number)
            # print(problem)

        problem_level2.append(problem)

math_problems = xlsxwriter.Workbook('math_problems.xlsx')
sheet = math_problems.add_worksheet('sheet1')

for i in range(0, len(problem_level1)):
    sheet.write(i, 0, problem_level1[i])
    sheet.write(i, 1, problem_level2[i])

math_problems.close()