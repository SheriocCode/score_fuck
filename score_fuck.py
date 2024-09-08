import os
import numpy as np
import pandas as pd
import pdfplumber


# 存放学生成绩单pdf的文件夹路径
PATH = r'C:\Users\Administrator\Desktop\score'
# 学期信息
semester_list = ['2023-2024-1学期', '2023-2024-2学期'] # [] 为全部学期
# 课程性质
course_list = ['选修', '必修', '实践课'] # ['必修', '实践课'] 为绩点课程


PDF2EXCEL_PATH = './pdf2excel'
RESULT_PATH = './result'

# 单科绩点计算保留小数位数
score_digit = 4
# 导出表格绩点保留小数位数
excel_digit = 4

def gpa(score):
    """
    绩点计算公式（score = (score - 60) / 10 + 1）
    """
    score = (score - 60) / 10 + 1
    return round(score, score_digit)


def filter_course(courses_info, semester_ls, course_ls):
    """
    按照学期信息、课程性质过滤课程
    """
    filtered_courses_info = [
        course for course in courses_info
        if course['学期'] in semester_ls
        and course['性质'] in course_ls
        if course['课程'] not in [np.nan, '以下空白']
    ]

    return filtered_courses_info


def fuck_score(path, stu_name, semester_list, course_list):
    # 读取pdf并保存提取结果
    pdf = pdfplumber.open(path)
    table = None
    for page in pdf.pages:
        table = page.extract_table()

    print(f'==读取【{stu_name}】成绩单，保存到 pdf2excel 文件夹==')
    if not os.path.exists(PDF2EXCEL_PATH):
        os.makedirs(PDF2EXCEL_PATH)
    df = pd.DataFrame(table)
    out_put_xlsx_path = os.path.join(PDF2EXCEL_PATH ,f'{stu_name}.xlsx')
    df.to_excel(out_put_xlsx_path, index=False)

    # 读取保存的excel文件
    df2 = pd.read_excel(out_put_xlsx_path, sheet_name='Sheet1', skiprows=2, nrows=32, header=None)
    # 去除pdf的空列
    columns_to_drop_indices = [1, 5, 6, 8, 9, 13, 15, 16]
    df2 = df2.drop(columns=df.columns[columns_to_drop_indices])

    new_columns = ['课程', '性质', '学分', '分数']
    new_df = pd.DataFrame(columns=new_columns)

    # 选取每4列的数据并添加到新的DataFrame，实现将整个多列表变成纵向四列
    for i in range(0, df2.shape[1], 4):
        group = df2.iloc[:, i:i+4]
        group.columns = new_columns
        new_df = pd.concat([new_df, group], ignore_index=True)

    # 存储课程信息
    courses_info = []
    current_semester = None


    # 遍历每一行数据
    for index, row in new_df.iterrows():
        if '学期' in str(row[0]):
            current_semester = row[0]
            # print(f'统计学期信息：{current_semester}')
            continue
        course_name = row[0]
        course_type = row[1]
        grade_point = row[2]
        score = row[3]
        courses_info.append({'学期': current_semester, '课程': course_name, '性质': course_type, '学分': grade_point, '分数': score})

    # print(f'学生成绩单全部课程数据：{courses_info}')
    print('==加载完毕学生全部课程数据，进行绩点计算==')

    # 统计方式（学年、课程性质）
    filtered_courses_info = filter_course(courses_info, semester_list, course_list)

    # 课程绩点计算
    # 计算每门课程的绩点，并添加到课程信息中
    for course in filtered_courses_info:
        course['绩点'] = gpa(course['分数'])

    # 计算总学分
    total_credits = sum(course['学分'] for course in filtered_courses_info)
    # 计算学年平均绩点
    total_gpa = sum(course['绩点'] * course['学分'] for course in filtered_courses_info)
    average_gpa = round(total_gpa / total_credits, excel_digit)

    print(f'【{stu_name}】成绩信息')
    [print(item) for item in filtered_courses_info]
    print(f"【{stu_name}】学年({semester_list}) 获得学分：{total_credits}, 平均绩点：{average_gpa}")

    return filtered_courses_info, total_credits, average_gpa


if __name__ == '__main__':
    result_ls = []
    for file_name in os.listdir(PATH):
        if file_name.endswith('.pdf'):
            # 打开PDF文件
            pdf_path = os.path.join(PATH, file_name)
            stu_name = file_name.split('.')[0] # 去除.pdf后缀
            filtered_courses_info, total_credits, average_gpa = fuck_score(pdf_path, stu_name, semester_list, course_list)
            result_ls.append({'学生': file_name, '平均绩点': average_gpa, '统计课程数':len(filtered_courses_info), '统计课程明细':filtered_courses_info, '课程总学分': total_credits})

    # 将结果保存到Excel文件中
    result_df = pd.DataFrame(result_ls)
    output_path = os.path.join(f'result.xlsx')
    result_df.to_excel(output_path, index=True)
    