# -*- coding: utf-8 -*-
# @Time : 2023/3/8 22:59
# @Author : Boyang Li
# @File : batchExport.py

import re
import os
import json
import pandas as pd
from tkinter import Frame, Text, END, Label
from assist import read_config, show_popup, calc_score

cwd = os.getcwd()  # 获取当前工作路径


def export_batch(path_scores, students_list_path, average_type, stu_info_col=None, lim_scores=200,
                 score_columns=None, frame: Frame = None, file_type=1, out_path=None):

    # 每次批量导出之前均会进行初始化
    set_batch_config({"average_type": average_type, "input_dir": path_scores})
    students_scores = os.listdir(path_scores)

    # 初始化导出成功/失败的文件列表
    success, failures = [], []

    # 对目标文件夹中符合条件的excel文档进行逐个处理
    for f in students_scores:
        if 'xls' in f:
            try:
                code = generate_files(students_list_path, path_scores, f, average_type, stu_info_col, lim_scores,
                                      score_columns, file_type)
                if code == -1:  # 没有找到学生列表或信息有误时直接结束函数
                    return
                if code == -2:  # 文件名不符合要求时跳过该文件
                    continue
                success.append(path_scores + '/' + f)
            except BaseException as e:  # 出现错误便会加入导出失败的列表
                print(e)
                failures.append(path_scores + '/' + f)

    # 输出导出文件的结果（有未成功导出的情况）
    if len(failures) > 0:
        label0 = Label(frame, text=' '*1000)
        label0.place(relx=0.28, rely=0.35, anchor='s')
        label_process = Label(frame, text='已将下列成绩表导出为文档：')
        label_process.place(relx=0.28, rely=0.35, anchor='s')
        text_process = Text(frame, height=20, width=35, font=("Times New Roman", 10))
        text_process.place(relx=0.28, rely=0.85, anchor='s')
        label_process_failed = Label(frame, text='未能导出文档的文件：')
        label_process_failed.place(relx=0.72, rely=0.35, anchor='s')
        text_process_failed = Text(frame, height=20, width=35, font=("Times New Roman", 10))
        text_process_failed.place(relx=0.72, rely=0.85, anchor='s')
        for f in failures:
            text_process_failed.insert(END, f + '\n')
        show_popup('提示', '部分文档导出失败，请检查右侧列出的成绩表')

    # 符合条件的excel文件全部导出成功
    else:
        label_process = Label(frame, text='已将下列成绩表导出为文档：')
        label_process.place(relx=0.5, rely=0.35, anchor='s')
        text_process = Text(frame, height=20, width=70, font=("Times New Roman", 10))
        text_process.place(relx=0.5, rely=0.85, anchor='s')
        show_popup('提示', '导出成功！')
    for f in success:
        text_process.insert(END, f + '\n')


def generate_files(students_list_path: str, student_score_path: str, file_name: str, average_type, stu_info_col=None,
                   lim_scores=200, score_columns=None, file_type=1):
    """

    :param student_score_path: 学生成绩表的路径
    :param students_list_path: 学生信息表的路径
    :param file_name: 成绩表的文件名
    :param average_type: 平均分计算方式（加权平均或算数平均）
    :param stu_info_col: 学生信息的列数分布（为内置变量）。
    :param lim_scores: 最大成绩条数
    :param score_columns: 成绩表中学分和得分信息所在列
    :param file_type: 生成文档的类型设置
    :return: None

    导入并填充学生信息的函数，因为学号对于每个学生都是唯一的，因此导入学生信息以学号为关键信息，只有通过学号才能导入学生的各项信息。
    """
    if stu_info_col is None:
        stu_info_col = [0, 1, 2, 3, 4, 5]
    i0, i1, i2, i3, i4, i5 = stu_info_col[0], stu_info_col[1], stu_info_col[2], stu_info_col[3], stu_info_col[4], \
                             stu_info_col[5]
    try:
        if students_list_path.split('.')[-1] in 'xlsx':
            stu_infos = pd.read_excel(students_list_path)
        else:
            stu_infos = pd.read_csv(students_list_path)

    # 两类严重错误，即学生信息表不在指定路径下，或给定的文件错误
    except FileNotFoundError:
        show_popup("未找到学生列表", "未找到学生列表，请尝试重新导入学生列表")
        return -1
    except AssertionError:
        show_popup("未找到学生列表", "未找到学生列表，请尝试重新导入学生列表")
        return -1

    stu_num = re.findall(r'\d{9}', file_name.split('.')[0])  # 寻找文件名中的九位连续数字
    if len(stu_num) == 0:  # 如未找到连续的9位数字将跳过后续步骤
        return -2
    stu_num = stu_num[0]
    stu_info = stu_infos[stu_infos.iloc[:, stu_info_col[1]] == int(stu_num)]
    print(stu_info.size)
    if stu_info.size != 0:
        print(stu_info)
        stu_info = stu_info.values[0]  # 找到根据学号定位的学生
        # 开始根据学生信息表填充各项学生信息，未提供所在行列的学生信息将不会进行填充

        stu_name = stu_info[i0].strip() if i0 >= 0 else ''
        stu_type_ = stu_info[i5] if i5 >= 0 else ''
        if len(stu_info[i2].strip()) == 1:
            stu_sex = stu_info[i2].strip() + '性' if i2 >= 0 else ''
        else:
            stu_sex = stu_info[i2].strip() if i2 >= 0 else ''

        config_general = read_config(cwd, 'config2.json')
        config_general['document_type'] = file_type
        if i3 >= 0 and str(stu_info[i3])[:4].isdigit():  # 当检测到专业或班级信息同时包含年级和班级时
            raw_string = stu_info[i3]
            raw_string = re.sub(r'\d*级', '', raw_string)
            raw_string = (stu_info[i3])[:4] + raw_string
            if len(re.findall(r'本科\d*班', stu_info[i3])) > 0:
                raw_string = re.sub(r'本科\d*班', '', stu_info[i3])
            raw_string = raw_string.replace('(', '（')
            raw_string = raw_string.replace(')', '）')
            # 以上均是对学生专业或班级信息的处理
            stu_major = raw_string[4:]
            stu_grade = stu_info[i3][:4]
            stu_major_en = config_general['ch_en'][raw_string[4:]] if i5 >= 0 else ''
        else:  # 检测到专业或班级信息中只包含专业信息时
            if len(re.findall(r'本科\d*班', stu_info[i3])) > 0:
                stu_info[i3] = re.sub(r'本科\d*班', '', stu_info[i3])
            raw_string = stu_info[i3]
            raw_string = raw_string.replace('(', '（')
            raw_string = raw_string.replace(')', '）')
            # 以上均是对学生专业或班级信息的处理
            stu_grade = stu_info[i4] if i4 >= 0 else ''
            stu_major = raw_string.strip() if i3 >= 0 else ''
            stu_major_en = config_general['ch_en'][raw_string.strip()] if i3 >= 0 else ''

        # 计算学生平均成绩并输出文档
        calc_score(stu_name, stu_num,
                   stu_sex, stu_grade,
                   stu_major, student_score_path + '/' + file_name,
                   stu_type_, average_type,
                   stu_major_en, True, lim_scores, score_columns,
                   None, config_general)

    # 找到的值为空值，说明表中不包含该学生或输入有误
    else:
        show_popup("提示", "未找到该学生，请正确输入学号")


# 处理批量导出文档的相关设置
def set_batch_config(config):
    with open(cwd + '/config-batch.json', 'w', encoding='utf-8') as f:
        json.dump(config, f)
