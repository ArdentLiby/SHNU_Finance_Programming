# -*- coding: utf-8 -*-
# @Time : 2023/3/8 23:03
# @Author : Boyang Li
# @File : assist.py

import os
import json
import pypinyin
import pandas as pd
import tkinter as tk
from tkinter import END, Entry, filedialog, Label
from genDocx import gen_document


def open_file(entry):
    # 打开文件选择窗口
    filetypes = (
        ("Excel files", "*.xlsx"), ("Excel 97-2003 files", "*.xls"), ("CSV files", "*.csv"), ("All files", "*.*"))
    # 打开文件对话框，并限制允许导入的文件类型
    file_path = filedialog.askopenfilename(filetypes=filetypes)

    # 如果用户选择了文件，则读取文件中的数据
    if file_path:
        entry.delete(0, END)
        entry.insert(END, f'{file_path}')


def open_directory(entry):
    # 打开文件选择对话框
    file_path = filedialog.askdirectory()

    # 获取选择的文件夹路径
    if file_path:
        entry.delete(0, END)
        entry.insert(END, f'{file_path}')


def clear_input(*args: Entry):
    """
    重置所有输入。
    """
    for ent in args:
        ent.delete(0, END)


def read_config(path_, file):
    """
    读取配置文件。
    """
    with open(path_ + '/%s' % file, 'r', encoding='utf-8') as f:
        config_ = json.load(f)
    if len(config_) <= 3:
        return config_

    if not config_['outputs_dir']:
        if not os.path.exists(path_ + '/outputs'):
            os.mkdir(path_ + '/outputs')
        config_['outputs_dir'] = path_ + '/outputs'
    return config_


def calc_score(stu_name, stu_num, stu_sex, stu_grade, stu_major, score_file, opt_1, opt_2, stu_major_en, output=False,
               lim_scores_=200, score_place=None, frame_: tk.Frame = None, config: dict = None):
    """
    :param stu_name: 学生姓名（中文）
    :param stu_num: 学生学号
    :param stu_sex: 学生性别
    :param stu_grade: 学生年级
    :param stu_major: 学生专业（中文）
    :param score_file: 学生成绩表路径
    :param opt_1: 学生的学籍状态：在读、毕业
    :param opt_2: 计算的平均分类型：算数、加权
    :param stu_major_en: 专业英文名称
    :param output: 是否输出文档：True表明需要输出；False将不会输出文档
    :param lim_scores_: 成绩条数上限，默认为200
    :param score_place: 成绩表中学分和分数的默认所在列
    :return: None

    是这个程序的主要函数之一，负责平均分的计算、字符串的解析和输出以及控制是否输出文档。
    在点击“文字结果”和“导出文档”时都会调用此函数。
    """
    if score_place is None:
        score_place = [3, 4]
    try:
        score_df = pd.read_excel(score_file)
        if score_df.shape[0] > lim_scores_:  # 如果成绩表的行数超过了限制，将会输出提示
            show_popup('提示', '成绩条数超过限制')
            return
    except:
        show_popup('提示', '请选择或输入正确的成绩表')  # 没有找到目标文件或目标文件无法识别将出现弹窗
        return

    # 字符串解析开始 ######
    if opt_2 == 2:
        type_2 = '加权'
        type_2_en = 'weighted '
        average_score = round(
            (score_df.iloc[:, score_place[1]].values * score_df.iloc[:, score_place[0]].values).sum()
            / score_df.iloc[:, score_place[0]].values.sum(), 2)
    else:
        average_score = round(score_df.iloc[:, score_place[1]].values.sum() / score_df.shape[0], 2)
        type_2 = '算数'
        type_2_en = ''
    if average_score == int(average_score):
        average_score = f'{average_score}0'

    if '毕业' in opt_1:
        stu_type_num = 2
        type_ = '现已毕业。该生就读期间'
        type_en = f'{config["sex_matchup"][stu_sex][1]} graduated from the school with the {type_2_en}average ' \
                  f'score of {average_score} out of 100.'
    else:
        stu_type_num = 1
        type_ = '截至目前，该生'
        type_en = f'The {type_2_en}average score of {config["sex_matchup"][stu_sex][2]} total study' \
                  f' by now is {average_score} out of 100.'

    pinyin_list = pypinyin.lazy_pinyin(stu_name)
    stu_name_en = (pinyin_list[0].upper() + ' ' + ''.join(pinyin_list[1:]).capitalize())

    text_zh_content = f"姓名：{stu_name}，性别：{stu_sex}，学号：{stu_num}，系我院{str(stu_grade)[:4]}级{stu_major}专业的" \
                      f"学生。{type_}所修读的所有课程的{type_2}平均分为{average_score}。"

    text_en_content_0 = f"This is to certify that the student {stu_name_en} ({config['sex_matchup'][stu_sex][0]})," \
                        f" Student Number: {stu_num}," \
                        f" has registered in {str(stu_grade)[:4]} in the major of {stu_major_en} at School of" \
                        f" Finance and Business of Shanghai Normal University."
    text_en_content = text_en_content_0 + ' ' + type_en

    if frame_ is not None:
        text_zh = tk.Text(frame_, height=10, width=30, font=("Times New Roman", 10))
        text_zh.place(relx=0.13, rely=0.6)
        text_zh.insert('end', text_zh_content)
        text_en = tk.Text(frame_, height=10, width=30, font=("Times New Roman", 10))
        text_en.insert('end', text_en_content)
        text_en.place(relx=0.52, rely=0.6)
        label_score = Label(frame_, text="平均分：")
        label_score.place(relx=0.593, rely=0.454, anchor='s')
        text_score = tk.Text(frame_, height=1, width=15, font=("Times New Roman", 10))
        text_score.insert('end', f'{average_score}')
        text_score.place(relx=0.72, rely=0.452, anchor='s')
    # 字符串解析结束 ######

    # 如果用户点击“导出文档”，将执行下列语句
    if output:
        outputs_format = config['document_type']
        outputs_dir = config['outputs_dir']
        gen_document(text_zh_content, text_en_content_0, type_en, stu_name, stu_num, outputs_dir, stu_type_num - 1,
                     outputs_format)


def show_popup(title, text, anchor='s'):
    """
    显示弹出窗口的函数。
    """
    popup = tk.Toplevel()
    screen_width_ = popup.winfo_screenwidth()
    screen_height_ = popup.winfo_screenheight()
    x_popup = (screen_width_ - 300) // 2
    y_popup = (screen_height_ - 100) // 2
    popup.geometry(f"300x100+{x_popup}+{y_popup}")
    popup.title(title)
    label_popup = Label(popup, text=text)
    label_popup.place(relx=0.5, rely=0.45, anchor=anchor)
    button_popup = tk.Button(popup, text=" 确定 ", command=popup.destroy)
    button_popup.place(relx=0.5, rely=0.88, anchor='s')
    popup.resizable(False, False)


def right_click(event, master, editor):
    """
    设置右键菜单的函数。
    """
    menubar = tk.Menu(master, tearoff=False)
    menubar.delete(0, END)
    menubar.add_command(label='复制', command=lambda: editor.event_generate("<<Copy>>"))
    menubar.add_command(label='剪切', command=lambda: editor.event_generate("<<Cut>>"))
    menubar.add_command(label='粘贴', command=lambda: editor.event_generate("<<Paste>>"))
    menubar.add_command(label='全选', command=lambda: editor.event_generate("<<SelectAll>>"))
    menubar.post(event.x_root, event.y_root)
