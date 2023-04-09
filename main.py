# -*- coding: utf-8 -*-
# @Time : 2023/3/4 0:34
# @Author : Boyang Li
# @File : main.py

import os
import json
import re

import pypinyin
import pandas as pd
from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
from tkinter.ttk import Combobox
from assist import open_file, clear_input, open_directory, read_config, calc_score, show_popup, right_click
from batchExport import export_batch


def import_list(stu_num, stu_info_col=None):
    """
    :param stu_num: 学生的学号。
    :param stu_info_col: 学生信息的列数分布（为内置变量）。
    :return: None

    导入并填充学生信息的函数，因为学号对于每个学生都是唯一的，因此导入学生信息以学号为关键信息，只有通过学号才能导入学生的各项信息。
    """
    if stu_info_col is None:
        stu_info_col = [0, 1, 2, 3, 4, 5]
    i0, i1, i2, i3, i4, i5 = stu_info_col[0], stu_info_col[1], stu_info_col[2], stu_info_col[3], stu_info_col[4], stu_info_col[5]
    try:
        if students_list_path.split('.')[-1] in 'xlsx':
            stu_infos = pd.read_excel(students_list_path)
        else:
            stu_infos = pd.read_csv(students_list_path)
    except FileNotFoundError:
        show_popup("未找到学生列表", "未找到学生列表，请尝试重新导入学生列表")
        return
    except AssertionError:
        show_popup("未找到学生列表", "未找到学生列表，请尝试重新导入学生列表")
        return
    try:
        stu_info = stu_infos[stu_infos.iloc[:, stu_info_col[1]] == int(stu_num)]
        if stu_info.size != 0:
            stu_info = stu_info.values[0]  # 找到根据学号定位的学生
            # 先对可能存在的原有内容进行清除后再进行填充
            clear_input(entry_name, entry_sex, entry_grade, entry_major, entry_major_en,
                        entry_name_en, entry_stu_type)

            # 开始根据学生信息表填充各项学生信息，未提供所在行列的学生信息将不会进行填充
            entry_name.insert(END, stu_info[i0].strip() if i0 >= 0 else '')
            pinyin_list = pypinyin.lazy_pinyin(stu_info[i0])
            stu_name_en = (pinyin_list[0].upper() + ' ' + ''.join(pinyin_list[1:]).capitalize())
            entry_name_en.insert(END, stu_name_en if i0 >= 0 else '')
            entry_stu_type.insert(END, stu_info[i5] if i5 >= 0 else '')
            if len(stu_info[i2].strip()) == 1:
                entry_sex.insert(END, stu_info[i2].strip() + '性' if i2 >= 0 else '')
            else:
                entry_sex.insert(END, stu_info[i2].strip() if i2 >= 0 else '')
            if i3 >= 0 and str(stu_info[i3])[:4].isdigit():  # 当检测到专业或班级信息同时包含年级和班级时
                raw_string = stu_info[i3]
                raw_string = re.sub(r'\d*级', '', raw_string)
                raw_string = (stu_info[i3])[:4] + raw_string
                if len(re.findall(r'本科\d*班', stu_info[i3])) > 0:
                    raw_string = re.sub(r'本科\d*班', '', stu_info[i3])
                raw_string = raw_string.replace('(', '（')
                raw_string = raw_string.replace(')', '）')
                entry_major.insert(END, raw_string[4:])
                entry_grade.insert(END, stu_info[i3][:4])
                entry_major_en.insert(END,
                                      major_ch_en[raw_string[4:]] if i5 >= 0 else '')
            else:  # 检测到专业或班级信息中只包含专业信息时
                if len(re.findall(r'本科\d*班', stu_info[i3])) > 0:
                    raw_string = re.sub(r'本科\d*班', '', stu_info[i3])
                else:
                    raw_string = stu_info[i3]

                raw_string = raw_string.replace('(', '（')
                raw_string = raw_string.replace(')', '）')
                entry_grade.insert(END, stu_info[i4] if i4 >= 0 else '')
                entry_major.insert(END, raw_string.strip() if i3 >= 0 else '')
                entry_major_en.insert(END, major_ch_en[raw_string.strip()] if i3 >= 0 else '')
        else:
            show_popup("提示", "未找到该学生，请正确输入学号")

    except BaseException as e:
        print(e)
        show_popup("提示", "学生信息不完整，请调整参数")


def set_config(default=False):
    """

    :param default: 控制是否将所填写或选择的配置设为默认。为True时将重写配置文件"config.json"，将目前所填的内容设为默认。
    :return: None

    当选择保存设置时，当前的相关设置仅对本次使用生效，当关闭程序再打开后，所有的设置将恢复默认值；
    当选择设为默认时，当前的相关设置将变为默认值，关闭程序后再次打开将保留最后的更改。
    """
    global students_list_path, outputs_dir, stu_info_columns, score_columns, config, lim_scores, outputs_format,\
        major_ch_en

    try:
        students_list_path = entry_stu_list_file.get()
        outputs_dir = entry_out_dir.get()
        stu_info_columns_raw = [entry_stu_info_place_1.get(), entry_stu_info_place_2.get(),
                                entry_stu_info_place_3.get(), entry_stu_info_place_4.get(),
                                entry_stu_info_place_5.get(), entry_stu_info_place_6.get()]
        stu_info_columns = [-1 if i in [0, ''] else int(i) - 1 for i in stu_info_columns_raw]
        lim_scores = int(entry_max_scores.get())
        score_columns = [int(entry_max_scores_detail_1.get()) - 1, int(entry_max_scores_detail_2.get()) - 1]
        outputs_format = option_output_format.get()
        if entry_add_major.get() != '':
            if entry_add_major_en.get() != '':
                major_ch_en[entry_add_major.get()] = entry_add_major_en.get()
            else:
                del major_ch_en[entry_add_major.get()]
                entry_add_major.delete(0, END)

        # 对配置的相关参数进行更改
        config['students_list_path'] = students_list_path
        config['outputs_dir'] = outputs_dir
        config['stu_info_columns'] = stu_info_columns
        config['score_columns'] = score_columns
        config['lim_scores'] = lim_scores
        config['document_type'] = outputs_format
        config['ch_en'] = major_ch_en
        show_popup("提示", "保存成功！")
        if default:  # 如果选择“设为默认”将重写配置文件
            with open(cwd + '/config2.json', 'w', encoding='utf-8') as f2:
                json.dump(config, f2)
            root.update()

    # 如果部分参数不符合要求
    except ValueError:
        show_popup('提示', "请填入正确的参数或参数类型")


# 本程序的主函数 #
if __name__ == '__main__':

    # 读取工作目录并初始化配置
    cwd = os.getcwd()
    if not (os.path.exists(cwd + '/config2.json') or os.path.exists(cwd + '/config-batch.json')):
        from initConfig import config, config_batch
        with open(cwd + '/config2.json', 'w', encoding='utf-8') as f:
            json.dump(config, f)
        with open(cwd + '/config-batch.json', 'w', encoding='utf-8') as f:
            json.dump(config_batch, f)

    # 读取配置文件
    config = read_config(cwd, 'config2.json')
    config_batch = read_config(cwd, 'config-batch.json')

    score_columns = config['score_columns']
    stu_info_columns = config['stu_info_columns']
    students_list_path = config['students_list_path']
    scores_limitation = config['lim_scores']
    outputs_dir = config['outputs_dir']
    lim_scores = config['lim_scores']
    input_dir = config_batch["input_dir"]
    major_ch_en = config['ch_en']

    # 创建窗口对象
    root = Tk()
    root.title('平均分计算')
    if os.path.exists('./icon.ico'):
        root.iconbitmap('./icon.ico')

    # 获取屏幕宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 窗口宽度和高度
    window_width = 550
    window_height = 650

    # 计算窗口居中时的左上角坐标
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    # 设置窗口大小和位置
    root.geometry("{}x{}+{}+{}".format(window_width, window_height, x, y))
    root.resizable(False, False)

    # 创建一个Notebook小部件
    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    # 创建第1个选项卡
    frame1 = ttk.Frame(notebook)
    notebook.add(frame1, text=' 基本功能 ')

    # 创建第2个选项卡
    frame2 = ttk.Frame(notebook)
    notebook.add(frame2, text=' 批量生成文档 ')

    # 创建第3个选项卡
    frame3 = ttk.Frame(notebook)
    notebook.add(frame3, text=' 自定义设置 ')

    # 如果目录下有“background.png”的图片，会将其加载为背景
    if os.path.exists(cwd + '/background.png'):
        image = Image.open("background.png").resize((545, 680))
        x, y = image.size  # 获得长和宽
        photo = ImageTk.PhotoImage(image)
        label_pic1 = Label(frame1, image=photo)
        label_pic1.pack()
        label_pic2 = Label(frame2, image=photo)
        label_pic2.pack()
        label_pic3 = Label(frame3, image=photo)
        label_pic3.pack()

    label2 = Label(frame2, text='学生成绩表文件名需包含学生学号\n请保证学生信息表中信息的完整')
    label2.place(relx=0.5, rely=0.11, anchor='s')
    label3 = Label(frame3, text='根据您的需求调整下列参数：')
    label3.place(relx=0.5, rely=0.09, anchor='s')

    ## 基本功能方面 ##
    # 学生基本信息输入
    label_name = Label(frame1, text="姓名")
    label_name.place(relx=0.07, rely=0.11, anchor='s')
    entry_name = Entry(frame1, width=10)
    entry_name.place(relx=0.19, rely=0.11, anchor='s')

    # 学号部分
    label_num = Label(frame1, text="学号")
    label_num.place(relx=0.37, rely=0.11, anchor='s')
    entry_num = Entry(frame1, width=12)
    entry_num.place(relx=0.5, rely=0.11, anchor='s')
    button_import = Button(frame1, text="导入..", command=lambda: import_list(entry_num.get(), stu_info_columns))
    button_import.place(relx=0.62, rely=0.115, anchor='s')
    entry_num.bind('<Button-3>', lambda a: right_click(a, frame1, entry_num))

    # 性别部分
    label_sex = Label(frame1, text="性别")
    label_sex.place(relx=0.7, rely=0.11, anchor='s')
    entry_sex = Combobox(frame1, values=config['options_sex'], width=10)
    entry_sex.place(relx=0.83, rely=0.11, anchor='s')

    # 年级部分
    label_grade = Label(frame1, text="年级")
    label_grade.place(relx=0.07, rely=0.20, anchor='s')
    entry_grade = Entry(frame1, width=10)
    entry_grade.place(relx=0.19, rely=0.20, anchor='s')
    entry_grade.insert(END, entry_num.get())

    # 专业名称部分
    label_major = Label(frame1, text="专业")
    label_major.place(relx=0.37, rely=0.20, anchor='s')
    entry_major = Combobox(frame1, values=list(major_ch_en.keys()), width=22)
    entry_major.place(relx=0.58, rely=0.20, anchor='s')

    # 姓名拼音部分
    label_name_en = Label(frame1, text="姓名（拼音）")
    label_name_en.place(relx=0.115, rely=0.29, anchor='s')
    entry_name_en = Entry(frame1, width=10)
    entry_name_en.place(relx=0.26, rely=0.29, anchor='s')

    # 专业的英文名称部分
    label_major_en = Label(frame1, text="专业（英文）")
    label_major_en.place(relx=0.44, rely=0.29, anchor='s')
    entry_major_en = Combobox(frame1, values=list(major_ch_en.values()), width=22)
    entry_major_en.place(relx=0.68, rely=0.29, anchor='s')

    # 选择学生成绩表所在位置（单个文件）
    label_scores = Label(frame1, text="选择成绩表")
    label_scores.place(relx=0.1, rely=0.56, anchor='s')
    entry_scores = Entry(frame1, width=49)
    entry_scores.place(relx=0.5, rely=0.56, anchor='s')
    browse_button = Button(frame1, text="浏览..", command=lambda: open_file(entry_scores))
    browse_button.place(relx=0.88, rely=0.568, anchor='s')

    # 创建学生学籍状态选项选择按钮
    option_stu_type = ['在读', '毕业']
    label_type = Label(frame1, text="学籍状态")
    label_type.place(relx=0.093, rely=0.37, anchor='s')
    entry_stu_type = Combobox(frame1, values=list(option_stu_type), width=10)
    entry_stu_type.place(relx=0.258, rely=0.37, anchor='s')

    # 创建计算平均数的方式选项选择按钮
    option_calc_type = IntVar()
    label_type = Label(frame1, text="计分方式")
    label_type.place(relx=0.093, rely=0.452, anchor='s')
    Radiobutton(frame1, text="算数平均", variable=option_calc_type, value=1).place(relx=0.25, rely=0.458, anchor='s')
    Radiobutton(frame1, text="加权平均", variable=option_calc_type, value=2).place(relx=0.4, rely=0.458, anchor='s')

    # 输出文字结果按钮
    submit_button = Button(frame1,
                           text=" 计算并输出文字结果 ",
                           command=lambda: calc_score(entry_name.get(), entry_num.get(),
                                                      entry_sex.get(), entry_grade.get(),
                                                      entry_major.get(), entry_scores.get(),
                                                      entry_stu_type.get(), option_calc_type.get(),
                                                      entry_major_en.get(), False, lim_scores, score_columns,
                                                      frame1, config))
    submit_button.place(relx=0.2, rely=0.93, anchor='s')

    # 导出成绩证明文档按钮
    submit_button2 = Button(frame1,
                            text=" 导出文档 ",
                            command=lambda: calc_score(entry_name.get(), entry_num.get(),
                                                       entry_sex.get(), entry_grade.get(),
                                                       entry_major.get(), entry_scores.get(),
                                                       entry_stu_type.get(), option_calc_type.get(),
                                                       entry_major_en.get(), True, lim_scores, score_columns,
                                                       frame1, config))
    submit_button2.place(relx=0.5, rely=0.93, anchor='s')

    # 重置所有输入
    clear_button = Button(frame1, text=" 重 置 ", command=lambda: clear_input(entry_name, entry_num, entry_sex,
                                                                            entry_grade, entry_major, entry_scores,
                                                                            entry_major_en, entry_name_en,
                                                                            entry_stu_type))
    clear_button.place(relx=0.8, rely=0.93, anchor='s')

    # 批处理选项卡
    label_input_dir = Label(frame2, text="需要处理的文件位置")
    label_input_dir.place(relx=0.15, rely=0.16, anchor='s')
    entry_input_dir = Entry(frame2, width=43)
    entry_input_dir.insert(END, input_dir)
    entry_input_dir.place(relx=0.55, rely=0.16, anchor='s')
    browse_button_input_dir = Button(frame2, text="选择文件夹", command=lambda: open_directory(entry_input_dir))
    browse_button_input_dir.place(relx=0.91, rely=0.168, anchor='s')

    # 设置输出文档的平均分计算方式
    option_calc_type_batch = IntVar()
    label_type_batch = Label(frame2, text="计分方式")
    label_type_batch.place(relx=0.093, rely=0.252, anchor='s')
    Radiobutton(frame2, text="算数平均", variable=option_calc_type_batch, value=1).place(relx=0.25, rely=0.258, anchor='s')
    Radiobutton(frame2, text="加权平均", variable=option_calc_type_batch, value=2).place(relx=0.4, rely=0.258, anchor='s')

    # 确认导出按钮
    button_confirm_export = Button(frame2, text=' 确认导出 ', command=lambda: export_batch(entry_input_dir.get(),
                                                                                       students_list_path,
                                                                                       option_calc_type_batch.get(),
                                                                                       stu_info_columns,
                                                                                       lim_scores,
                                                                                       score_columns, frame2,
                                                                                       option_output_format.get(),
                                                                                       outputs_dir))
    button_confirm_export.place(relx=0.5, rely=0.93, anchor='s')

    # 以下是设置部分（Frame3）
    label_stu_list_file = Label(frame3, text="学生信息表路径")
    label_stu_list_file.place(relx=0.12, rely=0.16, anchor='s')
    entry_stu_list_file = Entry(frame3, width=48)
    entry_stu_list_file.insert(END, students_list_path)
    entry_stu_list_file.place(relx=0.53, rely=0.16, anchor='s')
    browse_button_stu_list_file = Button(frame3, text=" 浏览.. ", command=lambda: open_file(entry_stu_list_file))
    browse_button_stu_list_file.place(relx=0.89, rely=0.168, anchor='s')

    label_stu_info_place = Label(frame3, text="学生信息表中姓名位于第      列；学号位于第      列；性别位于第      列\n\n"
                                              "专业（或班级）信息位于第      列；年级位于第      列；学籍状态位于第      列")
    label_stu_info_place.place(relx=0.53, rely=0.29, anchor='s')

    # 设置所选择的学生信息表各学生数据所在的列
    entry_stu_info_place_1 = Entry(frame3, width=2)
    entry_stu_info_place_1.insert(END, '' if stu_info_columns[0] < 0 else stu_info_columns[0] + 1)
    entry_stu_info_place_1.place(relx=0.442, rely=0.234, anchor='s')
    entry_stu_info_place_2 = Entry(frame3, width=2)
    entry_stu_info_place_2.insert(END, '' if stu_info_columns[1] < 0 else stu_info_columns[1] + 1)
    entry_stu_info_place_2.place(relx=0.639, rely=0.234, anchor='s')
    entry_stu_info_place_3 = Entry(frame3, width=2)
    entry_stu_info_place_3.insert(END, '' if stu_info_columns[2] < 0 else stu_info_columns[2] + 1)
    entry_stu_info_place_3.place(relx=0.837, rely=0.234, anchor='s')
    entry_stu_info_place_4 = Entry(frame3, width=2)
    entry_stu_info_place_4.insert(END, '' if stu_info_columns[3] < 0 else stu_info_columns[3] + 1)
    entry_stu_info_place_4.place(relx=0.431, rely=0.288, anchor='s')
    entry_stu_info_place_5 = Entry(frame3, width=2)
    entry_stu_info_place_5.insert(END, '' if stu_info_columns[4] < 0 else stu_info_columns[4] + 1)
    entry_stu_info_place_5.place(relx=0.629, rely=0.288, anchor='s')
    entry_stu_info_place_6 = Entry(frame3, width=2)
    entry_stu_info_place_6.insert(END, '' if stu_info_columns[5] < 0 else stu_info_columns[5] + 1)
    entry_stu_info_place_6.place(relx=0.869, rely=0.288, anchor='s')

    # 导出文档位置的相关设置
    label_out_dir = Label(frame3, text="导出文档位置")
    label_out_dir.place(relx=0.12, rely=0.36, anchor='s')
    entry_out_dir = Entry(frame3, width=48)
    entry_out_dir.insert(END, outputs_dir)
    entry_out_dir.place(relx=0.53, rely=0.36, anchor='s')
    browse_button_out_dir = Button(frame3, text="选择文件夹", command=lambda: open_directory(entry_out_dir))
    browse_button_out_dir.place(relx=0.91, rely=0.368, anchor='s')

    # 最大可输入成绩数量设置（虽然题目硬性要求了最大输出为200，但为了增加程序的灵活性设置为了默认值为200的可设置参数）
    label_max_scores = Label(frame3, text="最大成绩条数")
    label_max_scores.place(relx=0.12, rely=0.458, anchor='s')
    entry_max_scores = Entry(frame3, width=8)
    entry_max_scores.insert(END, lim_scores)
    entry_max_scores.place(relx=0.272, rely=0.456, anchor='s')

    # 成绩表的各信息所在位置设置
    label_max_scores_detail = Label(frame3, text="成绩表中学分位于第      列；分数位于第      列")
    label_max_scores_detail.place(relx=0.61, rely=0.458, anchor='s')
    entry_max_scores_detail_1 = Entry(frame3, width=2)
    entry_max_scores_detail_1.insert(END, score_columns[0] + 1)
    entry_max_scores_detail_1.place(relx=0.599, rely=0.456, anchor='s')
    entry_max_scores_detail_2 = Entry(frame3, width=2)
    entry_max_scores_detail_2.insert(END, score_columns[1] + 1)
    entry_max_scores_detail_2.place(relx=0.797, rely=0.456, anchor='s')

    # 调整设置输出文档的类型
    option_output_format = IntVar()
    label_output_format = Label(frame3, text="输出文档类型")
    label_output_format.place(relx=0.12, rely=0.533, anchor='s')
    Radiobutton(frame3, text="Word文档", variable=option_output_format, value=1).place(relx=0.289, rely=0.54, anchor='s')
    Radiobutton(frame3, text="PDF文档", variable=option_output_format, value=2).place(relx=0.438, rely=0.54, anchor='s')
    Radiobutton(frame3, text="二者都输出", variable=option_output_format, value=3).place(relx=0.588, rely=0.54, anchor='s')

    # 如果需要增加或减少专业可以进行设置
    label_add_major = Label(frame3, text="增加（或删除）专业")
    label_add_major.place(relx=0.15, rely=0.62, anchor='s')
    entry_add_major = Combobox(frame3, width=10, values=list(major_ch_en.keys()))
    entry_add_major.place(relx=0.35, rely=0.62, anchor='s')

    label_add_major_en = Label(frame3, text="英文名称（留空以删除）")
    label_add_major_en.place(relx=0.572, rely=0.62, anchor='s')
    entry_add_major_en = Entry(frame3, width=14)
    entry_add_major_en.place(relx=0.79, rely=0.619, anchor='s')
    entry_add_major.bind('<Button-3>', lambda a: right_click(a, frame3, entry_add_major))
    entry_add_major_en.bind('<Button-3>', lambda a: right_click(a, frame3, entry_add_major_en))

    label_notice_1 = Label(frame3, text="*设置参数对“ 基本功能” 和“ 批量生成文档” 功能均生效\n如需增加或删除可选专业，请点击“设为默认”后重启")
    label_notice_1.place(relx=0.5, rely=0.8, anchor='s')

    button_submit_config = Button(frame3, text=' 保存设置 ', command=lambda: set_config())
    button_submit_config.place(relx=0.3, rely=0.93, anchor='s')
    button_default_config = Button(frame3, text=' 设为默认 ', command=lambda: set_config(True))
    button_default_config.place(relx=0.7, rely=0.93, anchor='s')

    # 运行窗口的主事件循环
    root.mainloop()
