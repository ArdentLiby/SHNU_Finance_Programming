# -*- coding: utf-8 -*-
# @Time : 2023/3/11 13:18
# @Author : Boyang Li
# @File : initConfig.py

config = {
    "lim_scores": 200,
    "options_sex": ["男性", "女性", ""],
    "ch_en": {"投资学": "Investment",
              "金融学": "Finance",
              "信用管理": "Credit Management",
              "金融工程": "Financial Engineering",
              "金融学（拔尖人才实验班）": "Finance (Top talent experimental class)",
              "金融学(数学金融人工智能实验班)": "Finance (Mathematics, Finance and Artificial Intelligence Experimental Class)",
              "金融科技": "Financial Technology",
              "经济学": "Economics",
              "经济学（中美合作）": "Economics (Sino-US Cooperation)",
              "经济学（中法合作）": "Economics (Sino-French Cooperation)",
              "计算机科学与技术（中法合作）": "Computer Science and Technology (Sino-French Cooperation)",
              "广告学（中法合作）": "Advertising (Sino-French Cooperation)",
              "财务管理": "Financial Management",
              "资产评估": "Assets Appraisal",
              "电子商务": "Electronic Business"},
    "students_list_path": "",
    "outputs_dir": "",
    "score_columns": [3, 4],
    "stu_info_columns": [-1, -1, -1, -1, -1, -1],
    "sex_matchup": {"男性": ["Male", "He", "his"], "女性": ["Female", "She", "her"]},
    "document_type": 1
}

config_batch = {
    "average_type": 1,
    "input_dir": ""
}