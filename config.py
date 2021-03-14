
datas = {}

datas['common'] = {
    '分系统': r'\d+\.\d+',
    '子系统': r'\d+\.\d+\.\d+',
    '模块': r'\d+\.\d+\.\d+\.\d+',
    'lev': ['分系统', '子系统', '模块'],
}

datas['ads'] = {
    '分系统需求分析': r'第\d+章',
    '子系统': r'\d+\.\d+',
    '模块': r'\d+\.\d+\.\d+',
    'lev': ['分系统需求分析', '子系统', '模块'],
}
