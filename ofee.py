#!/bin/env python3

# @Author  : Liang Peng
# @Time    : 2021年 03月 13日 星期六 17:34:12 CST
# @Function: 从文档生成经费预算表

import re
import docx2txt
import xlwt
import xlrd
import os
import sys
import config


def get_doc(infile):
    line = docx2txt.process(infile)
    texts = re.findall(r'([^\n\r]+)', line)
    return texts


def read_old_table(oldfile):
    # {{{
    # 如果输入参考文件，则更新参考字典
    old = {}  # 初始化参考字典
    if len(sys.argv) > 3 and os.path.exists(oldfile):
        # 打开参考文件
        workread = xlrd.open_workbook(oldfile)
        # 获取参考文件
        table = workread.sheet_by_index(0)
        k = table.nrows
        # 遍历参考文件所有行
        for i in range(0, k):
            try:
                md = table.cell_value(rowx=i, colx=1)
                if re.match('.*模块', md):
                    # 存储所有的模块的经费值
                    old[md] = [
                        table.cell_value(rowx=i, colx=4),
                        table.cell_value(rowx=i, colx=5),
                        table.cell_value(rowx=i, colx=6),
                        table.cell_value(rowx=i, colx=7),
                    ]
                    print(md)
                    print(old[md])
            except ValueError:
                pass
    return old
    # }}}


def analysis_text(texts, data, lev):
    # {{{
    # 解析文件内容并更新输出表格内容
    # 初始化标题编号
    num = []
    for ilev in range(0, len(lev)):
        num.append(0)

    # 初始标题和编号
    nums = []
    titles = []

    # 遍历所有内容
    for t in texts:
        # 遍历所有级别规则寻找所有符合条件的标题
        for ilev in range(0, len(lev)):
            # 根据规则匹配标题
            if re.match(r'^\s*' + data[lev[ilev]] +
                        r'[^\d\.].*' + lev[ilev] + r'\s*$', t):
                # 标题编号值更新
                num[ilev] = num[ilev]+1  # 本级提升
                for jlev in range(ilev+1, len(lev)):
                    num[jlev] = 0  # 所有下级归零
                # 去掉标题的编号
                nt = re.sub(r'^\s*' + data[lev[ilev]]+r'\s*', '', t)
                # 生成新的编号
                n = ''
                for klev in range(0, len(lev)):
                    if num[klev] > 0:
                        n = n+str(num[klev])+'.'
                n = re.sub('.$', '', n)  # 去掉末尾的‘.’
                # 行号+1
                # 存储标题编号
                nums.append(n)
                titles.append(nt)
    return nums, titles
    # }}}


def gen_table(nums, titles, data, lev):
    # {{{
    # 初始总结公式
    st = ''
    tab = []
    i0 = 2
    for i in range(0, len(nums)+i0):
        tabi = []
        for j in range(0, 20):
            tabi.append(' ')
        tab.append(tabi)
    for i in range(0, len(nums)):
        i1 = i0 + i
        num = nums[i]
        si1 = str(i1+1)
        tab[i1][0] = nums[i]
        tab[i1][1] = titles[i]
        if re.match(r'\d+\.\d+\.\d+', num):
            tab[i1][4] = 1.5
            tab[i1][5] = 3
            tab[i1][6] = 1.5
            tab[i1][7] = 1
            tab[i1][8] = 'function=SUM(E'+si1+':H'+si1+')'
            tab[i1][9] = 2
            tab[i1][10] = 0
            tab[i1][11] = 'function=I'+si1+'*J'+si1

    # 添加公式
    for i in range(i0, i1):
        si = str(i+1)
        ni = nums[i-i0]
        if re.match(r'^\d+$', ni):
            st = st+'L'+si+','  # 收集分系统系统行号

        # 遍历所有行寻找当前标题的下一级编号
        ss = ''  # 初始分系统子系统公式
        for j in range(i0, i1):
            sj = str(j+1)
            nj = nums[j-i0]
            if re.match(ni+r'\.\d+$', nj):
                # 如果和上级匹配，则收集行号
                ss = ss + 'L'+sj+','
        ss = re.sub(',$', '', ss)  # 去掉行末的‘,’

        # 如果公式参数非空，则录入
        if re.match(r'[^\s]', ss):
            ss = 'SUM('+ss+')'
            # 将子系统的‘，’改成‘：’
            if re.match(r'\d+.\d+', ni):
                ss = re.sub(r',.*,', ':', ss)
                ss = re.sub(r',', ':', ss)
            # 录入公式（分系统，子系统）
            tab[i][11] = 'function='+ss

    # 录入总的公式(系统)
    st = re.sub(',$', '', st)
    st = 'SUM('+st+')'
    tab[i0-1][11] = 'function='+st
    return tab
    # }}}


def output_table(outfile, tab):
    # {{{
    # 打开输出文件
    workbook = xlwt.Workbook()
    # 打开sheet
    worksheet = workbook.add_sheet('My sheet')

    for i in range(0, len(tab)):
        for j in range(0, len(tab[i])):
            tabi = tab[i][j]
            if re.match(r'^\s*function=', str(tabi)):
                tabi = re.sub(r'\s*function=', '', tabi)
                worksheet.write(
                    i, j, xlwt.Formula(tabi))
            else:
                worksheet.write(i, j, tabi)
    workbook.save(outfile)
    os.system("wps " + outfile)
    # }}}


def add_old_table(old, tab):
    # {{{
    for i in range(0, len(tab)):
        if re.match(r'[^\s]', tab[i][1]):
            if tab[i][1] in old:
                tab[i][4:8] = old[tab[i][1]]
    return tab
    pass  # }}}


# 参数设置
# {{{
# 标题级别和级别顺序
tp = sys.argv[1]
data = config.datas[tp]
lev = config.datas[tp]['lev']
# }}}

# 输入输出文件名
# {{{
infile = sys.argv[2]  # 输入的docx文件
if not(os.path.exists(infile)):
    print('Error: "'+infile+'" do not exists')
    exit()
outfile = sys.argv[3]  # 输出的xlsx文件
if os.path.exists(outfile):
    print('Error: output file "'+outfile+'" exists')
    exit()
oldfile = 'xxx'  # 用以参考的旧的xlsx文件
if len(sys.argv) > 4:
    oldfile = sys.argv[4]
# }}}

# 读取文件内容
texts = get_doc(infile)

# 解析文件内容
nums, titles = analysis_text(texts, data, lev)

# 生成表格内容
tab = gen_table(nums, titles, data, lev)

# 读取旧表格内容
old = read_old_table(oldfile)

# 添加旧表格内容
tab = add_old_table(old, tab)

# 输出表格内容
output_table(outfile, tab)
