# @Author  : Liang Peng
# @Time    : 2021年 03月 12日 星期五 17:34:12 CST
# @Function: 生成经费预算表

import re
import docx2txt
import xlwt
import xlrd
import os
import sys
import config


# 参数设置
# {{{
# 标题级别和级别顺序
case = 0
if case == 0:
    data = {
        '分系统': r'\d+\.\d+',
        '子系统': r'\d+\.\d+\.\d+',
        '模块': r'\d+\.\d+\.\d+\.\d+',
            }
    lev = ['分系统', '子系统', '模块']
if case == 1:
    data = {
        '分系统需求分析': r'第\d+章',
        '子系统': r'\d+\.\d+',
        '模块': r'\d+\.\d+\.\d+',
            }
    lev = ['分系统需求分析', '子系统', '模块']
# }}}

# 输入输出文件名
# {{{
infile = sys.argv[1]  # 输入的docx文件
if not(os.path.exists(infile)):
    print('Error: "'+infile+'" do not exists')
    exit()
outfile = sys.argv[2]  # 输出的xlsx文件
if os.path.exists(outfile):
    print('Error: output file "'+outfile+'" exists')
    exit()
oldfile = 'xxx'  # 用以参考的旧的xlsx文件
if len(sys.argv) > 3:
    oldfile = sys.argv[3]
# }}}

# 读取文件内容
# {{{
line = docx2txt.process(infile)
texts = re.findall(r'([^\n\r]+)', line)
# }}}

# 读取旧表格内容
# {{{
old = {}  # 初始化参考字典
# 如果输入参考文件，则更新参考字典
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
# }}}

# 解析文件内容
# {{{
# 打开文件
workbook = xlwt.Workbook()
# 打开sheet
worksheet = workbook.add_sheet('My sheet')
# 初始化标题编号
num = []
for ilev in range(0, len(lev)):
    num.append(0)
print(num)
k = 0
ns = []
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
            print(nt)
            # 行号+1
            k = k + 1
            # 存储标题编号
            ns.append(n)
            # 填写表格
            worksheet.write(k, 0, n)
            worksheet.write(k, 1, nt)
            # 填写人月等
            if re.match(r'^\s*'+data[lev[len(lev)-1]] +
                        '.*'+lev[len(lev)-1]+r'\s*$', t):
                m = 1.5
                # 如果旧的文件有内容，则利用旧文件的内容
                if nt in old:
                    worksheet.write(k, 4, old[nt][0])
                    worksheet.write(k, 5, old[nt][1])
                    worksheet.write(k, 6, old[nt][2])
                    worksheet.write(k, 7, old[nt][3])
                else:
                    worksheet.write(k, 4, m)
                    worksheet.write(k, 5, m*2)
                    worksheet.write(k, 6, m)
                    worksheet.write(k, 7, m-0.5)

                # 录入公式计算模块的经费
                sk = str(k+1)
                worksheet.write(
                    k, 8, xlwt.Formula('SUM(E'+sk+':H'+sk+')'))
                worksheet.write(k, 9, 2)
                worksheet.write(k, 10, 0)
                worksheet.write(
                    k, 11, xlwt.Formula('I'+sk+'*J'+sk))
# }}}

# 编辑总结公式
# {{{
st = ''
# 遍历所有行获取当前标题编号
for i in range(0, k):
    si = str(i+2)
    ni = ns[i]
    ss = ''
    if re.match(r'^\d+$', ni):
        st = st+'L'+si+','
    # 遍历所有行寻找当前标题的下一级编号
    for j in range(0, k):
        sj = str(j+2)
        nj = ns[j]
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
        worksheet.write(i+1, 11, xlwt.Formula(ss))

# 录入系统公式
st = re.sub(',$', '', st)
st = 'SUM('+st+')'
worksheet.write(0, 11, xlwt.Formula(st))
# }}}

# 输出表格并打开
# {{{
workbook.save(outfile)
os.system("wps " + outfile)
# }}}
