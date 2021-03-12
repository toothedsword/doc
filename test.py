import docx2txt
import re
import xlwt
import xlrd
import os
import sys


# 参数设置
# {{{
data = {
    '分系统': r'\d+\.\d+',
    '子系统': r'\d+\.\d+\.\d+',
    '模块': r'\d+\.\d+\.\d+\.\d+',
        }
lev = ['分系统', '子系统', '模块']
# }}}

# 输入输出文件名
# {{{
infile = sys.argv[1]
outfile = sys.argv[2]
oldfile = 'xxx'
if len(sys.argv) > 3:
    oldfile = sys.argv[3]
# }}}

# 读取文件内容
# {{{
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My sheet')
text = docx2txt.process(infile)
texts = re.findall(r'([^\n\r]+)', text)
# }}}

# 读取旧表格内容
# {{{
old = {}
if len(sys.argv) > 3 and os.path.exists(oldfile):
    workread = xlrd.open_workbook(oldfile)
    table = workread.sheet_by_index(0)
    k = table.nrows
    for i in range(0, k):
        try:
            md = table.cell_value(rowx=i, colx=1)
            print(md)
            if re.match('.*模块', md):
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
num = []
for ilev in range(0, len(lev)):
    num.append(0)
print(num)
k = 0
ns = []
for t in texts:
    for ilev in range(0, len(lev)):
        if re.match(r'^\s*'+data[lev[ilev]]+r'[^\d\.].*'+lev[ilev]+r'\s*$', t):
            num[ilev] = num[ilev]+1
            for jlev in range(ilev+1, len(lev)):
                num[jlev] = 0
            nt = re.sub(r'^[\d\.]+', '', t)
            n = ''
            for klev in range(0, len(lev)):
                if num[klev] > 0:
                    n = n+str(num[klev])+'.'
            n = re.sub('.$', '', n)
            print(nt)
            k = k + 1
            ns.append(n)
            worksheet.write(k, 0, n)
            worksheet.write(k, 1, nt)
            if re.match(r'^\s*'+data[lev[len(lev)-1]] +
                        '.*'+lev[len(lev)-1]+r'\s*$', t):
                m = 1.5
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
n1 = '1'
n2 = '0.1'
n3 = '0.0.1'
st = ''
for i in range(0, k):
    si = str(i+2)
    ni = ns[i]
    ss = ''
    if re.match(r'^\d+$', ni):
        st = st+'L'+si+','
    for j in range(0, k):
        sj = str(j+2)
        nj = ns[j]
        if re.match(ni+r'\.\d+$', nj):
            n2 = n
            ss = ss + 'L'+sj+','
    ss = re.sub(',$', '', ss)
    if re.match(r'[^\s]', ss):
        ss = 'SUM('+ss+')'
        if re.match(r'\d+.\d+', ni):
            ss = re.sub(r',.*,', ':', ss)
            ss = re.sub(r',', ':', ss)
        worksheet.write(i+1, 11, xlwt.Formula(ss))
st = re.sub(',$', '', st)
st = 'SUM('+st+')'
worksheet.write(0, 11, xlwt.Formula(st))
# }}}

# 输出打开表格
workbook.save(outfile)
os.system("wps " + outfile)
