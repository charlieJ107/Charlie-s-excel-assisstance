import xlwt
import  xdrlib ,sys
import xlrd

#打开excel文件
def open_excel(file= 'data.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print (str(e))

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_name：Sheet1名称
def excel_table_byname(file= 'data2.xls', colnameindex=0, by_name=u'Sheet2'):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    nrows = table.nrows #行数
    colnames = table.row_values(colnameindex) #某一行数据
    list =[] #装读取结果的序列
    for rownum in range(0, nrows): #遍历每一行的内容
         row = table.row_values(rownum) #根据行号获取行
         if row: #如果行存在
             app = [] #一行的内容
             for i in range(len(colnames)): #一列列地读取行的内容
                app.append(row[i])
             list.append(app) #装载数据
    return list

def testXlwt(file='2017-2018学年第1周校园服务站签到表.xls'):
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet('签到表') #在其中创建一个名为签到表的sheet

    #表头内容
    row_1='2017-2018学年第1周校园服务站签到表'
    row_2 = ['班次','地点','姓名','联系方式','年级','签到','签退']
    line_1 = [
               '周六上午  8:30-11：30',
               '周六中午  11：30-14:30',
               '周六下午  14：30-17:30',
               '周日上午  8:30-11：30',
               '周日中午  11：30-14:30',
               '周日下午  14：30-17:30'
               ]
    line_2 = ['三家村','科艺中心','大南门固定','西校门固定']

    #创建几种样式
    style1 = xlwt.XFStyle() #第一行（标题）样式及文字属性
    font1 = xlwt.Font()
    font1.name = '宋体'
    font1.bold = True
    font1.height = 360
    style1.font = font1

    style2 = xlwt.XFStyle() #第二行（表头）样式及文字属性
    font2 = xlwt.Font()
    font2.name = '宋体'
    font2.bold = True
    font2.height = 240
    style2.font = font2
    style2.border = 1

    style3 = xlwt.XFStyle() #第一、二列样式及文字属性
    font3 = xlwt.Font()
    font3.name = '宋体'
    font3.bold = False
    font3.height = 220
    style3.font = font3

    #对齐设置
    alignment = xlwt.Alignment() #创建居中
    alignment.horz = xlwt.Alignment.HORZ_CENTER #可取值: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER # 可取值: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT # 自动换行
    style1.alignment = alignment # 给样式添加文字居中属性
    style2.alignment = alignment
    style3.alignment = alignment

    #框线设置
    borders = xlwt.Borders() # Create Borders
    borders.left = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    style1.borders = borders # Add Borders to Style
    style2.borders = borders
    style3.borders = borders

    #填充第一行
    sheet1.write_merge(0,0,0,6,row_1,style1)

    #填充第二行
    for i in range(2):
         sheet1.write(1,i,row_2[i],style2)

    #填充第二列
    for i in range(6):
	    sheet1.write_merge(10*i+2,10*i+3,1,1,line_2[0],style3)

    for i in range(6):
	    sheet1.write_merge(10*i+4,10*i+5,1,1,line_2[1],style3)

    for i in range(6):
	    sheet1.write_merge(10*i+6,10*i+8,1,1,line_2[2],style3)

    for i in range(6):
	    sheet1.write_merge(10*i+9,10*i+11,1,1,line_2[3],style3)

    #填充第一列
    for i in range(6):
	    sheet1.write_merge(10*i+2,10*(i+1)+1,0,0,line_1[i],style3)

    #设置列宽
    widths = [10,10,10,18,10,8,8]
    for i in range(6):
	    sheet1.col(i).width = 273*widths[i] # 273为估计比例

    #设置行高
    heights = [45,35,21]
    n=20   # n为估计比例
    for i in range(62):
	    if i==0:
		    sheet1.row(i).height_mismatch = True
		    sheet1.row(i).height = n*heights[0]
	    elif i==1:
		    sheet1.row(i).height_mismatch = True
		    sheet1.row(i).height = n*heights[1]
	    else:
		    sheet1.row(i).height_mismatch = True
		    sheet1.row(i).height = n*heights[2]

    #填充空白部分框线,注意不要填充要填入数据的单元格，否则会报错
    for i in range(2,62):
        for j in range(5,7):
            sheet1.write(i,j,'',style4)

    # 核心部分：数据填充
    time_nums = [0,0,0,0,0,0] # 记录每个班次已排班人数的列表

    for row in tables: # 按照顺利历遍列表
        for i in range(4,10):  # 历遍各个时段
            if row[11]=='*':  # 如果已经被排班
                break
            elif row[i]==1:   # 如果是上岗时段
                if time_nums[i-4]==10:  # 测试该班次是否已经排满
                    break
                sheet1.write(10*(i-4)+2+time_nums[i-4],2,row[1],style4)  # 数据填充
                sheet1.write(10*(i-4)+2+time_nums[i-4],3,row[3],style4)
                sheet1.write(10*(i-4)+2+time_nums[i-4],4,row[2],style4)
                row[i]=0  # 消除该上岗时段
                row[11]='*'  # 排班之后标记*号
                time_nums[i-4]+=1  # 增加该班次已排班人数

book.save(file) #创建保存文件

#主函数
def main():
   testXlwt()

if __name__=="__main__":
    main()