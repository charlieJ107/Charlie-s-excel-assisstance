# -*- coding: utf-8 -*-
import  xdrlib
import  sys
import xlrd
print('Opening file...')
#打开excel文件
def open_excel(file='G:/Charlie/Documents/xmu.edu.cn/大一/社团与学生工作/自律/信息学院宿舍信息表.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

print('Reading file...')
#根据名称获取Excel表格中的数据 参数:file：Excel文件路径 colnameindex：表头列名所在行的索引 ，by_name：Sheet1名称
def excel_table_byname(file='G:/Charlie/Documents/xmu.edu.cn/大一/社团与学生工作/自律/信息学院宿舍信息表.xlsx', colnameindex=0, by_name=u'思明海韵06'):
    data = open_excel(file) #打开excel文件                                              别
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet               碰
    nrows = table.nrows #行数                                                             这
    colnames = table.row_values(colnameindex) #某一行数据                                  段
    list = [] #装读取结果的序列                                                             破
    for rownum in range(0, nrows): #遍历每一行的内容                                         代
         row = table.row_values(rownum) #根据行号获取行                                       码
         if row: #如果行存在                                                                   !
             rowcontain = [] #一行的内容                                                         !
             for i in range(len(colnames)): #一列列地读取行的内容                                  !
               rowcontain.append(row[i])#                                                            !
             list.append(rowcontain) #装载数据                                                         !
    return list



tables = excel_table_byname()
print('Data-reading finished!')

#开始创建表格
import xlwt


def testXlwt(file = 'G:/Charlie/Documents/xmu.edu.cn/大一/社团与学生工作/自律/test.xls'):
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet('testsheet2') #在其中创建一个名为testsheet的sheet
     
   

    #开始填表
    print('Start writing...')

    writerow=0 #要填的行
    writerowbefore=0
    writeline=0 #要填的列
    linedefalt=0
    
    

    roomnumberbefore=100#int(tables[0][2])#定义默认的前一个房间号
    textcontain_name=str(tables[0][3])#定义默认的要填的名字
    
    
    textcontain_sex=str(tables[0][4])
    textcontain_subject=str(tables[0][5])
    textcontain_grade=str(tables[0][6])
    
    #一下都是定义默认的要填的性别\专业\年级啥的
    
    
    sexbefore=str(tables[0][4])
    subjectbefore=str(tables[0][5])
    gradebefore=str(tables[0][6])
    
    print("正在迭代数据...")
    for row in tables:#迭代列表
        roomnumber=int(row[2])#取房间号
        roomname=str(row[3])#取名字并强行转换成字符串
        
        sex=str(row[4])
        subject=str(row[5])
        grade=str(row[6])
        
        
        #print(writeline)
        
         
        

        if roomnumber!=roomnumberbefore:#如果房间号不同
            
            writeline=linedefalt#列码归位

            #这里还要加上填专业和年级的

            sheet1.write(writerow,writeline,roomnumber)#写房间号
            writeline=writeline+1
            sheet1.write(writerow,writeline,textcontain_name)#吧taxtcontain_name写进单元格
            #填性别专业和年级
            writeline=writeline+1
            sheet1.write(writerow,writeline,textcontain_sex)
            writeline=writeline+1
            sheet1.write(writerow,writeline,textcontain_subject)
            writeline=writeline+1
            sheet1.write(writerow,writeline,textcontain_grade)
            

            
            
            writerow=writerow+1#跳出本行,填写下一行
           
            #这个时候已经在准备下一行的内容了
            
            textcontain_name=roomname #先写上肯定不一样的每个宿舍第一个人
            
            textcontain_sex=sex
            textcontain_subject=subject
            textcontain_grade=grade
            


            roomnumberbefore=roomnumber
            
            sexbefore=sex
            subjectbefore=subject
            gradebefore=grade
 
        else: #如果房间号相同

            #三个判断,分别判断"上一个人的性别/专业/年级是否与这个人相同
            #如果相同,就判断下一个条件,都判断完了就填名字
            #如果不相同,就往要填的性别/专业/年级的字符串里面添一个内容
            
            if sex!=sexbefore:#这个是判断性别的
                textcontain_sex=textcontain_sex+'、'+sex
            else:
                if subject!=subjectbefore:#判断专业
                    textcontain_subject=textcontain_subject+'、'+subject
                else:
                
                    if grade!=gradebefore:#判断年级
                        tetextcontain_grade=textcontain_grade+'、'+grade

            textcontain_name=str(textcontain_name+'、'+roomname) #往textcontain里填一个名字

            
            roomnumberbefore=roomnumber#标记这次的房间号为"上次的房间号"
            
            sexbefore=sex
            subjectbefore=subject
            gradebefore=grade
                
            

            
    print('迭代完成!')
    print('正在填写数据...')
    print('正在生成表格文件')   

            
       #我也不知道为啥这个代码放在这里,好像放在else里面也没啥毛病[无奈]

    
    print('Saving file...')
    book.save(file) #创建保存文件

#主函数
def main_creat():
   testXlwt()

if __name__=="__main__":
    main_creat()
print('Finish!')