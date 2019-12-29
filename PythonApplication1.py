#import win32com.client
import docx #获取文档对象
import xlrd #引入excel读取模块

#打开包含空白表格的word，括号内为文件路径和文件名
file=docx.Document('D:\\test1.docx')
#测试一下打开是否成功
#print('测试打开文档内容:段落数:'+str(len(file.paragraphs)))
##输出段落编号及段落内容
#for i in range(len(file.paragraphs)): 
#    print('第'+str(i)+'段的内容是：'+file.paragraphs[i].text)

#打开需要转换的excel，括号内为文件路径和文件名
workbook = xlrd.open_workbook('D:\\test1.xlsx')  #获取数据
worksheet = workbook.sheet_by_name('Sheet1')
nrows = worksheet.nrows
#测试一下打开是否成功，注意单元格计数从0开始的
#print(worksheet.row_values(1)[1])

#未成功的表格复制，后面再说吧
#word = win32com.client.Dispatch('Word.Application')
#doc = word.Documents.Open('D:\\test1.docx')
#word.Content.Copy()

#填写单数表
file.tables[0].rows[0].cells[1].text=str(worksheet.row_values(1)[0])
file.tables[0].rows[0].cells[3].text=worksheet.row_values(1)[1]
file.tables[0].rows[1].cells[1].text=worksheet.row_values(1)[2]

#填写双数表
file.tables[1].rows[0].cells[1].text=str(worksheet.row_values(1)[0])
file.tables[1].rows[0].cells[3].text=worksheet.row_values(1)[1]
file.tables[1].rows[1].cells[1].text=worksheet.row_values(1)[3]

#for i in range(nrows):
#    file.add_sections(file.sections[0])
file.save('D:\\test1.docx')
