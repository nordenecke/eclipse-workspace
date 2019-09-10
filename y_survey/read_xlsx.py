'''
Created on 2019年9月10日

@author: 静火
'''
import xlrd


def readExcel(filename, cn, check_province, check_time, FileType):
    # 读取
    workbook = xlrd.open_workbook(filename)
    # 获取sheet
    sheet_name = workbook.sheet_names()[0]
    sheet = workbook.sheet_by_name(sheet_name)

    check_Item = 'a'
    
    itemCount = 0
    score = 0
    
    second = sheet.cell(7, 1).value.encode('utf-8')

    for i in range(7, sheet.nrows):
        if sheet.cell(i, 1).value.encode('utf-8') == second:
            check_Item = sheet.cell(i, 0).value.encode('utf-8')
            continue

    temp = []
    for j in range(0, sheet.ncols):
        temp.append(sheet.cell(i, j).value.encode('utf-8'))

    answer = sheet.cell(i, 7).value.encode('utf-8')

    if answer == "yes" or answer == "no":
        score = score + 1

    if answer == "other":
        print("!!!Failed to import'%s'" % (filename))
        print("!!!Please Choose an Right Answer for '%s'--------" % (filename))
        break
    else:
        cn.execute("insert into TB_CHECK (ITEM,FIELD,TYPE,CONTENT,"
                     "ATTRIBUTE,CHECKPOINT,REMARKS,ANSWER,DESCRIPTION,"
                     "SUGGESTION,PROVINCE,TIME,STYLE) "
                     "values('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                     "" % (temp[0], temp[1], temp[2], temp[3], temp[4], temp[5], temp[6], temp[7], temp[8], temp[9], check_province, check_time, check_Item))

        itemCount = itemCount + 1
    if itemCount != 0:
        score = round(score * (100 / itemCount), 2)
        cn.execute("insert into TB_SCORE (PROVINCE,TIME,FILETYPE,SCORE) "
               "values('%s','%s','%s','%.2f')" % (check_province, check_time, FileType, score))
        print("Successful for'%s'--------" % (filename))
    cn.commit()
