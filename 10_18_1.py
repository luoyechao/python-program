#-*-coding:utf-8-*-
import xlrd

import xlwt
def main():
    #写文件
    data_input = xlwt.Workbook()
    sheet = data_input.add_sheet("sheet1")
    #读文件
    string = "201555160"
    for j in range(3,42):
        #string1= string+str(j)+".xls"
        string1 = "20155516%02d.xls" % j
        data = xlrd.open_workbook(string1)
        table = data.sheets()[0]
        nrows = table.nrows
        print "the nrow is ",nrows
        
        row = 0 
        for i in table.row_values(0):
            sheet.write(j,row,i)
            row=row+1
            if i =='':
                continue
            #print i,' ',
    data_input.save("result.xls")

def create():
    data = xlrd.open_workbook("2015551601.xls")
    table = data.sheets()[0]
        
    string = "201555160"
    for i in range(3,42):
        #string1 = string +str(i)+".xls"
        string1 = "20155516%02d.xls" % i
        
        wordbook = xlwt.Workbook()
        sheet = wordbook.add_sheet("sheet1")
        row = 0 
        for i in table.row_values(5):
              sheet.write(0,row,i)
              row=row+1
              print i
        wordbook.save(string1)



def println():
    for i in range(1,42):
        string = "20155516%02d.xls" % i
        print string
        


if __name__ == '__main__':
    #println()
    main()
    #create()

