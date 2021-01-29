import os
import re
import xlrd
import xlwt


count = 0


rootdir = 'D:\\XJTLU UPD DCW\\UPD Data Property' #change here

def get_files(path=rootdir,rule='.xls'):
    all = []
    for fpathe,dirs,fs in os.walk(path):   # os.walk是获取所有的目录
        for f in fs:
            filename = os.path.join(fpathe,f)
            if filename.endswith(rule):  # 判断是否是"xxx"结尾
                all.append(filename)
    return all


if __name__ == "__main__":


    b = get_files()
    for i in b:
        '''
        print i
        list = os.listdir(rootdir) #列出文件夹下所有的目录与文件
        for i in range(0,len(list)):
        path = os.path.join(rootdir,list[i])
        if os.path.isfile(path):#你想对文件的操作
        '''

        data= xlrd.open_workbook(i)

        nums = len(data.sheets())

        sheet1 = data.sheets()[0]

        nrows = sheet1.nrows
        ncols = sheet1.ncols

        rows_get = []

        for i in range(nrows):
            A0 = sheet1.cell(i,0).value
            A0 = A0.strip()

            if i<2:
                rows_get.append(i)
            else:
                p = r'[\u4e00-\u9fa5]'

                pattern = re.compile(p)

                try:
                    chk_first = re.findall(pattern,A0)[0]

                    if A0[0:4] == '当页汇总':
                        pass
                    else:
                        rows_get.append(i)
                except:
                    continue

        workbook = xlwt.Workbook('ascii')
        sheet_w = workbook.add_sheet('write')

        wx = 0
        for x in rows_get:
            for y in range(ncols):
                sheet_w.write(wx, y, sheet1.cell(x,y).value)
            wx += 1
        workbook.save(str(i))
        count+=1
        print(str(count)+' success')


