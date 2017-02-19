#coding : utf-8
#author : 垚望

import xlwt,xlrd,os,json,time

start = time.time()#记录程序开始运行时间


def getPaper(content):# 卷子分为ABCD;用于获取字母
    title = content.row_values(0)[0]
    titleList = list(title)
    if '卷' in title:
        ndx = titleList.index('卷')-1
        paper = titleList[ndx]
        if paper =='A' or paper =='B' or paper =='C' or paper =='D':
            return paper

    return 'E'

def getJson(paper):#答案存于json文件中,用于解析答案

    with open('answer.json',encoding='utf-8') as file:
        content = json.load(file)[paper]

    return content

def logging(fileName,data): #生成改卷日志

    with open( './logs/' +  fileName + '.txt','w+',encoding='utf-8') as f:
        for line in data:
            f.write(line + '\n')


if __name__ == '__main__':
    count = 0 #记录已改的卷数
    score = 0   #记录单人成绩
    stu_info = '' #储存答题卡获取到的个人信息
    stuInfo = []    #List类型的个人信息
    paper = ''  #卷子的字母 /A B C D
    for exl in os.listdir('./files'):   #遍历答题卡

        if os.path.splitext(exl)[1] == '.xls' :
            score = 0
            judgeLog = []
            workbook = xlrd.open_workbook('./files/' + exl) #读取答题卡
            content = workbook.sheet_by_index(0)
            fileName = os.path.splitext(exl)[0]
            if content.nrows != 17 or content.ncols != 11:
                print('加载 %s 时出错' % (fileName))
                continue
            else:

                className,stuName,stuID = content.cell_value(2,1),content.cell_value(2,4),str(content.cell_value(2,7))
                stu_info = (className + stuName + stuID).replace(' ', '')
                paper = getPaper(content)
                if paper == 'E':
                    print('加载 %s 时出错,文件名 : %s'%(stu_info,fileName))
                    continue

                answer = getJson(paper)

                for raw in range(3,16,2):#遍历答题卡选项
                    for col in range(1,11):
                        number = content.cell_value(raw,col)
                        if isinstance(number,float):
                            numInt = int(number)
                            if 1<=numInt<=70:
                                if content.cell_value(raw+1,col) == answer[numInt-1]:
                                    score += 1
                                    judgeLog.append(str(numInt) + ' : √' + ' ')

                                else:
                                    judgeLog.append(str(numInt)  + ' : ×' + ' ')

                logging(stu_info,judgeLog)

                stuInfo.append([[className],[stuName],[stuID],[str(score)]])

                count += 1
            print('\033[1;34;40m%s   得分:%d' % (stu_info, score))

    summary = xlwt.Workbook()
    sheet = summary.add_sheet('summary',cell_overwrite_ok=True)
    row0 = ['班级','姓名','学号','成绩']
    cols = len(row0)
    for i in range(cols):
        sheet.write(0,i,row0[i])

    for raw in range(1,count+1):
        for col in range(cols):
            sheet.write(raw,col,stuInfo[raw-1][col])

    summary.save('score.xls')

    print('改卷用时:%ds'%(time.time()-start))


