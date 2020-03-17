import xlwings as xw
import time
from PIL import Image
import pytesseract
import re

# 用的pytesseract开源库，需要事先安装tesseract，并且对pytesseract进行配置
def OCR(urls):
    matchList = []
    # 匹配所有带括号的字符串(括号里是学号)
    pattern =  re.compile('.*\((.*)\).*')
    i = 1
    for url in urls:
    	print('正在对' + str(i) + '张图片进行OCR识别...')
    	list = pytesseract.image_to_string(Image.open(url))
    	i += 1
    	print('正在提取学号...')
    	#得到学号集
    	matchObjs = pattern.findall(list)
    	for obj in matchObjs:
    		matchList.append(obj)
    print(matchList)
    return matchList


def search(allpageId, totalId, sht2):
    res = []
    have = False
    notHave = []
    print('正在根据学生信息表进行比对...')
    for studentsId in allpageId:
    	for _id in studentsId:
            for i in range(0, len(totalId)):
                if _id == str(totalId[i]):
                    res.append(i + 3)
                    have = True
            if have == False:
        	    notHave.append(_id)
            have = False

        
    if len(notHave) > 0:
    	for studentsId in notHave:
    		print(studentsId + "不存在于系统中！")
    return get(res, sht2)


def get(res, sht2):
    _res = []
    tmpRes = []
    for i in res:
    	# 根据结果集里的第i列得到信息
        tmpRes = sht2.range(sht2.range('B' + str(i)), sht2.range('E' + str(i) + ':F' + str(i))).value
        _res.append(tmpRes)
    print('得到相关同学信息...')
    print(_res)
    return _res


def add(sht1, res):
    if len(res) == 0:
        return
    i = 3
    print('正在添加信息...')
    for li in res:
        sht1.range('A' + str(i)).value = li[0]
        sht1.range('B' + str(i)).value = li[1]
        sht1.range('C' + str(i)).value = li[3]
        sht1.range('D' + str(i)).value = li[4]
        i += 1
        print('进度: ' + str(int((i - 3) / len(res) * 100)) + '%')
    sht1.range('F1').value = "最近一次修改时间：" + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    return


def main():
    nowDate = time.strftime("%Y-%m-%d", time.localtime())
    allpageId = []
    print('正在启动excel...')
    # 记录全院学生信息的表格
    wb1 = xw.Book('check.xlsx')
    # 新建一个excel表格用于存储信息
    wb2 = xw.Book()
    str = ""
    urls = []
    print('疫情打卡辅助填表脚本')
    str = input('请输入图片地址(在本文件夹直接输入名字(需要后缀, end结束))\n')
    while str != 'end':
        urls.append(str)
        str = input('继续输入: ')
    print(urls)
    sht1 = wb2.sheets['Sheet1']
    sht2 = wb1.sheets['Sheet2']
    # 添加相关信息
    sht1.range('A1').value = '未按时打卡学生信息'
    sht1.range('A2').value = '学号'
    sht1.range('B2').value = '姓名'
    sht1.range('C2').value = '专业'
    sht1.range('D2').value = '班级'
    sht1.range('E2').value = '打卡日期'
    totalId = sht2.range('B3:B473').value
    #将float类型转化为int类型方便后面与字符串进行比较
    for i in range(0, len(totalId)):
    	totalId[i] = int(totalId[i])
    #得到未按时打卡学生的学号
    allpageId.append(OCR(urls))
    add(sht1, search(allpageId, totalId, sht2))
    wb2.save(nowDate + '.xlsx')


if __name__ == '__main__':
    main()

