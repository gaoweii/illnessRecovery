import xlwings as xw
import time
import datetime
from PIL import Image
import pytesseract
import re

def binary_search(id, totalId, featureCode):
 #二分查找算法 返回ID的featureCode + 3
	low = 0	
	high = len(totalId) - 1
	while low <= high:
		mid = (low + high) // 2
		if id > totalId[mid]:
			low = mid + 1
		elif id < totalId[mid]:
			high = mid - 1
		else:
			return hash(mid, featureCode);
	return -1


def hash(loc, featureCode): #返回对应位置的featureCode
	if loc == -1:
		return -1
	return featureCode[loc]


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
    		matchList.append(int(obj))
    print(matchList)
    return matchList


def search(allpageId, totalId, sht2, featureCode):
    res = []
    have = False
    notHave = []
    loc = -1
    print('正在根据学生信息表进行比对...')
    for studentsId in allpageId:
    	for _id in studentsId:
    		# 原来给的截图就已经排序好，不需要再排序，这里直接二分查找就行了
    		loc = binary_search(_id, totalId, featureCode)
    		if loc != -1:
    		    have = True
    		    res.append(loc) #这里的loc为sheet1对应的行号
    		if have == False:
    		    notHave.append(_id)
    		have = False
        
    if len(notHave) > 0:
    	for studentsId in notHave:
    		print(studentsId + "不存在于系统中！")
    return res
    

def asciiToChar(x):
    x = x + 5
    split = []
    res = ""
    #观察原表规律 A, B, C …… AA, AB, AC, …… 可以等效为一个26进制数从1开始增大 因此可以当作数字得到其个位，十位，百位……
    while x != 0:
    	split.append(x % 26) 
    	x = x // 26
    for ch in split:
    	res += chr(ch + 64)
    return res

def add(sht1, res):
    startTime = datetime.date(2020, 3, 17)
    # nowTime = datetime.date(2020, 3, 17)
    nowTime = datetime.date.today()
    x = nowTime.__sub__(startTime).days #x为时间偏移量
    if len(res) == 0:
        return
    print('正在添加信息...')
    start = asciiToChar(x)
    i = 0
    for li in res:
    	sht1.range(start + str(li)).value = '未打卡'
    	i += 1
    	print('进度: ' + str(int((i) / len(res) * 100)) + '%')
    sht1.range('F1').value = "最近一次修改时间：" + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    return


def main():
    nowDate = time.strftime("%Y-%m-%d", time.localtime())
    allpageId = []
    print('正在启动excel...')
    # 记录全院学生信息的表格
    wb1 = xw.Book('疫情打卡信息表.xlsx')
    # 新建一个excel表格用于存储信息
    str = ""
    urls = []
    print('疫情打卡辅助填表脚本')
    str = input('请输入图片地址(在本文件夹直接输入名字(需要后缀, end结束))\n')
    while str != 'end':
        urls.append(str)
        str = input('继续输入: ')
    print(urls)
    sht1 = wb1.sheets['Sheet1']
    sht2 = wb1.sheets['Sheet2']
    totalId = sht2.range('A2:A472').value #取出第二张表的ID(已经排好序)
    featureCode = sht2.range('B2:B472').value #取出featureCode
    for i in range(0, len(featureCode)):
    	featureCode[i] = int(featureCode[i])
    #将float类型转化为int类型方便后面与字符串进行比较
    for i in range(0, len(totalId)):
    	totalId[i] = int(totalId[i])
    #得到未按时打卡学生的学号
    allpageId.append(OCR(urls))
    add(sht1, search(allpageId, totalId, sht2, featureCode))
    wb1.save()


if __name__ == '__main__':
    main()

