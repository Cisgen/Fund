# -*- coding:utf-8 -*-
import os,time
import shutil
import json
import ctypes
import sys
import multiprocessing
import openpyxl
from datetime import datetime
import ctypes,sys

STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12

# Windows CMD命令行 字体颜色定义 text colors
FOREGROUND_BLACK = 0x00 # black.
FOREGROUND_DARKBLUE = 0x01 # dark blue.
FOREGROUND_DARKGREEN = 0x02 # dark green.
FOREGROUND_DARKSKYBLUE = 0x03 # dark skyblue.
FOREGROUND_DARKRED = 0x04 # dark red.
FOREGROUND_DARKPINK = 0x05 # dark pink.
FOREGROUND_DARKYELLOW = 0x06 # dark yellow.
FOREGROUND_DARKWHITE = 0x07 # dark white.
FOREGROUND_DARKGRAY = 0x08 # dark gray.
FOREGROUND_BLUE = 0x09 # blue.
FOREGROUND_GREEN = 0x0a # green.
FOREGROUND_SKYBLUE = 0x0b # skyblue.
FOREGROUND_RED = 0x0c # red.
FOREGROUND_PINK = 0x0d # pink.
FOREGROUND_YELLOW = 0x0e # yellow.
FOREGROUND_WHITE = 0x0f # white.

# Windows CMD命令行 背景颜色定义 background colors
BACKGROUND_BLUE = 0x10 # dark blue.
BACKGROUND_GREEN = 0x20 # dark green.
BACKGROUND_DARKSKYBLUE = 0x30 # dark skyblue.
BACKGROUND_DARKRED = 0x40 # dark red.
BACKGROUND_DARKPINK = 0x50 # dark pink.
BACKGROUND_DARKYELLOW = 0x60 # dark yellow.
BACKGROUND_DARKWHITE = 0x70 # dark white.
BACKGROUND_DARKGRAY = 0x80 # dark gray.
BACKGROUND_BLUE = 0x90 # blue.
BACKGROUND_GREEN = 0xa0 # green.
BACKGROUND_SKYBLUE = 0xb0 # skyblue.
BACKGROUND_RED = 0xc0 # red.
BACKGROUND_PINK = 0xd0 # pink.
BACKGROUND_YELLOW = 0xe0 # yellow.
BACKGROUND_WHITE = 0xf0 # white.

# get handle
std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
def set_cmd_text_color(color, handle=std_out_handle):
	Bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
	return Bool

#reset white
def resetColor():
	set_cmd_text_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)

###############################################################

#暗蓝色
#dark blue
def printDarkBlue(mess):
	set_cmd_text_color(FOREGROUND_DARKBLUE)
	sys.stdout.write(mess)
	resetColor()

#暗绿色
#dark green
def printDarkGreen(mess):
	set_cmd_text_color(FOREGROUND_DARKGREEN)
	sys.stdout.write(mess)
	resetColor()

#暗天蓝色
#dark sky blue
def printDarkSkyBlue(mess):
	set_cmd_text_color(FOREGROUND_DARKSKYBLUE)
	sys.stdout.write(mess)
	resetColor()

#暗红色
#dark red
def printDarkRed(mess):
	set_cmd_text_color(FOREGROUND_DARKRED)
	sys.stdout.write(mess)
	resetColor()

#暗粉红色
#dark pink
def printDarkPink(mess):
	set_cmd_text_color(FOREGROUND_DARKPINK)
	sys.stdout.write(mess)
	resetColor()

#暗黄色
#dark yellow
def printDarkYellow(mess):
	set_cmd_text_color(FOREGROUND_DARKYELLOW)
	sys.stdout.write(mess)
	resetColor()

#暗白色
#dark white
def printDarkWhite(mess):
	set_cmd_text_color(FOREGROUND_DARKWHITE)
	sys.stdout.write(mess)
	resetColor()

#暗灰色
#dark gray
def printDarkGray(mess):
	set_cmd_text_color(FOREGROUND_DARKGRAY)
	sys.stdout.write(mess)
	resetColor()

#蓝色
#blue
def printBlue(mess):
	set_cmd_text_color(FOREGROUND_BLUE)
	sys.stdout.write(mess)
	resetColor()

#绿色
#green
def printGreen(mess):
	set_cmd_text_color(FOREGROUND_GREEN)
	sys.stdout.write(mess)
	resetColor()

#天蓝色
#sky blue
def printSkyBlue(mess):
	set_cmd_text_color(FOREGROUND_SKYBLUE)
	sys.stdout.write(mess)
	resetColor()

#红色
#red
def printRed(mess):
	set_cmd_text_color(FOREGROUND_RED)
	sys.stdout.write(mess)
	resetColor()

#粉红色
#pink
def printPink(mess):
	set_cmd_text_color(FOREGROUND_PINK)
	sys.stdout.write(mess)
	resetColor()

#黄色
#yellow
def printYellow(mess):
	set_cmd_text_color(FOREGROUND_YELLOW)
	sys.stdout.write(mess)
	resetColor()

#白色
#white
def printWhite(mess):
	set_cmd_text_color(FOREGROUND_WHITE)
	sys.stdout.write(mess)
	resetColor()

##################################################

#白底黑字
#white bkground and black text
def printWhiteBlack(mess):
	set_cmd_text_color(FOREGROUND_BLACK | BACKGROUND_WHITE)
	sys.stdout.write(mess)
	resetColor()

#白底黑字
#white bkground and black text
def printWhiteBlack_2(mess):
	set_cmd_text_color(0xf0)
	sys.stdout.write(mess)
	resetColor()


#黄底蓝字
#white bkground and black text
def printYellowRed(mess):
	set_cmd_text_color(BACKGROUND_YELLOW | FOREGROUND_RED)
	sys.stdout.write(mess)
	resetColor()

LogFileName = "log.txt"
def func(xls_name):
	printGreen("xls_name :\t" + xls_name)
	path = os.getcwd()
	timeStamp = time.time()
	timeArray = time.localtime(timeStamp)
	StyleTime = time.strftime("%Y-%m-%d_%H%M%S", timeArray)
	ouputname = xls_name.split(".")[0] + "_" + StyleTime + "_out.xlsx"
	
	nowtime = datetime.now()
	timeLog = nowtime.strftime("%Y-%m-%d, %H:%M:%S")

	printGreen("执行中....")
	#错误列表
	fileLog = open(os.path.join(path, LogFileName),'a+')
	fileLog.write("本次时间：\t" + timeLog + "\n")
	fileLog.write(ouputname + "\n")
	error_list = []
	file_name = os.path.join(path, xls_name)
	sheet_names = openpyxl.load_workbook(file_name)
	for worksheet in sheet_names:
		print(worksheet)
		for row in worksheet:
			row_value = row[0].value
			if isinstance(row_value, int):
				continue
			elif isinstance(row_value, float):
				row_value = int(row_value)
				info = timeLog + "\t" + str(worksheet) + "\t转换格式 ：" + "\t" + str(row[0].coordinate) + "\t" + str(row[0].value) + " ==> " + str(row_value) + "\n"
				printGreen(info)
				fileLog.write(info)
			elif isinstance(row_value, str):
				info = timeLog + "\t" + str(worksheet) + "\t无效字符格式："  + "\t" + str(row[0].coordinate) + "\t" + row_value + "\n"
				printRed(info)
				fileLog.write(info)
				error_list.append(info)
				continue
			row[0].value = row_value
	sheet_names.save(filename = ouputname)
	
	#打印下错误的列表信息
	if error_list:
		printRed("=================================================================================================\n")
		printRed(xls_name + " 以下列表未成功转换，请检查\n")
		fileLog.write(xls_name + " 以下列表未成功转换，请检查" + "\n")
		for i in error_list:
			printRed(i)
			fileLog.write(i + "\n")
	fileLog.close()

def main():
	printGreen("转换开始\n")
	filenamelist = []
	path = os.getcwd()
	for file_data in os.listdir(path):
		if os.path.isfile(file_data):
			file_type = os.path.splitext(file_data)
			if file_type[1] in [".xlsx", ".xls"]:
				if "out" not in file_data:
					filenamelist.append(file_data);
	
	if os.path.exists(LogFileName): 
		os.remove(LogFileName)  
		
	for target_path in filenamelist:
		func(target_path)
	
	# pool = multiprocessing.Pool(processes = 4)
	# for i in filenamelist:
		# pool.apply_async(func, (i,))
	# pool.close()
	# pool.join()
	printGreen("转换结束\n")
	os.system('pause')
if __name__ == '__main__':
	main()