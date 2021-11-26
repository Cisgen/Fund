# -*- coding:utf-8 -*-
import os,time
import shutil
import json
import ctypes
import sys
import multiprocessing
import openpyxl
from datetime import datetime


LogFileName = "log.txt"
def func(xls_name):
	print("xls_name :", xls_name)
	path = os.getcwd()
	timeStamp = time.time()
	timeArray = time.localtime(timeStamp)
	StyleTime = time.strftime("%Y-%m-%d_%H%M%S", timeArray)
	ouputname = xls_name.split(".")[0] + "_" + StyleTime + "_out.xlsx"
	
	nowtime = datetime.now()
	timeLog = nowtime.strftime("%Y-%m-%d, %H:%M:%S")

	print("执行中....")
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
				info = timeLog + "\t" + str(worksheet) + "\t转换格式 ：" + "\t" + str(row[0].coordinate) + "\t" + str(row[0].value) + " ==> " + str(row_value)
				print(info)
				fileLog.write(info + "\n")
			elif isinstance(row_value, str):
				info = timeLog + "\t" + str(worksheet) + "\t无效字符格式："  + "\t" + str(row[0].coordinate) + "\t" + row_value
				print(info)
				fileLog.write(info + "\n")
				error_list.append(info)
				continue
			row[0].value = row_value
	sheet_names.save(filename = ouputname)
	
	#打印下错误的列表信息
	if error_list:
		print("=======================================================")
		print(xls_name + " 以下列表未成功转换，请检查")
		fileLog.write(xls_name + " 以下列表未成功转换，请检查" + "\n")
		for i in error_list:
			print(i)
			fileLog.write(i + "\n")
	fileLog.close()

def main():
	print("主进程开始")
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
	if filenamelist:
		print("主进程结束....")
		time.sleep(3)
	else :
		print("当前没有需要转换的文件...")
		time.sleep(3)
if __name__ == '__main__':
	main()