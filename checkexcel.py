# -*- coding:utf-8 -*-
import os,time
import shutil
import json
import ctypes
import sys
import multiprocessing
import openpyxl

def func(xls_name):
	print("xls_name :", xls_name)
	path = os.getcwd()
	timeStamp = time.time()
	timeArray = time.localtime(timeStamp)
	StyleTime = time.strftime("%Y-%m-%d_%H%M%S", timeArray)
	ouputname = StyleTime + "_out.xlsx"
	
	print("执行中....")
	#错误列表
	error_list = []
	file_name = os.path.join(path, xls_name)
	sheet_names = openpyxl.load_workbook(file_name)
	sheet = sheet_names.active
	for row in sheet:
		row_value = row[0].value
		if isinstance(row_value, int):
			continue
		elif isinstance(row_value, float):
			row_value = int(row_value)
		elif isinstance(row_value, str):
			error_list.append(str(row[0].coordinate)+"\t" + row_value)
			continue
		row[0].value = row_value
	sheet_names.save(filename = ouputname)
	
	#打印下错误的列表信息
	for i in error_list:
		print(i)

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
	
	for target_path in filenamelist:
		func(target_path)
	
	# pool = multiprocessing.Pool(processes = 4)
	# for i in filenamelist:
		# pool.apply_async(func, (i,))

	# pool.close()
	# pool.join()
	if filenamelist:
		print("主进程结束....")
		time.sleep(5)
	else :
		print("当前没有需要转换的文件...")
		time.sleep(2)
if __name__ == '__main__':
	main()