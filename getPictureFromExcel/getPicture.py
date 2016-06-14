# -*- coding: utf-8 -*- 
import xlrd
import os,urllib

def main():
	# get Url from Excel
	try:
		data = xlrd.open_workbook("test.xlsx") # Open excel file
	except:
		print("No file in directory")
		return
	table = data.sheet_by_index(0) # Get sheet1
	nrows = table.nrows 			# number of rows	
	prefix = "http://open.smart-nt.com/image/get/"
	for i in range(nrows):
		suffix = table.cell(i, 1).value
		fileName = table.cell(i, 0).value
		image_url = prefix + suffix
		getAndSaveImg(image_url, fileName)
	
def createFileWithFileName(localPathParam,fileName):  
    totalPath=localPathParam+'\\'+fileName  
    if not os.path.exists(totalPath):  
        file=open(totalPath,'a+')  
        file.close()  
        return totalPath 
        
def getAndSaveImg(imgUrl, fileNameParam):  
    if( len(imgUrl)!= 0 ):  
        fileName=fileNameParam+'.jpg'  
        urllib.urlretrieve(imgUrl,createFileWithFileName(localPath,fileName))  

localPath = "F:\myRepos\Miscellaneous\getPictureFromExcel\images"
if __name__ == "__main__":
	main()