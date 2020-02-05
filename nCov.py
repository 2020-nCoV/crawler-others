import json
import os
import re
import pandas as pd
from openpyxl import load_workbook


fileList = os.listdir('C:/Users/DELL/Documents/WeChat Files/yhb_1392895092/FileStorage/File/2020-02/stand')

regions = []
districts = []

for file in fileList:
	c = re.match('(.*?)-point', file)
	regions.append(c.group(1))
	#print(file)

	with open('C:/Users/DELL/Documents/WeChat Files/yhb_1392895092/FileStorage/File/2020-02/stand/{0}'.format(file), 'r', encoding='utf-8') as fp:
		json_data = json.load(fp)
		#print(json_data)
		for qu in json_data['features']:
			districts.append(qu['properties']['name'])

print(len(regions))
print(len(districts))
a = regions + districts

df = pd.read_excel('C:/Users/DELL/Documents/WeChat Files/yhb_1392895092/FileStorage/File/2020-02/Wuhan_nCoV汇总-2.2零时.xlsx', None)
#df[['省份', '地级行政单位', '县级行政单位']]
#print(len(df['省份']))
writer = pd.ExcelWriter('D:/new_Wuhan_nCoV0202.xlsx', engin='openpyxl')
for sheet_name in df.keys():
    if sheet_name == '省地级累计' or sheet_name == '消息来源参考' or sheet_name == '各地医务工作者驰援武汉信息':
        df[sheet_name].to_excel(excel_writer = writer, sheet_name=sheet_name, encoding="utf-8", index = False)
        print(sheet_name, '导入完毕！')
        
    else:
        for i in range(len(df[sheet_name]['省份'])):
            if pd.isnull(df[sheet_name]['省份'][i]):
                continue
            else:
                for region in regions:
                    if df[sheet_name]['省份'][i][0:2] in region:
                    	if df[sheet_name]['省份'][i][0:2] == '吉林':
                    		df[sheet_name]['省份'][i] = '吉林省'
                    	else:
                    		df[sheet_name]['省份'][i] = region
                    	break
                        
        for i in range(len(df[sheet_name]['地级行政单位'])):  
            if pd.isnull(df[sheet_name]['地级行政单位'][i]):
                continue
            else:
            	for region in a:
                    if df[sheet_name]['地级行政单位'][i][0:2] in region:
                        df[sheet_name]['地级行政单位'][i] = region
                        break
                        
        for i in range(len(df[sheet_name]['县级行政单位'])):
            if pd.isnull(df[sheet_name]['县级行政单位'][i]):
                continue
            else:
                for district in districts:
                    if df[sheet_name]['县级行政单位'][i][0:2] in district:
                        df[sheet_name]['县级行政单位'][i] = district
                        break
        df[sheet_name].to_excel(excel_writer = writer, sheet_name=sheet_name, encoding="utf-8", index = False)
        print(sheet_name, '导入完毕！')

writer.save()
writer.close()
        
    
