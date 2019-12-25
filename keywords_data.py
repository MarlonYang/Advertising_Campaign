import numpy as np
import pandas as pd
import os

def file_name():

    try:
        L=[]
        for root, dirs, files in os.walk(path):
            for file in files:
                if os.path.splitext(file)[1] == '.xlsx':
                    L.append(os.path.join(root, file))
        return L

    except Exception as e:
        print("获取各文件路径失败", e)

def read_data():

    try:
        L = file_name()
        df_week = pd.read_excel(transform_path, sheet_name='Sheet1')
        df_exchange = pd.read_excel(transform_path, sheet_name='Sheet2')
        upload_file = pd.read_excel(final_path, sheet_name='Sheet1')
        
#         df = pd.read_excel(L[0])
#         df["Group"] = L[0][-13:-12]
#         df["Country"] = L[0][-7:-5]
#         df["Station"] = L[0][-11:-5]
#         df.columns = list(upload_file.columns)
#         upload_file = pd.concat([upload_file, df], axis = 0)
        m = 0
        n = 0
        for l in L:
            df = pd.read_excel(l)
            df["Group"] = l[-13:-12]
            df["Country"] = l[-7:-5]
            df["Station"] = l[-11:-5]
            df.columns = list(upload_file.columns)
            print(l[-11:-5], df.shape)
            upload_file = pd.concat([upload_file, df], axis = 0)
            n += 1
            m += df.shape[0]
        print(n, m)
        print(upload_file.shape)
            

        upload_file = pd.merge(upload_file,df_week)
        upload_file = pd.merge(upload_file,df_exchange, on = 'Country', how = 'left')
        print(upload_file.shape)
        
        
        upload_file['Spend$'] = upload_file['Spend'] * upload_file['Price']
        upload_file['Sales$'] = upload_file['Sales'] * upload_file['Price']
        for i in ['Sales', 'Orders', '7 Day Total Units (#)', '7 Day Advertised SKU Units (#)', '7 Day Other SKU Units (#)', '7 Day Advertised SKU Sales', '7 Day Other SKU Sales']:
            upload_file[i] = upload_file[i].replace(0,np.nan)
        final_df = upload_file.drop(columns = 'Price')
        print(final_df.shape)
        final_df.to_excel(excel_writer = upload_path, index = None)

    except Exception as e:
        print("读取数据失败", e)
        
    else:
        print("文件已创建，可上传")

def main():

    global path, transform_path, final_path, upload_path

    path = r"F:\Marlon\Summary\关键词\20191221\B"
    upload_path = r"F:\Marlon\Summary\关键词\upload\upload-{0}{1}.xlsx".format(path[-10:-2],path[-1])
    transform_path = r"F:\Marlon\Summary\关键词\exchange.xlsx"
    final_path = r"F:\Marlon\Summary\关键词\final.xlsx"

#     path = r"F:\Marlon\Summary\关键词\其他\20191224"
#     upload_path = r"F:\Marlon\Summary\关键词\upload\upload-{}.xlsx".format(path[-8:])

    read_data()
    

if __name__ == '__main__':
    main()
