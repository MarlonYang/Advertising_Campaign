# -*- coding:UTF-8 -*- 
import pandas as pd
import os, time

def file_name():

    try:
        L=[]
        for root, dirs, files in os.walk(path):
            for file in files:
                if os.path.splitext(file)[1] == '.csv':
                    L.append(os.path.join(root, file))
        return L

    except Exception as e:
        print("获取各文件路径失败", e)

def read_data():

    try:
        L = file_name()
        df_exchange = pd.read_excel(transform_path, sheet_name='Sheet2')
        upload_file = pd.read_excel(final_path, sheet_name='Sheet1')
        
        for l in L:
            df = pd.read_csv(l)
            df["Group"] = l[-18]
            df["Country"] = l[-12:-10]
            df["Station"] = l[-16:-10]
            df["Month"] = l[-9:-4]

            if df.shape[1] > 17:
                cols=[x for i,x in enumerate(df.columns) if i in [10,12,14,16]]
                df = df.drop(cols, axis = 1)
                
            print(df["Station"][0], df["Month"][0], df.shape)
            
            for i in ["Sessions",'Page Views',"Units Ordered","Total Order Items"]:
                if df[i].dtype == "object":
                    df[i] = df[i].apply(lambda x: "".join(x.split(','))).astype('int64')
            
            upload_file = pd.concat([upload_file, df], axis = 0)

        upload_file = pd.merge(upload_file,df_exchange, on = 'Country', how = 'left')
        upload_file = upload_file.drop(columns = 'Title') 
        
        for i in ['Can$','€','$','£',',']:
            upload_file["Ordered Product Sales"] = upload_file["Ordered Product Sales"].apply(lambda x: "".join(x.split(i)))
            
#         for i in ["Ordered Product Sales", "Month"]:
#             upload_file[i] = upload_file[i].astype('float64')
        upload_file["Ordered Product Sales"] = upload_file["Ordered Product Sales"].astype('float64')
            
        for i in ['Session Percentage', 'Page Views Percentage', 'Buy Box Percentage', 'Unit Session Percentage']:
            upload_file[i] = upload_file[i].apply(lambda x: "".join(x.split(',')))
            upload_file[i] = upload_file[i].apply(lambda x: "".join(x.split('%'))).astype('float64') / 100

        upload_file['Sales$'] = upload_file['Ordered Product Sales'] * upload_file['Price']
        final_df = upload_file.drop(columns = 'Price')
        print(final_df.shape)
        final_df.to_excel(excel_writer = upload_path, index = None)

    except Exception as e:
        print("读取数据失败", e)
        
    else:
        print("文件已创建，可上传")

def main():

    global path, upload_path, transform_path, final_path    
    
    path = r"D:\Marlon\Summary\Orders\201911"
    today_date = time.strftime("%Y%m%d", time.localtime(time.time()))
    upload_path = r"D:\Marlon\Summary\Orders\upload\upload {}.xlsx".format(today_date)
    
    
    transform_path = r"D:\Marlon\Summary\关键词\exchange.xlsx"
    final_path = r"D:\Marlon\Summary\Orders\final.xlsx"

    read_data()

if __name__ == '__main__':
    main()