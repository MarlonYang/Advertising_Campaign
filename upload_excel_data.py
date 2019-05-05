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

        for l in L:
            country = l[-7:-5]
            station = l[-11:-5]
            group = l[-13:-12]
            df = pd.read_excel(l)
            df["Group"] = group
            df["Country"] = country
            df["Station"] = station

            new_df = pd.merge(df,df_week)
            new_df = pd.merge(new_df,df_exchange)
            new_df['Spend$'] = new_df['spend'] * new_df['Price']
            new_df['Sales$'] = new_df['sales'] * new_df['Price']

            final_df = new_df.drop(columns = 'Price')
            upload_file = pd.concat([upload_file, final_df], axis = 0)

        upload_file.to_excel(excel_writer = upload_path, index = None)

    except Exception as e:
        print("读取数据失败", e)

def main():

    global path, transform_path, final_path, upload_path

    path = r"/Users/yang/Downloads/20190428"
    transform_path = r"/Users/yang/Downloads/Data.xlsx"
    final_path = r"/Users/yang/Downloads/final.xlsx"
    upload_path = r"/Users/yang/Downloads/upload-{}.xlsx".format(path[-8:])

    read_data()

if __name__ == '__main__':
    main()