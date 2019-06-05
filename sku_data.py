import pandas as pd
import os, time

def file_name():
    #获取路径下待上传文件名
    try:
        L=[]
        for root, dirs, files in os.walk(path):
            for file in files:
                if os.path.splitext(file)[1] == '.xlsx':
                    L.append(os.path.join(root, file))
        return L

    except Exception as e:
        print("获取各文件路径失败", e)
        
def previous_file_name():
    #获取路径下已上传过的文件名
    try:
        L=[]
        for root, dirs, files in os.walk(previous_path):
            for file in files:
                if os.path.splitext(file)[1] == '.xlsx':
                    L.append(os.path.join(root, file))
        return L

    except Exception as e:
        print("获取各文件路径失败", e)

def read_data():

    try:
        #待上传模板
        upload_file = pd.read_excel(final_path, sheet_name='Sheet1')
        #数据title
        upload_title = pd.read_excel(final_path, sheet_name='Sheet2')
        
        #新数据汇总
        L = file_name()
        for l in L:
            df = pd.read_excel(l)
            df["Group"] = l[-13:-12]
            df["Country"] = l[-7:-5]
            df["Station"] = l[-11:-5]
            #源数据替换title
            df.columns = list(upload_title.columns)
            #根据新数据所在列获取数据
            new_data = df[upload_file.columns]
            print("新数据：",l[-11:-5], new_data.shape)
            #根据带上传模板拼接每个新数据
            new_data = pd.concat([upload_file, new_data], axis = 0)
        print("新数据汇总：", new_data.shape)
        
        
        #已上传数据汇总
        previous_L = previous_file_name()
        for l in previous_L:
            df = pd.read_excel(l)
            print("已上传：", l[50:-5], df.shape)
            #拼接每个已上传
            previous_file = pd.concat([upload_file, df], axis = 0) 
        print("已上传汇总：", previous_file.shape)
        
        #拼接已上传和新数据
        upload = pd.concat([previous_file, new_data], axis = 0)
        #行数据去重（已上传数据中没有重复项，因此去重的是新数据中重复项）
        upload = upload.drop_duplicates()
        #获取新数据中非重复行
        upload = upload[previous_file.shape[0]:]
        print("非重复：", upload.shape)
        
        #非重复新数据生成excel上传
        upload.to_excel(excel_writer = upload_path, index = None)

    except Exception as e:
        print("读取数据失败", e)
        
    else:
        print("文件已创建，可上传")

def main():

    global path, transform_path, final_path, upload_path, previous_path

    #新数据路径
    path = r"C:\Users\Administrator\Desktop\Summary\SKU\201906"
    #模板路径
    final_path = r"C:\Users\Administrator\Desktop\Summary\SKU\final.xlsx"
    #已上传数据文件路径
    previous_path = r"C:\Users\Administrator\Desktop\Summary\SKU\upload"
    
    date = time.strftime("%Y%m%d", time.localtime(time.time()))
    #待上传数据路径
    upload_path = r"C:\Users\Administrator\Desktop\Summary\SKU\sku_upload {0} {1}.xlsx".format(path[-6:], date)


    read_data()
    

if __name__ == '__main__':
    main()