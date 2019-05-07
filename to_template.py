import pandas as pd


def to_template():
    try:
        total_row = []
        
        for i in sku_info:
            row = ["美西一仓","上架存储","否",station_name,i[0],"","",i[1],i[2],i[3],"","","","","",""]
            total_row.append(row)
        
        final_df = pd.DataFrame(total_row, columns = US_df.columns)

        excel_name = "{0}\{1}.xlsx".format(to_path, station_name)

        final_df.to_excel(excel_writer = excel_name, sheet_name = "批量导入模板", index = None, )
    
    except Exception as e:
        print("写入失败", e)


def get_data():
    try:
        global US_df, sku_info, station_name
        
        df = pd.read_excel(data_path)
        
        US_df = pd.read_excel(US_template_path, sheet_name = '批量导入模板')
        US_df2 = pd.read_excel(US_template_path, sheet_name = '清单表')
        EU_df = pd.read_excel(EU_template_path, sheet_name = '批量导入模板')
        EU_df2 = pd.read_excel(EU_template_path, sheet_name = '清单表')
        
        station_set = set(map(lambda x: x[-6:], df['渠道来源']))
        
        for station in station_set:
            station_name = ''.join(['Amazon-Z01', station])
            sku_set = list(df.loc[df['渠道来源'] == station_name]['SellSKU'])

            sku_info = []

            for s in sku_set:
                FNSKU = list(df.loc[df['SellSKU'] == s]['FNSKU'])[0]
                ProductNumber = list(df.loc[df['SellSKU'] == s]['转仓数量确认'])[0]
                ProductSku = list(df.loc[df['SellSKU'] == s]['SKU'])[0]
                ProductName = list(df.loc[df['SellSKU'] == s]['产品英文名称'])[0]

                sku_info.append([FNSKU, ProductNumber, ProductSku, ProductName])
                
            to_template()

    except Exception as e:
        print("读取失败", e)


def main():

    global US_template_path, EU_template_path, data_path, to_path, expect_date

    US_template_path = r"C:\Users\Administrator\Desktop\上传模板\新建文件夹\1\美国退货换标模板.xlsx"
    EU_template_path = r"C:\Users\Administrator\Desktop\上传模板\新建文件夹\1\欧洲退货换标模板.xlsx"
    
    data_path = r"C:\Users\Administrator\Desktop\上传模板\新建文件夹\1\转仓数据.xlsx"
    
    to_path = r"C:\Users\Administrator\Desktop\上传模板\新建文件夹"
    
    expect_date = "2019/1/1"

    get_data()


if __name__ == '__main__':
    main()