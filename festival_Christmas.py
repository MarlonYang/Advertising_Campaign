import xlwt, time, xlrd
import pandas as pd


def main():
    path = r"D:\Marlon\上传模板\周四批量\20191128\新建文件夹\圣诞FBM.xlsx"
    campaign_budget = 200

    read_excel(path, campaign_budget)


#读批量广告中sku
def read_excel(path, campaign_budget):

    try:

        df = pd.read_excel(path, sheet_name = 0)
        station_set = set(map(lambda x: x[-6:], df['渠道来源']))
        for station in station_set:
            station_name = ''.join(['Amazon-Z01', station])
            sku = list(df.loc[df['渠道来源'] == station_name]['SellSKU'])
            sku_len = len(sku) * 2
            station_values(campaign_budget, station, sku, sku_len)

    except Exception as e:
        print("读取失败", e)


#根据站点获取信息填写方式
def station_values(campaign_budget, station, sku, sku_len):

    try:
        station_type = station[-2:]
        today_date = time.time()
        timeArray = time.localtime(today_date)
        campaign_name_date = time.strftime("%Y%m%d", timeArray)

        station_info = {'US': {'Date': '09/12/2019', 'Auto': 'Auto', 'default_bid':0.08, 'Status': 'Enabled',
                               'title_name': ['Campaign', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group', 'Max Bid',
                                              'SKU', 'Keyword or Product Targeting', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bidding strategy']},
                        'CA': {'Date': '09/12/2019', 'Auto': 'Auto', 'default_bid':0.08, 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group', 'Max Bid',
                                              'SKU', 'Keyword or Product Targeting', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bidding strategy']},
                        'UK': {'Date': '12/09/2019', 'Auto': 'Auto', 'default_bid':0.06, 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group Name',
                                              'Max Bid', 'SKU', 'Keyword or Product Targeting', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bid+']},
                        'DE': {'Date': '12/09/2019', 'Auto': 'automatisch', 'default_bid':0.06, 'Status': 'aktiviert',
                               'title_name': ['Kampagne', 'Tagesbudget Kampagne', 'Startdatum der Kampagne',
                                              'Enddatum der Kampagne', 'Ausrichtungstyp der Kampagne', 'Anzeigengruppe',
                                              'Maximales Gebot', 'SKU', 'Schlüsselwort- oder Produktausrichtung',
                                              'übereinstimmungstyp', 'Kampagnenstatus', 'Anzeigengruppe Status', 'Status',
                                              'gebot+']},
                        'FR': {'Date': '12/09/2019', 'Auto': 'Automatique', 'default_bid':0.04, 'Status': 'Activé',
                               'title_name': ['Nom de la Campagne', 'Budget quotidien de la Campagne', 'Date de début de la Campagne',
                                              'Date de fin de la Campagne', 'Type de Ciblage de la Campagne',
                                              "Nom du groupe d'annonces", 'Enchère Max', 'SKU',
                                              'Ciblage de mots-clés ou de produits', 'Type de correspondance',
                                              'Statut de la campagne', 'Statut du groupe d’annonces', 'Statut', '']},
                        'IT': {'Date': '12/09/2019', 'Auto': 'Automatico', 'default_bid':0.04, 'Status': 'attivo',
                               'title_name': ['Nome della campagna', 'Budget giornaliero campagna',
                                              'Data di inizio della campagna', 'Data di fine della campagna',
                                              'Tipo di targeting della campagna', 'Nome del gruppo di annunci',
                                              'Offerta massima', 'SKU', 'Targeting per parola chiave o prodotto',
                                              'Tipo di corrispondenza', 'Stato della campagna', 'Stato del gruppo', 'Stato']},
                        'ES': {'Date': '12/09/2019', 'Auto': 'Auto', 'default_bid':0.04, 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group Name',
                                              'Max Bid', 'SKU', 'Keyword or Product Targeting', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bid+']}}

        campaign_date = station_info[station_type]['Date']
        campaign_auto = station_info[station_type]['Auto']
        campaign_status = station_info[station_type]['Status']
        row0 = station_info[station_type]['title_name']
        default_bid = station_info[station_type]['default_bid']


        write_excel(station, sku, sku_len, campaign_name_date, campaign_date, campaign_auto, campaign_status, row0, campaign_budget, default_bid)

    except Exception as e:
        print("获取失败", e)


#写Excel
def write_excel(station, sku, sku_len, campaign_name_date, campaign_date, campaign_auto, campaign_status, row0, campaign_budget, default_bid):

    try:
        f = xlwt.Workbook()
        sheet1 = f.add_sheet('Template - Sponsored Product', cell_overwrite_ok=True)
        campaign_name = '{0}圣诞节'.format(station)
        row1 = [campaign_name, campaign_budget, campaign_date, "", campaign_auto, "",  "",  "",  "",  "", campaign_status]
        excel_name = '{0}.xls'.format(campaign_name)

        # 写第一行
        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i])

        #写第二行
        for i in range(0, len(row1)):
            sheet1.write(1, i, row1[i])

        #写A列
        for i in range(0,sku_len):
            sheet1.write(i+2, 0, campaign_name)

        #写F列 - 1
        for i in range(0,len(sku)):
            sheet1.write(i+2, 5, sku[i])

        #写F列 - 2
        for i in range(0,len(sku)):
            sheet1.write(i+len(sku)+2, 5, sku[i])

        #写G列
        for i in range(0,len(sku)):
            sheet1.write(i+2, 6, default_bid)

        #写H列
        for i in range(0,len(sku)):
            sheet1.write(i+len(sku)+2, 7, sku[i])

        #写L列
        for i in range(0,len(sku)):
            sheet1.write(i+2, 11, campaign_status)

        #写M列
        for i in range(0,len(sku)):
            sheet1.write(i+len(sku)+2, 12, campaign_status)


        f.save(excel_name)

    except Exception as e:
        print("写入失败", e)


if __name__ == '__main__':
    main()