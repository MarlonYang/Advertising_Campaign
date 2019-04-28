import time
import pandas as pd


#写Excel
def write_excel():
    try:
        pd_save = pd.DataFrame([sheet_head,], columns=sheet_head)

        for i in sku_info:

            campaign_name = '{0} {1}'.format(i[1], i[3])

            row1 = [campaign_name, campaign_budget, campaign_date, "", campaign_auto, "", "", "", "", "", campaign_status,'','','']
            row2 = [campaign_name,'','','','',i[0],i[2],'','','','',campaign_status,'','']
            row3 = [campaign_name,'','','','',i[0],'',i[0],'','','','',campaign_status,'']

            pd1 = pd.DataFrame([row1, row2, row3], columns = sheet_head)

            pd_save = pd.concat([pd_save, pd1])

        to_path = "/Users/yang/Downloads/{0} 新品广告.xls".format(station)

        pd_save.to_excel(excel_writer = to_path, header = None, index = None)

    except Exception as e:
        print("写入失败", e)


#根据站点获取信息填写方式
def station_values():

    global campaign_date, campaign_auto, campaign_status, sheet_head

    try:
        station_type = station[-2:]
        timeArray = time.localtime(time.time())

        station_info = {'US': {'Date': time.strftime("%Y/%m/%d", timeArray), 'Auto': 'Auto', 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group', 'Max Bid',
                                              'SKU', 'Keyword', 'Match Type', 'Campaign Status', 'Ad Group Status',
                                              'Status', 'Bid+']},
                        'CA': {'Date': time.strftime("%Y/%m/%d", timeArray), 'Auto': 'Auto', 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group', 'Max Bid',
                                              'SKU', 'Keyword', 'Match Type', 'Campaign Status', 'Ad Group Status',
                                              'Status', 'Bid+']},
                        'UK': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Auto', 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group Name',
                                              'Max Bid', 'SKU', 'Keyword', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bid+']},
                        'DE': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'automatisch', 'Status': 'aktiviert',
                               'title_name': ['Kampagne', 'Tagesbudget Kampagne', 'Startdatum der Kampagne',
                                              'Enddatum der Kampagne', 'Ausrichtungstyp der Kampagne', 'Anzeigengruppe',
                                              'Maximales Gebot', 'SKU', 'Keyword', 'Übereinstimmungstyp',
                                              'Kampagnenstatus', 'Anzeigengruppe Status', 'Status', 'gebot+']},
                        'FR': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Auto', 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group Name',
                                              'Max Bid', 'SKU', 'Keyword', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bid+']},
                        'IT': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Automatico', 'Status': 'attivo',
                               'title_name': ['Nome della campagna', 'Budget giornaliero campagna',
                                              'Data di inizio della campagna', 'Data di fine della campagna',
                                              'Tipo di targeting della campagna', 'Nome del gruppo di annunci',
                                              'Offerta massima', 'SKU', 'Parola chiave', 'Tipo di corrispondenza',
                                              'Stato della campagna', 'Stato del gruppo', 'Stato']},
                        'ES': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Auto', 'Status': 'Enabled',
                               'title_name': ['Campaign Name', 'Campaign Daily Budget', 'Campaign Start Date',
                                              'Campaign End Date', 'Campaign Targeting Type', 'Ad Group Name',
                                              'Max Bid', 'SKU', 'Keyword', 'Match Type', 'Campaign Status',
                                              'Ad Group Status', 'Status', 'Bid+']}}

        campaign_date = station_info[station_type]['Date']
        campaign_auto = station_info[station_type]['Auto']
        campaign_status = station_info[station_type]['Status']
        sheet_head = station_info[station_type]['title_name']

        write_excel()

    except Exception as e:
        print("获取失败", e)


#读批量广告中sku
def read_excel():

    global sku_info, station

    try:
        df = pd.read_excel(path, sheet_name = 0)
        station_set = set(map(lambda x: x[-6:], df['渠道来源']))
        for station in station_set:
            station_name = ''.join(['Amazon-Z01', station])
            sku_set = list(df.loc[df['渠道来源'] == station_name]['SellSKU'])

            sku_info = []
            for s in sku_set:
                asin = list(df.loc[df['SellSKU'] == s]['ASIN'])[0]
                default_bid = float(df.loc[df['SellSKU'] == s]['建议出价'])
                name = list(df.loc[df['SellSKU'] == s]['产品中文名称'])[0]

                sku_info.append([s, asin, default_bid, name])

            station_values()

    except Exception as e:
        print("读取失败", e)


def main():

    global path, campaign_budget

    path = r"/Users/yang/Downloads/test1.xls"
    campaign_budget = 5

    read_excel()


if __name__ == '__main__':
    main()