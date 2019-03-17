import xlwt, time, xlrd


def main():
    path = r"/Users/yang/文件/000 - 资料/Python/test/test1.xlsx"
    campaign_budget = 200
    default_bid = 0.02

    read_excel(path, campaign_budget, default_bid)


#读批量广告中sku
def read_excel(path, campaign_budget, default_bid):

    try:
        campaign_excel = xlrd.open_workbook(path)

        for i in range(len(campaign_excel.sheets())):
            station = campaign_excel.sheets()[i].name
            sku = campaign_excel.sheets()[i].col_values(0, 0)
            sku_len = len(sku)*2
            station_values(campaign_budget, default_bid, station, sku, sku_len)

    except Exception as e:
        print("读取失败", e)


#根据站点获取信息填写方式
def station_values(campaign_budget, default_bid, station, sku, sku_len):

    try:
        station_type = station[-2:]
        today_date = time.time()
        timeArray = time.localtime(today_date)
        campaign_name_date = time.strftime("%Y%m%d", timeArray)

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
                        'FR': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Automatique', 'Status': 'Activé',
                               'title_name': ['Nom de la Campagne', 'Budget quotidien de la Campagne',
                                              'Date de début de la Campagne', 'Date de fin de la Campagne',
                                              'Type de Ciblage de la Campagne', "Nom du groupe d'annonces",
                                              'Enchère Max', 'SKU', 'Mot-clé', 'Type de correspondance',
                                              'Statut de la campagne', 'Statut du groupe d’annonces', 'Statut']},
                        'IT': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Automatico', 'Status': 'attivo',
                               'title_name': ['Nome della campagna', 'Budget giornaliero campagna',
                                              'Data di inizio della campagna', 'Data di fine della campagna',
                                              'Tipo di targeting della campagna', 'Nome del gruppo di annunci',
                                              'Offerta massima', 'SKU', 'Parola chiave', 'Tipo di corrispondenza',
                                              'Stato della campagna', 'Stato del gruppo', 'Stato']},
                        'ES': {'Date': time.strftime("%d/%m/%Y", timeArray), 'Auto': 'Automático', 'Status': 'Habilitado',
                               'title_name': ['Nombre de la campaña', 'Presupuesto diario de la campaña',
                                              'Fecha de inicio de la campaña', 'Fecha de finalización de la campaña',
                                              'tipo de segmentación de la campaña', 'Grupo de anuncios', 'Puja Máxima',
                                              'SKU', 'Palabra clave', 'Tipo de concordancia', 'Estado de la campaña',
                                              'Estado del grupo de anuncios', 'Estado']}}

        campaign_date = station_info[station_type]['Date']
        campaign_auto = station_info[station_type]['Auto']
        campaign_status = station_info[station_type]['Status']
        row0 = station_info[station_type]['title_name']


        write_excel(station, sku, sku_len, campaign_name_date, campaign_date, campaign_auto, campaign_status, row0, campaign_budget, default_bid)

    except Exception as e:
        print("获取失败", e)


#写Excel
def write_excel(station, sku, sku_len, campaign_name_date, campaign_date, campaign_auto, campaign_status, row0, campaign_budget, default_bid):

    try:
        f = xlwt.Workbook()
        sheet1 = f.add_sheet('Template - Sponsored Product', cell_overwrite_ok=True)
        campaign_name = '{0} {1}'.format(station, campaign_name_date)
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
