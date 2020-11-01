import json
import csv
import xlsxwriter
from datetime import datetime, timedelta
import dateutil.parser
import os
import shutil
import xlrd


def convertseconds(n):
    return str(timedelta(seconds = n))

def copy_template(filename):
    if not os.path.isfile(f'settings/{filename}.xlsx'):
        shutil.copy(f'settings/manual_template.xlsx', f'settings/{filename}.xlsx')
        print (filename)

def parse_manual_entries(filename):
    data = {}
    if os.path.isfile(f'settings/{filename}.xlsx'):
        # To open Workbook
        wb = xlrd.open_workbook(f'settings/{filename}.xlsx')
        sheet = wb.sheet_by_index(0)

        for i in range(sheet.nrows):
            data[sheet.row_values(i)[0]] = sheet.row_values(i)[1:]
    return data

def parse(filename):
    efter_coldata = {}
    fore_coldata = {}
    total_data = {}
    compare_bedtime = {}
    with open('settings/spelschema.csv', 'r', encoding='utf-8') as csvfile:
        spelschema = csv.reader(csvfile, delimiter=',')
        for row in spelschema:
            fore_coldata[row[0]] = {'match': row[1] if row[1] != 'Malmö' else row[2]}
            efter_coldata[row[0]] = {'match': row[1] if row[1] != 'Malmö' else row[2]}

    copy_template(filename)


    with open (os.path.join("players",f"{filename}.json")) as f:
        player_data = json.load(f)
       # print(json.dumps(player_data, indent=4))
        for i in player_data['sleep']:
            bedtime_start = dateutil.parser.parse(i['bedtime_start'])
            bedtime_end = dateutil.parser.parse(i['bedtime_end'])

            if str(datetime.date(bedtime_end)) in fore_coldata:
                fore_coldata[str(datetime.date(bedtime_end))]['tid'] = convertseconds(i['duration'])
                fore_coldata[str(datetime.date(bedtime_end))]['djup'] = convertseconds(i['deep'])
                fore_coldata[str(datetime.date(bedtime_end))]['rem'] = convertseconds(i['rem'])
                fore_coldata[str(datetime.date(bedtime_end))]['hypnogram_5min'] = (i['hypnogram_5min'])
                fore_coldata[str(datetime.date(bedtime_end))]['score'] = i['score']
                fore_coldata[str(datetime.date(bedtime_end))]['restfullness'] = i['score_disturbances']
                fore_coldata[str(datetime.date(bedtime_end))]['lägsta_puls'] = i['hr_lowest']
                fore_coldata[str(datetime.date(bedtime_end))]['SnittHRV'] = i['hr_average']
                fore_coldata[str(datetime.date(bedtime_end))]['temp'] = i['temperature_deviation']
                fore_coldata[str(datetime.date(bedtime_end))]['Utvilad'] = ''
                fore_coldata[str(datetime.date(bedtime_end))]['pigg_fräsch'] = ''
                fore_coldata[str(datetime.date(bedtime_end))]['samling'] = ''
                fore_coldata[str(datetime.date(bedtime_end))]['insats'] = ''
                fore_coldata[str(datetime.date(bedtime_end))]['övrigt'] = ''
                fore_coldata[str(datetime.date(bedtime_end))]['matchstart'] = ''

            #GÅR OCH LÄGGER SIG INOM SAMMA DYGN
            if str(datetime.date(bedtime_start)) in efter_coldata:
                efter_coldata[str(datetime.date(bedtime_start))]['tid'] = convertseconds(i['duration'])
                efter_coldata[str(datetime.date(bedtime_start))]['djup'] = convertseconds(i['deep'])
                efter_coldata[str(datetime.date(bedtime_start))]['rem'] = convertseconds(i['rem'])
                efter_coldata[str(datetime.date(bedtime_start))]['hypnogram_5min'] = (i['hypnogram_5min'])
                efter_coldata[str(datetime.date(bedtime_start))]['score'] = i['score']
                efter_coldata[str(datetime.date(bedtime_start))]['restfullness'] = i['score_disturbances']
                efter_coldata[str(datetime.date(bedtime_start))]['lägsta_puls'] = i['hr_lowest']
                efter_coldata[str(datetime.date(bedtime_start))]['SnittHRV'] = i['hr_average']
                efter_coldata[str(datetime.date(bedtime_start))]['temp'] = i['temperature_deviation']
                efter_coldata[str(datetime.date(bedtime_start))]['Utvilad'] = ''
                efter_coldata[str(datetime.date(bedtime_start))]['pigg_fräsch'] = ''
                efter_coldata[str(datetime.date(bedtime_start))]['samling'] = ''
                efter_coldata[str(datetime.date(bedtime_start))]['insats'] = ''
                efter_coldata[str(datetime.date(bedtime_start))]['övrigt'] = ''
                efter_coldata[str(datetime.date(bedtime_start))]['matchstart'] = ''

                # Går och lägger sig samma dygn som matchen spelades
                #print(bedtime_start)

        # Hämta dagar då läggtid är efter matchdag
        for k, v in efter_coldata.items():
            if len(v) != 16:
                compare_bedtime[k] = []

        # Hämta alla objekt efter matchdag kan vara 1 eller 2
        for i in player_data['sleep']:
            x = str (datetime.date(dateutil.parser.parse(i['bedtime_start'])) - timedelta(days=1))
            if x in compare_bedtime:
                if len (compare_bedtime[x]) == 0:
                    compare_bedtime[x] = [i]
                else: compare_bedtime[x].append(i)

        # Gå igenom objekten och uppdatera efter_coldata med korrekt läggningsdata
        for k,v  in compare_bedtime.items():
            if len(v) > 1:
                l_bedtime = dateutil.parser.parse(v[0]['bedtime_start'])
                r_bedtime = dateutil.parser.parse(v[1]['bedtime_start'])
                if l_bedtime > r_bedtime:
                    efter_coldata[k]['tid'] = convertseconds(v[1]['duration'])
                    efter_coldata[k]['djup'] = convertseconds(v[1]['deep'])
                    efter_coldata[k]['rem'] = convertseconds(v[1]['rem'])
                    efter_coldata[k]['hypnogram_5min'] = (v[1]['hypnogram_5min'])
                    efter_coldata[k]['score'] = v[1]['score']
                    efter_coldata[k]['restfullness'] = v[1]['score_disturbances']
                    efter_coldata[k]['lägsta_puls'] = v[1]['hr_lowest']
                    efter_coldata[k]['SnittHRV'] = v[1]['hr_average']
                    efter_coldata[k]['temp'] = v[1]['temperature_deviation']
                    efter_coldata[k]['Utvilad'] = ''
                    efter_coldata[k]['pigg_fräsch'] = ''
                    efter_coldata[k]['samling'] = ''
                    efter_coldata[k]['insats'] = ''
                    efter_coldata[k]['övrigt'] = ''
                    efter_coldata[k]['matchstart'] = ''


                else:
                    efter_coldata[k]['tid'] = convertseconds(v[0]['duration'])
                    efter_coldata[k]['djup'] = convertseconds(v[0]['deep'])
                    efter_coldata[k]['rem'] = convertseconds(v[0]['rem'])
                    efter_coldata[k]['hypnogram_5min'] = (v[0]['hypnogram_5min'])
                    efter_coldata[k]['score'] = v[0]['score']
                    efter_coldata[k]['restfullness'] = v[0]['score_disturbances']
                    efter_coldata[k]['lägsta_puls'] = v[0]['hr_lowest']
                    efter_coldata[k]['SnittHRV'] = v[0]['hr_average']
                    efter_coldata[k]['temp'] = v[0]['temperature_deviation']
                    efter_coldata[k]['Utvilad'] = ''
                    efter_coldata[k]['pigg_fräsch'] = ''
                    efter_coldata[k]['samling'] = ''
                    efter_coldata[k]['insats'] = ''
                    efter_coldata[k]['övrigt'] = ''
                    efter_coldata[k]['matchstart'] = ''
            elif len(v) == 0:
                pass
            else:
                # måste kolla om personen haft ring på t.ex. genom att bed_start och end är samma dygn
                # Se 2020-02-02T22:30:35+01:00
                bs = datetime.date(dateutil.parser.parse(v[0]['bedtime_start']))
                be = datetime.date(dateutil.parser.parse(v[0]['bedtime_end']))
                if bs != be:
                    pass


                else:
                    efter_coldata[k]['tid'] = convertseconds(v[0]['duration'])
                    efter_coldata[k]['djup'] = convertseconds(v[0]['deep'])
                    efter_coldata[k]['rem'] = convertseconds(v[0]['rem'])
                    efter_coldata[k]['hypnogram_5min'] = (i['hypnogram_5min'])
                    efter_coldata[k]['score'] = v[0]['score']
                    efter_coldata[k]['restfullness'] = v[0]['score_disturbances']
                    efter_coldata[k]['lägsta_puls'] = v[0]['hr_lowest']
                    efter_coldata[k]['SnittHRV'] = v[0]['hr_average']
                    efter_coldata[k]['temp'] = v[0]['temperature_deviation']
                    efter_coldata[k]['Utvilad'] = ''
                    efter_coldata[k]['pigg_fräsch'] = ''
                    efter_coldata[k]['samling'] = ''
                    efter_coldata[k]['insats'] = ''
                    efter_coldata[k]['övrigt'] = ''
                    efter_coldata[k]['matchstart'] = ''

    workbook = xlsxwriter.Workbook(f'{filename}-{datetime.now().date()}.xlsx')
    header_format = workbook.add_format({'bold': True, 'font_size':14})
    hour_format = workbook.add_format({'num_format': 'hh:mm'})
    int_format = workbook.add_format({'num_format': '0'})
    dec_format = workbook.add_format({'num_format': '0.0'})
    red = workbook.add_format({})
    sam_sheet = workbook.add_worksheet('Sammanfattning')
    fore_sheet = workbook.add_worksheet('Före')
    efter_sheet = workbook.add_worksheet('Efter')
    hypno_fore_data_sheet = workbook.add_worksheet('HypnoDataFöre')
    hypno_efter_data_sheet = workbook.add_worksheet('HypnoDataEfter')
    total_data_sheet = workbook.add_worksheet('TotalData')

    chart = workbook.add_chart({'type': 'radar'})
    chart.set_title ({'name': 'Sömn'})
    chart.set_x_axis({'label_position': 'none'})
    chart.set_y_axis({'label_position': 'none'})
    chart.add_series({
        'name': 'Efter',
        'values':     f'=Efter!C2:C{len(fore_coldata)}',
        'line':       {'color': 'green'},
    })

    chart.add_series({
        'name': 'Före',
        'values':     f'=Före!C2:C{len(fore_coldata)}',
        'line':       {'color': 'blue'},
    })

    sam_sheet.write('A1',filename,header_format)
    #sam_sheet.insert_chart('A15', chart)

    sam_sheet.write('A5', 'Alla', header_format)
    sam_sheet.write('A7', 'Före', header_format)
    sam_sheet.write('A9', 'Efter', header_format)
    sam_sheet.write('B3', 'Tid', header_format)
    sam_sheet.write('D3', 'Score', header_format)
    sam_sheet.write('F3', 'Djup', header_format)
    sam_sheet.write('H3', 'HRV', header_format)
    sam_sheet.write('J3', 'Temp', header_format)
    sam_sheet.write('L3', 'Utvilad', header_format)
    sam_sheet.write('N3', 'Pigg/Fräsh', header_format)
    sam_sheet.write('P3', 'Insats', header_format)

    sam_sheet.write_formula('B5', '=AVERAGE(Före!C2:C99,Efter!C2:C99)', hour_format) # tid
    sam_sheet.write_formula('B7', '=AVERAGE(Före!C2:C99)', hour_format)
    sam_sheet.write_formula('B9)', '=AVERAGE(Efter!C2:C99)', hour_format)

    sam_sheet.write_formula('D5', '=AVERAGE(Före!F2:F99,Efter!F2:F99)', int_format)  # Score
    sam_sheet.write_formula('D7', '=AVERAGE(Före!F2:F99)', int_format)
    sam_sheet.write_formula('D9', '=AVERAGE(Efter!F2:F99)', int_format)

    sam_sheet.write_formula('F5', '=AVERAGE(Före!D2:D99,Efter!D2:D99)', hour_format) # Djup
    sam_sheet.write_formula('F7', '=AVERAGE(Före!D2:D99)', hour_format)
    sam_sheet.write_formula('F9', '=AVERAGE(Efter!D2:D99)', hour_format)

    sam_sheet.write_formula('H5', '=AVERAGE(Före!I2:I99,Efter!I2:I99)', int_format)  # HRV
    sam_sheet.write_formula('H7', '=AVERAGE(Före!I2:I99)', int_format)
    sam_sheet.write_formula('H9', '=AVERAGE(Efter!I2:I99)', int_format)

    sam_sheet.write_formula('J5', '=AVERAGE(Före!J2:J99,Efter!J2:J99)', dec_format)  # Temp
    sam_sheet.write_formula('J7', '=AVERAGE(Före!J2:J99)', dec_format)
    sam_sheet.write_formula('J9', '=AVERAGE(Efter!J2:J99)', dec_format)

    sam_sheet.write_formula('L5', '=AVERAGE(Före!K2:K99,Efter!K2:K99)')  # Utvilad
    sam_sheet.write_formula('L7', '=AVERAGE(Före!K2:K99)')
    sam_sheet.write_formula('L9', '=AVERAGE(Efter!K2:K99)')

    sam_sheet.write_formula('N5', '=AVERAGE(Före!L2:L99,Efter!L2:L99)')  # Pigg
    sam_sheet.write_formula('N7', '=AVERAGE(Före!L2:L99)')
    sam_sheet.write_formula('N9', '=AVERAGE(Efter!L2:L99)')

    sam_sheet.write_formula('P5', '=AVERAGE(Före!N2:N99,Efter!N2:N99)')  # Insats
    sam_sheet.write_formula('P7', '=AVERAGE(Före!N2:N99)')
    sam_sheet.write_formula('P9', '=AVERAGE(Efter!N2:N99)')


    # SET COlUMN PROPS
    fore_sheet.set_column(0,0,10)
    fore_sheet.set_column(1,1,16)
    fore_sheet.set_column(6,7,14)
    fore_sheet.set_column(8,8,11)
    fore_sheet.set_column(11,11,14)
    fore_sheet.set_column(12,12,11)
    fore_sheet.set_column(14,14,30)
    fore_sheet.set_column(15,15,16)
    fore_sheet.set_column(16,16,30)
    # SET COlUMN PROPS
    efter_sheet.set_column(0,0,10)
    efter_sheet.set_column(1,1,16)
    efter_sheet.set_column(6,7,14)
    efter_sheet.set_column(8,8,11)
    efter_sheet.set_column(11,11,14)
    efter_sheet.set_column(12,12,11)
    efter_sheet.set_column(14,14,30)
    efter_sheet.set_column(15,15,16)
    efter_sheet.set_column(16,16,30)

    #SET HEADER
    fore_sheet.write(0,0,'Datum', header_format)
    fore_sheet.write(0,1,'Match', header_format)
    fore_sheet.write(0,2,'Tid', header_format)
    fore_sheet.write(0,3,'Djup', header_format)
    fore_sheet.write(0,4,'REM', header_format)
    fore_sheet.write(0,5,'Score', header_format)
    fore_sheet.write(0,6,'Restfulness', header_format)
    fore_sheet.write(0,7,'LägstaPuls', header_format)
    fore_sheet.write(0,8,'SnittHRV', header_format)
    fore_sheet.write(0,9,'Temp', header_format)
    fore_sheet.write(0,10,'Utvilad', header_format)
    fore_sheet.write(0,11,'Pigg/Fräsch', header_format)
    fore_sheet.write(0,12,'Samling', header_format)
    fore_sheet.write(0,13,'Insats', header_format)
    fore_sheet.write(0,14,'Övrigt', header_format)
    fore_sheet.write(0,15,'Matchstart', header_format)
    fore_sheet.write(0,16,'Hypnogram', header_format)

    efter_sheet.write(0,0,'Datum', header_format)
    efter_sheet.write(0,1,'Match', header_format)
    efter_sheet.write(0,2,'Tid', header_format)
    efter_sheet.write(0,3,'Djup', header_format)
    efter_sheet.write(0,4,'REM', header_format)
    efter_sheet.write(0,5,'Score', header_format)
    efter_sheet.write(0,6,'Restfulness', header_format)
    efter_sheet.write(0,7,'LägstaPuls', header_format)
    efter_sheet.write(0,8,'SnittHRV', header_format)
    efter_sheet.write(0,9,'Temp', header_format)
    efter_sheet.write(0,10,'Utvilad', header_format)
    efter_sheet.write(0,11,'Pigg/Fräsch', header_format)
    efter_sheet.write(0,12,'Samling', header_format)
    efter_sheet.write(0,13,'Insats', header_format)
    efter_sheet.write(0,14,'Övrigt', header_format)
    efter_sheet.write(0,15,'Matchstart', header_format)
    efter_sheet.write(0,16,'Hypnogram', header_format)

    total_data_sheet.write(0,0,'summary_date', header_format)
    total_data_sheet.write(0,1,'awake', header_format)
    total_data_sheet.write(0,2,'light', header_format)
    total_data_sheet.write(0,3,'deep', header_format)
    total_data_sheet.write(0,4,'rem', header_format)

    manual_entries = parse_manual_entries(filename)

    # FÖRE SHEET
    row = 1
    col = 0
    for k, v in fore_coldata.items():
        if len(v) == 16:
            fore_sheet.write(row, col, k)
            fore_sheet.write(row, col + 1, v['match'])
            fore_sheet.write(row, col + 2, datetime.strptime(v['tid'],"%H:%M:%S"), hour_format)
            fore_sheet.write(row, col + 3, datetime.strptime(v['djup'],"%H:%M:%S"), hour_format)
            fore_sheet.write(row, col + 4, datetime.strptime(v['rem'],"%H:%M:%S"), hour_format)
            fore_sheet.write(row, col + 5, v['score'])
            fore_sheet.write(row, col + 6, v['restfullness'])
            fore_sheet.write(row, col + 7, v['lägsta_puls'])
            fore_sheet.write(row, col + 8, v['SnittHRV'])
            fore_sheet.write(row, col + 9, v['temp'])
            fore_sheet.write(row, col + 10, manual_entries[k][1])
            fore_sheet.write(row, col + 11, manual_entries[k][2])
            fore_sheet.write(row, col + 12, manual_entries[k][3])
            fore_sheet.write(row, col + 13, manual_entries[k][4])
            fore_sheet.write(row, col + 14, manual_entries[k][5])
            fore_sheet.write(row, col + 15, manual_entries[k][6])

            fore_sheet.add_sparkline(row, col + 16, {'range': f'HypnoDataFöre!B{row}:FN{row}'})
            row += 1
        else:
            fore_sheet.write(row, col, k)
            fore_sheet.write(row, col + 1, v['match'])
            fore_sheet.write(row, col + 14, 'Data saknas')
            row += 1

    # EFTER SHEET
    row = 1
    col = 0
    for k, v in efter_coldata.items():
        if len(v) == 16:
            efter_sheet.write(row, col, k)
            efter_sheet.write(row, col + 1, v['match'])
            efter_sheet.write(row, col + 2, datetime.strptime(v['tid'],"%H:%M:%S"), hour_format)
            efter_sheet.write(row, col + 3, datetime.strptime(v['djup'],"%H:%M:%S"), hour_format)
            efter_sheet.write(row, col + 4, datetime.strptime(v['rem'],"%H:%M:%S"), hour_format)
            efter_sheet.write(row, col + 5, v['score'])
            efter_sheet.write(row, col + 6, v['restfullness'])
            efter_sheet.write(row, col + 7, v['lägsta_puls'])
            efter_sheet.write(row, col + 8, v['SnittHRV'])
            efter_sheet.write(row, col + 9, v['temp'])
            efter_sheet.write(row, col + 10, manual_entries[k][1])
            efter_sheet.write(row, col + 11, manual_entries[k][2])
            efter_sheet.write(row, col + 12, manual_entries[k][3])
            efter_sheet.write(row, col + 13, manual_entries[k][4])
            efter_sheet.write(row, col + 14, manual_entries[k][5])
            efter_sheet.write(row, col + 15, manual_entries[k][6])
            efter_sheet.add_sparkline(row, col + 16, {'range': f'HypnoDataEfter!B{row}:FN{row}'})
            row += 1
        else:
            efter_sheet.write(row, col, k)
            efter_sheet.write(row, col + 1, v['match'])
            efter_sheet.write(row, col + 14, 'Data saknas')
            row += 1

    def splitHypno(data):
        return [char for char in data]

    # HYPNOGRAM FÖRE
    r = 0
    for k, v in fore_coldata.items():
        if len(v) == 16:
            c = 0
            hypno_fore_data_sheet.write(r, c, k)
            for num in splitHypno(v['hypnogram_5min']):
                c += 1
                if num == "1":
                    num = -2
                elif num == "3":
                    num = -3
                elif num == "4":
                    num = 4
                else:
                    num = 0
                hypno_fore_data_sheet.write(r, c, num, int_format)

        r += 1
    hypno_fore_data_sheet.write(r, c, num, int_format)

    # HYPNOGRAM EFTER
    r = 0
    for k, v in efter_coldata.items():
        if len(v) == 16:
            c = 0
            hypno_efter_data_sheet.write(r, c, k)
            for num in splitHypno(v['hypnogram_5min']):
                c += 1
                if num == "1":
                    num = -2
                elif num == "3":
                    num = -3
                elif num == "4":
                    num = 4
                else:
                    num = 0

                hypno_efter_data_sheet.write(r, c, num, int_format)

        r += 1
    hypno_efter_data_sheet.write(r, c, num, int_format)
    row = 1
    col = 0


    with open (os.path.join('players', f"{filename}.json")) as f:
        player_data = json.load(f)

        for i in player_data['sleep']:
            total_data_sheet.write(row,col,i['summary_date'])
            total_data_sheet.write(row, col + 1, i['awake'])
            total_data_sheet.write(row, col + 2, i['light'])
            total_data_sheet.write(row, col + 3, i['deep'])
            total_data_sheet.write(row, col + 4, i['rem'])
            row += 1


    #hypno_fore_data_sheet.hide()
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        print(e, "Stäng all öppna exceldokument!")

for (dirpath, dirnames, filenames) in os.walk('players'):
    for f in filenames:
        if f.endswith('.json'):
            parse (os.path.splitext(f)[0])