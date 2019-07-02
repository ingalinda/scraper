import xlsxwriter
import mysql.connector
from datetime import date, timedelta
import datetime

start_date = date(2018, 1, 1)
end_date = date.today()
day_range_yesterday = range((end_date - start_date).days)
list_of_all_days=[]

for i in day_range_yesterday:
    #print(start_date + timedelta(days=i))
    list_of_all_days.append(start_date + timedelta(days=i))

#print(len(list_of_all_days))
#print(str(end_date))

mydb = mysql.connector.connect(user='test', password='test',
                              host='test',
                              database='test')

mycursor = mydb.cursor()

sql_data = "select distinct id, type, sectionName, substring(webPublicationDate,1,10) AS Datum from guardian"
mycursor.execute(sql_data)
clean_data = mycursor.fetchall()

workbook = xlsxwriter.Workbook('Guardian.xlsx')

bold = workbook.add_format({'bold': True})
format_date=workbook.add_format({'num_format':'yyyy-mm-dd'})
format_number= workbook.add_format({'bold': True, 'num_format': '0.00'})
format_white= workbook.add_format({'color': 'white'})

worksheet1 = workbook.add_worksheet("Daten")
worksheet1.write(0, 0, "Datum", bold)
worksheet1.write(0, 1, "Media Type", bold)
worksheet1.set_column('{0}:{0}'.format(chr(1 + ord('A'))), len("Media Type") + 2)
worksheet1.write(0, 2, "Section Name", bold)
worksheet1.set_column('{0}:{0}'.format(chr(2 + ord('A'))), len("Section Name") + 2)
worksheet1.write(0, 3, "ID", bold)

row = 1
col = 0


for id, type, section, datum in clean_data:
    worksheet1.write(row, col,     datum, format_date)
    worksheet1.write(row, col + 1, type)
    worksheet1.write(row, col + 2, section)
    worksheet1.write(row, col + 3, id)
    worksheet1.set_column('{0}:{0}'.format(chr(col + ord('A'))), len(str(datum)) + 2)
    row += 1

###########
sql = "select substring(webPublicationDate,1,10) AS Datum, count(distinct id) AS `Anzahl Artikel` from guardian where type='article' group by substring(webPublicationDate,1,10)"

mycursor.execute(sql)

myresult = mycursor.fetchall()

worksheet2 = workbook.add_worksheet("Anzahl_Artikel_pro_Tag")
worksheet2.write(0, 0,  "Datum", bold)
worksheet2.write(0, 1, "Anzahl Artikel", bold)
worksheet2.set_column('{0}:{0}'.format(chr(1 + ord('A'))), len("Anzahl Artikel") + 2)

row2 = 1
col2 = 0

for datum, anzahl in myresult:
    worksheet2.set_column('{0}:{0}'.format(chr(col2 + ord('A'))), len(str(datum)) + 2)
    worksheet2.write(row2, col2,     datetime.datetime.strptime(datum, "%Y-%m-%d").date() , format_date)
    worksheet2.write(row2, col2 + 1, anzahl)
    row2 += 1


worksheet2.write(row2, 0, 'Total', bold)
worksheet2.write(row2, 1, '=SUM(INDIRECT(ADDRESS(1,COLUMN())&":"&ADDRESS(ROW()-1,COLUMN())))', bold)
row2 += 1
worksheet2.write(row2, 0, 'Durchschnitt', bold)
worksheet2.write(row2, 1, '=AVERAGE(INDIRECT(ADDRESS(1,COLUMN())&":"&ADDRESS(ROW()-2,COLUMN())))', format_number)
row2 += 1

sql = "select count(distinct id) AS `Anzahl Artikel` from guardian"

mycursor.execute(sql)

Total_anzahl = mycursor.fetchall()
all_days= len(list_of_all_days)
for a in Total_anzahl:
    AVG_anzahl=a[0]/all_days

worksheet2.write(row2, 0, 'Durchschnitt Alle Tage', bold)
worksheet2.write(row2, 1, AVG_anzahl, format_number)
row2 += 1


################
worksheet3 = workbook.add_worksheet("Alle_Tage")
worksheet3.write(0, 0,  "Datum", bold)
worksheet3.set_column('{0}:{0}'.format(chr(1 + ord('A'))), len("Anzahl Artikel") + 2)
worksheet3.write(0, 1,  "Anzahl Artikel", bold)

row3 = 1
col3 = 0



for datum in list_of_all_days:
    worksheet3.set_column('{0}:{0}'.format(chr(col3 + ord('A'))), len(str(datum)) + 2)
    worksheet3.write(row3, col3, datum, format_date )
    worksheet3.write(row3, col3 + 1, '=IFERROR(VLOOKUP(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),Anzahl_Artikel_pro_Tag!A:B,2,0),0)')
    row3 += 1

row_number=row3 -1
worksheet3.write(row3, 0, 'Durchschnitt', bold)
worksheet3.write(row3, 1, '=AVERAGE(INDIRECT(ADDRESS(1,COLUMN())&":"&ADDRESS(ROW()-1,COLUMN())))', format_number)
row3 += 1

worksheet3.write(2, 4, 'Quartile 1', bold)
worksheet3.write(3, 4, 'Quartile 3', bold)
worksheet3.write(4, 4, 'IQR', bold)
worksheet3.write(5, 4, 'Upper fence', bold)
worksheet3.write(2, 5, '=QUARTILE(B2:B531,1)', bold)
worksheet3.write(3, 5, '=QUARTILE(B2:B531,3)', bold)
worksheet3.write(4, 5, '=F4-F3', bold)
worksheet3.write(5, 5, '=F4+(F5*1.5)', bold)
worksheet3.set_column('{0}:{0}'.format(chr(4 + ord('A'))), len("Upper fence") + 2)

worksheet3.write(1, 2, '=F4+(F5*1.5)',format_white)
worksheet3.write(int(row_number), 2, '=F4+(F5*1.5)',format_white)

chart_dates = workbook.add_chart({'type': 'column'})
chart_dates.add_series({
    'name': 'Anzahl Artikel pro Tag',
    'categories': ['Alle_Tage', 1, 0, int(row_number), 0],
    'values': ['Alle_Tage', 1, 1, int(row_number), 1],
    'line': {'color': 'black'},
    'fill': {'color': 'black'},
})

line_chart = workbook.add_chart({'type': 'line'})
line_chart.add_series({
    'categories': ['Alle_Tage', 1, 0, int(row_number), 0],
    'values': ['Alle_Tage', 1, 2, int(row_number), 2],
    'line': {'color': 'red'},
    'dash_type': 'dash'
})

# Combine the charts.
chart_dates.combine(line_chart)
chart_dates.show_blanks_as('span')

chart_dates.set_size({'x_scale': 2, 'y_scale': 1.5})
chart_dates.set_legend({'none': True})
chart_dates.set_chartarea({
    'border': {'none': True},
    'fill':   {'color': '#C0C0C0'}
})
chart_dates.set_plotarea({
    'border': {'none': True},
    'fill':   {'color': '#808080'}
})

worksheet3.insert_chart('D8', chart_dates)

###########
sql = "select sectionName, count(distinct id) AS `Anzahl Artikel` from guardian where type='article' group by sectionName"

mycursor.execute(sql)

section_data = mycursor.fetchall()
#print(len(section_data))
#print('[%s]' % ', '.join(map(str, section_data)))

section_data_sorted = [[x[0], int(x[1])] for x in section_data[0:]]
section_data_sorted.sort(key = lambda x : x[1], reverse = True)
#print(len(section_data_sorted))
#print(section_data_sorted)


worksheet4 = workbook.add_worksheet("Section")
worksheet4.write(0, 0,  "Section", bold)
worksheet4.set_column('{0}:{0}'.format(chr(0 + ord('A'))), len("Anzahl Artikel") + 2)
worksheet4.set_column('{0}:{0}'.format(chr(1 + ord('A'))), len("Anzahl Artikel") + 2)
worksheet4.write(0, 1,  "Anzahl Artikel", bold)
row4 = 1
col4 = 0


for sectionName, anzahl in section_data_sorted:
    worksheet4.write(row4, col4,     sectionName)
    worksheet4.write(row4, col4 + 1, anzahl)
    row4 += 1

row_number=str(len(section_data)+1)

chart = workbook.add_chart({'type': 'column'})
chart.add_series({
    'name': '=Section!$A$1',
    'categories': ['Section', 1, 0, int(row_number), 0],
    'values': ['Section', 1, 1, int(row_number), 1],
    'line': {'color': 'black'},
    'fill': {'color': 'black'},
})

chart.set_size({'x_scale': 2, 'y_scale': 1.5})
chart.set_legend({'none': True})
chart.set_chartarea({
    'border': {'none': True},
    'fill':   {'color': '#C0C0C0'}
})
chart.set_plotarea({
    'border': {'none': True},
    'fill':   {'color': '#808080'}
})

worksheet4.insert_chart('D2', chart)

workbook.close()
mydb.close()

print("xlsx file created")
