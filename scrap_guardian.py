import json
import requests
from os import makedirs
from os.path import join, exists
import datetime
from datetime import date, timedelta
import mysql.connector

mydb = mysql.connector.connect(user='test', password='test',
                              host='test',
                              database='test')


mycursor = mydb.cursor()

sql = "INSERT INTO guardian(id, type, sectionName, webPublicationDate, webURL ) VALUES (%s, %s, %s, %s, %s)"


directory = join('tempdata', 'articles')
makedirs(directory, exist_ok=True)

URL = 'http://content.guardianapis.com/search?q=Trudeau%20AND%20Justin%20OR%20(Trudeau%20AND%20(Canadian%2C%20Minister)%20AND%20NOT%20(Pierre))&'
my_parameters = {
    'from-date': "",
    'to-date': "",
    'order-by': "oldest",
    'page-size': 200,
    'api-key': 'b812f58e-691c-4f3e-ad35-164d0ea4ee6c'
}

#start_date = date(2018, 1, 1) #enable for initial load
#start_date = datetime.datetime.now().date()
start_date = date.today() - timedelta(1) #enable for daily updates
end_date = date.today() - timedelta(1)
day_range = range((end_date - start_date).days + 1)
for day_count in day_range:
    next_date = start_date + timedelta(days=day_count)
    date_str = next_date.strftime('%Y-%m-%d')
    file_name = join(directory, date_str + '.json')
    if not exists(file_name):
        print("Downloading", date_str)
        all_results = []
        my_parameters['from-date'] = date_str
        my_parameters['to-date'] = date_str
        current_page = 1
        total_pages = 1
        while current_page <= total_pages:
            my_parameters['page'] = current_page
            resp = requests.get(URL, my_parameters)
            data = resp.json()
            all_results.extend(data['response']['results'])
            current_page += 1
            total_pages = data['response']['pages']
            #print('[%s]' % ', '.join(map(str, all_results)))

            for row in all_results:
                values = (row['id'], row['type'], row['sectionName'], row['webPublicationDate'], row['webUrl'])
                mycursor.execute(sql, values)
                mydb.commit()
                print(mycursor.rowcount, "record inserted.")

        with open(file_name, 'w') as f:
            print("Writing to", file_name)
            f.write(json.dumps(all_results, indent=2))

mycursor.close()
mydb.close()
print("Data downloaded")
