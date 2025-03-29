import requests
import xlwt
from xlwt import Workbook
from cities import cities
import concurrent.futures
from tqdm import tqdm
from datetime import datetime

USER_AGENT='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36'
REQUEST_HEADER= {
    'User-Agent':USER_AGENT,
    'Accept-Language':'en-US, en; q=0.5',
}
NO_OF_THREADS=10
def get_weather_data(city):
    URL=f'http://api.weatherapi.com/v1/current.json?key=3cd9d6bc52f042fb935133741240708&q={city}&aqi=yes'
    res=requests.get(url=URL,headers=REQUEST_HEADER)
    return res.json()

def get_cities_weather(city,data):
    weather_detail=get_weather_data(city)
    filtered_data = {
    "Name": weather_detail["location"]["name"],
    "Region": weather_detail["location"]["region"],
    "Country": weather_detail["location"]["country"],
    "Local Time": weather_detail["location"]["localtime"],
    "temp Celsius": f"{weather_detail["current"]["temp_c"]}°C",
    "Temp Fahrenheit": f"{weather_detail["current"]["temp_f"]}°F",
    "Sun Status": "Day" if weather_detail["current"]["is_day"]==1 else "Night",
    "Sky Status": weather_detail["current"]["condition"]["text"],
    "Humidity": f"{weather_detail["current"]["humidity"]}%",
    "Cloud": f"{weather_detail["current"]["cloud"]}%"
    }
    data.append(filtered_data)

def output_file_xls(data):
    wb=Workbook()
    excel_sheet=wb.add_sheet("Cities Weather Info")
    bold_style = xlwt.easyxf('font: bold 1')
    headers=list(data[0].keys())
    for i in range(0,len(headers)):
        excel_sheet.write(0,i,headers[i], bold_style)
    for i in range(0,len(data)):
        city=data[i]
        values=list(city.values())
        for x in range(0,len(values)):
            excel_sheet.write(i+1,x,values[x])
    output_file_name=f"weathers-{datetime.today().strftime('%d-%m-%Y')}.xls"
    wb.save(output_file_name)
    print("File created Successfully.")


    
    


if __name__=='__main__':
    weather_data=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=NO_OF_THREADS) as executor:
        for city in tqdm(cities):
            executor.submit(get_cities_weather,city,weather_data)
    output_file_xls(weather_data)

    
