#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Nov 16 16:34:57 2018

@author: omkar
"""

import requests
import urllib.parse
import xlwt

files = ['dec1.xls','dec2.xls','dec3.xls']
places = ['16.232787, 73.672795','18.740016, 73.116334','19.216674, 72.981171']
keys = ['cd8744e1bcb54c79ada114432181911']
#url to fetch the data
for j in range(3):
    main_url = 'https://api.worldweatheronline.com/premium/v1/past-weather.ashx?key='+keys[0]+'&q='+places[j]+'&format=json&'
    date = ['2009','2010','2011','2012','2013','2014','2015','2016','2017']
    sheet1 = ['Sheet1','Sheet2','Sheet3','Sheet4','Sheet5','Sheet6','Sheet7','Sheet8','Sheet9']


    #create a book
    book = xlwt.Workbook(encoding ="utf-8")

    for m in range(9):
        variable_url = 'date='+date[m]+'-12-01&enddate='+date[m]+'-12-31&tp=24'
        url = main_url + variable_url

    #fetch the json data
        json_data = requests.get(url).json()

    #store the fetched data in variable.
        data = json_data

    #create different sheets for different years in same excel file.
        sheet = book.add_sheet(sheet1[m])
        sheet.write(0,0,"Date")
        sheet.write(0,1,"Max Temp")
        sheet.write(0,2,"Min Temp")
        sheet.write(0,3,"UV Index")
        sheet.write(0,4,"Precipitation")
        sheet.write(0,5,"Humidity")
        sheet.write(0,6,"Visibility")
        sheet.write(0,7,"Pressure")
        sheet.write(0,8,"CloudCover")
        sheet.write(0,9,"Dew Point")
        sheet.write(0,10,"Heat Index")
       

        print("Date" +"\t\t" + "Min Temp" +"\t"+ "Max Temp"+ "\t" +"UV Index"+ "\t" + "Precipitation"+"\t"+"Humdiidty"+"\t"+"Visibility" + "\t" + "Pressure" + "\t" + "Cloud Cover" + "\t" + "Dew Point" + "\t" + "Heat Index")
        for i in range(31):
            dates= data['data']['weather'][i]['date']
            maxtemp = data['data']['weather'][i]['maxtempC']
            mintemp = data['data']['weather'][i]['mintempC']
            uvindex = data['data']['weather'][i]['uvIndex']
            precip = data['data']['weather'][i]['hourly'][0]['precipMM']
            humidity = data['data']['weather'][i]['hourly'][0]['humidity']
            vis = data['data']['weather'][i]['hourly'][0]['visibility']
            pressure = data['data']['weather'][i]['hourly'][0]['pressure']
            cloud = data['data']['weather'][i]['hourly'][0]['cloudcover']
            dewpoint = data['data']['weather'][i]['hourly'][0]['DewPointC']
            heat = data['data']['weather'][i]['hourly'][0]['HeatIndexC']
            print(dates,"\t",maxtemp,"\t\t",mintemp,"\t\t",uvindex,"\t\t",precip,"\t\t",humidity,"\t\t",vis,"\t\t",pressure,"\t\t",cloud,"\t\t",dewpoint,"\t\t",heat)
            sheet.write(i+1,0,dates)
            sheet.write(i+1,1,maxtemp)
            sheet.write(i+1,2,mintemp)
            sheet.write(i+1,3,uvindex)
            sheet.write(i+1,4,precip)
            sheet.write(i+1,5,humidity)
            sheet.write(i+1,6,vis)
            sheet.write(i+1,7,pressure)
            sheet.write(i+1,8,cloud)
            sheet.write(i+1,9,dewpoint)
            sheet.write(i+1,10,heat)
            
        book.save("/home/omkar/"+files[j])
