# -*- encoding:utf-8 -*-
from math import radians, cos, sin, asin, sqrt
import urllib
import urllib2
import json
import xlrd
import xlwt
from xlutils.copy import copy
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import random
import thread
import time

#http://api.map.baidu.com/geocoder/v2/?address=%E4%B8%8A%E6%B5%B7%E8%8D%A3%E5%92%8C%E5%AE%B6%E5%9B%AD%E5%8D%97%E9%97%A8&output=json&ak=W93D8huOyugs9ixKoBQHQhVpQad8E5e9
output="json"
ak='W93D8huOyugs9ixKoBQHQhVpQad8E5e9'
startAddress=u'上海市杨浦区齐齐哈尔路536号'
endAddress=u'上海市长寿路200号'

def getLocationInfo(address):
    questURL='http://api.map.baidu.com/geocoder/v2/?address='+address+'&output='+output+'&ak='+ak
    request=urllib2.Request(questURL)
    request.add_header('User-Agent','Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36')
    response=urllib2.urlopen(request)
    decodeJson=json.loads(response.read())
    lon, lat=decodeJson['result']['location']['lng'],decodeJson['result']['location']['lat']
    return lon,lat

def getDistance(lon1, lat1, lon2, lat2):
    lon1, lat1, lon2, lat2 = map(radians,[lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # 地球平均半径，单位为公里
    return c * r

def readExcel(filePaht):
    data=xlrd.open_workbook(filePaht)
    table=data.sheet_by_index(0)
    wb=copy(data)
    ws=wb.get_sheet(0)
    colList=table.col_values(0)

   #workColSet=set(table.col_values(1))
   #workColDict={}
   #for setItem in workColSet:
   #    workColDict[setItem]={}
   #    workLon, workLat=getLocationInfo(setItem)   #遍历工作地点，获取所有工作地点的经纬度
   #    workColDict[setItem]['lon'], workColDict[setItem]['lat']=workLon,workLat
   #    print 'work address info is: lon'+str(workLon)+',lat:'+str(workLat)+', address='+setItem
#
   #for workItem in workColSet:
   #    for addkey in workColDict.keys():
   #        if addkey==workItem:
   #            addlon,addlat=workColDict[addkey]['lon'],workColDict[addkey]['lat']
   #        else:
   #            addlon,addlat=getLocationInfo(workItem)

    #lon1,lat1=addlon,addlat
    for i in range(len(colList)):
        print 'No:'+str(i)+'/'+str(len(colList)-1)+' is running...'
        try:
            lon,lat=getLocationInfo(str(colList[i]).encode('gbk'))
            #distance=getDistance(lon,lat,lon1,lat1)
            #ws.write(i,4,distance)
            ws.write(i,2,lon)
            ws.write(i,3,lat)
        except:
            ws.write(i,2,0)
            ws.write(i,3,0)
    wb.save(filePaht)
    print 'All work done!'

#print getDistance(121.53,31.26,121.44,31.25)
#print getLocationInfo('上海市杨浦区齐齐哈尔路536号')
#readExcel('./location2.xlsx')

def calculateDistance(lon1, lat1, lon2, lat2):
    lon1, lat1, lon2, lat2 = map(radians,[lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # 地球平均半径，单位为公里
    return c * r

def processFils(filepath):
    data=xlrd.open_workbook(filepath)
    table=data.sheet_by_index(0)
    wb=copy(data)
    ws=wb.get_sheet(0)
    lon1=table.col_values(2)
    lat1=table.col_values(3)
    lon2=table.col_values(4)
    lat2=table.col_values(5)
    for i in range(len(lon1)):
        distance=calculateDistance(lon1[i],lat1[i],lon2[i],lat2[i])
        ws.write(i,6,distance)
    wb.save(filepath)
    print 'done!'
processFils('./location1.xlsx')