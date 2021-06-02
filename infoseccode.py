import glob
import pandas as pd
import pprint
import os
from datetime import date, timedelta
import statistics
import scipy.stats
import math
import xlrd
import xlwt
import datetime

#Creating a workbook and sheet to save output in the end
wbook = xlwt.Workbook()
outsheet = wbook.add_sheet("sheet1")

#Reading all Excel files from folder location
excel_files = glob.glob('abc/*.xlsx')

#Creating Lists of weeks with desired dates
week2 = ['11','12','13','14','15']
week1 = ['04','05','06','07','08']
weeks = [week1,week2] 

#Creating two different lists d and d1 with time interval starting from 8 am to 5 pm for all the dates in weeks
d=[]
d1=[]
def funcdelta(start, end, delta):
        curr = start
        while curr < end:
            yield curr
            curr += delta

s1 = datetime.datetime(2013,2, 4,8,0,0,0) 
e1 = datetime.datetime(2013,2, 15,17,0,0,0)
s2 = datetime.datetime(2013,2, 4,8,5,0,0)
e2 = datetime.datetime(2013,2, 15,17,5,0,0)    

for result in funcdelta(s1, e1, timedelta(seconds=300)):
    b=result.strftime("%Y-%m-%d %H:%M:%S.%f")
    d.append(b)
    
for result in funcdelta(s2,e2 , timedelta(seconds=300)):
    b=result.strftime("%Y-%m-%d %H:%M:%S.%f")
    d1.append(b)  

# Function which will do all the required operations of Data Cleaning and Checking for timestamp in given interval and give a list of average of oct/duration
#Function 1
def formatdata(csv,week,window,strttime,endtime):
    
    tdf=pd.DataFrame()
    timeStamp=[]
    #Reading columns with Real First Packet, doctets and Duration from excel and making a dataframe
    dd = pd.read_excel(csv,usecols = [4,5,9])
    
    # Calculating the internet usage per person by dividing octets/duration and adding it to the dataframe
    dd = dd[dd['Duration'] != 0]
    dd['Internet Usage'] = dd['doctets/dpkts']/dd['Duration']

    #Converting Epoch value of Real First Packet column to Timestamp    
    for i in dd['Real First Packet']:
        timeStamp.append(datetime.datetime.fromtimestamp((i) / 1000))
    dd['Time Stamp']=timeStamp
    flag=[]
    
    #Cleaning and keeping data of 2 weeks starting 4 to 15 excluding saturday sunday and time from 8am to 5pm data only, dropping rest of the data
    for i in dd['Time Stamp']:
        if i.strftime('%d') in week and (int(i.strftime('%H')) >=8) and (int(i.strftime('%H'))<17):
            flag.append(0)
        else:
            flag.append(1)
    dd['flag'] = flag

    dd.drop(dd[dd.flag > 0].index, inplace=True)
    dd.drop_duplicates(subset=['Time Stamp'], inplace=True)
    dd.sort_values(by=['Time Stamp'], inplace=True)
    dd.reset_index(drop=True,inplace=True)
  
    averages=[]
    flag =[]
    #New Dataframe with starttime and endtime columns with the required interval (10,227 and 300 secs)
    tdf['Startslot'] = strttime
    tdf['Endslot'] = endtime

    for i in tdf['Startslot']:
            val=datetime.datetime.strptime(i,"%Y-%m-%d %H:%M:%S.%f")
            if val.strftime('%d') in week and (int(val.strftime('%H')) >=8) and (int(val.strftime('%H'))<17):
                flag.append(0)
            else:
                flag.append(1)
    tdf['flag'] = flag
    tdf.drop(tdf[tdf.flag > 0].index, inplace=True)
    
    #Converting Dataframes to list to iterate
    starttime = pd.to_datetime(tdf['Startslot']).tolist()
    endtime = pd.to_datetime(tdf['Endslot']).tolist()
    time= pd.to_datetime(dd['Time Stamp']).tolist()
    usage=dd['Internet Usage'].tolist()
    duration = dd['Duration'].tolist()
    n=[]

    #Using Zip function to iterate parallely over columns of two dataframes dd and tdf
    for st,et in zip(starttime,endtime):
        for t, d, u in zip(time, duration, usage):
                if (t >= st and t<et and d <=300*1000):
                    n.append(u)
        if len(n) != 0:
                averages.append(statistics.mean(n))
        else:
                averages.append(0)
                
        n.clear() 
    return(averages)

#Calculates and returns Z value  
#Function 2
def calculateZ(r1a2a, r1a2b, r2a2b, N):

            if (r2a2b == 1):
                r2a2b = 0.99

            rm2 = ((r1a2a ** 2) + (r1a2b ** 2)) / 2
            a = (1 - r2a2b) / (2 * (1 - rm2))
            b = (1 - (a * rm2)) / (1 - rm2)

            z1a2a = 0.5 * (math.log10((1 + r1a2a) / (1 - r1a2a)))
            z1a2b = 0.5 * (math.log10((1 + r1a2b) / (1 - r1a2b)))

            if (r2a2b == 1):
                r2a2b = 0.99

            z = (z1a2a - z1a2b) * ((N - 3) ** 0.5) / (2 * (1 - r2a2b) * b)
            return z

# Calculates and returns P value
#Function 3
def calculateP(z):
            p = 0.3275911
            a1 = 0.254829592
            a2 = -0.284496736
            a3 = 1.421413741
            a4 = -1.453152027
            a5 = 1.061405429

            if z < 0.0:
                sign = -1
            else:
                sign = 1

            x = abs(z) / (2 ** 0.5)
            t = 1 / (1 + p * x)
            erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)
            
            return 0.5 * (1 + sign * erf)


#Looping over all the excel files , calculate Spearman’s correlation coefficient , calculate z and p values and saving to an output excel file
#Function calls here
for i in range(0,len(excel_files)):
    
    usera = excel_files[i]
    wk1a = formatdata(usera,weeks[0],300,d,d1)
    wk2a = formatdata(usera,weeks[1],300,d,d1)
    
    for j in range(0,len(excel_files)):
        
        userb = excel_files[j]
        print(usera,userb)
        
        wk2b = formatdata(userb,weeks[1],300,d,d1)
        
        #Spearman's calculation 
        r1a2a = scipy.stats.spearmanr(wk1a, wk2a)[0]
        r1a2b = scipy.stats.spearmanr(wk1a, wk2b)[0]
        r2a2b = scipy.stats.spearmanr(wk2a, wk2b)[0]
        
        #Length of List of average of oct/duration column
        N=len(wk1a)
        z = calculateZ(r1a2a, r1a2b, r2a2b, N)
        calculateP(z)
        
        outsheet.write(i, j, calculateP(z))

#Saving P values in the excel file.
wbook.save("outputofP.xls")

            
