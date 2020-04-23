# -*- coding: utf-8 -*-
"""
Created on Sun Dec 03 21:54:41 2017

@author: TUNTUNGAN
"""

from datetime import timedelta,datetime
import numpy as np
import re
import os.path
from scipy import signal
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

def step_removal_process(data,threshold,comp):
    #Smooth the data using a median filter to remove spikes and the Savitzky-Golay running average filter
    Binterp = data.interpolate(method = 'linear')
    medfiltB = signal.medfilt(Binterp.values[:], 3);
    vAvgResampB = signal.savgol_filter(medfiltB,17,1);
    
    diffB = np.diff(vAvgResampB)
    
    steps = [i for i,v in enumerate(abs(diffB)) if v > threshold]

    if np.array(steps).size:
        breakpoints = [i for i,v in enumerate(np.diff(steps)) if v > 1]
        print('Number of ',comp,' breakpoints: '+str(len(breakpoints)))

        #Go through each break point and sum up the accumulated change
        breakpoints =  [-1]+breakpoints+[len(steps)-1]  # put a zero at the start for each of looping and the final point
        for i in range(0,len(breakpoints)-1):
            st = breakpoints[i]+1; en =  breakpoints[i+1]
            #Extend the edges of the step detection back and forward a few
            #points and check you haven't gone over the edge of the time-series
            stp=steps[st]-2 if steps[st]>3 else steps[st]
            enp=steps[en]+2 if steps[en]<(len(diffB)-3) else steps[en]
            #Compute the change in the field across the step by cumulatively
            #summing up the differences from point-to-point
            change = np.cumsum(diffB[stp:enp+1])
            #Now remove it from the time series starting at the revelant point
            vAvgResampB[stp:enp+1] = vAvgResampB[stp:enp+1] - change
            vAvgResampB[enp+1:-1] = vAvgResampB[enp+1:-1] - change[-1]
            
    return vAvgResampB
          
def step_removal(date1,threshold,namefolderlemi,namefolder_out):           
    #Define data
    tahun1 = int((date1-timedelta(days=1)).year)
    bulan1 = int((date1-timedelta(days=1)).month)
    tanggal1 = int((date1-timedelta(days=1)).day)
    tahun2 = int(date1.year)
    bulan2 = int(date1.month)
    tanggal2 = int(date1.day)
    hh=np.zeros(1440)
    dd=np.zeros(1440)
    zz=np.zeros(1440)
    ff=np.zeros(1440)
    xx=np.zeros(1440)
    yy=np.zeros(1440)
    h_mean=np.zeros(1440)
    d_mean=np.zeros(1440)
    z_mean=np.zeros(1440)
    f_mean=np.zeros(1440)
    x_mean=np.zeros(1440)
    y_mean=np.zeros(1440)
    
    #Koefisien Gaussian filter
    gaussFilter=[0.00045933,0.00054772,0.00065055,0.00076964,0.00090693,
        0.00106449,0.00124449,0.00144918,0.00168089,0.00194194,0.00223468,
        0.0025614,0.0029243,0.00332543,0.00376666,0.00424959,0.00477552,
        0.00534535,0.00595955,0.00661811,0.00732042,0.0080653,0.0088509,
        0.00967467,0.01053338,0.01142303,0.01233892,0.01327563,0.01422707,
        0.01518651,0.01614667,0.01709976,0.01803763,0.01895183,0.01983377,
        0.0206748,0.02146643,0.02220039,0.02286881,0.02346437,0.0239804,
        0.02441104,0.02475132,0.02499727,0.02514602,0.0251958,0.02514602,
        0.02499727,0.02475132,0.02441104,0.0239804,0.02346437,0.02286881,
        0.02220039,0.02146643,0.0206748,0.01983377,0.01895183,0.01803763,
        0.01709976,0.01614667,0.01518651,0.01422707,0.01327563,0.01233892,
        0.01142303,0.01053338,0.00967467,0.0088509,0.0080653,0.00732042,
        0.00661811,0.00595955,0.00534535,0.00477552,0.00424959,0.00376666,
        0.00332543,0.0029243,0.0025614,0.00223468,0.00194194,0.00168089,
        0.00144918,0.00124449,0.00106449,0.00090693,0.00076964,0.00065055,
        0.00054772,0.00045933]
    
    #Parameter Stasiun
    with open('station.ini') as f_init:
        content_init = f_init.readlines()
        datasta = list(np.tile(np.nan,len(content_init)))		
        for i in range(0,len(content_init)):
            datasta[i] = re.split('\=|\n',content_init[i])
                
    #Import data Lemi   
    filename1 = '%4.0f\\%02.0f\\%4.0f %02.0f %02.0f 00 00 00.txt' % (tahun1,bulan1,tahun1,bulan1,tanggal1)
    filename2 = '%4.0f\\%02.0f\\%4.0f %02.0f %02.0f 00 00 00.txt' % (tahun2,bulan2,tahun2,bulan2,tanggal2)
    
    try:
#    if os.path.isfile(namefolderlemi + '/' + filename1)==True and os.path.isfile(namefolderlemi + '/' + filename2)==True:
        with open(namefolderlemi + '/' + filename1) as f_lemi1:
            content1 = f_lemi1.readlines()
            dates1 = list(np.tile(np.nan,len(content1)))
            for i in range(0,len(content1)):
                data1 = re.split('\s+',content1[i])
                Bx1 = float(data1[6]) #+ 41300
                By1 = float(data1[7])
                Bz1 = float(data1[8]) #- 7100
                dates1[i] = list((data1[0]+'-'+data1[1]+'-'+data1[2]+' '+data1[3]+':'+data1[4]+':'+data1[5],Bx1,By1,Bz1))
            datadates1 = pd.DataFrame(dates1,columns=['dt', 'Bx', 'By', 'Bz'])
    except:
        print('File '+filename1+' not exist')
        
    try:
        with open(namefolderlemi + '/' + filename2) as f_lemi2:
            content2 = f_lemi2.readlines()
            dates2 = list(np.tile(np.nan,len(content2)))
            for i in range(0,len(content2)):
                data2 = re.split('\s+',content2[i])
                Bx2 = float(data2[6]) #+ 41300
                By2 = float(data2[7])
                Bz2 = float(data2[8]) #- 7100
                dates2[i] = list((data2[0]+'-'+data2[1]+'-'+data2[2]+' '+data2[3]+':'+data2[4]+':'+data2[5],Bx2,By2,Bz2))
            datadates2 = pd.DataFrame(dates2,columns=['dt', 'Bx', 'By', 'Bz'])
    except:
        print('File '+filename2+' not exist'    )
    
    timefull = pd.date_range(date1-timedelta(days=1), periods=172800, freq='s')
    timeadd = pd.concat([datadates1,datadates2], ignore_index=True)
    timeadd.index = pd.DatetimeIndex(timeadd.dt)
    datafull = timeadd.reindex(timefull)

    #Now look for steps and remove them ...
    #Threshold is how large in nT/sample do you want to consider a step? it is important to set this correctly

    #Removing step 
    vAvgResampX = step_removal_process(datafull.Bx,threshold,'X')
    vAvgResampY = step_removal_process(datafull.By,threshold,'Y')
    vAvgResampZ = step_removal_process(datafull.Bz,threshold,'Z')

    #Smooth the data using a median filter to remove final spikes 
    finalX = signal.medfilt(vAvgResampX, 3)
    finalY = signal.medfilt(vAvgResampY, 3)
    finalZ = signal.medfilt(vAvgResampZ, 3)  

    finalF = np.sqrt((finalX**2)+(finalY**2)+(finalZ**2))
    finalH = np.sqrt(finalX**2+finalY**2)
    finalD = 60*np.degrees(np.arcsin(finalY/finalH))   
    
    #1 Minute mean value using gaussian filter
    gaussFilter=gaussFilter/np.sum(gaussFilter)
    for i in range(0,1440):
        H_conv=np.convolve(finalH[86354+(60*i):86445+(60*i)],gaussFilter)
        D_conv=np.convolve(finalD[86354+(60*i):86445+(60*i)],gaussFilter)
        Z_conv=np.convolve(finalZ[86354+(60*i):86445+(60*i)],gaussFilter)
        F_conv=np.convolve(finalF[86354+(60*i):86445+(60*i)],gaussFilter)
        X_conv=np.convolve(finalX[86354+(60*i):86445+(60*i)],gaussFilter)
        Y_conv=np.convolve(finalY[86354+(60*i):86445+(60*i)],gaussFilter)
        hh[i]=H_conv[90]
        dd[i]=D_conv[90]
        zz[i]=Z_conv[90]
        ff[i]=F_conv[90]
        xx[i]=X_conv[90]
        yy[i]=Y_conv[90] 

    #Hitung rata-rata satu menit
    for k in range(0,1440):
        if np.isnan(hh[k])==True:
            h_mean[k]=99999.00
        else:
            h_mean[k]=hh[k]
        if np.isnan(dd[k])==True:
            d_mean[k]=99999.00
        else:
            d_mean[k]=dd[k]
        if np.isnan(zz[k])==True:
            z_mean[k]=99999.00
        else:
            z_mean[k]=zz[k]
        if np.isnan(ff[k])==True:
            f_mean[k]=99999.00
        else:
            f_mean[k]=ff[k]
        if np.isnan(xx[k])==True:
            x_mean[k]=99999.00
        else:
            x_mean[k]=xx[k]
        if np.isnan(yy[k])==True:
            y_mean[k]=99999.00
        else:
            y_mean[k]=yy[k]

    #Menyimpan file output
    fileout = 'TUN'+str(date1.year)+str(date1.month).zfill(2)+str(date1.day).zfill(2)+'pmin'+'.min'
    namefolder = namefolder_out
    f_iaga = open(namefolder + '/' + fileout, 'w')
    
    f_iaga.write(' Format                 IAGA-2002                                    |\n')
    f_iaga.write(' Source of Data         BMKG                                         |\n')
    f_iaga.write(' Station Name           %-45s|\n'%datasta[1][1])
    f_iaga.write(' IAGA Code              %-45s|\n'%datasta[2][1])
    f_iaga.write(' Geodetic Latitude      %-45s|\n'%datasta[8][1])
    f_iaga.write(' Geodetic Longitude     %-45s|\n'%datasta[9][1])
    f_iaga.write(' Elevation              %-45s|\n'%datasta[5][1])
    f_iaga.write(' Reported               DHZF                                         |\n')
    f_iaga.write(' Sensor Orientation     XYZ                                          |\n')
    f_iaga.write(' Digital Sampling       1 second                                     |\n')
    f_iaga.write(' Data Interval Type     Filtered 1-minute (00:15-01:45)              |\n')
    f_iaga.write(' Data Type              provisional                                  |\n')
    f_iaga.write(' #F=sqrt((Bx^2)+(By^2)+(Bz^2)),H=sqrt(Bx^2+By^2),D=60*asind(By/H)    |\n')
    f_iaga.write('DATE       TIME         DOY     TUND      TUNH      TUNZ      TUNF   |\n')
    
    for j in range(0,1440):
        body_iaga = '%s    %9.2f %9.2f %9.2f %9.2f\n' %((datetime(year=int(date1.year),month=int(date1.month),
                    day=int(date1.day))+timedelta(minutes=j)).strftime("%Y-%m-%d %H:%M:00.000 %j"),
                    d_mean[j],h_mean[j],z_mean[j],f_mean[j])
        f_iaga.write(body_iaga)
    f_iaga.close()
    
    print(fileout+' has been created')
    print('Creating graphic, please wait...')

    #PLOT GRAFIK
    t=pd.date_range(date1.strftime('%d-%B-%Y 00:00'),date1.strftime('%d-%B-%Y 23:59'),freq="1min")
    plt.rcParams['axes.linewidth'] = 0.3
    xfmt = mdates.DateFormatter('%H:%M')
    fig, axs = plt.subplots(3, 2, sharex=True)
    fig.subplots_adjust(hspace=0)
    fig.suptitle(str.upper(date1.strftime('%d %B %Y')), fontsize=16, fontweight='bold', color='k')
    fig.set_size_inches(12.5, 8.5, forward=True)
    
    #PLOT Sebelum
    dataX = datafull.Bx[86400:172800]
    dataY = datafull.By[86400:172800]
    dataZ = datafull.Bz[86400:172800]
    
    #X Sebelum
    axs[0,0].plot(dataX, linewidth=1, color='r')
    axs[0,0].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[0,0].set_ylabel('X COMPONENT\n (nT)', fontsize=10)
    axs[0,0].set_title('SEBELUM', fontsize=12, fontweight='bold', color='r')
    
    #Y Sebelum
    axs[1,0].plot(dataY, linewidth=1, color='g')
    axs[1,0].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[1,0].set_ylabel('Y COMPONENT\n (nT)', fontsize=10)
    
    #Z Sebelum
    axs[2,0].plot(dataZ, linewidth=1, color='b')
    axs[2,0].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[2,0].set_ylabel('Z COMPONENT\n (nT)', fontsize=10)
    axs[2,0].set_xlabel('Time', fontsize=10)
    
    
    #PLOT Sesudah
    #X Sesudah
    axs[0,1].plot(t, x_mean, linewidth=1, color='r')
    axs[0,1].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[0,1].set_title('SESUDAH', fontsize=12, fontweight='bold', color='r')
    axs[0,1].set_ylabel('X COMPONENT\n (nT)', fontsize=10)
    axs[0,1].yaxis.set_label_position("right")
    axs[0,1].yaxis.tick_right()
    
    #Y Sesudah
    axs[1,1].plot(t, y_mean, linewidth=1, color='g')
    axs[1,1].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[1,1].set_ylabel('Y COMPONENT\n (nT)', fontsize=10)
    axs[1,1].yaxis.set_label_position("right")
    axs[1,1].yaxis.tick_right()
    
    #Z Sesudah
    axs[2,1].plot(t, z_mean, linewidth=1, color='b')
    axs[2,1].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[2,1].set_ylabel('Z COMPONENT\n (nT)', fontsize=10)
    axs[2,1].yaxis.set_label_position("right")
    axs[2,1].yaxis.tick_right()
    axs[2,1].set_xlabel('Time', fontsize=10)
    axs[2,1].xaxis.set_major_formatter(xfmt)
    axs[2,1].set_xlim(min(t),max(t))
	
    plt.show()

    print('Finish\n')