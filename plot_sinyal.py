# -*- coding: utf-8 -*-
"""
Created on Sun Nov 20 21:44:51 2016

@author: YOSI
"""
import numpy as np
import glob, os, re, time
from calendar import monthrange
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
import pandas as pd
import xml.etree.ElementTree
import lxml.etree

def plot_K(figname,start_date,length_date,k_i,A_i):
    cs=[]
    for value in k_i:
        if value>=0 and value<=3:
            cs.append('g')
        elif value==4:
            cs.append('yellow')
        elif value>=5 and value<=6:
            cs.append('r')
        elif value>=7 and value<=9:
            cs.append('purple')
        else:
            cs.append('k')
    "Plot K Index"        
    tt=['' for x in range(length_date*8)]
    for i in range(0,length_date*8):
        tt[i]=start_date+timedelta(hours=i*3)+timedelta(hours=1.5)
    ax5 = plt.subplot(211)
    plt.xlim(start_date,start_date+timedelta(days=length_date))
    plt.setp(ax5.xaxis.set_minor_locator(mdates.DayLocator()))
    plt.setp(ax5.xaxis.set_major_formatter(mdates.DateFormatter('%d %b')))
    plt.setp(ax5.get_xticklabels(), fontsize=8)
    plt.ylim(-1,9)
    plt.ylabel('K-Index',fontsize=8,weight='bold')
    plt.yticks(np.arange(0,10,1))
    plt.setp(ax5.get_yticklabels(), visible=True, fontsize=8)
    ax5.set_title(str.upper(start_date.strftime('%B %Y')), fontsize=12, fontweight='bold', color='r')
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.bar(tt,k_i,0.12,color=cs,linewidth=0.2,edgecolor='gold')
    
    "Plot K Legend"
    box = ax5.get_position()
    ax5.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax5.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    purple_patch = mpatches.Patch(color='purple', label='K = 7 - 9')
    red_patch = mpatches.Patch(color='red', label='K = 5 - 6')
    yellow_patch = mpatches.Patch(color='yellow', label='K = 4')
    green_patch = mpatches.Patch(color='green', label='K = 0 - 3')
    black_patch = mpatches.Patch(color='black', label='No data')
    plt.legend(handles=[purple_patch,red_patch,yellow_patch,green_patch,black_patch],bbox_to_anchor=(1, 1.04), loc='upper left', ncol=1,fontsize=9,labelspacing=1.74,borderpad=0.8)
    
    "Plot A Index"
    colormap = np.array(['g', 'yellow', 'r', 'darkred', 'purple', 'k'])
    csa = np.array(np.tile(5,length_date))
    for i in range(0,len(A_i)):
        if A_i[i]>=0 and A_i[i]<20:
            csa[i]=0
        elif A_i[i]>=20 and A_i[i]<30:
            csa[i]=1
        elif A_i[i]>=30 and A_i[i]<50:
            csa[i]=2
        elif A_i[i]>=50 and A_i[i]<100:
            csa[i]=3
        elif A_i[i]>=100:
            csa[i]=4
        else:
            csa[i]=5
        
    ttt=['' for x in range(length_date)]
    for i in range(0,length_date):
        ttt[i]=start_date+timedelta(days=i)
    ax6 = plt.subplot(212)
    plt.ylim(-3,100)
    plt.xlim(start_date-timedelta(hours=12),start_date+timedelta(days=length_date)-timedelta(hours=12))
    plt.setp(ax6.xaxis.set_minor_locator(MultipleLocator(1)))
    plt.xlabel("Date",weight='bold',fontsize=8)
    plt.setp(ax6.get_xticklabels(), fontsize=8)
    plt.setp(ax6.xaxis.set_major_formatter(mdates.DateFormatter('%d %b')))
    plt.ylabel('A-Index',fontsize=8,weight='bold')
    plt.setp(ax6.get_yticklabels(), visible=True, fontsize=8)
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)

    plt.scatter(ttt,A_i,marker="s",color=colormap[csa],edgecolor='k',zorder=2)
    plt.plot(ttt,A_i,'--r',zorder=1,dashes=(3, 1),linewidth=0.4)

    "Plot A Legend"
    box = ax6.get_position()
    ax6.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax6.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    purple_patch = mpatches.Patch(color='purple', label='A >= 100')
    darkred_patch = mpatches.Patch(color='darkred', label='A = 50 - 99')
    red_patch = mpatches.Patch(color='red', label='A = 30 - 49')
    yellow_patch = mpatches.Patch(color='yellow', label='A = 20 - 29')
    green_patch = mpatches.Patch(color='green', label='A = 0 - 19')
    black_patch = mpatches.Patch(color='black', label='No data')
    plt.legend(handles=[purple_patch,darkred_patch,red_patch,yellow_patch,green_patch,black_patch],bbox_to_anchor=(1, 1.03), loc='upper left', ncol=1,fontsize=8,labelspacing=1.67,borderpad=0.8)
    plt.savefig(str(figname)+'.png',dpi=150,bbox_inches='tight')
    plt.close()

def plot_baseline(filebaseline,filexml,tahun):
    def poly2latex(poly, variable="x", width=3):
        t = ["{0:0.{width}f}"]
        t.append(t[-1] + " {variable}")
        t.append(t[-1] + "^{1}")
    
        def f():
            for i, v in enumerate(reversed(poly)):
                idx = i if i < 2 else 2
                yield t[idx].format(v, i, variable=variable, width=width)
    
        return "${}$".format("+".join(f()))
    
    #Setting Figure properties
    fig, axs = plt.subplots(3, 1, sharex=True)
    fig.subplots_adjust(hspace=0)
    fig.suptitle('Baseline Data for TUNTUNGAN %s' %tahun, fontsize=10, fontweight='bold', color='k')
    xfmt = mdates.DateFormatter('%b')
    font = {'color':  'black',
            'size': 6,
            'ha':'center',
            'va': 'top'
            }
    
    #Read baseline file
    with open(filebaseline) as f_base:
        content = f_base.readlines()
        date = list(np.tile(np.nan,len(content)))
        Hobs = list(np.tile(np.nan,len(content)))
        Dobs = list(np.tile(np.nan,len(content)))
        Zobs = list(np.tile(np.nan,len(content)))
        for i in range(0,len(content)):
            data = re.split('\s+',content[i])
            date[i] = datetime.strptime(data[1]+' '+data[2], '%d-%b-%Y %H:%M:%S')
            #Remove 99999.9 value
            Hobs[i] = float(data[10]) if data[10] != '99999.9' else np.nan
            Hobs[i] = Hobs[i] if content[i][129] == 'H' else np.nan
            Dobs[i] = float(data[4]) if data[4] != '99999.9' else np.nan
            Dobs[i] = Dobs[i] if content[i][130] == 'D' else np.nan
            Zobs[i] = float(data[11]) if data[11] != '99999.9' else np.nan
            Zobs[i] = Zobs[i] if content[i][131] == 'Z' else np.nan
        #Create pandas dataframe from baseline file
        abs_obs = pd.DataFrame({'dt':date,'H':Hobs,'D':Dobs,'Z':Zobs})
        #Set date as index
        abs_obs = abs_obs.set_index(['dt'])
    
    #Plot absolute observation data
    axs[0].plot(date, Hobs, 'bs',markersize=3, marker='o', alpha=0.4)
    axs[0].set_ylabel('H (nT)', fontsize=8)
    axs[0].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[0].tick_params(axis='both', which='major', labelsize=8)
    
    axs[1].plot(date, Dobs, 'bs',markersize=3,marker='o',alpha=0.4)
    axs[1].set_ylabel('D (min)', fontsize=8)
    axs[1].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[1].tick_params(axis='both', which='major', labelsize=8)
    
    axs[2].plot(date, Zobs, 'bs',markersize=3,marker='o',alpha=0.4)
    axs[2].grid(b=True, which='major', color='c', linestyle=':', linewidth=0.5)
    axs[2].tick_params(axis='both', which='major', labelsize=8)
    axs[2].set_ylabel('Z (nT)', fontsize=8)
    axs[2].set_xlabel('Month', fontsize=8)
    axs[2].xaxis.set_major_formatter(xfmt)
    
    #Read functions XML file
    tree = xml.etree.ElementTree.parse(filexml)
    root = tree.getroot()
    #Count SD_entry, it must not read
    doc = lxml.etree.parse(filexml)
    S_count = int(doc.xpath('count(//SD_entry)'))
    #Declare variables
    component=list(np.tile(np.nan,len(root)-S_count))
    start_date=list(np.tile(np.nan,len(root)-S_count))
    end_date=list(np.tile(np.nan,len(root)-S_count))
    dates=list(np.tile(np.nan,len(root)-S_count))
    coefficient=list(np.tile(np.nan,len(root)-S_count))
    
    #Read each component, start_date,end_date, and coefficients
    for k in range(0,len(root)-S_count):
        component[k]=root[k][0].text
        start_date[k]=datetime.strptime(root[k][1].text, '%d-%b-%Y %H:%M:%S')
        end_date[k]=datetime.strptime(root[k][2].text, '%d-%b-%Y %H:%M:%S')
        #Create date range from start_date to end_date
        dates[k] = [start_date[k] + timedelta(days=x) for x in range(0, (end_date[k]-start_date[k]).days)]
        #Declare coefficient variable
        coefficient[k]=list(np.tile(np.nan,len(root[k])-3))
        #Read each coefficient from each component and date range
        for j in range(0,len(root[k])-3):
            coefficient[k][j]=float(root[k][j+3].text)
    
        #Calculate mean of date range, used for whitening later
        mu = mdates.date2num(dates[k]).mean()
        #Read each component observed data inside date range
        comp_value = abs_obs.loc[start_date[k]:end_date[k]][component[k]]
        #Declare x and y data, x axis is whitened, used for calculating coefficients
        x = mdates.date2num(comp_value.index.to_pydatetime())-mu
        y = comp_value
        #Clean NaN value
        idx = np.isfinite(x) & np.isfinite(y)
        #Calculate coefficients using polyfit, degree is len(root[k])-4
        coeffs = np.polyfit(x[idx], y[idx],len(root[k])-4)
        #Create polynomial function using coefficients
        f = np.poly1d(coeffs)
    
        #Plot line from polynomial functions
        if component[k]=='H':
            axs[0].plot(mdates.date2num(dates[k]), f(mdates.date2num(dates[k])-mu), 'r-')
            axs[0].text(mu, f((mdates.date2num(dates[k])-mu).mean()), poly2latex(coeffs), fontdict=font)
        elif component[k]=='D':
            axs[1].plot(mdates.date2num(dates[k]), f(mdates.date2num(dates[k])-mu), 'r-')
            axs[1].text(mu, f((mdates.date2num(dates[k])-mu).mean()), poly2latex(coeffs), fontdict=font)
        elif component[k]=='Z':
            axs[2].plot(mdates.date2num(dates[k]), f(mdates.date2num(dates[k])-mu), 'r-')
            axs[2].text(mu, f((mdates.date2num(dates[k])-mu).mean()), poly2latex(coeffs), fontdict=font)
            
    fig.savefig('baseline.png',dpi=150,bbox_inches='tight')
    
def plot_sinyal(pathIAGA):
    with open('station.ini') as f_init:
        content_init = f_init.readlines()
        datasta = list(np.tile(np.nan,len(content_init)))		
        for i in range(0,len(content_init)):
            datasta[i] = re.split('\=|\n',content_init[i])

    pathIMFV = pathIAGA + '\\IMFV'
    first_data = os.path.basename(glob.glob(os.path.join(str(pathIAGA), '*.min'))[0])
    tahun1 = int(first_data[3:7])
    bulan1 = int(first_data[7:9])
    tanggal1 = 1
    date1 = datetime(year=tahun1,month=bulan1,day=tanggal1)
    hari_bulan = monthrange(tahun1, bulan1)[1]
    Bx = list(np.tile(np.nan,hari_bulan*1440))
    By = list(np.tile(np.nan,hari_bulan*1440))
    Bz = list(np.tile(np.nan,hari_bulan*1440))
    Bf = list(np.tile(np.nan,hari_bulan*1440))
    Bh = list(np.tile(np.nan,hari_bulan*1440))
    Bd = list(np.tile(np.nan,hari_bulan*1440))
    Bi = list(np.tile(np.nan,hari_bulan*1440))
    k_i = [-1]*(hari_bulan+7)*8
    a_i = [-1]*(hari_bulan+7)*8
    sk = ['']*(hari_bulan+7)
    A_i = [-1]*(hari_bulan+7)

    "READ IAGA DATA"
    for filename in glob.glob(os.path.join(str(pathIAGA), '*.min')):
        with open(filename) as f_iaga:
            content = f_iaga.readlines()
            for num, line in enumerate(content, 1):
                if 'DATE' in line:
                    skip_line = num
        with open(filename) as f_iaga:
            content = f_iaga.readlines()[skip_line:]
            for i in range(0,len(content)):
                data = re.split('\s+',content[i])
                tanggal = datetime.strptime(data[0], '%Y-%m-%d')
                tanggal_data = tanggal.day-1
                if data[3]!='99999.00':
                    Bh[i+(1440*tanggal_data)] = float(data[3])
                else:
                    Bh[i+(1440*tanggal_data)] = np.nan
                if data[4]!='99999.00':
                    Bd[i+(1440*tanggal_data)] = float(data[4])
                else:
                    Bd[i+(1440*tanggal_data)] = np.nan
                if data[5]!='99999.00':
                    Bz[i+(1440*tanggal_data)] = float(data[5])
                else:
                    Bz[i+(1440*tanggal_data)] = np.nan
                if data[6]!='99999.00':
                    Bf[i+(1440*tanggal_data)] = float(data[6])
                else:
                    Bf[i+(1440*tanggal_data)] = np.nan
                Bx[i+(1440*tanggal_data)] = Bh[i+(1440*tanggal_data)]*np.cos(np.deg2rad(Bd[i+(1440*tanggal_data)]/60))
                By[i+(1440*tanggal_data)] = Bh[i+(1440*tanggal_data)]*np.sin(np.deg2rad(Bd[i+(1440*tanggal_data)]/60))
                Bi[i+(1440*tanggal_data)] = np.rad2deg(np.arctan(Bz[i+(1440*tanggal_data)]/Bh[i+(1440*tanggal_data)]))
    
    "READ K INDEX"
    with open(pathIMFV + '\\data.dka') as f:
        content = f.readlines()
        for j in range(6,len(content)):
            data_k = re.split(' ',content[j])
            for i in range(0,len(data_k)-14):
                data_k.remove('')
            index_kk = time.strptime(data_k[0],'%d-%b-%y').tm_mday-1
            sk[index_kk] = int(data_k[10])
            if sk[index_kk] == -1:
                sk[index_kk] = ''
            for m in range(0,8):
                k_i[m+((index_kk)*8)] = int(data_k[m+2])
                if (k_i[m+((index_kk)*8)]==0):
                    a_i[m+((index_kk)*8)]=0.0
                elif (k_i[m+((index_kk)*8)]==1):
                    a_i[m+((index_kk)*8)]=3.0
                elif (k_i[m+((index_kk)*8)]==2):
                    a_i[m+((index_kk)*8)]=7.0
                elif (k_i[m+((index_kk)*8)]==3):
                    a_i[m+((index_kk)*8)]=15.0	
                elif (k_i[m+((index_kk)*8)]==4):
                    a_i[m+((index_kk)*8)]=27.0	
                elif (k_i[m+((index_kk)*8)]==5):
                    a_i[m+((index_kk)*8)]=48.0	
                elif (k_i[m+((index_kk)*8)]==6):
                    a_i[m+((index_kk)*8)]=80.0
                elif (k_i[m+((index_kk)*8)]==7):
                    a_i[m+((index_kk)*8)]=140.0	    
                elif (k_i[m+((index_kk)*8)]==8):
                    a_i[m+((index_kk)*8)]=200	    
                elif (k_i[m+((index_kk)*8)]==9):
                    a_i[m+((index_kk)*8)]=300.0
                elif (k_i[m+((index_kk)*8)]==-1):
                    a_i[m+((index_kk)*8)]=-1
                    
        for n in range(0,hari_bulan):
            a_i_array = np.array(a_i[0+(n*8):8+(n*8)])
            if list(a_i_array).count(-1)>0:
                A_i[n] = -1
            else:
                A_i[n] = a_i_array[a_i_array != -1].mean()

    t=np.arange(0,hari_bulan,1/1440.0)
    kelas=4
    now=datetime.utcnow()
    plt.rcParams['axes.linewidth'] = 0.3
    "PLOT X"
    ax1 = plt.subplot(411)
    ax1.set_title(str.upper(date1.strftime('%B %Y')), fontsize=12, fontweight='bold', color='r')
    plt.plot(t, Bx, linewidth=0.3, color='k')
    plt.setp(ax1.get_xticklabels(), visible=False)
    plt.setp(ax1.get_yticklabels(), fontsize=8)
    plt.ylabel('X COMPONENT\n (nT)', fontsize=8)
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.yticks(np.arange(np.nanmin(Bx), np.nanmax(Bx), (np.nanmax(Bx)-np.nanmin(Bx))/kelas))

    "PLOT Y"
    ax2 = plt.subplot(412, sharex=ax1)
    plt.plot(t, By, linewidth=0.3, color='k')
    plt.setp(ax2.yaxis.set_major_formatter(FormatStrFormatter('%.2f')))
    plt.setp(ax2.get_xticklabels(), visible=False)
    plt.setp(ax2.get_yticklabels(), fontsize=8)
    plt.ylabel('Y COMPONENT\n (nT)', fontsize=8)
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.yticks(np.arange(np.nanmin(By), np.nanmax(By), (np.nanmax(By)-np.nanmin(By))/kelas))

    "PLOT Z"
    ax3 = plt.subplot(413, sharex=ax1)
    plt.plot(t, Bz, linewidth=0.3, color='k')
    plt.setp(ax3.yaxis.set_major_formatter(FormatStrFormatter('%.2f')))
    plt.setp(ax3.get_xticklabels(), visible=False)
    plt.setp(ax3.get_yticklabels(), fontsize=8)
    plt.ylabel('Z COMPONENT\n (nT)', fontsize=8)
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.yticks(np.arange(np.nanmin(Bz), np.nanmax(Bz), (np.nanmax(Bz)-np.nanmin(Bz))/kelas))

    "PLOT F"
    ax4 = plt.subplot(414, sharex=ax1)
    ax4.text(0,-0.5,datasta[0][1], fontsize=8, fontweight='bold',ha='left', va='bottom', transform=ax4.transAxes)
    ax4.text(1,-0.5,'[Created at %s]' %now.strftime('%Y-%m-%d %H:%M UT'), fontsize=8, fontweight='bold',ha='right', va='bottom', transform=ax4.transAxes)
    plt.plot(t, Bf, linewidth=0.3, color='k')
    plt.setp(ax4.get_xticklabels(), fontsize=8)
    plt.setp(ax4.get_yticklabels(), fontsize=8)
    plt.yticks(np.arange(np.nanmin(Bf), np.nanmax(Bf), (np.nanmax(Bf)-np.nanmin(Bf))/kelas))
    plt.xticks(np.arange(1,hari_bulan,5))
    ax4.xaxis.set_minor_locator(MultipleLocator(1))
    plt.ylabel('F COMPONENT\n (nT)', fontsize=8)
    plt.xlabel('Date', fontsize=8)
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.xlim(0,hari_bulan)
    plt.savefig('sinyal.png',dpi=150,bbox_inches='tight')

    "PLOT K INDEX"
    for i in range(0,int(np.ceil(monthrange(tahun1, bulan1)[1]/7.0))):
        plot_K(i,date1+timedelta(days=i*7),7,k_i[i*7*8:(i+1)*7*8],A_i[i*7:(i+1)*7])
    plot_K('k_index',date1,hari_bulan,k_i[0:(8*hari_bulan)],A_i[0:hari_bulan])