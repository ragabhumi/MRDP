# -*- coding: utf-8 -*-
"""
Created on Sun Nov 20 21:44:51 2016

@author: YOSI
"""
import numpy as np
import glob
import os
import re
import time
from calendar import monthrange
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import matplotlib.patches as mpatches
import matplotlib.dates as mdates

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
        tt[i]=start_date+timedelta(hours=i*3)
    ax5 = plt.subplot(211)
    plt.xlim(start_date,start_date+timedelta(days=length_date))
    plt.setp(ax5.xaxis.set_minor_locator(MultipleLocator(1)))
    plt.setp(ax5.xaxis.set_major_formatter(mdates.DateFormatter('%d %b')))
    plt.setp(ax5.get_xticklabels(), fontsize=8)
    plt.ylim(-1,9)
    plt.ylabel('K-Index',fontsize=8,weight='bold')
    plt.yticks(np.arange(0,10,1))
    plt.setp(ax5.get_yticklabels(), visible=True, fontsize=8)
    ax5.set_title(str.upper(start_date.strftime('%B %Y')), fontsize=12, fontweight='bold', color='r')
    plt.grid(b=True, which='major', color='c', linestyle='-', linewidth=0.1)
    plt.grid(b=True, which='minor', color='c', linestyle='-', linewidth=0.1)
    plt.bar(tt,k_i,0.12,color=cs,linewidth=0.2)
 
    "Plot K Legend"
    box = ax5.get_position()
    ax5.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax5.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    purple_patch = mpatches.Patch(color='purple', label='K = 7 - 9')
    red_patch = mpatches.Patch(color='red', label='K = 5 - 6')
    yellow_patch = mpatches.Patch(color='yellow', label='K = 4')
    green_patch = mpatches.Patch(color='green', label='K = 1 - 3')
    black_patch = mpatches.Patch(color='black', label='Lost data')
    plt.legend(handles=[purple_patch,red_patch,yellow_patch,green_patch,black_patch],bbox_to_anchor=(1, 1.03), loc='upper left', ncol=1,fontsize=9,labelspacing=2.75,borderpad=0.8)
    
    "Plot A Index"
    colormap = np.array(['g', 'yellow', 'r', 'purple', 'k'])
    csa = np.array(np.tile(4,length_date))
    for i in range(0,len(A_i)):
        if A_i[i]>=0 and A_i[i]<=30:
            csa[i]=0
        elif A_i[i]>30 and A_i[i]<=50:
            csa[i]=1
        elif A_i[i]>50 and A_i[i]<=100:
            csa[i]=2
        elif A_i[i]>100:
            csa[i]=3
        else:
            csa[i]=4
        
    ttt=['' for x in range(length_date)]
    for i in range(0,length_date):
        ttt[i]=start_date+timedelta(days=i)
    ax6 = plt.subplot(212)
    plt.ylim(0,100)
    plt.xlim(start_date,start_date+timedelta(days=length_date))
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
    purple_patch = mpatches.Patch(color='purple', label='Great\nStorm')
    red_patch = mpatches.Patch(color='red', label='Intermediate\nStorm')
    yellow_patch = mpatches.Patch(color='yellow', label='Small\nStorm')
    green_patch = mpatches.Patch(color='green', label='Relatively\nQuiet')
    black_patch = mpatches.Patch(color='black', label='Lost data')
    plt.legend(handles=[purple_patch,red_patch,yellow_patch,green_patch,black_patch],bbox_to_anchor=(1, 1.03), loc='upper left', ncol=1,fontsize=8,labelspacing=1.7,borderpad=0.8)
    plt.savefig(str(figname)+'.png',dpi=150)
    plt.close()

def plot_sinyal(pathIAGA):
    pathIMFV = pathIAGA + '\\IMFV'
    Bx = list(np.tile(np.nan,31*1440))
    By = list(np.tile(np.nan,31*1440))
    Bz = list(np.tile(np.nan,31*1440))
    Bf = list(np.tile(np.nan,31*1440))
    Bh = list(np.tile(np.nan,31*1440))
    Bd = list(np.tile(np.nan,31*1440))
    Bi = list(np.tile(np.nan,31*1440))
    first_data = os.path.basename(glob.glob(os.path.join(str(pathIAGA), '*.min'))[0])
    tahun1 = int(first_data[3:7])
    bulan1 = int(first_data[7:9])
    tanggal1 = 1
    date1 = datetime(year=tahun1,month=bulan1,day=tanggal1)
    hari_bulan = monthrange(tahun1, bulan1)[1]
    k_i = [-1]*(hari_bulan+7)*8
    a_i = [-1]*(hari_bulan+7)*8
    sk = ['']*(hari_bulan+7)
    A_i = [-1]*(hari_bulan+7)
    p = 0

    "READ IAGA DATA"
    for filename in glob.glob(os.path.join(str(pathIAGA), '*.min')):
        with open(filename) as f_lemi:
            content = f_lemi.readlines()
            for num, line in enumerate(content, 1):
                if 'DATE' in line:
                    skip_line = num
        with open(filename) as f_lemi:
            content = f_lemi.readlines()[skip_line:]
            for i in range(0,len(content)):
                data = re.split('\s+',content[i])
                if data[3]!='99999.00':
                    Bh[i+(1440*p)] = float(data[3])
                else:
                    Bh[i+(1440*p)] = np.nan
                if data[4]!='99999.00':
                    Bd[i+(1440*p)] = float(data[4])
                else:
                    Bd[i+(1440*p)] = np.nan
                if data[5]!='99999.00':
                    Bz[i+(1440*p)] = float(data[5])
                else:
                    Bz[i+(1440*p)] = np.nan
                if data[6]!='99999.00':
                    Bf[i+(1440*p)] = float(data[6])
                else:
                    Bf[i+(1440*p)] = np.nan
                Bx[i+(1440*p)] = Bh[i+(1440*p)]*np.cos(np.deg2rad(Bd[i+(1440*p)]/60))
                By[i+(1440*p)] = Bh[i+(1440*p)]*np.sin(np.deg2rad(Bd[i+(1440*p)]/60))
                Bi[i+(1440*p)] = np.rad2deg(np.arctan(Bz[i+(1440*p)]/Bh[i+(1440*p)]))
        p = p+1
    
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
                    a_i[m+((index_kk)*8)]=0
                elif (k_i[m+((index_kk)*8)]==1):
                    a_i[m+((index_kk)*8)]=3
                elif (k_i[m+((index_kk)*8)]==2):
                    a_i[m+((index_kk)*8)]=7	
                elif (k_i[m+((index_kk)*8)]==3):
                    a_i[m+((index_kk)*8)]=15	
                elif (k_i[m+((index_kk)*8)]==4):
                    a_i[m+((index_kk)*8)]=27	
                elif (k_i[m+((index_kk)*8)]==5):
                    a_i[m+((index_kk)*8)]=48	
                elif (k_i[m+((index_kk)*8)]==6):
                    a_i[m+((index_kk)*8)]=80
                elif (k_i[m+((index_kk)*8)]==7):
                    a_i[m+((index_kk)*8)]=140	    
                elif (k_i[m+((index_kk)*8)]==8):
                    a_i[m+((index_kk)*8)]=200	    
                elif (k_i[m+((index_kk)*8)]==9):
                    a_i[m+((index_kk)*8)]=300
                elif (k_i[m+((index_kk)*8)]==-1):
                    a_i[m+((index_kk)*8)]=-1

        for n in range(0,hari_bulan):
            A_i[n] = sum(a_i[0+(n*8):8+(n*8)])/8

    t=np.arange(0,31,1/1440.0)
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
    ax4.text(0,-0.5,'Stasiun Geofisika Tuntungan', fontsize=8, fontweight='bold',ha='left', va='bottom', transform=ax4.transAxes)
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
    plt.savefig('sinyal.png',dpi=150)

    "PLOT K INDEX"
    for i in range(0,int(np.ceil(monthrange(tahun1, bulan1)[1]/7.0))):
        plot_K(i,date1+timedelta(days=i*7),7,k_i[i*7*8:(i+1)*7*8],A_i[i*7:(i+1)*7])
    plot_K('k_index',date1,hari_bulan,k_i[0:(8*hari_bulan)],A_i[0:hari_bulan])