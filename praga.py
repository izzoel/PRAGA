#   _____  _____            _____          
#  |  __ \|  __ \     /\   / ____|   /\    
#  | |__) | |__) |   /  \ | |  __   /  \   
#  |  ___/|  _  /   / /\ \| | |_ | / /\ \  
#  | |    | | \ \  / ____ \ |__| |/ ____ \ 
#  |_|    |_|  \_\/_/    \_\_____/_/    \_\
#
# PROGRAM ABSENSI PEGAWAI
# 16 Oktober 2022 | Zulfahmi Ramadhani


#Library yg diperlukan
import pandas as pd
from pandas import ExcelWriter, date_range
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

print('memproses..')

#buka file
file = pd.read_excel('big_data.xlsx')

#set tgl awal dan akhir
awal = datetime(2022, 1, 1)
akhir = datetime(2022, 12, 31)

#variabel penyimpan dataframe
tl = []

#fungsi mewarnai baris excel
def hightlight(cell):
    if cell["KET"] == "telat < 1 jam":
        return ['background-color: yellow'] * len(cell)
    elif cell["KET"] == "CLOSED":
        return ['background-color: grey'] * len(cell)
    elif cell["KET"] == "telat > 1 jam":
        return ['background-color: red'] * len(cell)
    else :
        return ['background-color: none'] * len(cell)
    
#cari datanya dari tgl awal s/d akhir
for hari in date_range(awal,akhir):
    hari = str(datetime.date(hari))
    
    data= file.loc[(file['CHECKTIME'] >= hari) & (file['CHECKTIME'] <= hari+' 23:59:59')]
    
    df = pd.DataFrame(tl)
    
    tl = pd.concat([df,data.sort_values('CHECKTIME').groupby('USERID').first(),data.sort_values('CHECKTIME').groupby('USERID').last()]).drop_duplicates()
    
    tl.loc[(tl['CHECKTIME']  >= hari+' 00:00:00') & (tl['CHECKTIME']  <= hari+' 06:59:59'), ['KET','DENDA']] = ['CLOSED', 0]
    tl.loc[(tl['CHECKTIME']  >= hari+' 07:00:00') & (tl['CHECKTIME']  <= hari+' 07:45:59'), ['KET','DENDA']] = ['on time', 0]
    tl.loc[(tl['CHECKTIME']  >= hari+' 07:46:00') & (tl['CHECKTIME']  <= hari+' 08:00:59'), ['KET','DENDA']] = ['telat < 1 jam', 22500]
    tl.loc[(tl['CHECKTIME']  >= hari+' 08:01:00') & (tl['CHECKTIME']  <= hari+' 10:00:59'), ['KET','DENDA']] = ['telat > 1 jam', 40000]
    tl.loc[(tl['CHECKTIME']  >= hari+' 10:01:00') & (tl['CHECKTIME']  <= hari+' 14:59:59'), ['KET','DENDA']] = ['CLOSED', 0]
    tl.loc[(tl['CHECKTIME']  >= hari+' 15:00:00') & (tl['CHECKTIME']  <= hari+' 17:00:59'), ['KET','DENDA']] = ['on time', 0]
    tl.loc[(tl['CHECKTIME']  >= hari+' 17:01:00') & (tl['CHECKTIME']  <= hari+' 23:59:59'), ['KET','DENDA']] = ['CLOSED', 0]

#urutkan USERID dan CHECKTIME secara asc
res = tl.sort_values(['USERID', 'CHECKTIME'], ascending = [True, True])

#membuat worksheet utk data perbulan
nama_bulan = ['Januari', 'Februari','Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
i = 1
writer = ExcelWriter('Praga_rekap ('+str(date.today())+').xlsx')
for bulan in nama_bulan:
    tl = res.loc[(res['CHECKTIME'] >= datetime(2022,1,1)+relativedelta(months=i-1)) & (res['CHECKTIME'] <= datetime(2022,1,1)+relativedelta(months=+i))]

    styler = tl.reset_index(drop=False).style.apply(hightlight, axis=1)
    styler.to_excel(writer, sheet_name=bulan,index=False)

    i += 1
writer.save()