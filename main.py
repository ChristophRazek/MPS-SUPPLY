import pandas as pd
import warnings
from tkinter import messagebox
from datetime import datetime
import os

warnings.filterwarnings('ignore')

link = r'S:\ShipmentList\MPS-SUPPLY.xlsx'

#Zeitstempel Letzte Änderung!
c_time = os.path.getmtime(link)
dt_c = datetime.fromtimestamp(c_time)

#Kopieren von MPS File zu Laufwerk
bestellungen = pd.read_excel(link)


bestellungen[['FIXPOSNR','BELEGART','BELEGNR']] = bestellungen[['FIXPOSNR','BELEGART','BELEGNR']].fillna(0).astype('int64')
bestellungen.to_csv(r'L:\Supply\MPS-SUPPLY.csv', sep=';', index=False)

#Log File
with open(r'S:\EMEA\Kontrollabfragen\MPS-SUPPLY.txt', 'w') as f:
    f.write(f'Last MPS-SUPPLY copied at: {dt_c}')
    f.close()


messagebox.showinfo('Update Erfolgreich!', f'Das MPS-SUPPLY Update mit letzter Änderung vom {dt_c} wurde erfolgreich durchgeführt.')