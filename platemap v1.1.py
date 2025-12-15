import pandas as pd
import datetime
from fpdf import FPDF
import re
import os
import logging

logging.basicConfig(
    filename = 'errors.log',
    level = logging.DEBUG,
    format= '%(asctime)s %(levelname)s %(name)s %(message)s'
)

fpath = 'L:/Micro/AR-TB/AR/ABI7500 import - export files/ABI7500 Importer/'
ipath = 'L:/Micro/AR-TB/AR/ABI7500 import - export files/Imports/'
ppath = 'L:/Micro/AR-TB/AR/ABI7500 import - export files/Platemaps/'
try:
    rnumber = input('Please enter PCR run number: ')
except:
    logging.error('Error entering run number')
reString = r"\.xlsx$"
try:
    flist = []
    for file in os.listdir(fpath):
        if re.search(reString, file):
            flist.append(file)
            fname = file
            print(f"Found file: {fname}")
    if len(flist) > 1:
        logging.warning('Too many excel files found, please rerun with only one excel file')
    elif len(flist) == 0:
        logging.warning('No .xlsx file found')
    fullpath = fpath+fname
except:
    logging.error("Error finding data file")
pcols = range(1,13)
prows = ['A','B','C','D','E','F','G','H']
tryal = pd.read_excel(fullpath)
class RunProgram:
    def __init__(self):
        self.start()
        try:    
            self.getTemplate()
            self.getData()
            self.chooseTemplate()
        except:
            logging.warning('Run Program class failed to initialize')
        try:   
            for plate in self.plates:
                MakeExport(plate)
                makePDF(plate.platemap,plate.name)
        except:
            logging.warning('Unable to create plates')
    def start(self):
        self.isrunning = True
    def getTemplate(self):
        data = {}
        for x in pcols:
            tmplist = []
            for y in prows:
                tmplist.append(f'{y}{x}')
            data[x] = tmplist
        self.template = Platemap(data,'Template')
    def getData(self):
        self.samples = pd.read_excel(fullpath)['Specimen'].to_list()
    def chooseTemplate(self):
        self.plates = []
        if len(self.samples) <=19:
            self.plates = self.quarter()
            print(self.plates)
        elif len(self.samples)<=43:
            self.plates = self.half()
            print(self.plates)
        elif len(self.samples)<=91:
            self.plates = self.whole()
            print(self.plates)
        else:
            print("too many samples please choose a different sample file")
    def quarter(self):
        print(f'Creating Quarter Plate')
        plates = [pd.DataFrame(columns = self.template.platemap.columns, index = self.template.platemap.index)]
        i=0
        s=0
        for plate in plates:
            for col in plate[1:3].columns:
                for row in plate.index:
                    if i==0:
                        print(f'Adding NTC to {row}{col}')
                        plate.at[row,col] = 'NTC'
                        
                    elif i==1:
                         plate.at[row,col] = '1706'
                         print(f'Adding 1706 to {row}{col}')
                    elif i ==22:
                        plate.at[row,col] = '1705'
                        print(f'Adding 1705 to {row}{col}')
                    elif i == 23:
                        plate.at[row,col] = '2146'
                        print(f'Adding 2146 to {row}{col}')
                    elif s < len(self.samples):
                        print(f'Adding {self.samples[s]} to {row}{col}')
                        plate.at[row,col] = self.samples[s]
                        s+=1
                    else:
                        plate.at[row,col] = ''
                    i+=1
        plate[4] = plate[1].copy()
        plate[5] = plate[2].copy()
        plate[6] = plate[3].copy()
        plate[7] = plate[1].copy()
        plate[8] = plate[2].copy()
        plate[9] = plate[3].copy()
        plate[10] = plate[1].copy()
        plate[11] = plate[2].copy()
        plate[12] = plate[3].copy()
        plate.at['G',6] = '#0039'
        plate.at['H',6] = '#0054'
        plate.at['G',9] = '#0034'
        plate.at['H',9] = '#0092'
        plate.at['F',12] = '#0045'
        plate.at['G',12] = '#0036'
        plate.at['H',12] = '#0052'
        return [Platemap(plate,'KNVOIA')]
    def half(self):
        print(f'Creating Half Plate')
        plate = pd.DataFrame(columns = self.template.platemap.columns, index = self.template.platemap.index)
        i=0
        s=0
        
        for col in plate[1:6].columns:
            for row in plate.index:
                if i==0:
                    print(f'Adding NTC to {row}{col}')
                    plate.at[row,col] = 'NTC'
                        
                elif i==1:
                        plate.at[row,col] = '1706'
                        print(f'Adding 1706 to {row}{col}')
                elif i ==46:
                    plate.at[row,col] = '1705'
                    print(f'Adding 1705 to {row}{col}')
                elif i == 47:
                    plate.at[row,col] = '2146'
                    print(f'Adding 2146 to {row}{col}')
                elif s < len(self.samples):
                    print(f'Adding {self.samples[s]} to {row}{col}')
                    plate.at[row,col] = self.samples[s]
                    s+=1
                else:
                    plate.at[row,col] = ''
                i+=1
        plate[7] = plate[1].copy()
        plate[8] = plate[2].copy()
        plate[9] = plate[3].copy()
        plate[10] = plate[4].copy()
        plate[11] = plate[5].copy()
        plate[12] = plate[6].copy()
        plate.at['G',12] = '#0039'
        plate.at['H',12] = '#0054'
        plate2 = plate.copy()
        plate2.at['G', 6] = '#0034'
        plate2.at['H', 6] = '#0092'
        plate2.at['F', 12] = '#0045'
        plate2.at['G', 12] = '#0036'
        plate2.at['H', 12] = '#0052'
        return [Platemap(plate,'KV'),Platemap(plate2,'IO')]
    def whole(self):
        print(f'Creating Full Platemap')
        plate = pd.DataFrame(columns = self.template.platemap.columns, index = self.template.platemap.index)
        i=0
        s=0
        
        for col in plate.columns:
            for row in plate.index:
                if i==0:
                    print(f'Adding NTC to {row}{col}')
                    plate.at[row,col] = 'NTC'
                        
                elif i==1:
                        plate.at[row,col] = '1706'
                        print(f'Adding 1706 to {row}{col}')
                elif i ==94:
                    plate.at[row,col] = '1705'
                    print(f'Adding 1705 to {row}{col}')
                elif i == 95:
                    plate.at[row,col] = '2146'
                    print(f'Adding 2146 to {row}{col}')
                elif s < len(self.samples):
                    print(f'Adding {self.samples[s]} to {row}{col}')
                    plate.at[row,col] = self.samples[s]
                    s+=1
                else:
                    plate.at[row,col] = ''
                i+=1
        plate2 = plate.copy()
        plate2.at['G', 12] = '#0039'
        plate2.at['H', 12] = '#0054'
        plate3 = plate.copy()
        plate3.at['G', 12] = '#0034'
        plate3.at['H', 12] = '#0092'
        plate4 = plate.copy()
        plate4.at['F', 12] = '#0045'
        plate4.at['G', 12] = '#0036'
        plate4.at['H', 12] = '#0052'
        return [Platemap(plate,'KN'),Platemap(plate2,'VO'),Platemap(plate3,'IMP'),Platemap(plate4,'AO')]
class Platemap:
    def __init__(self, _df, _name):
        self.name = _name
        self.platemap = pd.DataFrame(_df,columns=pcols, index=prows, dtype='U20')
    def __repr__(self):
        return f'Platemap object: {self.name}'
    def __str__(self):
        return self.name
class MakeExport:
    def __init__(self, _plate):
            self.time = datetime.datetime.now().strftime("%Y%m%d")
            self.plate = _plate
            tmpname = f'{rnumber}_AR-CRO-PCR-ABI-7500_{self.time} - {self.plate.name}.txt'
            tmppath = ipath+tmpname
            self.file = open(tmppath,'w')
            self.data = self.plate.platemap
            self.headers = [f'*** SDS Setup File Version\t3\n',f'*** Output Plate Size\t96\n',
                            f'*** Output Plate ID\t{self.plate.name}\n',f'*** Number of Detectors\t13\n',
                            f'Detector\tReporter\tQuencher\tDescription\tComments\n',f'KPC\tFAM\n',f'NDM\tVIC\n'
                            f'16s (KN)\tCY5\n',f'OXA-48-like\tFAM\n',f'VIM\tVIC\n',f'16s (VO)\tCY5\n',f'IMP 1\tFAM\n',
                            f'IMP 2\tVIC\n',f'16s (I)\tCY5\n',f'OXA-23-like\tFAM\n',f'OXA-24/40-like\tVIC\n',
                            f'OXA-58-like\tTEXAS RED\n',f'16s (AO)\tCY5\n',f'Well\tSample Name\tDetector\tTask\tQuantity\n']
            if self.plate.name == 'KN':
                self.KN()
            elif self.plate.name == 'VO':
                self.VO()
            elif self.plate.name == 'IMP':
                self.IMP()
            elif self.plate.name == 'AO':
                self.AO()
            elif self.plate.name == 'KV':
                self.KV()
            elif self.plate.name == 'IO':
                self.IO()
            elif self.plate.name == 'KNVOIA':
                self.KNVOIA()
    def KN(self):
        for head in self.headers:
            self.file.write(head)
            i=1
        for row in self.data.index:
            for col in self.data.columns:
                print(self.data.loc[row,col])
                if self.data.loc[row,col] != '':
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tKPC\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tNDM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (KN)\tUNKN\n')
                i+=1
    def VO(self):
        for head in self.headers:
            self.file.write(head)
            i=1
        for row in self.data.index:
            for col in self.data.columns:
                print(self.data.loc[row,col])
                if self.data.loc[row,col] != '':
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tVIM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-48-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (VO)\tUNKN\n')
                i+=1
    def IMP(self):
        for head in self.headers:
            self.file.write(head)
            i=1
        for row in self.data.index:
            for col in self.data.columns:
                print(self.data.loc[row,col])
                if self.data.loc[row,col] != '':
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tIMP 1\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tIMP 2\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (I)\tUNKN\n')
                i+=1
    def AO(self):
        for head in self.headers:
            self.file.write(head)
        i=1
        for row in self.data.index:
            for col in self.data.columns:
                print(self.data.loc[row,col])
                if self.data.loc[row,col] != '':
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tOXA-23-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-24/40-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-58-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (AO)\tUNKN\n')
                i+=1
    def KV(self):
        for head in self.headers:
            self.file.write(head)
        i=1
        for row in self.data.index:
            for col in self.data.columns:
                if col in range(1,7):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tKPC\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tNDM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (KN)\tUNKN\n')
                elif col in range(7,13):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tVIM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-48-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (VO)\tUNKN\n')
                i+=1

    def IO(self):
        for head in self.headers:
            self.file.write(head)
        i=1
        for row in self.data.index:
            for col in self.data.columns:
                if col in range(1,7):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tIMP 1\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tIMP 2\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (I)\tUNKN\n')
                elif col in range(7,13):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tOXA-23-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-24/40-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-58-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (AO)\tUNKN\n')
                i+=1
    def KNVOIA(self):
        for head in self.headers:
            self.file.write(head)
        i=1
        for row in self.data.index:
            for col in self.data.columns:
                if col in range(1,4):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tKPC\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tNDM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (KN)\tUNKN\n')
                elif col in range(4,7):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tVIM\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-48-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (VO)\tUNKN\n')
                elif col in range(7,10):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tIMP 1\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tIMP 2\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (I)\tUNKN\n')
                elif col in range(10,13):
                    self.file.write(f'{i}\t{self.data.loc[row,col]}\tOXA-23-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-24/40-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\tOXA-58-like\tUNKN\n'+
                                    f'{i}\t{self.data.loc[row,col]}\t16s (AO)\tUNKN\n')
            i+=1
class makePDF:
    def __init__(self, _df,_name):
        self.cellsize = 20
        self.xstart = 13.5
        self.data = _df
        self.name = f'{rnumber}_AR-CRO-PCR-ABI-7500_ {datetime.datetime.now().strftime("%Y%m%d")} - {_name}'
        self.cellorigins = [(15*x+13.5,15*y) for x in range(3,15) for y in range(3,11)]
        self.pdf = FPDF('L', 'mm', 'A4')
        
        self.pdf.add_page()
        self.pdf.set_font("Arial", 'B', size=24)
        self.pdf.cell(265,10,'ABI 7500 Fast DX CRO Platemap',0,1,'C')
        self.pdf.set_font("Arial", size=12)
        self.usecells()
        self.createlabels()
        self.output(self.name)
    def usecells(self):
        point = (25,25)
        self.pdf.set_xy(point[0],point[1])
        self.pdf.set_font("Arial", 'B',size=8)
        for i,x in enumerate(self.data.index):
            for j,y in enumerate(self.data.columns):
                txt = self.data.loc[x,y]
                self.pdf.cell(self.cellsize, self.cellsize, txt, border=1, ln=0, align='C')
            point = (point[0], point[1] + self.cellsize)
            self.pdf.set_xy(point[0],point[1])

    def creategrid(self):
        for c in self.cellorigins:
            self.pdf.rect(c[0], c[1], self.cellsize, self.cellsize)
    def createlabels(self):
        tmpx = 50
        tmpy = 37.5
        self.pdf.set_font("Arial", 'B', size=18)
        self.pdf.set_xy(10,25)
        for i,y in enumerate(range(3, 11)):
            self.pdf.cell(20,20,f'{chr(64+i+1)}',0,1,'C')
        self.pdf.set_xy(20,12)
        for i,x in enumerate(range(4, 16)):
            self.pdf.cell(20,20,f'{i+1}',0,0,'C')
        self.pdf.set_font("Arial", 'B',size=8)
        self.pdf.text(25, 190, f'Machine:_________________________________________\n')
        self.pdf.text(25, 195, f'Operator:________________________________________\n')
        self.pdf.text(25, 200, f'File Name: {self.name}\n')
        self.pdf.text(125,190,f'Notes:___________________________________________________________________________________\n')
        self.pdf.text(125,195,f'_________________________________________________________________________________________\n')
        self.pdf.text(125,200,f'_________________________________________________________________________________________\n')
    def addsamples(self,):
        self.pdf.set_font("Arial", 'B', size=6)
        for i,row in enumerate(self.data.index):
            for j,col in enumerate(self.data.columns):
                sample = self.data.loc[row,col]
                if sample != '':
                    xpos = self.xstart + (col+2)*self.cellsize
                    ypos = self.cellsize*(i+3)+3
                    self.pdf.text(xpos, ypos, sample)

    def output(self,_pName):
        self.pdf.output(f'{ppath}{_pName}.pdf')
        print(f'PDF saved to {ppath}{_pName}.pdf')
        print('Sending to printer...')
        os.startfile(f'{ppath}{_pName}.pdf','print')
PROG = RunProgram()
npath = f'{fpath}processed/{fname}'
os.rename(fpath+fname, npath)

