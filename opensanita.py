import xlrd
import glob
import os

files =  glob.glob("xls/*.xls")
regioni = {'010': 'PIEMONTE', '020': 'VALLE D\'AOSTA', '030':'LOMBARDIA', '041': 'PROVINCIA AUTONOMA BOLZANO',
		   '042': 'PROVINCIA AUTONOMA TRENTO', '050': 'VENETO', '060': 'FRIULI VENEZIA GIULIA', '070': 'LIGURIA',
		   '080': 'EMILIA ROMAGNA', '090': 'TOSCANA', '100': 'UMBRIA', '110': 'MARCHE', '120': 'LAZIO',
		   '130': 'ABRUZZO', '140': 'MOLISE', '150': 'CAMPANIA', '160': 'PUGLIA', '170': 'BASILICATA',
		   '180': 'CALABRIA', '190': 'SICILIA', '200': 'SARDEGNA'}
values = {}
	
for f in files:
	anno = f[-11:-7]
	regione = regioni[f[-7:-4]]
	if regione not in values.keys():
		values[regione] = {}
		
	values[regione][anno] = {}
	values[regione][anno]['nomeFile'] = f
	wb = xlrd.open_workbook(f)
	sh = wb.sheet_by_index(1)
	asls = {}
	values[regione][anno] = asls
	
	for rownum in range(1, sh.nrows):
		aslCode = int(sh.cell(rownum,0).value)
		if aslCode != '':
			asls[aslCode] = {} 
			asls[aslCode]['struttura'] = str(sh.cell(rownum,1).value)
	
	sh = wb.sheet_by_index(0)
	rowHead = 0
	colHead = 1
	numCols = 0
	for rownum in range(0, sh.nrows):
		labelValue = sh.cell(rownum,0).value
		labelValueNext = sh.cell(rownum,1).value
		if labelValue == '' and labelValueNext != '':
			rowHead = rownum
			numCols = len(sh.row_values(rownum))
			break
		
	for rownum in range(rowHead+1, sh.nrows):
		firstCol = sh.cell(rownum,0).value
		if '9999' in firstCol:
			tipoVoce =  firstCol[0]
			for col in range(colHead, numCols):
				asl = sh.cell(rowHead, col).value
				if asl != '':
					asl = int(asl)
					if asl in asls.keys():
						importo = sh.cell(rownum, col).value
						if importo != '':
							asls[asl][tipoVoce] = importo
						else:	
							asls[asl][tipoVoce] = 0

print 'REGIONE;ANNO;STRUTTURA;Totale valore della produzione (A); \
	   Totale costi della produzione (B);Totale proventi e oneri finanziari (C); \
	   Totale rettifiche di valore di attivita\' finanziarie (D);Totale proventi e oneri straordinari (E); \
	   Totale imposte e tasse;RISULTATO DI ESERCIZIO'


for regione, value in values.iteritems():
	for anno, value2 in value.iteritems():
		for struttura, value3 in value2.iteritems():
			struttura = value3['struttura'].strip()
			a=b=c=d=e=y=z= 0
			if 'A' in value3.keys():
				a =  value3['A']
			if 'B' in value3.keys():
				b = value3['B']
			if 'C' in value3.keys():
				c = value3['C']
			if 'D' in value3.keys():
				d = value3['D']
			if 'E' in value3.keys():
				e = value3['E']
			if 'Y' in value3.keys():
				y = value3['Y']
			if 'Z' in value3.keys():
				z = value3['Z']
			if a == 0 and b == 0 and c == 0 and d == 0 and e == 0 and y == 0 and z == 0:
				continue
			print("%s;%s;%s;%d;%d;%d;%d;%d;%d;%d" % (regione, anno, struttura, a,b,c,d,e,y,z))
			