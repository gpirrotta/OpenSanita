## Opensanità, dati sanitari aperti 2001-2010 

Script (**NON OTTIMIZZATO**) per la generazione in un unico file dei dati sanitari in formato csv


Dataset originali: [Ministero della Salute](http://www.salute.gov.it/portale/temi/p2_6.jsp?lingua=italiano&id=1314&area=programmazioneSanitariaLea&menu=vuoto)

La directory xls contiene i file originali normalizzati.

Per generare il file:

``` bash
python extract.py > opensanita.csv
```