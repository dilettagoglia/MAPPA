<p align="center">
    <img align="center" src='https://www.farm.unipi.it/wp-content/uploads/2021/11/Logo_Farmacia-e1638259967379.png' width="50px">
    <img align='center' src='https://dscm.dcci.unipi.it/images/news/logo_unipi_blu.jpg' width='100px'>
</p>


# MAPPA project: Medicinal and Aromatic Plant Procedures for Authentication.
<a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/3.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-nc-nd/3.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/3.0/">Creative Commons Attribution-NonCommercial-NoDerivs 3.0 Unported License</a>.

_This repository contains the code that retrieves the content of the MAPPA database and
transform it in structured cards for each plant in the database, both in Word and in PDF format._

<a href="https://github.com/dilettagoglia/MAPPA/blob/main/LICENSE"><img src="https://img.shields.io/github/license/dilettagoglia/MAPPA" alt="License" /></a>
<a href="https://github.com/dilettagoglia/MAPPA/stargazers"><img src="https://img.shields.io/github/stars/dilettagoglia/MAPPA" alt="GitHub stars" /></a>
<a href="https://github.com/dilettagoglia/MAPPA/network/members"><img alt="GitHub forks" src="https://img.shields.io/github/forks/dilettagoglia/MAPPA" /></a>

## Description
The database in .xlsx format and the code contained in this repo have been developed for the 
MAPPA project (Medicinal and Aromatic Plant Procedures for Authentication) of Pharmacy Department, University of Pisa.

## Directory structure (main elements)
```
MAPPA
  │── src
  │    │── tables.py                   
  │    │── params.py                          
  │    │── utilities.py
  │    └── main.py                    # file to run
  └── database
  │    └── MAPPA_database.xlsx        # access allowed only to project members
  └── export
  │   │── doc      
  │   │    └── ...                    # produced cards in .docx format
  │   └── pdf     
  │        └── ...                    # produced carda in .pdf format
  └── img          
  │   └── logo.jpg                    # project logo 
  └── material          
  │   │── Guida.docx                  # project guide for collaborators 
  │   │── MAPPA_poster_en.pdf
  │   └── MAPPA_poster_it.pdf
  └── requirements.txt
  └── README.md
  └── LICENSE  
```

## Quick start
Install Python:<br>
`sudo apt install python3`

Install pip:<br>
`sudo apt install --upgrade python3-pip`

Install requirements:<br>
`python -m pip install --requirement requirements.txt`

Execute [main](src/main.py)
```
cd src/
python main.py
```

## Corresponding author
**Dr. Diletta Goglia** <a href="https://orcid.org/0000-0002-2622-7495"><img alt="ORCID logo" src="https://info.orcid.org/wp-content/uploads/2019/11/orcid_16x16.png" width="16" height="16" /></a> <br/>
**Postgraduate Student in MSc in Artificial Intelligence** <br/>
**Computer Science department, University of Pisa, Italy** <br/>
[d.goglia@studenti.unipi.it](mailto:d.goglia@studenti.unipi.it) <br/>
[dilettagoglia.netlify.app](http://www.dilettagoglia.netlify.app) 

_Candidata vincitrice dell'incarico di collaborazione a seguito del [bando dedicato 
all'attività di supporto all’interno del progetto MAPPA](https://www.farm.unipi.it/wp-content/uploads/2022/04/PD-253-prot-1986-del-12-04-2022-Bando-PSd-Bertoli_150-ore.pdf) per lo sviluppo di una banca 
dati di settore on-line dedicata all’inquadramento di materie prime e derivati 
a base di piante officinali mediante catalogazione di dati etno-botanico-farmaceutici selezionati_


## Project info
**Coordinator**: Alessandra Bertoli, associated professor at Pharmacy Department, University of Pisa
via Bonanno 33 56126 PISA (Italy), [alessandra.bertoli@unipi.it](alessandra.bertoli@unipi.it)
