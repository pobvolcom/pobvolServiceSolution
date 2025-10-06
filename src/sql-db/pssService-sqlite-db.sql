/* 
File:         sql-db-pobvol.sql
Task:         Create or reset the database for pobvol Service Solution    

This file is part of the software solution 'pobvol Service Solution'. 
'pobvol Service Solution' is Free Software, delivered as open source. 
You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. The solution is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with the solution. If not, see <http://www.gnu.org/licenses/>. 
Copyright © 2025 Volker Pobloth
Web: https://pobvol.com/
*/

DROP TABLE IF EXISTS serviceaccounting;
DROP TABLE IF EXISTS tripreports;
DROP TABLE IF EXISTS servicesE;
DROP TABLE IF EXISTS servicesP;
DROP TABLE IF EXISTS checkpointstext;
DROP TABLE IF EXISTS checkpoints;
DROP TABLE IF EXISTS flexfieldstext;
DROP TABLE IF EXISTS flexfields;
DROP TABLE IF EXISTS responsetypes;
DROP TABLE IF EXISTS checkliststext;
DROP TABLE IF EXISTS checklists;
DROP TABLE IF EXISTS statustext;
DROP TABLE IF EXISTS status;
DROP TABLE IF EXISTS services;
DROP TABLE IF EXISTS servicearticles;
DROP TABLE IF EXISTS servicetypestext;
DROP TABLE IF EXISTS servicetypes;
DROP TABLE IF EXISTS articletypestext;
DROP TABLE IF EXISTS articletypes;
DROP TABLE IF EXISTS contractsP;
DROP TABLE IF EXISTS contracts;
DROP TABLE IF EXISTS contracttypestext;
DROP TABLE IF EXISTS contracttypes;
DROP TABLE IF EXISTS devices;
DROP TABLE IF EXISTS inspectioncycletext;
DROP TABLE IF EXISTS inspectioncycles;
DROP TABLE IF EXISTS devicetypestext;
DROP TABLE IF EXISTS devicetypes;
DROP TABLE IF EXISTS contacts;
DROP TABLE IF EXISTS comtypestext;
DROP TABLE IF EXISTS comtypes;
DROP TABLE IF EXISTS salutationstext;
DROP TABLE IF EXISTS salutations;
DROP TABLE IF EXISTS customers;
DROP TABLE IF EXISTS countriestext;
DROP TABLE IF EXISTS countries;
DROP TABLE IF EXISTS translations;
DROP TABLE IF EXISTS settings;
DROP TABLE IF EXISTS languagestext;
DROP TABLE IF EXISTS languages;

---------------------------------------------------------------------------------------
--
--
-- GENERAL TABLES
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS languages;
---------------------------------------------------------------------------------------
CREATE TABLE languages (
lang TEXT(2) PRIMARY KEY NOT NULL UNIQUE);
/**
@table: languages
@location: 161.68826 400.38937
*/
-- Default entries 
INSERT INTO languages (lang) VALUES ('de');
INSERT INTO languages (lang) VALUES ('en');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS languagestext;
---------------------------------------------------------------------------------------
CREATE TABLE languagestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
lang TEXT(2) NOT NULL,
lang2 TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: languagestext
@location: 162.0635 482.81882
@columnsDescription:  id() lang() lang2() text()
*/
-- Default entries 
INSERT INTO languagestext (id,lang,lang2,text) VALUES (1,'en','en','English');
INSERT INTO languagestext (id,lang,lang2,text) VALUES (2,'en','de','Englisch');
INSERT INTO languagestext (id,lang,lang2,text) VALUES (3,'de','en','German');
INSERT INTO languagestext (id,lang,lang2,text) VALUES (4,'de','de','Deutsch');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS settings;
---------------------------------------------------------------------------------------
CREATE TABLE settings (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
type TEXT(1) NOT NULL,
setting TEXT NOT NULL UNIQUE,
value TEXT NOT NULL);
/**
@table: settings
@location: 161.72122 238.68332
@columnsDescription:  id() type(system or user) setting() value()
*/
-- Default entries 
INSERT INTO settings (id,type,setting,value) VALUES (1,'S','Environment','DEV');
INSERT INTO settings (id,type,setting,value) VALUES (2,'S','Currency','EUR');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS translations;
---------------------------------------------------------------------------------------
CREATE TABLE translations (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
word TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: translations
@location: 162.02478 641.8734
@columnsDescription:  id() word() lang() text()
*/
-- Default entries 
INSERT INTO translations (id,word,lang,text) VALUES (1,'Abbrechen','de','Abbrechen');
INSERT INTO translations (id,word,lang,text) VALUES (2,'Abbrechen','en','Cancel');
---------------------------------------------------------------------------------------
--
--
-- CUSTOMERS
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS countries;
---------------------------------------------------------------------------------------
CREATE TABLE countries (
country TEXT(2) PRIMARY KEY NOT NULL,
countryISO31661 TEXT(2) NOT NULL);
/**
@table: countries
@location: 600.86804 40.19214
@columnsDescription:  country() countryISO31661()
*/
-- Default entries 
INSERT INTO countries (country,countryISO31661) VALUES ('A','AT');
INSERT INTO countries (country,countryISO31661) VALUES ('CH','CH');
INSERT INTO countries (country,countryISO31661) VALUES ('D','DE');
INSERT INTO countries (country,countryISO31661) VALUES ('L','LU');
INSERT INTO countries (country,countryISO31661) VALUES ('US','US');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS countriestext;
---------------------------------------------------------------------------------------
CREATE TABLE countriestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
country TEXT(2) NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(country) REFERENCES countries(country),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: countriestext
@location: 600.37494 134.70532
@columnsDescription:  id() country() lang() text()
*/
-- Default entries 
INSERT INTO countriestext (id,country,lang,text) VALUES (1,'A','en','Austria');
INSERT INTO countriestext (id,country,lang,text) VALUES (2,'A','de','Österreich');
INSERT INTO countriestext (id,country,lang,text) VALUES (3,'CH','en','Switzerland');
INSERT INTO countriestext (id,country,lang,text) VALUES (4,'CH','de','Schweiz');
INSERT INTO countriestext (id,country,lang,text) VALUES (5,'D','en','Germany');
INSERT INTO countriestext (id,country,lang,text) VALUES (6,'D','de','Deutschland');
INSERT INTO countriestext (id,country,lang,text) VALUES (7,'L','en','Luxembourg');
INSERT INTO countriestext (id,country,lang,text) VALUES (8,'L','de','Luxemburg');
INSERT INTO countriestext (id,country,lang,text) VALUES (9,'L','en','USA');
INSERT INTO countriestext (id,country,lang,text) VALUES (10,'L','de','USA');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS customers;
---------------------------------------------------------------------------------------
CREATE TABLE customers (
cno TEXT(16) PRIMARY KEY NOT NULL UNIQUE,
customer TEXT NOT NULL,
street TEXT NOT NULL,
zip INTEGER NOT NULL,
city TEXT NOT NULL,
country TEXT(2) NOT NULL DEFAULT D,
countryISO31661 TEXT(2) NOT NULL DEFAULT DE,
latitude DECIMAL(9,6),
longtitude DECIMAL(9,6),
FOREIGN KEY(country) REFERENCES countries(country));
/**
@table: customers
@location: 1111.4204 110.440506
*/
-- Default entries 
INSERT INTO customers (cno,customer,street,zip,city,country,countryISO31661) VALUES ('demo1','Demo1 GmbH','Wolfskaulstr. 84',66292,'Riegelsberg','D','DE');
---------------------------------------------------------------------------------------
--
--
-- CONTACTS
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS salutations;
---------------------------------------------------------------------------------------
CREATE TABLE salutations  (
sal TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: salutations 
@location: 601.6507 321.6034
*/
-- Default entries 
INSERT INTO salutations (sal) VALUES ('Mr.');
INSERT INTO salutations (sal) VALUES ('Ms.');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS salutationstext;
---------------------------------------------------------------------------------------
CREATE TABLE salutationstext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
sal TEXT(3) NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(sal) REFERENCES salutations (sal),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: salutationstext
@location: 602.0072 400.4694
@columnsDescription:  id() sal() lang() text()
*/
-- Default entries 
INSERT INTO salutationstext (id,sal,lang,text) VALUES (1,'Mr.','en','Mr.');
INSERT INTO salutationstext (id,sal,lang,text) VALUES (2,'Mr.','de','Herr');
INSERT INTO salutationstext (id,sal,lang,text) VALUES (3,'Ms.','en','Ms.');
INSERT INTO salutationstext (id,sal,lang,text) VALUES (4,'Ms.','de','Frau');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS comtypes;
---------------------------------------------------------------------------------------
CREATE TABLE comtypes (
comtype TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: comtypes
@description: Types of communication
@location: 600.7637 582.43396
*/
-- Default entries 
INSERT INTO comtypes (comtype) VALUES ('elektronisch (per E-Mail)');
INSERT INTO comtypes (comtype) VALUES ('postalisch (per Brief)');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS comtypestext;
---------------------------------------------------------------------------------------
CREATE TABLE comtypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
comtype TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(comtype) REFERENCES comtypes(comtype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: comtypestext
@location: 600.95703 698.9914
@columnsDescription:  id() comtype() lang() text()
*/
-- Default entries 
INSERT INTO comtypestext (id,comtype,lang,text) VALUES (1,'elektronisch (per E-Mail)','de','elektronisch (per E-Mail)');
INSERT INTO comtypestext (id,comtype,lang,text) VALUES (2,'elektronisch (per E-Mail)','en','electronically (by email)');
INSERT INTO comtypestext (id,comtype,lang,text) VALUES (3,'postalisch (per Brief)','de','postalisch (per Brief)');
INSERT INTO comtypestext (id,comtype,lang,text) VALUES (4,'postalisch (per Brief)','en','by post (by letter)');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS contacts;
---------------------------------------------------------------------------------------
CREATE TABLE contacts (
--id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL,
contact TEXT NOT NULL,
sal TEXT(3),
phone TEXT,
email TEXT,
lang TEXT(2) NOT NULL DEFAULT de,
comtype TEXT(8) NOT NULL,
comments TEXT,
PRIMARY KEY (cno,contact),
FOREIGN KEY(cno) REFERENCES customers(cno),
FOREIGN KEY(sal) REFERENCES salutations (sal),
FOREIGN KEY(comtype) REFERENCES comtypes(comtype));
/**
@table: contacts
@location: 1115.0328 462.9214
@columnsDescription:  id() cno(Customer number) contact(Contact name) sal(Salutation) phone(Phone number) email(Email address) lang(Language) comtype(Prefers communication by email or letter) comments()
*/
-- Default entries 
--INSERT INTO contacts (id,cno,contact,sal,phone,email,lang,comtype,comments) VALUES (1,'demo1','Herr Demo1','Mr.',NULL,NULL,'de','elektronisch (per E-Mail)',NULL);
INSERT INTO contacts (cno,contact,sal,phone,email,lang,comtype,comments) VALUES ('demo1','Herr Demo1','Mr.',NULL,NULL,'de','elektronisch (per E-Mail)',NULL);
---------------------------------------------------------------------------------------
--
--
-- DEVICES
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS devicetypes;
---------------------------------------------------------------------------------------
CREATE TABLE devicetypes (
dtype TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: devicetypes
@location: 601.29175 934.98865
*/
-- Default entries 
INSERT INTO devicetypes (dtype) VALUES ('Flurfoerderzeuge');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS devicetypestext;
---------------------------------------------------------------------------------------
CREATE TABLE devicetypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
dtype TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(dtype) REFERENCES devicetypes(dtype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: devicetypestext
@location: 602.3889 1021.98334
@columnsDescription:  id() dtype() lang() text()
*/
-- Default entries 
INSERT INTO devicetypestext (id,dtype,lang,text) VALUES (1,'Flurfoerderzeuge','de','Flurförderzeug');
INSERT INTO devicetypestext (id,dtype,lang,text) VALUES (2,'Flurfoerderzeuge','en','Industrial truck');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS inspectioncycles;
---------------------------------------------------------------------------------------
CREATE TABLE inspectioncycles (
icycle TEXT PRIMARY KEY NOT NULL UNIQUE,
icyclemm INTEGER(2) NOT NULL);
/**
@table: inspectioncycles
@location: 602.8508 1203.277
*/
-- Default entries 
INSERT INTO inspectioncycles (icycle,icyclemm) VALUES ('monthly',1);
INSERT INTO inspectioncycles (icycle,icyclemm) VALUES ('quarterly',3);
INSERT INTO inspectioncycles (icycle,icyclemm) VALUES ('semiannual',6);
INSERT INTO inspectioncycles (icycle,icyclemm) VALUES ('yearly',12);
INSERT INTO inspectioncycles (icycle,icyclemm) VALUES ('every 2nd year',24);
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS inspectioncycletext;
---------------------------------------------------------------------------------------
CREATE TABLE inspectioncycletext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
icycle TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(icycle) REFERENCES inspectioncycles(icycle),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: inspectioncycletext
@location: 601.22943 1323.3466
@columnsDescription:  id() icycle() lang() text()
*/
-- Default entries 
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (1,'monthly','de','monatlich');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (2,'monthly','en','monthly');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (3,'quarterly','de','vierteljährlich');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (4,'quarterly','en','quarterly');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (5,'semiannual','de','halbjährlich');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (6,'semiannual','en','semiannual');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (7,'yearly','de','jährlich');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (8,'yearly','en','yearly');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (9,'every 2nd year','de','alle 2 Jahre');
INSERT INTO inspectioncycletext (id,icycle,lang,text) VALUES (10,'every 2nd year','en','every 2nd year');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS devices;
---------------------------------------------------------------------------------------
CREATE TABLE devices (
--id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL,
dno TEXT(16) NOT NULL,
manufacturer TEXT NOT NULL,
device TEXT NOT NULL,
dtype TEXT NOT NULL,
articleno TEXT,
serialno TEXT,
custdeviceno TEXT,
yearmodel INTEGER(4),
poh INTEGER(6),
software TEXT,
hardware TEXT,
moduls TEXT,
location TEXT,
lat DECIMAL(9,6),
lon DECIMAL(9,6),
icycle TEXT,
icyclemm INTEGER(2),
nextinspection DATE,
warend DATE,
img BLOB,
PRIMARY KEY (cno,dno),
FOREIGN KEY(cno) REFERENCES customers(cno),
FOREIGN KEY(dtype) REFERENCES devicetypes(dtype),
FOREIGN KEY(icycle) REFERENCES inspectioncycles(icycle));
/**
@table: devices
@location: 1116.8599 803.78064
*/
-- Default entries 
--INSERT INTO devices (id,cno,dno,manufacturer,device,dtype,yearmodel,poh,location,icycle,icyclemm) VALUES (1,'demo1','bmw01','BMW','Mini Basic rot','Flurfoerderzeuge',1996,45000,'Garage links','every 2nd year',24);
INSERT INTO devices (cno,dno,manufacturer,device,dtype,yearmodel,poh,location,icycle,icyclemm)
  VALUES ('demo1','bmw01','BMW','Mini Basic rot','Flurfoerderzeuge',1996,45000,'Garage links','every 2nd year',24);
---------------------------------------------------------------------------------------
--
--
-- CONTRACTS
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS contracttypes;
---------------------------------------------------------------------------------------
CREATE TABLE contracttypes (
contracttype TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: contracttypes
@location: 601.6571 1522.3699
*/
-- Default entries 
INSERT INTO contracttypes (contracttype) VALUES ('Abo');
INSERT INTO contracttypes (contracttype) VALUES ('FV');
INSERT INTO contracttypes (contracttype) VALUES ('Garantie');
INSERT INTO contracttypes (contracttype) VALUES ('PV');
INSERT INTO contracttypes (contracttype) VALUES ('STK');
INSERT INTO contracttypes (contracttype) VALUES ('SWV');
INSERT INTO contracttypes (contracttype) VALUES ('UVV');
INSERT INTO contracttypes (contracttype) VALUES ('VWV');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS contracttypestext;
---------------------------------------------------------------------------------------
CREATE TABLE contracttypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
contracttype TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(contracttype) REFERENCES contracttypes(contracttype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: contracttypestext
@location: 600.464 1602.5759
@columnsDescription:  id() contracttype() lang() text()
*/
-- Default entries 
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (1,'Abo','de','Abonnement mit automatischer Verlängerung');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (2,'Abo','en','Auto-renewing subscription');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (3,'FV','de','Finanzierungsvertrag');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (4,'FV','en','Financing Agreement');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (5,'Garantie','de','Garantie');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (6,'Garantie','en','Guarantee');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (7,'PV','de','Preisvereinbarung');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (8,'PV','en','Price agreement');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (9,'STK','de','Sicherheitstechnische Kontrolle');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (10,'STK','en','Security related control');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (11,'SWV','de','Standard-Wartungsvertrag');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (12,'SWV','en','Standard Maintenance Contract');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (13,'UVV','de','UVV-Wartungsvertrag');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (14,'UVV','en','UVV InspectionContract');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (15,'VWV','de','Vollwartungsvertrag');
INSERT INTO contracttypestext (id,contracttype,lang,text) VALUES (16,'VWV','en','Full Maintenance Contract');
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS contracts;
---------------------------------------------------------------------------------------
CREATE TABLE contracts (
--id INTEGER  AUTOINCREMENT NOT NULL,
contractno TEXT(16) NOT NULL PRIMARY KEY UNIQUE,
cno TEXT(16) NOT NULL,
dno TEXT(16) NOT NULL,
contracttype TEXT NOT NULL,
comments TEXT,
startdatetime DATETIME NOT NULL,
enddatetime DATETIME NOT NULL,
duration INTEGER(2) NOT NULL,
terminationuntil DATETIME NOT NULL,
terminationrreceivedon DATE,
inclsrv TEXT NOT NULL,
intervalmm INTEGER(2) NOT NULL DEFAULT 0,
pdfid INTEGER,
attachment BLOB,
autorenewal BOOLEAN NOT NULL DEFAULT false,
childno TEXT,
parentno TEXT NOT NULL,
FOREIGN KEY(cno) REFERENCES devices(cno),
FOREIGN KEY(dno) REFERENCES devices(dno),
FOREIGN KEY(contracttype) REFERENCES contracttypes(contracttype));
/**
@table: contracts
@location: 1120.5042 1441.9677
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS contractsP;
---------------------------------------------------------------------------------------
CREATE TABLE contractsP (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
contractno TEXT(16) NOT NULL,
dtype TEXT NOT NULL,
arttype TEXT NOT NULL,
srvtype TEXT NOT NULL,
price DECIMAL(8,2) NOT NULL,
FOREIGN KEY(contractno) REFERENCES contracts(contractno),
FOREIGN KEY(dtype) REFERENCES servicearticles(dtype),
FOREIGN KEY(arttype) REFERENCES servicearticles(arttype),
FOREIGN KEY(srvtype) REFERENCES servicearticles(srvtype));
/**
@table: contractsP
@location: 1122.097 1962.1068
*/
---------------------------------------------------------------------------------------
--
--
-- SERVICES
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS articletypes;
---------------------------------------------------------------------------------------
CREATE TABLE articletypes (
arttype TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: articletypes
@location: 601.01843 1799.8224
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS articletypestext;
---------------------------------------------------------------------------------------
CREATE TABLE articletypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
arttype TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(arttype) REFERENCES articletypes(arttype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: articletypestext
@location: 600.69135 1882.7952
@columnsDescription:  id() arttype() lang() text()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS servicetypes;
---------------------------------------------------------------------------------------
CREATE TABLE servicetypes (
srvtype TEXT NOT NULL UNIQUE
);
/**
@table: servicetypes
@location: 2742.3652 -210.92082
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS servicetypestext;
---------------------------------------------------------------------------------------
CREATE TABLE servicetypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(srvtype) REFERENCES servicetypes(srvtype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: servicetypestext
@location: 2742.2805 -118.43279
@columnsDescription:  id() srvtype() lang() text()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS servicearticles;
---------------------------------------------------------------------------------------
CREATE TABLE servicearticles (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
dtype TEXT NOT NULL,
arttype TEXT NOT NULL,
srvtype TEXT NOT NULL UNIQUE,
price DECIMAL(8,2) NOT NULL,
duration INTEGER(4) NOT NULL,
FOREIGN KEY(dtype) REFERENCES devicetypes(dtype),
FOREIGN KEY(arttype) REFERENCES articletypes(arttype),
FOREIGN KEY(srvtype) REFERENCES servicetypes(srvtype));
/**
@table: servicearticles
@location: 601.1197 2081.0369
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS services;
---------------------------------------------------------------------------------------
CREATE TABLE services (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL,
customer TEXT NOT NULL,
street TEXT,
zip INTEGER,
city TEXT,
country TEXT(2),
contact TEXT NOT NULL,
sal TEXT(3),
phone TEXT,
email TEXT,
lang TEXT(2) NOT NULL,
comtype TEXT(8) NOT NULL,
dno TEXT(16) NOT NULL,
manufacturer TEXT NOT NULL,
device TEXT NOT NULL,
dtype TEXT NOT NULL,
articleno TEXT,
serialno TEXT,
custdeviceno TEXT,
yearmodel INTEGER(4),
poh INTEGER(6),
countryISO31661 TEXT(2) NOT NULL,
countrytext TEXT NOT NULL,
saltext TEXT,
dtypetext TEXT NOT NULL,
KEY TEXT NOT NULL UNIQUE,
version INTEGER(2) NOT NULL DEFAULT 1,
technician  TEXT NOT NULL,
software TEXT,
hardware TEXT,
moduls TEXT,
location TEXT,
difflocation TEXT NOT NULL,
lat DECIMAL(9),
lon DECIMAL(9),
warend DATE,
reportcity TEXT NOT NULL,
reportrequested BOOLEAN NOT NULL DEFAULT false,
reportcreated BOOLEAN NOT NULL DEFAULT false,
reportapproved BOOLEAN NOT NULL DEFAULT false,
reportsent BOOLEAN NOT NULL DEFAULT false,
srvtype TEXT(8) NOT NULL,
srvtypetext TEXT NOT NULL,
srvtypeicon TEXT,
chklst TEXT(8) NOT NULL,
chklsttext TEXT NOT NULL,
srvdate DATE NOT NULL,
timestamp  DATETIME NOT NULL,
srvstart DATETIME NOT NULL,
srvend DATETIME NOT NULL,
arrivaldate DATE,
barcode TEXT,
nfccode TEXT,
custorderno TEXT,
custorderdate DATE,
custorderfrom TEXT,
badgeassigned BOOLEAN NOT NULL DEFAULT false,
badgeicon TEXT,
badgemmyy TEXT,
invrequested BOOLEAN NOT NULL DEFAULT false,
invcreated BOOLEAN NOT NULL DEFAULT false,
invapproved BOOLEAN NOT NULL DEFAULT false,
invsent BOOLEAN NOT NULL DEFAULT false,
invamount DECIMAL(9,2) DEFAULT 0,
invcurrency TEXT(3) NOT NULL DEFAULT EUR,
defectclass TEXT(1) DEFAULT 0,
defectclasstext TEXT,
seccomments TEXT,
FOREIGN KEY(cno) REFERENCES customers(cno),
FOREIGN KEY(sal) REFERENCES contacts(sal),
FOREIGN KEY(dno) REFERENCES devices(dno),
FOREIGN KEY(dtype) REFERENCES devices(dtype),
FOREIGN KEY(srvtype) REFERENCES servicetypes(srvtype));
/**
@table: services
@location: 1877.4722 -1.724884

Maengelklasse.0	de			
Maengelklasse.0	en			
Maengelklasse.1	de	Ohne festgestellte Mängel		
Maengelklasse.1	en	Without identified defects		
Maengelklasse.1.1	de	Mängel beseitigt		
Maengelklasse.1.1	en	defects fixed		
Maengelklasse.2	de	Geringe Mängel		
Maengelklasse.2	en	Minor defects		
Maengelklasse.3	de	Erhebliche Mängel		
Maengelklasse.3	en	Significant deficiencies		
Maengelklasse.4	de	Gefährliche Mängel		
Maengelklasse.4	en	Dangerous deficiencies		
MaengelklassePlakette.0	de			
MaengelklassePlakette.0	en			
MaengelklassePlakette.1	de	Ohne festgestellte Mängel		
MaengelklassePlakette.1	en	Without identified defects		
MaengelklassePlakette.1.1	de	Mängel beseitigt		
MaengelklassePlakette.1.1	en	defects fixed		
MaengelklassePlakette.2	de	Geringe Mängel		
MaengelklassePlakette.2	en	Minor defects		
MaengelklassePlakette.3	de	Erhebliche Mängel		
MaengelklassePlakette.3	en	Significant deficiencies		
MaengelklassePlakette.4	de	Gefährliche Mängel		
MaengelklassePlakette.4	en	Dangerous deficiencies		
MaengelklassePlaketteText.0	de			
MaengelklassePlaketteText.0	en			
MaengelklassePlaketteText.1	de	Wir haben keine Mängel gefunden. Einer Nutzung stehen keine Bedenken entgegen. Sie bekommen die neue Prüfplakette.		
MaengelklassePlaketteText.1	en	We did not find any defects. There are no objections to its use. You will receive the new inspection sticker.		
MaengelklassePlaketteText.1.1	de	Wir haben Mängel beseitigt. Einer Nutzung stehen keine Bedenken entgegen. Sie bekommen die neue Prüfplakette.		
MaengelklassePlaketteText.1.1	en	We eliminated any defects found. There are no objections to its use. You will receive the new inspection sticker.		
MaengelklassePlaketteText.2	de	Wir haben wenige kleine Mängel gefunden. Einer Nutzung stehen keine Bedenken entgegen. Sie bekommen die neue Prüfplakette. Gekennzeichnete Mängel sollten jedoch zeitnah beseitigt werden!		
MaengelklassePlaketteText.2	en	We found a few small defects. There are no objections to its use. You will receive the new inspection sticker. However, marked defects should be eliminated promptly!		
MaengelklassePlaketteText.3	de	Wegen erheblichen Sicherheitsmängeln stehen einer Nutzung Bedenken entgegen! Sie erhalten daher keine Plakette.		
MaengelklassePlaketteText.3	en	Due to considerable safety deficiencies, there are concerns about its use! Therefore, you will not receive a sticker.		
MaengelklassePlaketteText.4	de	Wegen gefährlichen Sicherheitsmängeln, die eine direkte und unmittelbare Gefährdung darstellen oder die Umwelt beeinträchtigen, stehen einer Nutzung Bedenken entgegen! Sie erhalten daher keine Plakette.		
MaengelklassePlaketteText.4	en	Due to dangerous safety deficiencies that pose a direct and immediate hazard or affect the environment, there are concerns about its use! Therefore, you will not receive a sticker.		
MaengelklasseText.0	de			
MaengelklasseText.0	en			
MaengelklasseText.1	de	Wir haben keine Mängel gefunden. Einer Nutzung stehen keine Bedenken entgegen.		
MaengelklasseText.1	en	We did not find any defects. There are no objections to its use.		
MaengelklasseText.1.1	de	Wir haben Mängel beseitigt. Einer Nutzung stehen keine Bedenken entgegen.		
MaengelklasseText.1.1	en	We eliminated any defects found. There are no objections to its use.		
MaengelklasseText.2	de	Wir haben wenige kleine Mängel gefunden. Einer Nutzung stehen keine Bedenken entgegen. Gekennzeichnete Mängel sollten jedoch zeitnah beseitigt werden!		
MaengelklasseText.2	en	We found a few small defects. There are no objections to its use. However, marked defects should be eliminated promptly!		
MaengelklasseText.3	de	Wegen erheblichen Mängeln stehen einer Nutzung Bedenken entgegen!		
MaengelklasseText.3	en	Due to considerable deficiencies, there are concerns about its use!		
MaengelklasseText.4	de	Wegen gefährlichen Mängeln stehen einer Nutzung Bedenken entgegen!		
MaengelklasseText.4	en	Due to dangerous deficiencies there are concerns about its use!		

*/

---------------------------------------------------------------------------------------
--
--
-- SERVICES: POSITIONS
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS status;
---------------------------------------------------------------------------------------
CREATE TABLE status (
status TEXT PRIMARY KEY NOT NULL UNIQUE,
sortno INTEGER(3) NOT NULL,
defectclass INTEGER(1) NOT NULL,
icon TEXT NOT NULL);
/**
@table: status
@location: 4306.845 435.98264
@columnsDescription:  status() sortno(Sort number) defectclass(Defect class) icon()
*/
/*
Status.abgeholt,Sort:00;Maengelklasse:0;Icon:CheckBadge;	de	Inventar abgeholt		
Status.abgeholt,Sort:00;Maengelklasse:0;Icon:CheckBadge;	en	Inventory picked up		
Status.beantwortet,Sort:00;Maengelklasse:0;Icon:CheckBadge;	de	beantwortet		
Status.beantwortet,Sort:00;Maengelklasse:0;Icon:CheckBadge;	en	answered		
Status.erhebliche Mängel,Sort:03;Maengelklasse:3;Icon:EmojiSad;	de	erhebliche Mängel		
Status.erhebliche Mängel,Sort:03;Maengelklasse:3;Icon:EmojiSad;	en	significant deficiencies		
Status.erledigt,Sort:00;Maengelklasse:0;Icon:CheckBadge;	de	erledigt		
Status.erledigt,Sort:00;Maengelklasse:0;Icon:CheckBadge;	en	done		
Status.gefährliche Mängel,Sort:04;Maengelklasse:4;Icon:Warning;	de	gefährliche Mängel		
Status.gefährliche Mängel,Sort:04;Maengelklasse:4;Icon:Warning;	en	dangerous deficiencies		
Status.geringe Mängel,Sort:02;Maengelklasse:2;Icon:EmojiSmile;	de	geringe Mängel		
Status.geringe Mängel,Sort:02;Maengelklasse:2;Icon:EmojiSmile;	en	minor defects		
Status.keine Mängel,Sort:01;Maengelklasse:1;Icon:EmojiHappy;	de	keine Mängel		
Status.keine Mängel,Sort:01;Maengelklasse:1;Icon:EmojiHappy;	en	no deficiencies		
Status.Mängel beseitigt,Sort:05;Maengelklasse:1.1;Icon:Enhance;	de	Mängel beseitigt		
Status.Mängel beseitigt,Sort:05;Maengelklasse:1.1;Icon:Enhance;	en	defects fixed		
Status.nicht beantwortet,Sort:01;Maengelklasse:0;Icon:Flag;	de	nicht beantwortet		
Status.nicht beantwortet,Sort:01;Maengelklasse:0;Icon:Flag;	en	not answered		
Status.nicht erledigt,Sort:01;Maengelklasse:0;Icon:Flag;	de	offen		
Status.nicht erledigt,Sort:01;Maengelklasse:0;Icon:Flag;	en	open		
Status.nicht relevant,Sort:01;Maengelklasse:0;Icon:Blocked;	de	nicht relevant		
Status.nicht relevant,Sort:01;Maengelklasse:0;Icon:Blocked;	en	not relevant		
Status.ohne erkennbare Mängel,Sort:01;Maengelklasse:1;Icon:EmojiHappy;	de	ohne festgestellte Mängel		
Status.ohne erkennbare Mängel,Sort:01;Maengelklasse:1;Icon:EmojiHappy;	en	without identified defects		
Status.relevant,Sort:00;Maengelklasse:0;Icon:CheckBadge;	de	relevant		
Status.relevant,Sort:00;Maengelklasse:0;Icon:CheckBadge;	en	relevant		
Status.Reparatur erforderlich,Sort:03;Maengelklasse:3;Icon:Medical;	de	Reparatur erforderlich		
Status.Reparatur erforderlich,Sort:03;Maengelklasse:3;Icon:Medical;	en	to be fixed		
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS statustext;
---------------------------------------------------------------------------------------
CREATE TABLE statustext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
status TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(status) REFERENCES status(status),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: statustext
@location: 4268.8047 950.72345
@columnsDescription:  id() status() lang() text()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS checklists;
---------------------------------------------------------------------------------------
CREATE TABLE checklists (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL,
FOREIGN KEY(srvtype) REFERENCES servicetypes(srvtype));
/**
@table: checklists
@location: 3744.5095 138.68855
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS checkliststext;
---------------------------------------------------------------------------------------
CREATE TABLE checkliststext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(srvtype) REFERENCES checklists(srvtype),
FOREIGN KEY(chklist) REFERENCES checklists(chklist),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: checkliststext
@location: 3730.385 338.16077
@columnsDescription:  id() srvtype() chklist() lang() text()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS responsetypes;
---------------------------------------------------------------------------------------
CREATE TABLE responsetypes (
responsetype TEXT PRIMARY KEY NOT NULL UNIQUE);
/**
@table: responsetypes
@location: 4786.9326 779.5831
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS flexfields;
---------------------------------------------------------------------------------------
CREATE TABLE flexfields (
flexfield TEXT PRIMARY KEY NOT NULL UNIQUE,
responsetype  TEXT NOT NULL,
choices TEXT NOT NULL,
FOREIGN KEY(responsetype ) REFERENCES responsetypes(responsetype));
/**
@table: flexfields
@location: 4293.6675 746.99567
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS flexfieldstext;
---------------------------------------------------------------------------------------
CREATE TABLE flexfieldstext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
flexfield TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(flexfield) REFERENCES flexfields(flexfield),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: flexfieldstext
@location: 4288.4673 1179.7695
@columnsDescription:  id() flexfield() lang() text()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS checkpoints;
---------------------------------------------------------------------------------------
CREATE TABLE checkpoints (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL,
chkpoint TEXT NOT NULL,
pos INTEGER NOT NULL,
required BOOLEAN NOT NULL DEFAULT true,
status1 TEXT NOT NULL,
status2 TEXT,
status3 TEXT,
status4 TEXT,
status5 TEXT,
field1 TEXT,
field2 TEXT,
field3 TEXT,
field4 TEXT,
field5 TEXT,
FOREIGN KEY(srvtype) REFERENCES checklists(srvtype),
FOREIGN KEY(chklist) REFERENCES checklists(chklist),
FOREIGN KEY(status1) REFERENCES status(status),
FOREIGN KEY(status2) REFERENCES status(status),
FOREIGN KEY(status3) REFERENCES status(status),
FOREIGN KEY(status4) REFERENCES status(status),
FOREIGN KEY(status5) REFERENCES status(status),
FOREIGN KEY(field1) REFERENCES flexfields(flexfield),
FOREIGN KEY(field2) REFERENCES flexfields(flexfield),
FOREIGN KEY(field3) REFERENCES flexfields(flexfield),
FOREIGN KEY(field4) REFERENCES flexfields(flexfield),
FOREIGN KEY(field5) REFERENCES flexfields(flexfield));
/**
@table: checkpoints
@location: 3735.4766 561.67224
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS checkpointstext;
---------------------------------------------------------------------------------------
CREATE TABLE checkpointstext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL,
chkpoint TEXT NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
description  TEXT,
FOREIGN KEY(srvtype) REFERENCES checkpoints(srvtype),
FOREIGN KEY(chklist) REFERENCES checkpoints(chklist),
FOREIGN KEY(chkpoint) REFERENCES checkpoints(chkpoint),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: checkpointstext
@location: 3730.0886 1077.2927
@columnsDescription:  id() srvtype() chklist() chkpoint() lang() text() description ()
*/
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS servicesP;
---------------------------------------------------------------------------------------
CREATE TABLE servicesP (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
parentid INTEGER NOT NULL,
KEY TEXT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL,
chkpoint TEXT NOT NULL,
pos INTEGER NOT NULL,
required  BOOLEAN NOT NULL,
status TEXT NOT NULL,
comments TEXT,
img BLOB,
defectclass  INTEGER(1) NOT NULL,
field1 TEXT,
field1text TEXT,
field1val TEXT,
field2 TEXT,
field3 TEXT,
field4 TEXT,
field5 TEXT,
field2text TEXT,
field3text TEXT,
field4text TEXT,
field5text TEXT,
field2val TEXT,
field3val TEXT,
field4val TEXT,
field5val TEXT,
FOREIGN KEY(parentid) REFERENCES services(id),
FOREIGN KEY(KEY) REFERENCES services(KEY),
FOREIGN KEY(srvtype) REFERENCES checkpoints(srvtype),
FOREIGN KEY(chklist) REFERENCES checkpoints(chklist),
FOREIGN KEY(chkpoint) REFERENCES checkpoints(chkpoint),
FOREIGN KEY(pos) REFERENCES checkpoints(pos),
FOREIGN KEY(required ) REFERENCES checkpoints(required),
FOREIGN KEY(field1) REFERENCES checkpoints(field1));
/**
@table: servicesP
@description: Services Positions
@location: 2921.7983 120.44517
@columnsDescription:  id() parentid(Parent Id) KEY() srvtype(Service type) chklist(Checklist) chkpoint(Checkpoint) pos(Position) required (Required?) status(Status / Answer) comments(Comments) img(Image) defectclass (Defect class) field1(Field 1) field1text(Field 1 label) field1val(Field 1 value) field2() field3() field4() field5() field2text() field3text() field4text() field5text() field2val() field3val() field4val() field5val()
*/
---------------------------------------------------------------------------------------
--
--
-- SERVICES: SPARE PARTS / ERSATZTEILE, ZUBEHOER, EINMALARTIKEL 
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS servicesE;
---------------------------------------------------------------------------------------
CREATE TABLE servicesE (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
KEY TEXT NOT NULL,
parentid INTEGER NOT NULL,
pos INTEGER(2) NOT NULL,
Kennung TEXT(1) NOT NULL DEFAULT E,
FehlerortId TEXT,
Fehlerort TEXT,
FehlerortAnzahl INTEGER,
FehlerortEinheit TEXT,
FehlerortEinheitText TEXT,
FehlerortChargenNr TEXT,
FehlerortUntergruppe TEXT,
Einbauort TEXT,
Fehlerart TEXT,
DurchgefuehrteMassnahme TEXT,
ErsatzteilId TEXT,
ErsatzteilChargenNr TEXT,
ZubehoerId TEXT,
Zubehoer TEXT,
ZubehoerChargenNr TEXT,
EinmalartikelId TEXT,
Einmalartikel TEXT,
Anzahl DECIMAL(8,2),
Einheit TEXT,
EinheitText TEXT,
Lagerort TEXT,
price DECIMAL(8,2) NOT NULL,
vatp DECIMAL(4,2) NOT NULL,
vat DECIMAL(8,2) NOT NULL,
FOREIGN KEY(KEY) REFERENCES services(KEY),
FOREIGN KEY(parentid) REFERENCES services(id));
/**
@table: servicesE
@description: Services Ersatzteile
@location: 2920.283 1404.491

ErsatzteilEinheit.cm,Sort:02	de	cm		
ErsatzteilEinheit.cm,Sort:02	en	cm		
ErsatzteilEinheit.Liter,Sort:01	de	Liter		
ErsatzteilEinheit.Liter,Sort:01	en	liter		
ErsatzteilEinheit.Stueck,Sort:00	de	Stück		
ErsatzteilEinheit.Stueck,Sort:00	en	piece		
ErsatzteilEinheit.Paar,Sort:03	de	Paar		
ErsatzteilEinheit.Paar,Sort:03	en	pair		

ErsatzteilKennung.1,Sort:02	de	Einmalartikel		
ErsatzteilKennung.1,Sort:02	en	Single-use items		
ErsatzteilKennung.E,Sort:00	de	Ersatzteil		
ErsatzteilKennung.E,Sort:00	en	Spare part		
ErsatzteilKennung.Z,Sort:01	de	Zubehör		
ErsatzteilKennung.Z,Sort:01	en	Accessory		
*/
---------------------------------------------------------------------------------------
--
--
-- SERVICES: TRIP REPORTS
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS tripreports;
---------------------------------------------------------------------------------------
CREATE TABLE tripreports (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
type TEXT NOT NULL,
KEY TEXT NOT NULL,
parentid INTEGER NOT NULL,
technician  TEXT NOT NULL,
startdatetime DATETIME NOT NULL,
enddatetime DATETIME NOT NULL,
duration FLOAT NOT NULL,
departure  TEXT NOT NULL,
destination TEXT NOT NULL,
purpose TEXT NOT NULL DEFAULT Kundendienst,
distance INTEGER NOT NULL,
costs DECIMAL(6,2),
othercosts DECIMAL(8,2),
FOREIGN KEY(KEY) REFERENCES customers(id));
/**
@table: tripreports
@description: Travels
@location: 2920.0159 921.27966
*/
---------------------------------------------------------------------------------------
--
--
-- SERVICES: ACCOUNTING
--
--
---------------------------------------------------------------------------------------
DROP TABLE IF EXISTS serviceaccounting;
---------------------------------------------------------------------------------------
CREATE TABLE serviceaccounting (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
KEY TEXT NOT NULL,
parentid INTEGER NOT NULL,
dtype TEXT NOT NULL,
arttype TEXT NOT NULL,
srvtype TEXT NOT NULL,
text TEXT NOT NULL,
comment TEXT,
currency  TEXT(3) NOT NULL,
duration INTEGER,
pricenet DECIMAL(8,2) NOT NULL,
price DECIMAL(8,2) NOT NULL,
vatp DECIMAL(4,2) NOT NULL DEFAULT 19,
vat DECIMAL(8,2) NOT NULL,
warranty BOOLEAN NOT NULL DEFAULT false,
contract BOOLEAN NOT NULL DEFAULT false,
contracttype TEXT,
contracttypetext TEXT,
contractno TEXT(16),
contractstartdatetime DATETIME,
FOREIGN KEY(KEY) REFERENCES services(KEY),
FOREIGN KEY(parentid) REFERENCES services(id),
FOREIGN KEY(dtype) REFERENCES services(dtype));
/**
@table: serviceaccounting
@description: Service accounting
@location: 2915.2805 2265.7622
*/
---------------------------------------------------------------------------------------
--
--
-- EOF / END OF FILE
--
--
---------------------------------------------------------------------------------------