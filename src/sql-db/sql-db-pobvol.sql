DROP TABLE IF EXISTS customers;
DROP TABLE IF EXISTS countries;
DROP TABLE IF EXISTS contacts;
DROP TABLE IF EXISTS devices;
DROP TABLE IF EXISTS devicetypes;
DROP TABLE IF EXISTS salutations ;
DROP TABLE IF EXISTS comtypes;
DROP TABLE IF EXISTS inspectioncycles;
DROP TABLE IF EXISTS services;
DROP TABLE IF EXISTS languages;
DROP TABLE IF EXISTS status;
DROP TABLE IF EXISTS statustext;
DROP TABLE IF EXISTS flexfields;
DROP TABLE IF EXISTS flexfieldstext;
DROP TABLE IF EXISTS checkpoints;
DROP TABLE IF EXISTS checkpointstext;
DROP TABLE IF EXISTS checklists;
DROP TABLE IF EXISTS checkliststext;
DROP TABLE IF EXISTS servicetypes;
DROP TABLE IF EXISTS languagestext;
DROP TABLE IF EXISTS countriestext;
DROP TABLE IF EXISTS devicetypestext;
DROP TABLE IF EXISTS salutationstext;
DROP TABLE IF EXISTS comtypestext;
DROP TABLE IF EXISTS inspectioncycletext;
DROP TABLE IF EXISTS servicetypestext;
DROP TABLE IF EXISTS responsetypes;
DROP TABLE IF EXISTS servicesP;
DROP TABLE IF EXISTS contracts;
DROP TABLE IF EXISTS contracttypes;
DROP TABLE IF EXISTS contracttypestext;
DROP TABLE IF EXISTS contractsP;
DROP TABLE IF EXISTS servicearticles;
DROP TABLE IF EXISTS articletypes;
DROP TABLE IF EXISTS articletypestext;
DROP TABLE IF EXISTS settings;
DROP TABLE IF EXISTS translations;
DROP TABLE IF EXISTS tripreports;
DROP TABLE IF EXISTS serviceaccounting;
DROP TABLE IF EXISTS servicesE;


CREATE TABLE customers (
id BIGINT PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL UNIQUE,
customer TEXT NOT NULL,
street TEXT NOT NULL,
zip INTEGER NOT NULL,
city TEXT NOT NULL,
country TEXT(2) NOT NULL DEFAULT de,
lat DECIMAL(9,6),
lon DECIMAL(9,6),
FOREIGN KEY(country) REFERENCES countries(country));
/**
@table: customers
@location: 1111.4204 110.440506
*/

CREATE TABLE countries (
country TEXT(2) PRIMARY KEY NOT NULL DEFAULT de,
countryISO31661 TEXT(2) NOT NULL);
/**
@table: countries
@location: 600.86804 40.19214
@columnsDescription:  country() countryISO31661()
*/

CREATE TABLE contacts (
id BIGINT PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL,
contact TEXT NOT NULL,
sal TEXT(3),
phone TEXT,
email TEXT,
lang TEXT(2) NOT NULL DEFAULT de,
comtype TEXT(8) NOT NULL,
comments TEXT,
FOREIGN KEY(cno) REFERENCES customers(cno),
FOREIGN KEY(sal) REFERENCES salutations (sal),
FOREIGN KEY(comtype) REFERENCES comtypes(comtype));
/**
@table: contacts
@location: 1115.0328 462.9214
@columnsDescription:  id() cno(Customer number) contact(Contact name) sal(Salutation) phone(Phone number) email(Email address) lang(Language) comtype(Prefers communication by email or letter) comments()
*/

CREATE TABLE devices (
id BIGINT PRIMARY KEY AUTOINCREMENT NOT NULL,
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
icycle TEXT(2),
icyclemm INTEGER(2),
nextinspection DATE,
warend DATE,
img BLOB,
FOREIGN KEY(cno) REFERENCES customers(cno),
FOREIGN KEY(dtype) REFERENCES devicetypes(dtype),
FOREIGN KEY(icycle) REFERENCES inspectioncycles(icycle));
/**
@table: devices
@location: 1116.8599 803.78064
*/

CREATE TABLE devicetypes (
dtype TEXT PRIMARY KEY NOT NULL);
/**
@table: devicetypes
@location: 601.29175 934.98865
*/

CREATE TABLE salutations  (
sal TEXT(3) PRIMARY KEY NOT NULL DEFAULT Mr);
/**
@table: salutations 
@location: 601.6507 321.6034
*/

CREATE TABLE comtypes (
comtype TEXT(8) PRIMARY KEY NOT NULL DEFAULT email);
/**
@table: comtypes
@description: Types of communication
@location: 600.7637 582.43396
*/

CREATE TABLE inspectioncycles (
icycle TEXT(2) PRIMARY KEY NOT NULL,
icyclemm INTEGER(2) NOT NULL);
/**
@table: inspectioncycles
@location: 602.8508 1203.277
*/

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
*/

CREATE TABLE languages (
lang TEXT(2) PRIMARY KEY NOT NULL DEFAULT de);
/**
@table: languages
@location: 161.68826 400.38937
*/

CREATE TABLE status (
status TEXT PRIMARY KEY NOT NULL,
sortno INTEGER(3) NOT NULL,
defectclass INTEGER(1) NOT NULL,
icon TEXT NOT NULL);
/**
@table: status
@location: 4306.845 435.98264
@columnsDescription:  status() sortno(Sort number) defectclass(Defect class) icon()
*/

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

CREATE TABLE flexfields (
flexfield TEXT PRIMARY KEY NOT NULL,
responsetype  TEXT NOT NULL,
choices TEXT NOT NULL,
FOREIGN KEY(responsetype ) REFERENCES responsetypes(responsetype));
/**
@table: flexfields
@location: 4293.6675 746.99567
*/

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

CREATE TABLE checklists (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
srvtype TEXT NOT NULL,
chklist TEXT NOT NULL UNIQUE,
FOREIGN KEY(srvtype) REFERENCES servicetypes(srvtype));
/**
@table: checklists
@location: 3744.5095 138.68855
*/

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

CREATE TABLE servicetypes (
srvtype TEXT NOT NULL);
/**
@table: servicetypes
@location: 2742.3652 -210.92082
*/

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

CREATE TABLE comtypestext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
comtype TEXT(8) NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(comtype) REFERENCES comtypes(comtype),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: comtypestext
@location: 600.95703 698.9914
@columnsDescription:  id() comtype() lang() text()
*/

CREATE TABLE inspectioncycletext (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
icycle TEXT(2) NOT NULL,
lang TEXT(2) NOT NULL,
text TEXT NOT NULL,
FOREIGN KEY(icycle) REFERENCES inspectioncycles(icycle),
FOREIGN KEY(lang) REFERENCES languages(lang));
/**
@table: inspectioncycletext
@location: 601.22943 1323.3466
@columnsDescription:  id() icycle() lang() text()
*/

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

CREATE TABLE responsetypes (
responsetype TEXT PRIMARY KEY NOT NULL);
/**
@table: responsetypes
@location: 4786.9326 779.5831
*/

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

CREATE TABLE contracts (
id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
cno TEXT(16) NOT NULL,
dno TEXT(16) NOT NULL,
contractno TEXT(16) NOT NULL UNIQUE,
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

CREATE TABLE contracttypes (
contracttype TEXT PRIMARY KEY NOT NULL);
/**
@table: contracttypes
@location: 601.6571 1522.3699
*/

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

CREATE TABLE articletypes (
arttype TEXT PRIMARY KEY NOT NULL);
/**
@table: articletypes
@location: 601.01843 1799.8224
*/

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
*/

