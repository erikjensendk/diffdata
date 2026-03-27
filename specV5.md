# Oracle–Salesforce Datasammenligningsværktøj v5.0

## Specifikation til Claude Code

---

## 1. Projektbeskrivelse

Et desktop-værktøj (Tkinter) der sammenligner data fra en Oracle-database (STADS) med data fra Salesforce. Formålet er at validere en datamigration — alle felter fra et Oracle-view skal findes og være identiske i Salesforce.

Værktøjet erstatter den tidligere CSV-baserede sammenligning (v4.0), hvor to CSV-filer blev sammenlignet. I v5.0 tilgås de to datakilder direkte via databaseforbindelser.

---

## 2. Arkitektur og dataflow

```
Oracle (STADS)                    Salesforce
     │                                 │
     │  oracledb (thin mode)           │  simple_salesforce
     │                                 │
     ▼                                 ▼
Hent feltliste fra view          Byg SOQL automatisk
Hent alle rækker                 Hent alle rækker (query_all)
     │                                 │
     └──────────────┬──────────────────┘
                    ▼
         Sammenlign række-for-række
         Join-nøgle: LegacyId_m__c
         Alle Oracle-felter sammenlignes
                    │
                    ▼
              Rapport + Excel
```

---

## 3. Teknisk stack

| Komponent      | Bibliotek              |
| -------------- | ---------------------- |
| Oracle         | `oracledb` (thin mode) |
| Salesforce     | `simple_salesforce`    |
| UI             | `tkinter`              |
| Excel-eksport  | `openpyxl`             |
| Databehandling | `pandas`               |
| Konfiguration  | `json` (gemt lokalt)   |

### requirements.txt skal indeholde:

```
oracledb
simple-salesforce
openpyxl
pandas
```

---

## 4. Funktionelle krav

### 4.1 Forbindelsesstyring

- Brugeren indtaster Oracle-forbindelsesoplysninger: host, port, service name, username, password
- Brugeren indtaster Salesforce-forbindelsesoplysninger: username, password, security token, domain
- Alle forbindelsesoplysninger gemmes i en lokal `config.json`
- Passwords gemmes i `config.json` (ingen keychain-integration i denne version)
- "Test forbindelse"-knap for begge systemer

### 4.2 Guidet opsætning

- Brugeren vælger Oracle-view fra en dropdown (populeret automatisk fra databasen)
- Salesforce-objekt indtastes manuelt
- `LegacyId_m__c` er fast join-nøgle — ikke konfigurerbar af brugeren

### 4.3 Sammenligning

- Alle felter fra Oracle-view hentes automatisk
- SOQL bygges automatisk ud fra feltlisten fra Oracle
- `attributes`-feltet fjernes fra Salesforce-records (Salesforce metadata)
- Sammenligning er celle-for-celle via join-nøglen `LegacyId_m__c`
- Numeriske felter normaliseres via `Decimal` (samme logik som v4.0)
- Feltnavne antages at være identiske i Oracle og Salesforce

### 4.4 Rapport

- Samme rapportformat som v4.0 (se afsnit 7)
- Eksport til `.txt`-fil
- Eksport til `sammenlign.xlsx` (samme kolonner som v4.0)

---

## 5. UI-layout (Tkinter)

Applikationen har to faner:

```
┌─────────────────────────────────────────────┐
│  Faner:  [Sammenligning]  [Forbindelser]    │
├─────────────────────────────────────────────┤
│  FORBINDELSER-FANE:                         │
│                                             │
│  Oracle:                                    │
│    Host:           [tekstfelt]              │
│    Port:           [tekstfelt]  (1521)      │
│    Service name:   [tekstfelt]              │
│    Username:       [tekstfelt]              │
│    Password:       [password-felt]          │
│    [Test Oracle-forbindelse]                │
│                                             │
│  Salesforce:                                │
│    Username:       [tekstfelt]              │
│    Password:       [password-felt]          │
│    Security token: [tekstfelt]              │
│    Domain:         [tekstfelt]  (test)      │
│    [Test Salesforce-forbindelse]            │
│                                             │
│    [Gem forbindelser]                       │
├─────────────────────────────────────────────┤
│  SAMMENLIGNING-FANE:                        │
│                                             │
│  Oracle view:       [dropdown]  [Hent views]│
│  Salesforce objekt: [tekstfelt]             │
│  Klynge:            [tekstfelt]             │
│                                             │
│  [Sammenlign]  [Ryd]  [Download rapport]   │
│  [Eksporter til sammenlign.xlsx]  [STATUS]  │
│                                             │
│  ┌─────────────────────────────────────┐   │
│  │  Rapport (scrollbar tekstboks)      │   │
│  └─────────────────────────────────────┘   │
└─────────────────────────────────────────────┘
```

---

## 6. Konfigurationsfil (`config.json`)

```json
{
  "oracle": {
    "host": "",
    "port": 1521,
    "service_name": "",
    "username": "",
    "password": ""
  },
  "salesforce": {
    "username": "",
    "password": "",
    "security_token": "",
    "domain": "test"
  }
}
```

- Filen oprettes automatisk hvis den ikke findes
- Indlæses ved opstart og udfylder felterne i UI

---

## 7. Rapportformat

Rapporten skal følge samme struktur som v4.0:

```
========================================
         DATA SAMMENLIGNINGSRAPPORT
========================================

Klynge: <klynge>
Oracle view: <view-navn>
Salesforce objekt: <objekt-navn>
Dato for sammenligning: <timestamp>
Diff-script version: 5.0

----------------------------------------
1. Rækketælling
----------------------------------------
Antal rækker i Oracle:      <n>
Antal rækker i Salesforce:  <n>

----------------------------------------
2. Numeriske kolonne-summer
----------------------------------------
<kolonne>:
  SUM Oracle:      <sum>
  SUM Salesforce:  <sum>

----------------------------------------
3. Kolonne- og cellematch
----------------------------------------
Antal kolonner sammenlignet:    <n>
Antal celler sammenlignet:      <n>
Antal afvigelser fundet:        <n>

----------------------------------------
4. Resultat
----------------------------------------
STATUS: IDENTISKE / FORSKELLE FUNDET

<liste af afvigelser>

----------------------------------------
5. Runtime-information
----------------------------------------
Script eksekveringstid: <ms> ms
Miljø/Maskine: <hostname>

========================================
                SLUT PÅ RAPPORT
========================================
```

### Excel-eksport (`sammenlign.xlsx`) kolonner:

- tidsstempel
- klynge
- oracle_view
- salesforce_objekt
- primaernoegle_kolonne
- raekker_oracle
- raekker_salesforce
- kolonner_sammenlignet
- celler_sammenlignet
- afvigelser
- status
- eksekveringstid_ms
- maskine
- script_version

---

## 8. Salesforce-integration (reference)

```python
from simple_salesforce import Salesforce

sf = Salesforce(
    username="bruger@domæne.dk",
    password="hemligt",
    security_token="token",
    domain="test"  # "test" for sandbox, udelad for production
)

# Byg SOQL automatisk fra Oracle-feltliste
fields = ", ".join(oracle_fields)
query = f"SELECT {fields} FROM {salesforce_object}"
result = sf.query_all(query)
records = result["records"]

# Fjern Salesforce metadata
rows = [{k: v for k, v in r.items() if k != "attributes"} for r in records]
```

---

## 9. Oracle-integration (reference)

```python
import oracledb

connection = oracledb.connect(
    user="username",
    password="password",
    dsn="host:port/service_name"
)

# Hent liste af views brugeren kan vælge imellem
cursor = connection.cursor()
cursor.execute("SELECT view_name FROM user_views ORDER BY view_name")
views = [row[0] for row in cursor.fetchall()]

# Hent feltliste fra valgt view
cursor.execute(f"SELECT * FROM {view_name} WHERE ROWNUM = 1")
fields = [col[0] for col in cursor.description]

# Hent alle rækker
cursor.execute(f"SELECT * FROM {view_name}")
rows = cursor.fetchall()
```

---

## 10. Out of scope (v5.0)

- Kolonnemapping (feltnavne antages ens i Oracle og Salesforce)
- Sikker password-opbevaring (keychain el. lign.)
- Flere Oracle-views på én gang
- Scheduling eller automatisering
- CSV-import (erstattet af direkte databaseforbindelser)
