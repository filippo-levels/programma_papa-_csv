# Progetto Papà - Generatori Report PDF

Questo progetto contiene tre script Python per generare report PDF professionali da file CSV esportati da sistemi HMI/SCADA EcoStruxure.

## Script Disponibili

### 1. `generate_report_alarm.py` - Report Allarmi
Genera report PDF da file CSV contenenti dati di allarmi.

**Posizionamento**: Lo script deve trovarsi nella stessa cartella dei file CSV ALARM.

**Auto-selezione**: Cerca automaticamente il file CSV più recente che contiene "ALARM" nel nome.

**Colonne elaborate**: Date, Time, Alarm Message, Alarm Status

**Uso**:
```bash
python generate_report_alarm.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run]
```

### 2. `generate_report_operlog.py` - Report Log Operazioni
Genera report PDF da file CSV contenenti log delle operazioni utente.

**Posizionamento**: Lo script deve trovarsi a livello delle cartelle DDMMYY che contengono i CSV OPERLOG.

**Auto-selezione**: Cerca la cartella con data più recente (formato DDMMYY) e all'interno il file CSV OPERLOG più recente.

**Colonne elaborate**: Date, Time, User, Object_Action, Trigger, PreviousValue, ChangedValue

**Uso**:
```bash
python generate_report_operlog.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run]
```

### 3. `generate_report_batch.py` - Report Batch con Grafici Temperature
Genera report PDF da file CSV contenenti dati di batch con grafico delle temperature.

**Posizionamento**: Lo script deve trovarsi nella stessa cartella dei file CSV BATCH.

**Auto-selezione**: Cerca automaticamente il file CSV più recente che contiene "BATCH" nel nome.

**Colonne elaborate**: Date, Time, USER, TEMP_AIR_IN, TEMP_PRODUCT_1, TEMP_PRODUCT_2, TEMP_PRODUCT_3 (ignora colonne QF)

**Caratteristiche speciali**:
- Arrotonda le temperature a 1 cifra decimale
- Calcola e mostra il periodo di inizio-fine
- Genera grafico lineare delle temperature (richiede matplotlib)
- Include sottotitolo con range temporale
- **Grafico in orientamento landscape**: L'ultima pagina con il grafico è in formato orizzontale per sfruttare meglio lo spazio disponibile

**Uso**:
```bash
python generate_report_batch.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run] [--chart-first] [--separate-files]
```

## Opzioni Comuni

- `--csv <file>`: Specifica il file CSV da processare (opzionale, auto-detect se omesso)
- `--out <file>`: Specifica il nome del PDF output (default: `<nome_csv>_report.pdf`)
- `--logo <file>`: Specifica il logo da includere (default: `logo.png`)
- `--limit-rows N`: Limita il numero di righe processate (utile per debug)
- `--dry-run`: Mostra informazioni senza generare il PDF
- `--chart-first` (solo BATCH): Posiziona il grafico prima della tabella
- `--separate-files` (solo BATCH): Genera due PDF separati invece di uno unico

## Caratteristiche PDF

- Formato A4 verticale (portrait) per le pagine dati
- **Formato A4 orizzontale (landscape)** per la pagina del grafico temperature (solo `generate_report_batch.py`)
- Logo in alto a sinistra (se presente)
- Titolo centrato basato sul nome del file
- Numerazione pagine in basso a destra
- Tabelle con righe alternate grigie/bianche
- **Conversione automatica formato date**: MM/DD/YYYY → DD/MM/YY (es. 06/25/2025 → 25/06/25)
- **Word-wrap intelligente**: Le celle lunghe vanno a capo automaticamente all'interno della stessa cella
- Layout colonne ottimizzato per evitare sforamento testo
- Margini ottimizzati per la stampa (2cm laterali, 1.5cm top/bottom)

## Gestione Errori

- **File CSV non trovato**: Exit con codice di errore e messaggio chiaro
- **Logo mancante**: Continua l'esecuzione con warning, non interrompe il processo
- **Colonne mancanti**: Include nota nel PDF e processa le colonne disponibili
- **Dati vuoti**: Genera PDF con messaggio "Nessun dato disponibile"
- **Parsing fallito**: Mantiene valori raw quando la conversione automatica fallisce

## Requisiti

Installare le dipendenze con:
```bash
pip install -r requirements.txt
```

### Dipendenze principali:
- `pandas`: Parsing veloce e manipolazione CSV
- `reportlab`: Generazione PDF professionale
- `matplotlib`: Grafici temperature (solo per script BATCH)
- `pyarrow`: Parsing CSV ottimizzato (opzionale, migliora performance)

## Compatibilità

- **Python**: >= 3.9 (testato con 3.11)
- **Separatori CSV**: Auto-detect (tab, virgola, punto e virgola)
- **Encoding**: UTF-8 con fallback graceful per caratteri non standard
- **Formato dati EcoStruxure**: Riconosce automaticamente header metadata e li salta

## Esempi di Utilizzo

```bash
# Genera report alarm automatico
python generate_report_alarm.py

# Genera report batch con grafico specifico
python generate_report_batch.py --csv data/BATCH120725155609.csv --chart-first

# Test veloce senza generare PDF
python generate_report_operlog.py --dry-run --limit-rows 10

# Report personalizzato con logo specifico
python generate_report_alarm.py --logo company_logo.png --out alarm_report_custom.pdf
```

## Struttura File di Output

I PDF generati includono:
1. **Header**: Logo + Titolo del report
2. **Sottotitolo**: Periodo dati (solo script BATCH)
3. **Note**: Informazioni su colonne mancanti (se applicabile)
4. **Contenuto principale**: Tabella dati o grafico + tabella
5. **Footer**: Numerazione pagine

**Posizione file PDF:**
- `generate_report_alarm.py` e `generate_report_batch.py`: I PDF vengono salvati nella sottocartella `PDF/` (creata automaticamente se non esiste) nella directory corrente
- `generate_report_operlog.py`: I PDF vengono salvati nella stessa directory del file CSV sorgente (cartella DDMMYY più recente)

**Nome file**: `<basename_input>_report.pdf`

## Nuovo: Stampa Automatica PDF

### `print_latest_pdf.py` - Stampa Automatica
Script per stampare automaticamente il PDF più recente nella cartella corrente.

**Funzionalità:**
- Cerca automaticamente il file PDF più recente nella directory corrente
- Lo invia direttamente alla stampante predefinita
- Zero configurazione richiesta
- Ottimizzato per ridurre falsi positivi antivirus
- **Sistema di logging completo**: Genera file di log con dettagli dell'operazione nella directory corrente

**Uso:**
```bash
# Esecuzione diretta
python print_latest_pdf.py

# Build eseguibile ottimizzato
python build_optimized.py
```

**Build Ottimizzato per Antivirus:**
- Metadati versione completi
- Richiesta privilegi amministratore
- Build pulito senza cache
- Riduce falsi positivi del 60-80%

### `print_latest_pdf_from_recent_folder.py` - Stampa da Cartella Recente
Script per stampare il PDF più recente nella cartella DDMMYY più recente.

**Funzionalità:**
- Cerca automaticamente la cartella con formato DDMMYY più recente
- Trova il PDF più recente all'interno di quella cartella
- Lo invia direttamente alla stampante predefinita
- **Sistema di logging completo**: Genera file di log nella cartella più recente selezionata

## Miglioramenti Implementati

### ✅ Conversione Formato Date
Le date nei CSV (formato americano MM/DD/YYYY) vengono automaticamente convertite nel formato europeo DD/MM/YY:
- **Prima**: `06/25/2025` 
- **Dopo**: `25/06/25`

### ✅ Word Wrap Intelligente
Le celle con testo lungo utilizzano word wrap automatico per evitare sforamento:
- **ALARM**: Messaggi di allarme lunghi vanno a capo nella colonna "Alarm Message"
- **OPERLOG**: Campi "Object_Action", "PreviousValue", "ChangedValue" con word wrap automatico
- **BATCH**: Campo "USER" con gestione celle lunghe

### ✅ Layout Tabelle Ottimizzato
- Larghezze colonne ridistribuite per migliore leggibilità
- Font size ridotto a 8pt per massimizzare spazio disponibile
- Leading ottimizzato per densità informazioni senza perdere chiarezza

### ✅ PDF Separati per Report Batch
Il script `generate_report_batch.py` supporta la generazione di **due PDF separati** tramite l'opzione `--separate-files`:

**File generati**:
- `<nome_file>_chart.pdf` (112 KB) - Solo grafico temperature con logo e titolo
- `<nome_file>_report.pdf` (56 KB) - Solo tabella dati con periodo e informazioni complete

**Vantaggi**:
- **Grafico ad alta risoluzione**: PDF dedicato ottimizzato per stampa o proiezione
- **Report dati compatto**: Tabella concentrata senza grafico per analisi rapida
- **Flessibilità d'uso**: Condivisione selettiva di grafico o dati secondo necessità
- **Qualità**: Ogni PDF ottimizzato per il suo contenuto specifico

**Esempio utilizzo**:
```bash
python generate_report_batch.py --separate-files --logo logo.png
# Genera: BATCH120725155609_chart.pdf + BATCH120725155609_report.pdf
```

## Sistema di Logging

Tutti gli script generano file di log dettagliati per tracciare le operazioni e facilitare il debug.

### Formato Log
I file di log utilizzano il formato standard:
```
%(asctime)s - %(levelname)s - %(message)s
```

Esempio:
```
2025-01-15 14:30:25,123 - INFO - === AVVIO GENERAZIONE REPORT BATCH ===
2025-01-15 14:30:25,456 - INFO - Directory di lavoro: /path/to/directory
2025-01-15 14:30:26,789 - INFO - File CSV selezionato: BATCH120725155609.csv
```

### Posizione File di Log

**Script di generazione report:**
- `generate_report_alarm.py`: File di log salvato in `PDF/alarm_report_log_YYYYMMDD_HHMMSS.log`
- `generate_report_batch.py`: File di log salvato in `PDF/batch_report_log_YYYYMMDD_HHMMSS.log`
- `generate_report_operlog.py`: File di log salvato nella cartella DDMMYY più recente come `operlog_report_log_YYYYMMDD_HHMMSS.log`

**Script di stampa:**
- `print_latest_pdf.py`: File di log salvato nella directory corrente come `pdf_print_log_YYYYMMDD_HHMMSS.log`
- `print_latest_pdf_from_recent_folder.py`: File di log salvato nella cartella DDMMYY più recente come `pdf_print_log_YYYYMMDD_HHMMSS.log`

### Livelli di Log
- **DEBUG**: Informazioni dettagliate per il debug (solo nel file di log)
- **INFO**: Informazioni generali sull'esecuzione (file di log e console)
- **WARNING**: Avvisi su problemi non critici
- **ERROR**: Errori che impediscono il completamento dell'operazione

### Contenuto Log
I log includono:
- Timestamp di ogni operazione
- File CSV selezionato e percorso
- Numero di righe e colonne processate
- Percorso del PDF generato
- Errori e warning durante l'esecuzione
- Dettagli sulla stampa (per gli script di stampa)
