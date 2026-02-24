# Excel to SPSS Converter

Script Python per convertire file Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) in formato SPSS (`.sav`).

---

## Requisiti

- Python 3.10 o superiore
- Le librerie elencate in `requirements.txt`

---

## Installazione

1. Clona o scarica questa repository.
2. Installa le dipendenze:

```bash
pip install -r requirements.txt
```

---

## Utilizzo

### Sintassi base

```bash
python convert_excel_to_spss.py <percorso_file.xlsx>
```

Il file `.sav` verrà salvato **nella stessa cartella** del file Excel di input, con lo stesso nome base.

**Esempio:**

```bash
python convert_excel_to_spss.py dati/questionario.xlsx
# → crea dati/questionario.sav
```

---

### Opzioni disponibili

| Opzione | Descrizione |
|---|---|
| `--sheet NOME_O_NUMERO` | Specifica il foglio da convertire (nome o indice 0-based). Default: primo foglio. |
| `--output PERCORSO.sav` | Percorso personalizzato per il file di output. |

---

### Esempi

**Convertire il primo foglio (comportamento predefinito):**

```bash
python convert_excel_to_spss.py dati.xlsx
```

**Specificare un foglio per nome:**

```bash
python convert_excel_to_spss.py dati.xlsx --sheet "Risposte"
```

**Specificare un foglio per indice (0-based):**

```bash
python convert_excel_to_spss.py dati.xlsx --sheet 1
```

**Salvare il file .sav in una cartella diversa:**

```bash
python convert_excel_to_spss.py dati.xlsx --output /home/utente/output/dati_spss.sav
```

---

## Comportamento della conversione

### Tipi di variabile

Lo script rileva automaticamente il tipo di ogni colonna:

| Tipo Excel / pandas | Tipo SPSS |
|---|---|
| Numeri interi (`int`) | Variabile numerica |
| Numeri decimali (`float`) | Variabile numerica |
| Booleani (`bool`) | Variabile numerica (0 / 1) |
| Testo (`object`, `str`) | Variabile stringa |
| Date / timestamp | Variabile stringa (formato ISO: `YYYY-MM-DD HH:MM:SS`) |

### Nomi delle variabili

I nomi delle colonne vengono automaticamente adattati alle regole SPSS:

- Spazi sostituiti con `_`
- Caratteri speciali sostituiti con `_`
- Nomi che iniziano con una cifra preceduti da `v_`
- Lunghezza massima: 64 caratteri
- Nomi duplicati rinominati con suffisso numerico (`_1`, `_2`, …)

I nomi originali delle colonne Excel vengono preservati come **etichette di variabile** nel file SPSS.

### Valori mancanti

- Valori `NaN` nelle colonne numeriche → missing system SPSS
- Valori `NaN` / `None` nelle colonne stringa → stringa vuota `""`

---

## Struttura del progetto

```
excel-to-spss/
├── convert_excel_to_spss.py   # Script principale
├── requirements.txt           # Dipendenze Python
└── README.md                  # Questa documentazione
```

---

## Dipendenze principali

| Libreria | Scopo |
|---|---|
| `pandas` | Lettura dei file Excel e manipolazione dei dati |
| `openpyxl` | Engine per leggere file `.xlsx` |
| `pyreadstat` | Scrittura del file SPSS `.sav` |

---

## Risoluzione dei problemi

**`ModuleNotFoundError`** → Assicurati di aver eseguito `pip install -r requirements.txt`.

**File `.sav` non aperto da SPSS** → Verifica che la versione di SPSS supporti i file generati da `pyreadstat` (compatibile con SPSS 15+).

**Colonne con nomi lunghi o caratteri speciali** → Lo script le rinomina automaticamente; i nomi originali restano nelle etichette.

**Foglio non trovato** → Controlla il nome esatto del foglio con `--sheet` oppure usa l'indice numerico.
