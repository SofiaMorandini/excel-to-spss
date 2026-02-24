#!/usr/bin/env python3
"""
convert_excel_to_spss.py
Converte un file Excel (.xlsx) in formato SPSS (.sav).

Uso:
    python convert_excel_to_spss.py <percorso_file.xlsx>
    python convert_excel_to_spss.py <percorso_file.xlsx> --sheet "NomeFoglio"
    python convert_excel_to_spss.py <percorso_file.xlsx> --output /altro/percorso.sav
"""

import argparse
import os
import sys

import pandas as pd
import pyreadstat


def rileva_tipo_variabile(serie: pd.Series) -> str:
    """
    Rileva il tipo di una colonna pandas:
      - 'numeric'  per tipi numerici (int, float, bool)
      - 'string'   per tutto il resto (object, stringhe, date, ecc.)
    """
    if pd.api.types.is_bool_dtype(serie):
        return "numeric"
    if pd.api.types.is_numeric_dtype(serie):
        return "numeric"
    return "string"


def prepara_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, dict, dict]:
    """
    Prepara il DataFrame per la scrittura SPSS:
      - converte le colonne booleane in 0/1
      - converte le date/datetime in stringhe ISO 8601
      - converte le colonne non numeriche in stringa (NaN → '')
      - tronca i nomi delle colonne a 64 caratteri (limite SPSS)
      - restituisce il df pulito, le larghezze delle stringhe e i meta-tipi
    """
    df = df.copy()

    # Rinomina eventuali colonne duplicate o troppo lunghe
    nuovi_nomi = {}
    nomi_visti: set[str] = set()
    for col in df.columns:
        nome = str(col)[:64].strip().replace(" ", "_")
        # Rimuove caratteri non validi per SPSS (ammette lettere, cifre, _)
        nome_pulito = "".join(c if c.isalnum() or c == "_" else "_" for c in nome)
        if not nome_pulito or nome_pulito[0].isdigit():
            nome_pulito = "v_" + nome_pulito
        # Gestisce duplicati
        originale = nome_pulito
        contatore = 1
        while nome_pulito in nomi_visti:
            suffisso = f"_{contatore}"
            nome_pulito = originale[: 64 - len(suffisso)] + suffisso
            contatore += 1
        nomi_visti.add(nome_pulito)
        nuovi_nomi[col] = nome_pulito

    df.rename(columns=nuovi_nomi, inplace=True)

    larghezze_stringhe: dict[str, int] = {}
    formati_variabili: dict[str, str] = {}

    for col in df.columns:
        tipo = rileva_tipo_variabile(df[col])

        if tipo == "numeric":
            # Booleani → 0/1 float
            if pd.api.types.is_bool_dtype(df[col]):
                df[col] = df[col].astype(float)
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            formati_variabili[col] = "numeric"

        else:
            # Date/datetime → stringa ISO
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")
            else:
                df[col] = df[col].astype(str)

            # Sostituisce 'nan' e 'NaT' con stringa vuota
            df[col] = df[col].replace({"nan": "", "NaT": "", "None": ""})

            # Calcola larghezza massima (minimo 1, massimo 32767)
            max_len = df[col].str.len().max()
            if pd.isna(max_len) or max_len == 0:
                max_len = 1
            larghezze_stringhe[col] = min(int(max_len), 32767)
            formati_variabili[col] = "string"

    return df, larghezze_stringhe, formati_variabili


def converti(
    percorso_input: str,
    foglio: str | None = None,
    percorso_output: str | None = None,
) -> str:
    """
    Legge un file Excel e lo salva in formato SPSS.

    Parametri
    ---------
    percorso_input  : percorso del file .xlsx
    foglio          : nome o indice del foglio (default: primo foglio)
    percorso_output : percorso di destinazione .sav (default: stessa cartella)

    Restituisce il percorso del file .sav creato.
    """
    if not os.path.isfile(percorso_input):
        raise FileNotFoundError(f"File non trovato: {percorso_input}")

    estensione = os.path.splitext(percorso_input)[1].lower()
    if estensione not in (".xlsx", ".xls", ".xlsm", ".xlsb"):
        raise ValueError(f"Formato non supportato: '{estensione}'. Usa un file Excel.")

    # Determina il percorso di output
    if percorso_output is None:
        cartella = os.path.dirname(os.path.abspath(percorso_input))
        nome_base = os.path.splitext(os.path.basename(percorso_input))[0]
        percorso_output = os.path.join(cartella, nome_base + ".sav")

    print(f"Lettura del file: {percorso_input}")
    kwargs_lettura: dict = {"engine": "openpyxl"}
    if foglio is not None:
        # Prova a interpretare come numero di foglio se possibile
        try:
            kwargs_lettura["sheet_name"] = int(foglio)
        except ValueError:
            kwargs_lettura["sheet_name"] = foglio

    df = pd.read_excel(percorso_input, **kwargs_lettura)

    if df.empty:
        raise ValueError("Il foglio Excel è vuoto.")

    righe, colonne = df.shape
    print(f"  Righe: {righe:,}  |  Colonne: {colonne:,}")

    # Prepara il DataFrame
    df, larghezze_stringhe, formati_variabili = prepara_dataframe(df)

    # Etichette di colonna (usa i nomi originali come etichette)
    etichette_colonne = list(df.columns)

    # Conta i tipi rilevati
    n_numeriche = sum(1 for t in formati_variabili.values() if t == "numeric")
    n_stringhe = sum(1 for t in formati_variabili.values() if t == "string")
    print(f"  Variabili numeriche: {n_numeriche}  |  Variabili stringa: {n_stringhe}")

    # Scrittura del file .sav
    print(f"Scrittura del file SPSS: {percorso_output}")
    pyreadstat.write_sav(
        df,
        percorso_output,
        column_labels=etichette_colonne,
        variable_value_labels={},
        **({"variable_display_width": {col: larghezze_stringhe[col]
             for col in larghezze_stringhe}} if larghezze_stringhe else {}),
    )

    dimensione = os.path.getsize(percorso_output)
    print(f"Conversione completata. Dimensione file: {dimensione:,} byte")
    return percorso_output


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Converte un file Excel (.xlsx) in formato SPSS (.sav).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  python convert_excel_to_spss.py dati.xlsx
  python convert_excel_to_spss.py dati.xlsx --sheet "Foglio2"
  python convert_excel_to_spss.py dati.xlsx --output /home/utente/output.sav
        """,
    )
    parser.add_argument(
        "input",
        help="Percorso del file Excel di input (.xlsx, .xls, .xlsm, .xlsb)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Nome o numero (0-based) del foglio da convertire (default: primo foglio)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Percorso del file .sav di output (default: stessa cartella del file input)",
    )

    args = parser.parse_args()

    try:
        percorso_sav = converti(
            percorso_input=args.input,
            foglio=args.sheet,
            percorso_output=args.output,
        )
        print(f"\nFile SPSS salvato in: {percorso_sav}")
    except (FileNotFoundError, ValueError) as e:
        print(f"\nErrore: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nErrore inatteso: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
