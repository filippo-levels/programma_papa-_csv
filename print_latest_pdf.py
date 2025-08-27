#!/usr/bin/env python3
"""
print_latest_pdf.py

Stampa l'ultimo PDF generato nella cartella corrente usando la stampante predefinita di Windows.

Copyright © 2024 Filippo Caliò
Version: 1.0.0
"""

import os
import sys
import glob
from pathlib import Path
from time import sleep


def find_latest_pdf(folder="."):
    """Trova il PDF più recente nella cartella."""
    pdf_files = list(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return None
    
    # Trova il PDF più recente
    latest_pdf = max(pdf_files, key=lambda f: f.stat().st_mtime)
    return latest_pdf


def print_pdf_windows(pdf_path):
    """Stampa un singolo PDF."""
    if pdf_path:
        try:
            print(f"Invio alla stampante: {pdf_path}")
            os.startfile(str(pdf_path), "print")
            # Attendi qualche secondo per evitare che il processo termini troppo presto
            sleep(2)
            print("Stampa inviata con successo.")
        except Exception as e:
            print(f"ERRORE durante la stampa di {pdf_path}: {e}", file=sys.stderr)
            sys.exit(1)


def main():
    # Cerca il PDF più recente
    latest_pdf = find_latest_pdf()
    
    if not latest_pdf:
        print("Nessun file PDF trovato nella cartella corrente.", file=sys.stderr)
        sys.exit(1)
    
    print(f"PDF più recente trovato: {latest_pdf}")
    
    # Stampa il PDF
    print(f"\nAvvio stampa...")
    print_pdf_windows(latest_pdf)
    
    print("Stampa completata.")


if __name__ == "__main__":
    main() 