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
import logging
from pathlib import Path
from time import sleep
from datetime import datetime


def setup_logging():
    """Configura il sistema di logging per file e console."""
    # Crea logger
    logger = logging.getLogger('pdf_printer')
    logger.setLevel(logging.DEBUG)
    
    # Handler per file
    log_file = f"pdf_print_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # Handler per console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Aggiungi handler
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def find_latest_pdf(folder=".", logger=None):
    """Trova il PDF più recente nella cartella."""
    if logger:
        logger.info(f"Ricerca PDF in: {os.path.abspath(folder)}")
    
    try:
        pdf_files = list(Path(folder).glob("*.pdf"))
        if logger:
            logger.info(f"Trovati {len(pdf_files)} file PDF")
            for pdf in pdf_files:
                stat = pdf.stat()
                logger.debug(f"  PDF: {pdf.name} - Dimensione: {stat.st_size} bytes - Modificato: {datetime.fromtimestamp(stat.st_mtime)}")
        
        if not pdf_files:
            if logger:
                logger.warning("Nessun file PDF trovato")
            return None
        
        # Trova il PDF più recente
        latest_pdf = max(pdf_files, key=lambda f: f.stat().st_mtime)
        
        if logger:
            stat = latest_pdf.stat()
            logger.info(f"PDF più recente selezionato: {latest_pdf}")
            logger.info(f"  Dimensione: {stat.st_size} bytes")
            logger.info(f"  Modificato: {datetime.fromtimestamp(stat.st_mtime)}")
            logger.info(f"  Percorso completo: {os.path.abspath(latest_pdf)}")
        
        return latest_pdf
    except Exception as e:
        if logger:
            logger.error(f"Errore nella ricerca PDF: {e}")
        return None


def print_pdf_windows(pdf_path, logger=None):
    """Stampa un singolo PDF."""
    if logger:
        logger.info(f"Tentativo di stampa: {pdf_path}")
    
    # Verifica che il file esista
    if not os.path.exists(pdf_path):
        if logger:
            logger.error(f"File PDF non trovato: {pdf_path}")
        else:
            print(f"ERRORE: File PDF non trovato: {pdf_path}", file=sys.stderr)
        return False
    
    # Verifica che sia un file PDF
    if not str(pdf_path).lower().endswith('.pdf'):
        if logger:
            logger.error(f"File non è un PDF: {pdf_path}")
        else:
            print(f"ERRORE: File non è un PDF: {pdf_path}", file=sys.stderr)
        return False
    
    if pdf_path:
        try:
            if logger:
                logger.info(f"Invio alla stampante: {pdf_path}")
            else:
                print(f"Invio alla stampante: {pdf_path}")
            
            os.startfile(str(pdf_path), "print")
            # Attendi qualche secondo per evitare che il processo termini troppo presto
            sleep(3)
            
            if logger:
                logger.info("Stampa inviata con successo alla coda di stampa")
            else:
                print("Stampa inviata con successo.")
            return True
        except Exception as e:
            if logger:
                logger.error(f"ERRORE durante la stampa: {e}")
            else:
                print(f"ERRORE durante la stampa di {pdf_path}: {e}", file=sys.stderr)
            return False
    
    return False


def main():
    # Setup logging
    logger = setup_logging()
    logger.info("=== AVVIO STAMPA PDF ===")
    logger.info(f"Directory di lavoro: {os.getcwd()}")
    
    try:
        # Cerca il PDF più recente
        latest_pdf = find_latest_pdf(logger=logger)
        
        if not latest_pdf:
            logger.error("Nessun file PDF trovato nella cartella corrente.")
            sys.exit(1)
        
        print(f"PDF più recente trovato: {latest_pdf}")
        
        # Stampa il PDF
        print(f"\nAvvio stampa...")
        success = print_pdf_windows(latest_pdf, logger=logger)
        
        if success:
            logger.info("=== STAMPA COMPLETATA CON SUCCESSO ===")
            print("Stampa completata.")
        else:
            logger.error("=== ERRORE DURANTE LA STAMPA ===")
            sys.exit(2)
            
    except Exception as e:
        logger.error(f"Errore imprevisto nel main: {e}")
        sys.exit(3)


if __name__ == "__main__":
    main() 