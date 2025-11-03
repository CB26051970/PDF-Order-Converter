import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
from openpyxl.styles import Font

class PDFToExcelConverter:
    def __init__(self):
        self.conversion_db = None
        self.conversion_dict = {}
        
    def load_conversion_db(self, excel_file_path):
        """Carica il database di conversione dal file Excel"""
        try:
            self.conversion_db = pd.read_excel(excel_file_path)
            self.conversion_dict = dict(zip(self.conversion_db.iloc[:, 1], self.conversion_db.iloc[:, 2]))
            print(f"Database di conversione caricato: {len(self.conversion_dict)} codici")
            return True
        except Exception as e:
            print(f"Errore nel caricamento del DB di conversione: {e}")
            return False
    
    def extract_data_from_pdf(self, pdf_path):
        """Estrae i dati dal PDF con parser migliorato"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                all_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"
                
                return self.smart_table_parser(all_text, pdf_path)
        except Exception as e:
            print(f"Errore nell'estrazione dal PDF: {e}")
            return None
    
    def smart_table_parser(self, text, pdf_path):
        """Parser intelligente che cerca tabelle in diversi formati"""
        data = {
            'po_number': '',
            'po_date': '',
            'delivery_date': '',
            'items': []
        }
        
        # Estrai informazioni header
        po_match = re.search(r'PO No:\s*(\d+)', text)
        if po_match:
            data['po_number'] = po_match.group(1)
        
        date_match = re.search(r'Date of PO:\s*(\d{2}/\d{2}/\d{4})', text)
        if date_match:
            data['po_date'] = date_match.group(1)
        
        delivery_match = re.search(r'Delivery Date.*?ON.*?:\s*(\d{2}/\d{2}/\d{4})', text)
        if delivery_match:
            data['delivery_date'] = delivery_match.group(1)
        
        print(f"Analizzando ordine {data['po_number']}...")
        
        # CERCA LA TABELLA DEGLI ARTICOLI
        items = []
        
        # Metodo 1: Cerca blocchi di articoli completi
        item_blocks = self.find_item_blocks(text)
        for block in item_blocks:
            item = self.parse_item_block(block)
            if item and item['customer_code']:
                items.append(item)
        
        # Metodo 2: Se il primo metodo non trova abbastanza articoli
        if len(items) < 3:
            items = self.alternative_parsing(text)
        
        data['items'] = items
        print(f"Trovati {len(items)} articoli")
        
        return data
    
    def find_item_blocks(self, text):
        """Trova blocchi completi di articoli nel testo"""
        blocks = []
        lines = text.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Se troviamo una riga che inizia con codice articolo
            if line.startswith('*'):
                block = {
                    'code_line': line,
                    'data_lines': []
                }
                
                # Cerca le righe dati successive (max 2 righe)
                for j in range(1, 3):
                    if i + j < len(lines):
                        next_line = lines[i + j].strip()
                        # Se è una riga dati (contiene numeri e prezzi)
                        if self.is_data_line(next_line):
                            block['data_lines'].append(next_line)
                        else:
                            break
                
                if block['data_lines']:
                    blocks.append(block)
                    i += len(block['data_lines'])  # Salta le righe già processate
            
            i += 1
        
        return blocks
    
    def is_data_line(self, line):
        """Determina se una riga contiene dati dell'articolo"""
        # Una riga dati tipicamente contiene: quantità UOM prezzo
        patterns = [
            r'^\d+\s+[^\s€]+\s+€',  # quantità UOM €prezzo
            r'^\d+\s+[^\s€]+\s+[^\s€]+\s+€',  # quantità UOM qualcosaltro €prezzo
            r'\d+\s+[^\s€]+\s+\d+[,.]\d+'  # quantità UOM numero,numero
        ]
        
        return any(re.search(pattern, line) for pattern in patterns)
    
    def parse_item_block(self, block):
        """Analizza un blocco articolo completo"""
        item = {'customer_code': '', 'description': '', 'quantity': '', 'uom': ''}
        
        # Estrai dalla riga codice
        code_line = block['code_line']
        code_match = re.match(r'(\*\d+\w*)\s+(.*)', code_line)
        if code_match:
            item['customer_code'] = code_match.group(1)
            item['description'] = code_match.group(2).strip()
        
        # Estrai dalla prima riga dati
        if block['data_lines']:
            data_line = block['data_lines'][0]
            
            # Pattern 1: quantità UOM prezzo
            pattern1 = r'^(\d+)\s+([^\s€]+)\s+€'
            match1 = re.search(pattern1, data_line)
            if match1:
                item['quantity'] = match1.group(1)
                item['uom'] = match1.group(2)
            else:
                # Pattern 2: quantità UOM (senza €)
                pattern2 = r'^(\d+)\s+([^\s]+)'
                match2 = re.search(pattern2, data_line)
                if match2:
                    item['quantity'] = match2.group(1)
                    item['uom'] = match2.group(2)
            
            # Pulisci la descrizione da eventuali dati rimasti
            item['description'] = self.clean_description(item['description'], item.get('quantity', ''))
        
        return item
    
    def alternative_parsing(self, text):
        """Metodo alternativo di parsing per PDF difficili"""
        items = []
        lines = text.split('\n')
        
        # Cerca sezioni che contengono la tabella articoli
        table_section = self.extract_table_section(lines)
        
        for line in table_section:
            line = line.strip()
            if line.startswith('*'):
                # Prova diversi pattern per estrarre i dati
                item = self.try_multiple_patterns(line)
                if item and item['customer_code']:
                    items.append(item)
        
        return items
    
    def extract_table_section(self, lines):
        """Estrae la sezione della tabella articoli"""
        start_idx = -1
        end_idx = len(lines)
        
        # Trova inizio tabella (dopo l'header)
        for i, line in enumerate(lines):
            if 'Item Code' in line and 'Item Description' in line:
                start_idx = i + 1
                break
        
        # Trova fine tabella (prima dei totali)
        for i in range(start_idx, len(lines)):
            if any(x in lines[i] for x in ['Total', 'Delivery to:', 'Note:']):
                end_idx = i
                break
        
        return lines[start_idx:end_idx] if start_idx != -1 else lines
    
    def try_multiple_patterns(self, line):
        """Prova diversi pattern per estrarre i dati - VERSIONE CORRETTA"""
        item = {'customer_code': '', 'description': '', 'quantity': '', 'uom': ''}
        
        # PATTERN 1 MIGLIORATO: *codice descrizione quantità confezione prezzo...
        # Cerca: *codice descrizione QUANTITÀ confezione €prezzo
        pattern1 = r'(\*\d+\w*)\s+(.*?)\s+(\d+)\s+(\d+\s*x\s*\d+[^\s€]*)\s+€'
        match1 = re.search(pattern1, line)
        if match1:
            item['customer_code'] = match1.group(1)
            item['description'] = match1.group(2).strip()
            item['quantity'] = match1.group(3)  # QUANTITÀ dopo la descrizione
            item['uom'] = match1.group(4)       # Confezione (12 x 75cl)
            print(f"  Pattern1 - Trovato: {item['customer_code']} Qty: {item['quantity']} UOM: {item['uom']}")
            return item
        
        # PATTERN 2: *codice descrizione quantità UOM semplice prezzo...
        pattern2 = r'(\*\d+\w*)\s+(.*?)\s+(\d+)\s+([^\s€]+)\s+€'
        match2 = re.search(pattern2, line)
        if match2:
            item['customer_code'] = match2.group(1)
            item['description'] = match2.group(2).strip()
            item['quantity'] = match2.group(3)
            item['uom'] = match2.group(4)
            print(f"  Pattern2 - Trovato: {item['customer_code']} Qty: {item['quantity']} UOM: {item['uom']}")
            return item
        
        # PATTERN 3: Cerca quantità nella stessa riga dopo la descrizione
        pattern3 = r'(\*\d+\w*)\s+(.*?)\s+(\d+)\s+'
        match3 = re.search(pattern3, line)
        if match3:
            item['customer_code'] = match3.group(1)
            remaining_text = match3.group(2) + " " + line[match3.end():]
            item['quantity'] = match3.group(3)
            
            # Cerca UOM dopo la quantità
            uom_match = re.search(r'\d+\s+([^\s€]+)\s+', remaining_text)
            if uom_match:
                item['uom'] = uom_match.group(1)
                item['description'] = re.sub(r'\s*\d+\s+[^\s€]+\s+.*', '', match3.group(2)).strip()
            else:
                item['description'] = match3.group(2).strip()
            
            print(f"  Pattern3 - Trovato: {item['customer_code']} Qty: {item['quantity']} UOM: {item['uom']}")
            return item
        
        # PATTERN 4: Solo codice e descrizione (fallback)
        pattern4 = r'(\*\d+\w*)\s+(.*)'
        match4 = re.search(pattern4, line)
        if match4:
            item['customer_code'] = match4.group(1)
            item['description'] = match4.group(2).strip()
            print(f"  Pattern4 - Solo codice: {item['customer_code']} (quantità non trovata)")
        
        return item
    
    def clean_description(self, description, quantity):
        """Pulisce la descrizione rimuovendo dati numerici"""
        if not quantity:
            return description.strip()
        
        # Rimuovi la quantità e tutto ciò che segue se presente
        pattern = r'(.+?)\s*' + re.escape(quantity) + r'.*'
        match = re.match(pattern, description)
        if match:
            return match.group(1).strip()
        
        return description.strip()
    
    def convert_to_internal_codes(self, order_data):
        """Converte i codici cliente in codici interni"""
        converted_items = []
        
        for item in order_data['items']:
            customer_code = item['customer_code']
            internal_code = self.conversion_dict.get(customer_code, "**NEW**")
            
            converted_item = {
                'internal_code': internal_code,
                'quantity': item.get('quantity', ''),
                'description': item.get('description', ''),
                'customer_code': customer_code,
                'uom': item.get('uom', '')
            }
            converted_items.append(converted_item)
        
        return converted_items
    
    def create_excel_output(self, order_data, converted_items, output_path):
        """Crea il file Excel di output"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_data = []
                for item in converted_items:
                    df_data.append({
                        'Codice Interno': item['internal_code'],
                        'Quantità': item['quantity'],
                        'Descrizione': item['description'],
                        'Codice Cliente': item['customer_code'],
                        'UOM': item.get('uom', '')
                    })
                
                df = pd.DataFrame(df_data)
                df.to_excel(writer, sheet_name='Ordine', index=False, startrow=4)
                
                workbook = writer.book
                worksheet = writer.sheets['Ordine']
                
                worksheet['A1'] = f"Numero Ordine: {order_data['po_number']}"
                worksheet['A2'] = f"Data Ordine: {order_data['po_date']}"
                worksheet['A3'] = f"Data Consegna: {order_data['delivery_date']}"
                
                column_widths = {'A': 15, 'B': 12, 'C': 40, 'D': 15, 'E': 10}
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                
                for row in range(5, len(converted_items) + 5):
                    cell_value = worksheet[f'A{row}'].value
                    if cell_value == "**NEW**":
                        worksheet[f'A{row}'].font = Font(bold=True, color="FF0000")
                
                print(f"File Excel creato: {output_path}")
                return True
                
        except Exception as e:
            print(f"Errore nella creazione del file Excel: {e}")
            return False
    
    def process_single_pdf(self, pdf_path, conversion_db_path, output_dir=None):
        """Elabora un singolo file PDF"""
        if output_dir is None:
            output_dir = os.path.dirname(pdf_path)
        
        if not self.load_conversion_db(conversion_db_path):
            return False
        
        order_data = self.extract_data_from_pdf(pdf_path)
        if not order_data or not order_data['items']:
            print("Nessun dato estratto dal PDF")
            return False
        
        print(f"Trovati {len(order_data['items'])} articoli nell'ordine {order_data['po_number']}")
        for i, item in enumerate(order_data['items'], 1):
            print(f"  {i}. {item['customer_code']} - Qty: {item.get('quantity', 'N/A')} - UOM: {item.get('uom', 'N/A')}")
        
        converted_items = self.convert_to_internal_codes(order_data)
        
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(output_dir, f"{pdf_name}_converted.xlsx")
        
        success = self.create_excel_output(order_data, converted_items, output_path)
        
        if success:
            print(f"Conversione completata: {output_path}")
            new_codes = sum(1 for item in converted_items if item['internal_code'] == "**NEW**")
            if new_codes > 0:
                print(f"ATTENZIONE: {new_codes} codici senza corrispondenza trovati")
        
        return success

def main():
    converter = PDFToExcelConverter()
    
    root = tk.Tk()
    root.withdraw()
    
    print("=== CONVERTITORE PDF ORDINI FORNITORI ===")
    print("Seleziona il file DB CONVERSION.xlsx")
    
    conversion_file = filedialog.askopenfilename(
        title="Seleziona il file DB CONVERSION.xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    if not conversion_file:
        return
    
    print("Seleziona i file PDF degli ordini da convertire")
    pdf_files = filedialog.askopenfilenames(
        title="Seleziona i file PDF degli ordini",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    
    if not pdf_files:
        return
    
    output_dir = filedialog.askdirectory(title="Seleziona la cartella di output")
    if not output_dir:
        output_dir = os.path.dirname(pdf_files[0])
    
    success_count = 0
    for pdf_file in pdf_files:
        print(f"\nElaborando: {os.path.basename(pdf_file)}")
        if converter.process_single_pdf(pdf_file, conversion_file, output_dir):
            success_count += 1
    
    print(f"\n=== CONVERSIONE COMPLETATA ===")
    print(f"File elaborati con successo: {success_count}/{len(pdf_files)}")
    
    messagebox.showinfo("Conversione Completata", 
                       f"Elaborati {success_count}/{len(pdf_files)} file.\n"
                       f"Output salvato in: {output_dir}")

if __name__ == "__main__":
    main()
