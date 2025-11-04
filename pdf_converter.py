import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
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
            'supplier': '',
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
        
        # Estrai nome fornitore
        supplier_match = re.search(r'To:\s*(.+?)\n', text)
        if supplier_match:
            data['supplier'] = supplier_match.group(1).strip()
        
        print(f"Analizzando ordine {data['po_number']}...")
        
        # CERCA LA TABELLA DEGLI ARTICOLI
        items = self.extract_items_from_text(text)
        data['items'] = items
        
        print(f"Trovati {len(items)} articoli")
        for i, item in enumerate(items, 1):
            print(f"  {i}. {item['customer_code']} - Qty: {item.get('quantity', 'N/A')} - UOM: {item.get('uom', 'N/A')}")
        
        return data
    
    def extract_items_from_text(self, text):
        """Estrae gli articoli dal testo con parser robusto"""
        items = []
        lines = text.split('\n')
        
        # Trova la sezione della tabella articoli
        table_start = -1
        table_end = len(lines)
        
        for i, line in enumerate(lines):
            if any(keyword in line for keyword in ['Item Code', 'Item Description', 'QTY']):
                table_start = i + 1
                break
        
        if table_start == -1:
            # Se non trova l'header, cerca direttamente i codici articolo
            return self.find_items_by_codes(lines)
        
        # Trova la fine della tabella
        for i in range(table_start, len(lines)):
            if any(keyword in lines[i] for keyword in ['Total', 'Delivery to:', 'Note:', 'Subtotal']):
                table_end = i
                break
        
        # Processa le righe della tabella
        for i in range(table_start, table_end):
            line = lines[i].strip()
            if not line:
                continue
                
            item = self.parse_item_line(line)
            if item and item['customer_code']:
                items.append(item)
        
        return items
    
    def find_items_by_codes(self, lines):
        """Trova articoli cercando direttamente i codici che iniziano con *"""
        items = []
        
        for i, line in enumerate(lines):
            line = line.strip()
            if line.startswith('*'):
                item = self.parse_item_line_advanced(line, lines, i)
                if item and item['customer_code']:
                    items.append(item)
        
        return items
    
    def parse_item_line(self, line):
        """Analizza una singola riga articolo - PARSER MIGLIORATO"""
        item = {'customer_code': '', 'description': '', 'quantity': '', 'uom': ''}
        
        # Cerca il codice articolo (inizia sempre con *)
        code_match = re.search(r'(\*\d+\w*)', line)
        if not code_match:
            return item
        
        item['customer_code'] = code_match.group(1)
        remaining_text = line[code_match.end():].strip()
        
        # SEPARA DESCRIZIONE E DATI NUMERICI
        # Cerca il punto dove iniziano i dati numerici (quantità)
        quantity_match = re.search(r'(\d+)\s+([^\s€]+)(?:\s+€)', remaining_text)
        if quantity_match:
            # Estrai quantità e UOM
            item['quantity'] = quantity_match.group(1)
            item['uom'] = quantity_match.group(2)
            
            # La descrizione è tutto prima della quantità
            description_end = remaining_text.find(quantity_match.group(0))
            if description_end > 0:
                item['description'] = remaining_text[:description_end].strip()
            else:
                # Fallback: prendi tutto fino al primo numero
                desc_match = re.match(r'(.*?)(?=\d+\s+[^\s€]+\s+€)', remaining_text)
                if desc_match:
                    item['description'] = desc_match.group(1).strip()
        else:
            # METODO ALTERNATIVO: cerca qualsiasi pattern con quantità
            words = remaining_text.split()
            for i, word in enumerate(words):
                if word.isdigit() and i < len(words) - 1:
                    item['quantity'] = word
                    item['uom'] = words[i + 1]
                    item['description'] = ' '.join(words[:i]).strip()
                    break
        
        # Pulisci la descrizione
        item['description'] = self.clean_description(item['description'])
        
        return item
    
    def parse_item_line_advanced(self, line, all_lines, current_index):
        """Parser avanzato per formati complessi come Word"""
        item = {'customer_code': '', 'description': '', 'quantity': '', 'uom': ''}
        
        # Estrai codice
        code_match = re.search(r'(\*\d+\w*)', line)
        if not code_match:
            return item
        
        item['customer_code'] = code_match.group(1)
        remaining = line[code_match.end():].strip()
        
        # Cerca pattern specifico per formato Word/tabella
        # Pattern: *codice descrizione quantità UOM prezzo...
        patterns = [
            # Pattern per formato Word con tabelle
            r'(\*\d+\w*)\s+(.*?)\s+(\d+)\s+([^\s€]+?)\s+€',
            # Pattern più generico
            r'(\*\d+\w*)\s+(.*?)\s+(\d+)\s+([^\s€]+)',
            # Pattern solo quantità
            r'(\*\d+\w*)\s+(.*?)\s+(\d+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                item['customer_code'] = match.group(1)
                item['description'] = match.group(2).strip()
                item['quantity'] = match.group(3)
                if len(match.groups()) >= 4:
                    item['uom'] = match.group(4)
                break
        
        # Se non trova con i pattern, prova a cercare nella riga successiva
        if not item['quantity'] and current_index + 1 < len(all_lines):
            next_line = all_lines[current_index + 1].strip()
            # Cerca quantità nella riga successiva
            qty_match = re.search(r'^\s*(\d+)\s+([^\s€]+)', next_line)
            if qty_match:
                item['quantity'] = qty_match.group(1)
                item['uom'] = qty_match.group(2)
        
        # Pulisci la descrizione
        item['description'] = self.clean_description(item['description'])
        
        return item
    
    def clean_description(self, description):
        """Pulisce la descrizione rimuovendo parti ripetute"""
        if not description:
            return ""
        
        # Rimuovi parti ripetute (es: "GSD PET 50cl Sprite GSD PET 50cl Sprite")
        words = description.split()
        if len(words) > 3:
            # Cerca sequenze ripetute
            for i in range(1, len(words) // 2 + 1):
                if words[:i] == words[i:2*i]:
                    return ' '.join(words[:i])
        
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
                # Prepara i dati per il DataFrame
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
                df.to_excel(writer, sheet_name='Ordine', index=False, startrow=6)
                
                workbook = writer.book
                worksheet = writer.sheets['Ordine']
                
                # Intestazione
                worksheet['A1'] = f"NUMERO ORDINE: {order_data['po_number']}"
                worksheet['A2'] = f"DATA ORDINE: {order_data['po_date']}"
                worksheet['A3'] = f"DATA CONSEGNA: {order_data['delivery_date']}"
                worksheet['A4'] = f"FORNITORE: {order_data.get('supplier', '')}"
                worksheet['A5'] = f"TOTALE ARTICOLI: {len(converted_items)}"
                
                # Stile per l'intestazione
                for row in range(1, 6):
                    worksheet[f'A{row}'].font = Font(bold=True)
                
                # Stile per i codici nuovi
                for row in range(7, len(converted_items) + 7):
                    cell_value = worksheet[f'A{row}'].value
                    if cell_value == "**NEW**":
                        worksheet[f'A{row}'].font = Font(bold=True, color="FF0000")
                
                # Larghezza colonne
                column_widths = {'A': 15, 'B': 12, 'C': 40, 'D': 15, 'E': 15}
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                
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

    def manual_order_entry(self, conversion_db_path, output_dir):
        """Interfaccia per inserimento manuale ordini"""
        print("\n=== INSERIMENTO MANUALE ORDINE ===")
        
        if not self.load_conversion_db(conversion_db_path):
            return False
        
        # Input dati ordine usando dialoghi grafici
        po_number = simpledialog.askstring("Inserimento Ordine", "Numero Ordine (PO):")
        if not po_number:
            return False
            
        po_date = simpledialog.askstring("Inserimento Ordine", "Data Ordine (dd/mm/yyyy):")
        if not po_date:
            return False
            
        delivery_date = simpledialog.askstring("Inserimento Ordine", "Data Consegna (dd/mm/yyyy):")
        if not delivery_date:
            return False
            
        supplier = simpledialog.askstring("Inserimento Ordine", "Fornitore:")
        if not supplier:
            return False
        
        order_data = {
            'po_number': po_number,
            'po_date': po_date,
            'delivery_date': delivery_date,
            'supplier': supplier,
            'items': []
        }
        
        # Input articoli
        while True:
            add_more = messagebox.askyesno("Inserimento Articoli", "Aggiungere un articolo?")
            if not add_more:
                break
            
            customer_code = simpledialog.askstring("Articolo", "Codice Cliente (es: *274077):")
            if not customer_code:
                continue
                
            quantity = simpledialog.askstring("Articolo", "Quantità:")
            description = simpledialog.askstring("Articolo", "Descrizione:")
            uom = simpledialog.askstring("Articolo", "UOM (es: 12 x 50cl):")
            
            order_data['items'].append({
                'customer_code': customer_code,
                'quantity': quantity or '',
                'description': description or '',
                'uom': uom or ''
            })
        
        if not order_data['items']:
            print("Nessun articolo inserito")
            return False
        
        # Conversione e creazione Excel
        converted_items = self.convert_to_internal_codes(order_data)
        
        output_path = os.path.join(output_dir, f"MANUALE_{po_number}_converted.xlsx")
        success = self.create_excel_output(order_data, converted_items, output_path)
        
        if success:
            print(f"\nOrdine manuale creato: {output_path}")
            new_codes = sum(1 for item in converted_items if item['internal_code'] == "**NEW**")
            if new_codes > 0:
                messagebox.showwarning("Codici Non Trovati", 
                                     f"ATTENZIONE: {new_codes} codici senza corrispondenza trovati")
        
        return success

def main():
    converter = PDFToExcelConverter()
    
    root = tk.Tk()
    root.withdraw()  # Nasconde la finestra principale
    
    # Scelta tra conversione PDF e inserimento manuale
    choice = messagebox.askquestion(
        "Selezione Modalità",
        "Scegli la modalità:\n\nSì = Converti PDF automaticamente\nNo = Inserimento manuale ordine",
        icon='question'
    )
    
    print("Seleziona il file DB CONVERSION.xlsx")
    conversion_file = filedialog.askopenfilename(
        title="Seleziona il file DB CONVERSION.xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    if not conversion_file:
        return
    
    output_dir = filedialog.askdirectory(title="Seleziona la cartella di output")
    if not output_dir:
        output_dir = os.getcwd()
    
    if choice == 'yes':
        # Conversione automatica PDF
        print("\nSeleziona i file PDF degli ordini da convertire")
        pdf_files = filedialog.askopenfilenames(
            title="Seleziona i file PDF degli ordini",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if not pdf_files:
            return
        
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
    
    else:
        # Inserimento manuale
        if converter.manual_order_entry(conversion_file, output_dir):
            messagebox.showinfo("Ordine Creato", 
                               f"Ordine manuale creato con successo!\n"
                               f"Output salvato in: {output_dir}")
        else:
            messagebox.showerror("Errore", "Errore nella creazione dell'ordine manuale")
    
    root.destroy()

if __name__ == "__main__":
    main()
