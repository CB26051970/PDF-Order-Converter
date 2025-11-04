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
        """Parser semplificato per formato tabellare"""
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
        
        print(f"Analizzando ordine {data['po_number']}...")
        
        # Estrai articoli con metodo semplificato
        items = self.extract_items_from_text_simple(text)
        data['items'] = items
        
        print(f"Trovati {len(items)} articoli")
        return data
    
    def extract_items_from_text_simple(self, text):
        """Estrae articoli con metodo semplice e affidabile - VERSIONE MIGLIORATA"""
        items = []
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Salta righe non rilevanti
            if (not line or 
                line.startswith('Codice') or 
                line.startswith('Item Code') or
                'Descrizione' in line or
                'QTY' in line or
                'Total' in line or
                'Delivery' in line or
                'Note:' in line):
                continue
                
            # Dividi per 2+ spazi
            columns = re.split(r'\s{2,}', line)
            
            # Deve avere almeno 4 colonne: Codice, Descrizione, QTY, UOM
            if len(columns) >= 4:
                # Prendi solo le prime 4 colonne
                code_col = columns[0].strip()
                desc_col = columns[1].strip()
                qty_col = columns[2].strip()
                uom_col = columns[3].strip()
                
                # Verifica che la quantità sia numerica e il codice sia valido
                # MODIFICA: Accetta anche "12" come quantità (per GSD CO2)
                if (qty_col.isdigit() or qty_col == '12') and (code_col.isdigit() or code_col.startswith('*')):
                    item = {
                        'customer_code': f"*{code_col}" if not code_col.startswith('*') else code_col,
                        'description': desc_col,
                        'quantity': qty_col,
                        'uom': uom_col
                    }
                    items.append(item)
                    print(f"  Articolo: {item['customer_code']} - Qty: {item['quantity']} - UOM: {item['uom']}")
        
        return items
    
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
