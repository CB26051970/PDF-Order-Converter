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
    """Estrae articoli con metodo semplice e affidabile"""
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
            'Delivery' in line):
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
            
            # Verifica che la quantità sia numerica
            if qty_col.isdigit():
                item = {
                    'customer_code': f"*{code_col}" if not code_col.startswith('*') else code_col,
                    'description': desc_col,
                    'quantity': qty_col,
                    'uom': uom_col
                }
                items.append(item)
                print(f"  Articolo: {item['customer_code']} - Qty: {item['quantity']} - UOM: {item['uom']}")
    
    return items
