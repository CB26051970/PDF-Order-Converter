def extract_data_from_pdf(self, pdf_path):
    """Estrae i dati dal PDF con parser migliorato"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            # SALVA IL TESTO ESTRATTO PER DEBUG
            debug_file = pdf_path.replace('.pdf', '_debug.txt')
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write(all_text)
            print(f"Testo estratto salvato in: {debug_file}")
            
            return self.smart_table_parser(all_text, pdf_path)
    except Exception as e:
        print(f"Errore nell'estrazione dal PDF: {e}")
        return None
