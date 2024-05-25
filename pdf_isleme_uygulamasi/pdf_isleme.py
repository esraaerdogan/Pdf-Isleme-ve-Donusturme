import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import json
import xml.etree.ElementTree as ET
import os
from bs4 import BeautifulSoup
from pdfminer.high_level import extract_text
import fitz  # PyMuPDF

def convertPdfToXml(sourceFile, hedefDosya):
    """Pdfplumber ve xml.etree.ElementTree kutuphanelerini kullanarak PDF'yi Xml'e dönüştürme"""
    try:
        with pdfplumber.open(sourceFile) as pdf:  #pdfplumber kütüphanesini kullanarak bir PDF dosyasını aç
            root = ET.Element("root") #XML belgesinin temelini oluşturma
            for i, page in enumerate(pdf.pages):
                # Sayfa metnini al
                page_text = page.extract_text()
                # XML ağacına sayfa düğümünü ekle
                page_element = ET.SubElement(root, f"page_{i+1}")
                page_element.text = page_text
                
                # Sayfadaki resimleri XML'e ekleme 
                images = page.images
                if images:
                    images_element = ET.SubElement(page_element, "images")
                    for img in images:
                        img_element = ET.SubElement(images_element, "image")
                        img_element.text = str(img)
                
                # Sayfadaki tabloları XML'e ekle
                tables = page.extract_tables()
                if tables:
                    tables_element = ET.SubElement(page_element, "tables")
                    for k, table in enumerate(tables):
                        table_element = ET.SubElement(tables_element, f"table_{k+1}")
                        for row in table:
                            row_element = ET.SubElement(table_element, "row")
                            for cell in row:
                                cell_element = ET.SubElement(row_element, "cell")
                                cell_element.text = str(cell)
                
                # Sayfadaki eşitlikleri XML'e ekle (Varsa)
                equations = page.extract_words()
                if equations:
                    equations_element = ET.SubElement(page_element, "equations")
                    for eq in equations:
                        eq_element = ET.SubElement(equations_element, "equation")
                        eq_element.text = eq["text"]
        
        # XML ağacını dosyaya yaz
        tree = ET.ElementTree(root)
        tree.write(hedefDosya)
        
        messagebox.showinfo("Başarılı", f"Sonuc dosyasi \"{hedefDosya}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def convertPdfToJson(sourceFile, hedefDosya):
    """PyMuPDF ve json kütüphanelerini kullanarak PDF'yi Json'a dönüştürme"""
    try:
        pdf_document = fitz.open(sourceFile)
        data = {}
        data["pages"] = []
        
        # pdfplumber ile tabloları çıkar
        with pdfplumber.open(sourceFile) as pdf:
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                tables = page.extract_tables()
                
                # Tabloları sayfa verilerine ekle
                page_tables = []
                for table in tables:
                    page_tables.append(table)
                    
                # PyMuPDF ile resimleri ve diğer verileri çıkar
                pdf_document = fitz.open(sourceFile)
                page_images = []
                page_objects = pdf_document.load_page(i)
                images = page_objects.get_images(full=True)
                for j, img_info in enumerate(images):
                    xref = img_info[0]
                    base_image = pdf_document.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    img_data = {
                        "image_number": j + 1,
                        "image_extension": image_ext
                    }
                    page_images.append(img_data)

                # Sayfa verilerini oluştur
                page_data = {
                    "page_number": i + 1,
                    "text": page_text,
                    "tables": page_tables,
                    "images": page_images
                }

                data["pages"].append(page_data)
                
        # JSON verisini dosyaya yaz
        with open(hedefDosya, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
        
        messagebox.showinfo("Başarılı", f"Sonuc dosyasi \"{hedefDosya}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def convertPdfToHtml(sourceFile, hedefDosya):
    """Pdfplumber ve BeautifulSoup kullanarak PDF'yi Html'e dönüştürme"""
    try:
        with pdfplumber.open(sourceFile) as pdf:
            print(f"Number of pages: {len(pdf.pages)}")
            html_content = "<html><body>"

            for page_num, page in enumerate(pdf.pages, start=1):
                html_content += f"<h2>Sayfa {page_num}</h2>"
                
                # Metni çıkar ve HTML'e ekle
                text = page.extract_text()
                if text:
                    html_content += f"<div>{text}</div>"
                    
                # Resimleri çıkar ve HTML'e ekle
                images = page.images
                if images:
                    for img_idx, img in enumerate(images):
                        # Resim verilerini base64 ile kodla
                        img_data = base64.b64encode(img["stream"].get_data()).decode("utf-8")
                        img_src = f"data:image/jpeg;base64,{img_data}"
                        html_content += f"<img src='{img_src}' alt='Resim {img_idx + 1}' /><br>"
                
                # Tabloları çıkar ve HTML'e ekle
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        html_content += "<table border='1' style='width:100%; border-collapse:collapse;'>"
                        for row in table:
                            html_content += "<tr>"
                            for cell in row:
                                html_content += f"<td style='border:1px solid black; padding:5px;'>{cell}</td>"
                            html_content += "</tr>"
                        html_content += "</table><br>"

            html_content += "</body></html>"

        soup = BeautifulSoup(html_content, "html.parser")
        formatted_html = soup.prettify()

        with open(hedefDosya, "w", encoding="utf-8") as file:
            file.write(formatted_html)
        
        messagebox.showinfo("Başarılı", f"Sonuç dosyası \"{hedefDosya}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def convertPdfToText(sourceFile, hedefDosya):
    """Pdfminer ve pdfplumber ile PDF'yi metin dosyasına dönüştürme"""
    try:
        # PDF'den metni çıkar
        text = extract_text(sourceFile)
        
        # pdfplumber kullanarak PDF'yi aç
        with pdfplumber.open(sourceFile) as pdf:
            for page in pdf.pages:
                # Tabloları çıkar
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        text += "\n\nTablo:\n"
                        for row in table:
                            text += "\t".join([str(cell) for cell in row]) + "\n"
                
                # Görselleri çıkar
                images = page.images
                if images:
                    for img in images:
                        text += f"\n\nResim: {img}\n"

        # Çıkarılan metni bir TXT dosyasına kaydet
        with open(hedefDosya, "w", encoding="utf-8") as file:
            file.write(text)
        
        messagebox.showinfo("Başarılı", f"Sonuc dosyası \"{hedefDosya}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))
        

def browsePdfFile():
    file_path = filedialog.askopenfilename(filetypes=[("PDF dosyaları", "*.pdf")])
    if file_path:
        entry_pdf_path.delete(0, tk.END)
        entry_pdf_path.insert(0, file_path)

def saveXmlFile():
    file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML dosyaları", "*.xml")])
    if file_path:
        entry_xml_path.delete(0, tk.END)
        entry_xml_path.insert(0, file_path)
        convertPdfToXml(entry_pdf_path.get(), file_path)

def saveJsonFile():
    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON dosyaları", "*.json")])
    if file_path:
        entry_json_path.delete(0, tk.END)
        entry_json_path.insert(0, file_path)
        convertPdfToJson(entry_pdf_path.get(), file_path)

def saveHtmlFile():
    file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML dosyaları", "*.html")])
    if file_path:
        entry_html_path.delete(0, tk.END)
        entry_html_path.insert(0, file_path)
        convertPdfToHtml(entry_pdf_path.get(), file_path)

def saveTextFile():
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("TEXT dosyaları", "*.txt")])
    if file_path:
        entry_text_path.delete(0, tk.END)
        entry_text_path.insert(0, file_path)
        convertPdfToText(entry_pdf_path.get(), file_path)

# Ana pencereyi oluşturma
root = tk.Tk()
root.title("PDF Dönüştürücü")

# Widget'ları oluşturma ve yerleştirme
label_pdf_path = tk.Label(root, text="PDF Dosya Yolu:")
label_pdf_path.grid(row=0, column=0, padx=10, pady=10, sticky="e")

entry_pdf_path = tk.Entry(root, width=50)
entry_pdf_path.grid(row=0, column=1, padx=10, pady=10)

button_browse_pdf = tk.Button(root, text="Göz at...", command=browsePdfFile)
button_browse_pdf.grid(row=0, column=2, padx=10, pady=10)

# XML dönüştürme bölümü
label_xml_path = tk.Label(root, text="XML'i Farklı Kaydet:")
label_xml_path.grid(row=1, column=0, padx=10, pady=10, sticky="e")

entry_xml_path = tk.Entry(root, width=50)
entry_xml_path.grid(row=1, column=1, padx=10, pady=10)

button_save_xml = tk.Button(root, text="XML'e Dönüştür ve Kaydet", command=saveXmlFile)
button_save_xml.grid(row=1, column=2, padx=10, pady=10)

# JSON dönüştürme bölümü
label_json_path = tk.Label(root, text="JSON'u Farklı Kaydet:")
label_json_path.grid(row=2, column=0, padx=10, pady=10, sticky="e")

entry_json_path = tk.Entry(root, width=50)
entry_json_path.grid(row=2, column=1, padx=10, pady=10)

button_save_json = tk.Button(root, text="JSON'a Dönüştür ve Kaydet", command=saveJsonFile)
button_save_json.grid(row=2, column=2, padx=10, pady=10)

# HTML dönüştürme bölümü
label_html_path = tk.Label(root, text="HTML'i Farklı Kaydet:")
label_html_path.grid(row=3, column=0, padx=10, pady=10, sticky="e")

entry_html_path = tk.Entry(root, width=50)
entry_html_path.grid(row=3, column=1, padx=10, pady=10)

button_save_html = tk.Button(root, text="HTML'e Dönüştür ve Kaydet", command=saveHtmlFile)
button_save_html.grid(row=3, column=2, padx=10, pady=10)

# TEXT dönüştürme bölümü
label_text_path = tk.Label(root, text="TEXT'i Farklı Kaydet:")
label_text_path.grid(row=4, column=0, padx=10, pady=10, sticky="e")

entry_text_path = tk.Entry(root, width=50)
entry_text_path.grid(row=4, column=1, padx=10, pady=10)

button_save_text = tk.Button(root, text="TEXT'e Dönüştür ve Kaydet", command=saveTextFile)
button_save_text.grid(row=4, column=2, padx=10, pady=10)

# Ana döngüyü başlat
root.mainloop()