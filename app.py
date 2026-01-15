import os
import io
import uuid
import base64
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

app = Flask(__name__)

# --- AYARLAR ---
LOGO_URL = "https://static.wixstatic.com/media/06f423_9d350d42007948448e351781e43950c1~mv2.png"

# GÜVENLİK: Şifreleri kodun içine yazmıyoruz, Render'ın kasasından (Environment Variables) çekiyoruz.
MY_EMAIL = os.environ.get("MY_EMAIL")
MY_PASSWORD = os.environ.get("MY_PASSWORD")

# --- YARDIMCI FONKSİYONLAR ---
def clean_tag(tag):
    if '}' in tag: return tag.split('}', 1)[1]
    return tag

def clean_attribs(attrib):
    return {clean_tag(k).lower(): v for k, v in attrib.items()}

def get_image_from_data(base64_string):
    try:
        clean_data = base64_string.replace('\n', '').replace('\r', '').replace(' ', '')
        if len(clean_data) < 200: return None
        img_bytes = base64.b64decode(clean_data)
        img_io = io.BytesIO(img_bytes)
        img = Image.open(img_io)
        img.verify()
        img_io.seek(0)
        return img_io
    except:
        return None

def apply_formatting(run, attrib):
    run.font.color.rgb = RGBColor(0, 0, 0)
    if attrib.get('bold') == 'true' or attrib.get('b') == 'true': run.bold = True
    if attrib.get('italic') == 'true' or attrib.get('i') == 'true': run.italic = True
    if attrib.get('underline') == 'true' or attrib.get('u') == 'true': run.underline = True
    f_size = attrib.get('size') or attrib.get('fontsize')
    if f_size and f_size.isdigit():
        try: run.font.size = Pt(int(f_size))
        except: pass

def generate_word_doc(udf_dosya_objesi):
    try:
        with zipfile.ZipFile(udf_dosya_objesi) as z:
            if 'content.xml' not in z.namelist(): return None, "content.xml bulunamadı"
            
            xml_content = z.read('content.xml').decode('utf-8', errors='ignore')
            root = ET.fromstring(xml_content)
            doc = Document()

            global_text = ""
            max_len = 0
            for elem in root.iter():
                if elem.text:
                    l = len(elem.text)
                    if l > max_len: max_len = l; global_text = elem.text
            if not global_text: doc.add_paragraph("[Metin içeriği boş]")

            elements_processed = False
            found_images = [] 
            elements_node = None
            for elem in root.iter():
                if 'elements' in clean_tag(elem.tag): elements_node = elem; break
            if elements_node is not None and global_text:
                current_p = doc.add_paragraph()
                for item in elements_node.iter():
                    attrib = clean_attribs(item.attrib)
                    tag = clean_tag(item.tag)
                    if tag == 'content' and 'startoffset' in attrib:
                        elements_processed = True
                        try:
                            start = int(attrib['startoffset'])
                            length = int(attrib['length'])
                            chunk = global_text[start : start + length]
                            if '\n' in chunk:
                                lines = chunk.split('\n')
                                for i, line in enumerate(lines):
                                    if line:
                                        run = current_p.add_run(line)
                                        apply_formatting(run, attrib)
                                    if i < len(lines) - 1: current_p = doc.add_paragraph()
                            else:
                                run = current_p.add_run(chunk)
                                apply_formatting(run, attrib)
                        except: pass
                    elif 'imagedata' in attrib:
                        v = attrib['imagedata']
                        if len(v) > 500:
                            img_obj = get_image_from_data(v)
                            if img_obj: found_images.append(img_obj)
            if not elements_processed and global_text:
                p = doc.add_paragraph()
                run = p.add_run(global_text)
                run.font.color.rgb = RGBColor(0, 0, 0)
                for elem in root.iter():
                    for k, v in elem.attrib.items():
                        if len(v) > 500:
                            img_obj = get_image_from_data(v)
                            if img_obj: found_images.append(img_obj)
            if found_images:
                doc.add_page_break()
                baslik = doc.add_heading('DOSYA İÇİNDEKİ GÖRSELLER', level=1)
                baslik.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for i, img_obj in enumerate(found_images):
                    doc.add_paragraph("\n")
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    try: run.add_picture(img_obj, width=Inches(4.5))
                    except: continue
                    lbl = doc.add_paragraph(f"#Görsel {i+1}")
                    lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    lbl.style.font.bold = True
                    lbl.style.font.color.rgb = RGBColor(0, 0, 0)
            if 'sign.sgn' in z.namelist():
                doc.add_paragraph("\n\n")
                p = doc.add_paragraph("e-imzalıdır")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = RGBColor(255, 0, 0)
                p.runs[0].font.size = Pt(14)
                p2 = doc.add_paragraph("(Bu belge 5070 sayılı Kanun uyarınca imzalanmıştır)")
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.style.font.size = Pt(9)
                p2.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            return doc, None
    except Exception as e:
        return None, str(e)

# --- E-POSTA GÖNDERME FONKSİYONU ---
def send_email_with_attachment(to_email, file_path, file_name):
    if not MY_EMAIL or not MY_PASSWORD:
        return False, "Sunucu e-posta ayarları yapılmamış."
        
    try:
        msg = MIMEMultipart()
        msg['From'] = MY_EMAIL
        msg['To'] = to_email
        msg['Subject'] = "UDF Dönüştürülmüş Dosyanız"
        
        body = "Merhaba,\n\nUDF Dönüştürücü kullanarak oluşturduğunuz dosya ektedir.\n\nİyi çalışmalar."
        msg.attach(MIMEText(body, 'plain'))
        
        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {file_name}")
            msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(MY_EMAIL, MY_PASSWORD)
        text = msg.as_string()
        server.sendmail(MY_EMAIL, to_email, text)
        server.quit()
        return True, "E-posta gönderildi"
    except Exception as e:
        return False, str(e)

# --- ROUTES ---
@app.route('/')
def anasayfa():
    return f'''
    <!doctype html>
    <html lang="tr">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>UDF Dönüştürücü</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: transparent; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; flex-direction: column; }}
            .card {{ background: #ffffff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); text-align: center; width: 450px; max-width: 90%; position: relative; }}
            .logo-img {{ max-width: 200px; margin-bottom: 20px; }}
            
            .file-upload {{ position: relative; display: inline-block; width: 100%; margin-bottom: 15px; }}
            input[type=file] {{ border: 2px dashed #ccc; padding: 20px; width: 88%; border-radius: 8px; background: #f0f0f0; color: #555; cursor: pointer; transition: 0.3s; }}
            input[type=file]:hover {{ border-color: #999; background: #e9e9e9; }}
            
            .email-section {{ margin-top: 10px; margin-bottom: 20px; text-align: left; }}
            .email-label {{ font-size: 13px; color: #666; display: block; margin-bottom: 5px; font-weight: 600; }}
            input[type=email] {{ width: 92%; padding: 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; outline: none; transition: 0.3s; }}
            input[type=email]:focus {{ border-color: #555; }}

            .btn-group {{ display: flex; gap: 15px; margin-bottom: 20px; }}
            .btn {{ flex: 1; padding: 14px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; color: white; transition: transform 0.1s, opacity 0.2s; display: flex; flex-direction: column; align-items: center; gap: 4px; }}
            .btn span {{ font-size: 11px; opacity: 0.8; font-weight: normal; }}
            .btn:active {{ transform: scale(0.98); }}
            .btn-word {{ background-color: #2b5797; }}
            .btn-pdf {{ background-color: #d32f2f; }}
            .btn:hover {{ opacity: 0.9; }}

            .progress-container {{ display: none; width: 100%; background-color: #d0d0d0; border-radius: 4px; margin: 20px 0; overflow: hidden; }}
            .progress-bar {{ width: 0%; height: 10px; background-color: #666; transition: width 0.4s ease; }}
            .loader-text {{ font-size: 13px; color: #777; margin-top: 5px; display: none; }}
            
            .result-area {{ display: none; margin-top: 20px; padding: 20px; background: #e0e0e0; border-radius: 8px; border: 1px solid #ccc; }}
            .success-msg {{ color: #333; font-weight: bold; font-size: 18px; margin-bottom: 10px; display: block; }}
            .mail-msg {{ font-size: 14px; color: #555; margin-bottom: 15px; display: block; }}
            .download-link {{ display: inline-block; text-decoration: none; background: #333; color: white; padding: 10px 20px; border-radius: 5px; font-weight: bold; transition: 0.2s; font-size: 13px; }}
            .download-link:hover {{ background: #000; }}
            .reset-link {{ cursor:pointer; color:#666; text-decoration:underline; margin-top: 15px; display: inline-block; }}
        </style>
    </head>
    <body>
        <div class="card">
            <img src="{LOGO_URL}" alt="Logo" class="logo-img">
            <form id="uploadForm">
                <div class="file-upload">
                    <input type="file" id="fileInput" name="dosya" required accept=".udf">
                </div>
                
                <div class="email-section">
                    <label class="email-label">E-posta ile gönder (İsteğe Bağlı):</label>
                    <input type="email" id="emailInput" placeholder="ornek@mail.com">
                </div>

                <div class="btn-group" id="btnGroup">
                    <button type="button" class="btn btn-word" onclick="startProcess('word')">
                        WORD'e Çevir
                        <span>(İndir veya Mail At)</span>
                    </button>
                    <button type="button" class="btn btn-pdf" onclick="startProcess('pdf')">
                        PDF'e Çevir
                        <span>(İndir veya Mail At)</span>
                    </button>
                </div>

                <div class="progress-container" id="progressContainer"><div class="progress-bar" id="progressBar"></div></div>
                <div class="loader-text" id="loaderText">İşleniyor ve gönderiliyor...</div>

                <div class="result-area" id="resultArea">
                    <span class="success-msg">İşlem Tamamlandı!</span>
                    <span id="mailResultMsg" class="mail-msg"></span>
                    <a href="#" id="downloadBtn" class="download-link">Dosyayı İndir</a><br>
                    <small class="reset-link" onclick="resetForm()">Yeni İşlem Yap</small>
                </div>
            </form>
        </div>

        <script>
            function startProcess(type) {{
                var fileInput = document.getElementById('fileInput');
                var emailInput = document.getElementById('emailInput');
                
                if (fileInput.files.length === 0) {{ alert("Lütfen önce bir UDF dosyası seçin."); return; }}

                document.getElementById('btnGroup').style.display = 'none';
                document.querySelector('.email-section').style.display = 'none';
                document.querySelector('.file-upload').style.display = 'none';
                document.getElementById('progressContainer').style.display = 'block';
                document.getElementById('loaderText').style.display = 'block';
                
                var progressBar = document.getElementById('progressBar');
                var width = 0;
                var speed = emailInput.value ? 100 : 50; 
                var interval = setInterval(function() {{ if (width >= 90) clearInterval(interval); else {{ width++; progressBar.style.width = width + '%'; }} }}, speed);

                var formData = new FormData();
                formData.append('dosya', fileInput.files[0]);
                formData.append('email', emailInput.value); 

                var url = type === 'word' ? '/islem_word' : '/islem_pdf';

                fetch(url, {{ method: 'POST', body: formData }})
                .then(response => {{ if (response.status !== 200) throw new Error("İşlem hatası"); return response.json(); }})
                .then(data => {{
                    clearInterval(interval); progressBar.style.width = '100%';
                    
                    var byteCharacters = atob(data.file_content);
                    var byteNumbers = new Array(byteCharacters.length);
                    for (var i = 0; i < byteCharacters.length; i++) {{ byteNumbers[i] = byteCharacters.charCodeAt(i); }}
                    var byteArray = new Uint8Array(byteNumbers);
                    var blob = new Blob([byteArray], {{type: data.mime_type}});
                    
                    var downloadUrl = window.URL.createObjectURL(blob);
                    var downloadBtn = document.getElementById('downloadBtn');
                    downloadBtn.href = downloadUrl;
                    downloadBtn.download = data.file_name;
                    downloadBtn.innerText = type === 'word' ? "Dosyayı İndir (Word)" : "Dosyayı İndir (PDF)";

                    var mailMsgSpan = document.getElementById('mailResultMsg');
                    if (data.email_sent) {{
                        mailMsgSpan.innerHTML = "✅ Dosya <b>" + data.email_address + "</b> adresine gönderildi.";
                        mailMsgSpan.style.color = "green";
                    }} else if (emailInput.value) {{
                        mailMsgSpan.innerHTML = "⚠️ Dosya hazırlandı ama mail atılamadı: " + data.email_error;
                        mailMsgSpan.style.color = "orange";
                    }} else {{
                        mailMsgSpan.innerHTML = "";
                    }}

                    setTimeout(function() {{
                        document.getElementById('progressContainer').style.display = 'none';
                        document.getElementById('loaderText').style.display = 'none';
                        document.getElementById('resultArea').style.display = 'block';
                    }}, 600);
                }})
                .catch(error => {{ alert("Bir hata oluştu: " + error); resetForm(); }});
            }}

            function resetForm() {{
                location.reload(); 
            }}
        </script>
    </body>
    </html>
    '''

@app.route('/islem_word', methods=['POST'])
def islem_word():
    if 'dosya' not in request.files: return jsonify({'error': 'Dosya yok'}), 400
    dosya = request.files['dosya']
    email = request.form.get('email', '').strip()
    
    doc, hata = generate_word_doc(dosya)
    if hata: return jsonify({'error': hata}), 500
    
    unique_id = str(uuid.uuid4())
    filename = dosya.filename.replace('.udf', '.docx')
    temp_path = f"temp_{unique_id}.docx"
    doc.save(temp_path)
    
    email_result = False
    email_error = ""
    
    if email:
        success, msg = send_email_with_attachment(email, temp_path, filename)
        email_result = success
        email_error = msg

    with open(temp_path, "rb") as f:
        file_data = f.read()
        b64_data = base64.b64encode(file_data).decode('utf-8')

    if os.path.exists(temp_path): os.remove(temp_path)
    
    return jsonify({
        'file_content': b64_data,
        'file_name': filename,
        'mime_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'email_sent': email_result,
        'email_error': email_error,
        'email_address': email
    })

@app.route('/islem_pdf', methods=['POST'])
def islem_pdf():
    if 'dosya' not in request.files: return jsonify({'error': 'Dosya yok'}), 400
    dosya = request.files['dosya']
    email = request.form.get('email', '').strip()

    doc, hata = generate_word_doc(dosya)
    if hata: return jsonify({'error': hata}), 500
    
    unique_id = str(uuid.uuid4())
    filename = dosya.filename.replace('.udf', '.pdf')
    temp_docx = f"temp_{unique_id}.docx"
    temp_pdf = f"temp_{unique_id}.pdf"
    
    try:
        doc.save(temp_docx)
        subprocess.run(['libreoffice', '--headless', '--invisible', '--convert-to', 'pdf', '--outdir', os.getcwd(), temp_docx], check=True)
        
        email_result = False
        email_error = ""

        if os.path.exists(temp_pdf):
            if email:
                success, msg = send_email_with_attachment(email, temp_pdf, filename)
                email_result = success
                email_error = msg
            
            with open(temp_pdf, "rb") as f:
                file_data = f.read()
                b64_data = base64.b64encode(file_data).decode('utf-8')
            
            return jsonify({
                'file_content': b64_data,
                'file_name': filename,
                'mime_type': 'application/pdf',
                'email_sent': email_result,
                'email_error': email_error,
                'email_address': email
            })
        else:
            return jsonify({'error': 'PDF oluşturulamadı'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(temp_docx): os.remove(temp_docx)
        if os.path.exists(temp_pdf): os.remove(temp_pdf)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
