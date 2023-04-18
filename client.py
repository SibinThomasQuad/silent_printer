from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import pdfkit
import os
import base64
import winreg

DRIVER_PATH = r'driver\bin\wkhtmltopdf.exe'
PDF_FILE_NAME  = 'print_out.pdf'

class Configure:
    
    def set_autostart_registry(self,app_name, key_data=None, autostart: bool = True) -> bool:
        
        with winreg.OpenKey(
                key=winreg.HKEY_CURRENT_USER,
                sub_key=r'Software\Microsoft\Windows\CurrentVersion\Run',
                reserved=0,
                access=winreg.KEY_ALL_ACCESS,
        ) as key:
            try:
                if autostart:
                    winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, key_data)
                else:
                    winreg.DeleteValue(key, app_name)
            except OSError:
                return False
        return True
    
    def check_autostart_registry(self,value_name):
        with winreg.OpenKey(
                key=winreg.HKEY_CURRENT_USER,
                sub_key=r'Software\Microsoft\Windows\CurrentVersion\Run',
                reserved=0,
                access=winreg.KEY_ALL_ACCESS,
        ) as key:
            idx = 0
            while idx < 1_000:     # Max 1.000 values
                try:
                    key_name, _, _ = winreg.EnumValue(key, idx)
                    if key_name == value_name:
                        return True
                    idx += 1
                except OSError:
                    break
        return False

class Message:
    def error(self):
        return '{"status":False,"message":"Print Failed"}'
    
    def success(self):
        return '{"status":True,"message":"Print Success"}'

class Document:
    
    def convert_to_pdf(self,html_value):
        config = pdfkit.configuration(wkhtmltopdf=DRIVER_PATH)
        pdfkit.from_string(html_value, PDF_FILE_NAME, configuration = config)

    def print_data(self):
        os.startfile(PDF_FILE_NAME, "print")

    def base_64_to_html(self,base_64_data):
        html_data = base64.b64decode(base_64_data).decode('utf-8')
        return html_data


class SimpleHTTPRequestHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        document_obj = Document()
        message_obj = Message()
        try:
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            html_data = data['html']
            html_data = document_obj.base_64_to_html(html_data)
            document_obj.convert_to_pdf(html_data)
            document_obj.print_data()
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            response = message_obj.success()
            self.wfile.write(response.encode('utf-8'))
        except:
            response = message_obj.error()
            self.wfile.write(response.encode('utf-8'))

def run():
    registery_obj = Configure()
    if registery_obj.check_autostart_registry('SilentPrinter'):
        server_address = ('', 8000)
        httpd = HTTPServer(server_address, SimpleHTTPRequestHandler)
        print('Silent Printer started')
        httpd.serve_forever()
    else:
        registery_obj.set_autostart_registry('SilentPrinter', r'C:\Users\Abel Jose\OneDrive\Desktop\SILENT\dist\silent_printer.exe')

if __name__ == '__main__':
    run()
 
