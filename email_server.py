"""
Simple HTTP server to handle email sending requests from web interface
Run: python email_server.py
"""
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import json
import subprocess
import os
import sys

class EmailHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/get-data':
            # Return student data from DS_KQ_WITH_QR.xlsx
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            try:
                if os.path.exists('DS_KQ_WITH_QR.xlsx'):
                    import pandas as pd
                    df = pd.read_excel('DS_KQ_WITH_QR.xlsx', dtype={'SBD': str})
                    
                    # Convert to list of dicts
                    students = []
                    for _, row in df.iterrows():
                        students.append({
                            'candidate': str(row.get('SBD', '')).strip(),
                            'fullName': str(row.get('FULL NAME', '')).strip(),
                            'dob': str(row.get('D.O.B', '')).strip(),
                            'pass': str(row.get('PASS', '')).strip(),
                            'grade': str(row.get('KHỐI', '')).strip(),
                            'school': str(row.get('TRƯỜNG', '')).strip(),
                            'area': str(row.get('KHU VỰC', '')).strip(),
                            'sdt': str(row.get('SĐT', '')).strip(),
                            'email': str(row.get('EMAIL', '')).strip(),
                            'toan': str(row.get('TOÁN', '')).strip(),
                            'kh': str(row.get('KHOA HỌC', '')).strip(),
                            'ta': str(row.get('TIẾNG ANH', '')).strip(),
                            'certCode': str(row.get('CERT CODE', '')).strip(),
                            'certCode2': str(row.get('CERT CODE FULL', '')).strip(),
                            'qrData': str(row.get('QR DATA', '')).strip(),
                            'lop': str(row.get('LỚP', '')).strip(),
                            'phhs': str(row.get('PHHS', '')).strip(),
                            'photo': str(row.get('PHOTO', '')).strip()
                        })
                    
                    response = {'status': 'success', 'data': students, 'count': len(students)}
                    print(f"📊 Sent {len(students)} students to web")
                else:
                    response = {'status': 'info', 'message': 'Chưa có file DS_KQ_WITH_QR.xlsx'}
                    
            except Exception as e:
                response = {'status': 'error', 'message': str(e)}
            
            self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()
    
    def do_POST(self):
        if self.path == '/upload-data':
            # Handle file upload
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            try:
                # Parse multipart form data
                import re
                boundary = self.headers['Content-Type'].split('boundary=')[1].encode()
                parts = post_data.split(b'--' + boundary)
                
                for part in parts:
                    if b'filename="DATA KQ.xlsx"' in part or b'filename="DATA' in part:
                        # Extract file content
                        file_start = part.find(b'\r\n\r\n') + 4
                        file_end = len(part) - 2  # Remove trailing \r\n
                        file_data = part[file_start:file_end]
                        
                        # Save to DATA KQ.xlsx
                        with open('DATA KQ.xlsx', 'wb') as f:
                            f.write(file_data)
                        
                        print(f"✅ Saved DATA KQ.xlsx ({len(file_data)} bytes)")
                        response = {'status': 'success', 'message': 'File uploaded successfully'}
                        break
                else:
                    response = {'status': 'error', 'message': 'No file found in upload'}
                    
            except Exception as e:
                print(f"❌ Upload error: {e}")
                response = {'status': 'error', 'message': str(e)}
            
            self.wfile.write(json.dumps(response).encode())
            
        elif self.path == '/process-and-send':
            # Generate QR codes from DATA KQ.xlsx
            print("📥 Received /process-and-send request")
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            print(f"📊 Request data: {data}")
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            try:
                print("📤 Generating QR codes from DATA KQ.xlsx...")
                
                # Check if DATA KQ.xlsx exists
                if not os.path.exists('DATA KQ.xlsx'):
                    response = {'status': 'error', 
                               'message': 'Không tìm thấy file DATA KQ.xlsx. Vui lòng copy file vào thư mục này.'}
                else:
                    print("🔄 Running create_qr_for_all_students.py...")
                    result = subprocess.run(['python', 'create_qr_for_all_students.py'], 
                                          capture_output=True, text=True, timeout=60)
                    
                    if result.returncode == 0:
                        print("✅ QR codes generated successfully!")
                        # Count students
                        import pandas as pd
                        df = pd.read_excel('DS_KQ_WITH_QR.xlsx')
                        count = len(df)
                        response = {'status': 'success', 
                                   'message': f'Đã tạo QR cho {count} học sinh. File DS_KQ_WITH_QR.xlsx đã sẵn sàng.',
                                   'count': count}
                    else:
                        response = {'status': 'error', 
                                   'message': f'Lỗi tạo QR: {result.stderr}'}
                
            except subprocess.TimeoutExpired:
                response = {'status': 'error', 'message': 'Timeout: Quá nhiều học sinh, vui lòng chạy script trực tiếp'}
            except Exception as e:
                response = {'status': 'error', 'message': str(e)}
            
            self.wfile.write(json.dumps(response).encode())
            
        elif self.path == '/send-email':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            action = data.get('action')
            print(f"📨 Received request: {action}")
            
            if action == 'send_all':
                # Run send_student_awards.py
                print("🚀 Starting bulk email send...")
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                
                try:
                    subprocess.Popen(['python', 'send_student_awards.py'])
                    response = {'status': 'success', 'message': 'Started sending emails to all students'}
                except Exception as e:
                    response = {'status': 'error', 'message': str(e)}
                
                self.wfile.write(json.dumps(response).encode())
                
            elif action == 'send_single':
                sbd = data.get('sbd')
                print(f"📧 Sending email for SBD: {sbd}")
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                
                try:
                    # Import and call function from send_student_awards
                    from send_student_awards import send_single_student_email
                    result = send_single_student_email(sbd)
                    response = result
                except Exception as e:
                    response = {'status': 'error', 'message': str(e)}
                
                self.wfile.write(json.dumps(response).encode())
        else:
            self.send_response(404)
            self.end_headers()
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

def run_server(port=8000):
    server_address = ('', port)
    httpd = ThreadingHTTPServer(server_address, EmailHandler)
    print(f'🚀 Email server running on http://localhost:{port}')
    print('📧 Ready to handle email requests from web interface')
    print('⏹️  Press Ctrl+C to stop')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\n👋 Stopping server...')
        httpd.shutdown()

if __name__ == '__main__':
    run_server()
