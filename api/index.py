from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        
        html_content = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🎓 ASMO Awards Processing System</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 800px;
            width: 100%;
            padding: 40px;
            text-align: center;
        }
        h1 {
            color: #2c3e50;
            font-size: 2.5em;
            margin-bottom: 20px;
        }
        .icon {
            font-size: 4em;
            margin-bottom: 20px;
        }
        .description {
            color: #7f8c8d;
            font-size: 1.2em;
            line-height: 1.8;
            margin-bottom: 30px;
        }
        .info-box {
            background: #ecf0f1;
            border-left: 4px solid #3498db;
            padding: 20px;
            margin: 20px 0;
            text-align: left;
            border-radius: 5px;
        }
        .info-box h3 {
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .info-box p {
            color: #555;
            line-height: 1.6;
        }
        .download-section {
            margin-top: 30px;
            padding-top: 30px;
            border-top: 2px solid #ecf0f1;
        }
        .btn {
            display: inline-block;
            padding: 15px 30px;
            margin: 10px;
            background: #27ae60;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            font-weight: bold;
            transition: all 0.3s;
            box-shadow: 0 4px 15px rgba(39, 174, 96, 0.3);
        }
        .btn:hover {
            background: #229954;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(39, 174, 96, 0.4);
        }
        .btn-secondary {
            background: #3498db;
            box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3);
        }
        .btn-secondary:hover {
            background: #2980b9;
            box-shadow: 0 6px 20px rgba(52, 152, 219, 0.4);
        }
        .features {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 30px;
        }
        .feature {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            border: 2px solid #e9ecef;
        }
        .feature-icon {
            font-size: 2em;
            margin-bottom: 10px;
        }
        .feature h4 {
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .feature p {
            color: #7f8c8d;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="icon">🎓</div>
        <h1>ASMO Awards Processing System</h1>
        <p class="description">
            Hệ thống xử lý mã CERT và quản lý giải thưởng ASMO với giao diện đồ họa hiện đại
        </p>
        
        <div class="info-box">
            <h3>ℹ️ Thông tin</h3>
            <p>
                <strong>Đây là ứng dụng Desktop (Tkinter)</strong>, không phải web application.<br>
                Để sử dụng, bạn cần tải về và chạy trên máy tính Windows.
            </p>
        </div>
        
        <div class="download-section">
            <h3 style="color: #2c3e50; margin-bottom: 20px;">📥 Tải xuống</h3>
            <a href="https://github.com/tranhuyktqd-debug/Award" class="btn" target="_blank">
                📦 Xem trên GitHub
            </a>
            <a href="https://github.com/tranhuyktqd-debug/Award/archive/refs/heads/main.zip" class="btn btn-secondary" target="_blank">
                ⬇️ Tải Source Code
            </a>
        </div>
        
        <div class="features">
            <div class="feature">
                <div class="feature-icon">📋</div>
                <h4>Xử lý Mã Cert</h4>
                <p>So sánh, xếp hạng và tạo mã CERT tự động</p>
            </div>
            <div class="feature">
                <div class="feature-icon">📦</div>
                <h4>Chia danh sách</h4>
                <p>Chia danh sách học sinh theo STT túi</p>
            </div>
            <div class="feature">
                <div class="feature-icon">🔍</div>
                <h4>Tra cứu</h4>
                <p>Tìm kiếm thông tin học sinh nhanh chóng</p>
            </div>
        </div>
        
        <div style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #ecf0f1; color: #7f8c8d;">
            <p>© 2026 ASMO Vietnam. All rights reserved.</p>
            <p style="margin-top: 10px; font-size: 0.9em;">
                Repository: <a href="https://github.com/tranhuyktqd-debug/Award" target="_blank" style="color: #3498db;">github.com/tranhuyktqd-debug/Award</a>
            </p>
        </div>
    </div>
</body>
</html>
        """
        
        self.wfile.write(html_content.encode('utf-8'))
        return
