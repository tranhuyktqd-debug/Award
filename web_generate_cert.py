# -*- coding: utf-8 -*-
"""
Web Form để tạo mã CERT cho học sinh
Sử dụng Flask
"""
from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from datetime import datetime
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

def map_result_to_code(result, subject):
    """Chuyển đổi kết quả thành mã"""
    if pd.isna(result) or result == '':
        return f'NULL-{subject}'
    
    result = str(result).strip().upper()
    
    if 'VÀNG' in result or 'VANG' in result:
        return f'V-{subject}'
    elif 'BẠC' in result or 'BAC' in result:
        return f'B-{subject}'
    elif 'ĐỒNG' in result or 'DONG' in result:
        return f'D-{subject}'
    elif 'KHUYẾN KHÍCH' in result or 'KHUYEN KHICH' in result or 'KK' in result:
        return f'KK-{subject}'
    elif 'CHỨNG NHẬN' in result or 'CHUNG NHAN' in result or 'CN' in result:
        return f'CN-{subject}'
    else:
        return f'NULL-{subject}'

def generate_cert_code_full(row):
    """Tạo mã Cert đầy đủ"""
    khoi = row['Khối']
    if pd.isna(khoi):
        khoi = 'X'
    else:
        khoi = str(int(khoi))
    
    math_code = map_result_to_code(row['KQ VQG TOÁN'], 'MATH')
    science_code = map_result_to_code(row['KQ VQG KHOA HỌC'], 'SCIENCE')
    english_code = map_result_to_code(row['KQ VQG TIẾNG ANH'], 'ENGLISH')
    
    return f"{khoi}*{math_code}*{science_code}*{english_code}"

def generate_cert_code_short(row):
    """Tạo mã Cert rút gọn"""
    khoi = row['Khối']
    if pd.isna(khoi):
        khoi = 'X'
    else:
        khoi = str(int(khoi))
    
    math_code = map_result_to_code(row['KQ VQG TOÁN'], 'M')
    science_code = map_result_to_code(row['KQ VQG KHOA HỌC'], 'S')
    english_code = map_result_to_code(row['KQ VQG TIẾNG ANH'], 'E')
    
    parts = [khoi]
    
    if not math_code.startswith('NULL'):
        parts.append(math_code)
    
    if not science_code.startswith('NULL'):
        parts.append(science_code)
    
    if not english_code.startswith('NULL'):
        parts.append(english_code)
    
    return '*'.join(parts)

@app.route('/')
def index():
    """Trang chủ"""
    return render_template('generate_cert.html')

@app.route('/generate', methods=['POST'])
def generate():
    """Xử lý tạo mã CERT"""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'Không có file được tải lên'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'Không có file được chọn'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'File phải là định dạng Excel (.xlsx hoặc .xls)'}), 400
        
        # Read Excel file
        df = pd.read_excel(file)
        
        # Validate columns
        required_cols = ['Khối', 'KQ VQG TOÁN', 'KQ VQG KHOA HỌC', 'KQ VQG TIẾNG ANH']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            return jsonify({
                'error': f'File thiếu các cột: {", ".join(missing_cols)}'
            }), 400
        
        # Generate CERT codes
        df['MÃ CERT ĐẦY ĐỦ'] = df.apply(generate_cert_code_full, axis=1)
        df['MÃ CERT'] = df.apply(generate_cert_code_short, axis=1)
        
        # Create output file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        
        # Statistics
        stats = {
            'total_students': len(df),
            'top_certs': df['MÃ CERT'].value_counts().head(5).to_dict()
        }
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'Awards_WITH_CERT_{timestamp}.xlsx'
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
        
    except Exception as e:
        return jsonify({'error': f'Có lỗi xảy ra: {str(e)}'}), 500

if __name__ == '__main__':
    # Create templates folder if not exists
    if not os.path.exists('templates'):
        os.makedirs('templates')
    
    print("="*60)
    print("🎓 WEB FORM TẠO MÃ CERT CHO HỌC SINH")
    print("="*60)
    print("\n🌐 Mở trình duyệt và truy cập: http://localhost:5000")
    print("\n⚠️  Nhấn Ctrl+C để dừng server\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
