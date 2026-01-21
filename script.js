// Global data storage
let studentsData = [];

// Admin state
const ADMIN_PASSWORD = 'Admin@2024';
let isAdminLoggedIn = false;

// Initialize date dropdowns
function initializeDateSelects() {
    const daySelect = document.getElementById('daySelect');
    const monthSelect = document.getElementById('monthSelect');
    const yearSelect = document.getElementById('yearSelect');

    // Populate days (1-31)
    for (let i = 1; i <= 31; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        daySelect.appendChild(option);
    }

    // Populate months (1-12)
    for (let i = 1; i <= 12; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        monthSelect.appendChild(option);
    }

    // Populate years (1990-2020)
    for (let i = 2020; i >= 1990; i--) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        yearSelect.appendChild(option);
    }
}

// File upload handler
document.getElementById('fileInput').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;

    // Check if it's DATA KQ.xlsx (without QR)
    if (!file.name.includes('WITH_QR') && (file.name.includes('DATA') || file.name.includes('KQ'))) {
        // This is the original DATA file - need to save it first
        
        const reader = new FileReader();
        reader.onload = async function(event) {
            try {
                // Upload file to server
                const formData = new FormData();
                formData.append('file', file);
                
                console.log('📤 Uploading file to server...');
                
                const response = await fetch('http://localhost:8000/upload-data', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.status === 'success') {
                    alert(`✅ Đã upload file: ${file.name}\n\n📊 File đã được lưu trên server.\n\n👉 Nhấn nút "🔲 Tạo QR Codes" để tạo QR cho tất cả học sinh.`);
                    
                    // Show the Generate QR button
                    document.getElementById('generateQRBtn').style.display = 'block';
                    document.getElementById('downloadBtn').style.display = 'none';
                } else {
                    alert(`❌ Lỗi upload: ${result.message}`);
                }
            } catch (error) {
                console.error('Upload error:', error);
                alert(`❌ Không thể upload file lên server.\n\nVui lòng đảm bảo server đang chạy:\npython email_server.py`);
            }
        };
        reader.readAsArrayBuffer(file);
        return;
    }

    // Normal file upload (DS_KQ_WITH_QR.xlsx)
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Read from first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            studentsData = jsonData.map(row => {
                const student = {
                    fullName: row['FULL NAME'] || row['Full Name'] || row['Họ tên'] || '',
                    candidate: row['SBD'] || row['Candidate'] || row['Số báo danh'] || '',
                    dob: row['Ngày sinh'] || row['D.O.B'] || row['D.O.B2'] || row['DOB'] || '',
                    grade: row['KHỐI'] || row['Grade'] || row['Lớp'] || row['Khối'] || '',
                    school: row['TRƯỜNG'] || row['School'] || row['Trường'] || '',
                    area: row['KHU VỰC'] || row['Area'] || row['Khu vực'] || '',
                    toan: row['KQ VQG TOÁN'] || row['TOÁN'] || row['Toán'] || '',
                    kh: row['KQ VQG KHOA HỌC'] || row['KHOA HỌC'] || row['Khoa học'] || row['KH'] || '',
                    ta: row['KQ VQG TIẾNG ANH'] || row['TIẾNG ANH'] || row['Tiếng Anh'] || row['TA'] || '',
                    certCode: row['MÃ CERT ĐẦY ĐỦ'] || row['CERT CODE FULL'] || row['Cert code'] || row['Mã chứng chỉ'] || '',
                    certCode2: row['MÃ CERT'] || row['CERT CODE'] || row['CERT CODE2'] || '',
                    lop: row['LỚP'] || '',
                    pass: row['PASS'] || '',
                    phhs: row['PHHS'] || '',
                    sdt: row['Số điện thoại liên hệ'] || row['SĐT'] || '',
                    email: row['Email liên hệ'] || row['EMAIL'] || '',
                    photo: row['PHOTO'] || row['Photo'] || row['Ảnh'] || '',
                    qrData: row['QR DATA'] || '' // Read QR DATA from Excel
                };
                
                // Generate QR data if not exists (fallback)
                if (!student.qrData) {
                    student.qrData = `STUDENT INFORMATION
Candidate: ${student.candidate}
Name: ${student.fullName}
Date of Birth: ${student.dob}
Grade ${student.grade} - ${student.school}

RESULTS:
Math: ${student.toan || 'N/A'}
Science: ${student.kh || 'N/A'}
English: ${student.ta || 'N/A'}

Certificate: ${student.certCode || student.certCode2 || 'N/A'}`;
                }
                
                return student;
            });

            console.log('Sheet đã đọc:', sheetName);
            console.log('Dữ liệu mẫu:', studentsData[0]);
            console.log('Tên các cột:', Object.keys(jsonData[0]));
            
            // Show download button
            document.getElementById('downloadBtn').style.display = 'block';
            
            // Update email count if admin logged in
            if (isAdminLoggedIn) {
                updateEmailCount();
            }
            
            alert(`Đã tải lên ${studentsData.length} học sinh thành công từ sheet ${sheetName}!`);
            clearForm();
        } catch (error) {
            alert('Lỗi khi đọc file Excel. Vui lòng kiểm tra định dạng file.');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Search function
function searchStudents() {
    const candidate = document.getElementById('candidateInput').value.trim().toLowerCase();
    const fullName = document.getElementById('fullNameInput').value.trim().toLowerCase();
    const day = document.getElementById('daySelect').value;
    const month = document.getElementById('monthSelect').value;
    const year = document.getElementById('yearSelect').value;

    let results = studentsData;

    // Filter by candidate number
    if (candidate) {
        results = results.filter(student => 
            String(student.candidate).toLowerCase().includes(candidate)
        );
    }

    // Filter by full name
    if (fullName) {
        results = results.filter(student => 
            student.fullName.toLowerCase().includes(fullName)
        );
    }

    // Filter by date of birth
    if (day || month || year) {
        results = results.filter(student => {
            if (!student.dob) return false;
            
            const dob = student.dob.toString();
            let match = true;
            
            // Support format: DD-MM-YYYY or DD/MM/YYYY
            if (day) {
                const dayStr = day.toString().padStart(2, '0');
                if (!dob.startsWith(dayStr)) {
                    match = false;
                }
            }
            if (month) {
                const monthStr = month.toString().padStart(2, '0');
                if (!dob.includes(`-${monthStr}-`) && !dob.includes(`/${monthStr}/`)) {
                    match = false;
                }
            }
            if (year) {
                if (!dob.endsWith(year.toString())) {
                    match = false;
                }
            }
            
            return match;
        });
    }

    displayResults(results);
    
    // If only one result, automatically show student details
    if (results.length === 1) {
        setTimeout(() => {
            const firstRow = document.querySelector('.results-table tbody tr');
            if (firstRow) {
                showStudentDetails(0, firstRow);
                firstRow.classList.add('selected');
            }
        }, 100);
    }
}

// Display search results
function displayResults(results) {
    const resultsBody = document.getElementById('resultsBody');
    const resultCount = document.getElementById('resultCount');
    
    resultCount.textContent = results.length;

    if (results.length === 0) {
        resultsBody.innerHTML = `<tr><td colspan="${isAdminLoggedIn ? '10' : '9'}" class="no-results">No results to display.</td></tr>`;
        return;
    }

    resultsBody.innerHTML = results.map((student, index) => `
        <tr onclick="showStudentDetails(${index}, this)">
            <td>${student.fullName}</td>
            <td>${student.candidate}</td>
            <td>${student.dob}</td>
            <td>${student.grade}</td>
            <td>${student.school}</td>
            <td>${student.toan || ''}</td>
            <td>${student.kh || ''}</td>
            <td>${student.ta || ''}</td>
            <td>${student.certCode2 || student.certCode || ''}</td>
            <td class="admin-only" style="display: ${isAdminLoggedIn ? '' : 'none'};">
                <button class="send-email-btn" onclick="event.stopPropagation(); sendSingleEmail('${student.candidate}')">
                    📧 Send
                </button>
            </td>
        </tr>
    `).join('');

    // Store current results for detail view
    window.currentResults = results;
}

// Show student details
function showStudentDetails(index, row) {
    // Remove previous selection
    document.querySelectorAll('.results-table tbody tr').forEach(tr => {
        tr.classList.remove('selected');
    });
    
    // Add selection to clicked row
    row.classList.add('selected');

    const student = window.currentResults[index];
    const placeholder = document.querySelector('.info-placeholder');
    
    placeholder.style.display = 'none';
    document.getElementById('studentDetails').style.display = 'block';
    
    // Get medal class based on score
    function getMedalClass(score) {
        if (!score || score === 'nan' || score === 'NaN') return '';
        
        const upperScore = score.toString().toUpperCase();
        if (upperScore.includes('VÀNG') || upperScore.includes('VANG') || upperScore.includes('GOLD')) return 'gold';
        if (upperScore.includes('BẠC') || upperScore.includes('BAC') || upperScore.includes('SILVER')) return 'silver';
        if (upperScore.includes('ĐỒNG') || upperScore.includes('DONG') || upperScore.includes('BRONZE')) return 'bronze';
        if (upperScore.includes('KHUYẾN KHÍCH') || upperScore.includes('KHUYEN KHICH') || upperScore.includes('KK')) return 'encouragement';
        if (upperScore.includes('CHỨNG NHẬN') || upperScore.includes('CHUNG NHAN') || upperScore.includes('CN')) return 'certificate';
        return '';
    }
    
    // Update info boxes
    document.getElementById('candidateBox').textContent = student.candidate;
    document.getElementById('nameBox').textContent = student.fullName;
    document.getElementById('dobBox').textContent = student.dob;
    document.getElementById('gradeSchoolBox').textContent = `Grade ${student.grade} - ${student.school}`;
    
    // Update scores with colors
    const mathBadge = document.getElementById('mathScore');
    mathBadge.textContent = student.toan || '';
    mathBadge.className = 'score-badge ' + getMedalClass(student.toan);
    
    const scienceBadge = document.getElementById('scienceScore');
    scienceBadge.textContent = student.kh || '';
    scienceBadge.className = 'score-badge ' + getMedalClass(student.kh);
    
    const englishBadge = document.getElementById('englishScore');
    englishBadge.textContent = student.ta || '';
    englishBadge.className = 'score-badge ' + getMedalClass(student.ta);
    
    // Update cert code and medal summary
    document.getElementById('certBox').textContent = student.certCode || student.certCode2 || '';
    
    // Update photo
    const photoDiv = document.getElementById('studentPhoto');
    photoDiv.innerHTML = '';
    
    const img = document.createElement('img');
    img.style.width = '100%';
    img.style.height = '100%';
    img.style.objectFit = 'cover';
    img.alt = student.fullName;
    
    // Ưu tiên: 1. Base64 từ Excel, 2. Tên file từ Excel, 3. SBD.jpg từ thư mục photos/
    if (student.photo) {
        if (student.photo.startsWith('data:image')) {
            // Base64 image từ Excel
            img.src = student.photo;
        } else {
            // Tên file từ Excel, tìm trong thư mục photos/
            img.src = `photos/${student.photo}`;
        }
    } else {
        // Tìm ảnh theo SBD trong thư mục photos/
        img.src = `photos/${student.candidate}.jpg`;
    }
    
    // Xử lý lỗi nếu không tìm thấy ảnh
    img.onerror = function() {
        photoDiv.innerHTML = '<span>No Photo Available</span>';
    };
    
    photoDiv.appendChild(img);
    
    // Generate QR Code with Student Info
    const qrDiv = document.getElementById('studentQR');
    qrDiv.innerHTML = '';
    
    try {
        // Create canvas for QR code
        const canvas = document.createElement('canvas');
        const qr = new QRious({
            element: canvas,
            value: student.qrData,
            size: 200,
            background: 'white',
            foreground: 'black',
            level: 'M'
        });
        qrDiv.appendChild(canvas);
    } catch (error) {
        qrDiv.innerHTML = '<span style="color: red;">Error generating QR</span>';
        console.error('QR Code generation error:', error);
    }
}

// Download Excel with QR data
function downloadExcelWithQR() {
    if (studentsData.length === 0) {
        alert('Vui lòng tải dữ liệu học sinh trước!');
        return;
    }
    
    // Prepare data for Excel
    const excelData = studentsData.map(student => ({
        'FULL NAME': student.fullName,
        'SBD': student.candidate,
        'Ngày sinh': student.dob,
        'KHỐI': student.grade,
        'TRƯỜNG': student.school,
        'KHU VỰC': student.area,
        'KQ VQG TOÁN': student.toan,
        'KQ VQG KHOA HỌC': student.kh,
        'KQ VQG TIẾNG ANH': student.ta,
        'MÃ CERT ĐẦY ĐỦ': student.certCode,
        'MÃ CERT': student.certCode2,
        'Số điện thoại liên hệ': student.sdt,
        'Email liên hệ': student.email,
        'QR DATA': student.qrData
    }));
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    // Add sheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'DATA_WITH_QR');
    
    // Generate file and download
    XLSX.writeFile(wb, 'DS_KQ_WITH_QR.xlsx');
    
    alert('File Excel đã được tải xuống với dữ liệu QR!');
}

// Clear form
function clearForm() {
    document.getElementById('candidateInput').value = '';
    document.getElementById('fullNameInput').value = '';
    document.getElementById('daySelect').value = '';
    document.getElementById('monthSelect').value = '';
    document.getElementById('yearSelect').value = '';
    
    document.getElementById('resultsBody').innerHTML = '<tr><td colspan="9" class="no-results">No results to display.</td></tr>';
    document.getElementById('resultCount').textContent = '0';
    
    document.getElementById('studentDetails').style.display = 'none';
    document.querySelector('.info-placeholder').style.display = 'block';
    
    // Clear photo and QR
    document.getElementById('studentPhoto').innerHTML = '<span>No Photo</span>';
    document.getElementById('studentQR').innerHTML = '<span>No QR Code</span>';
}

// Admin functions
function adminLogin() {
    const password = document.getElementById('adminPassword').value;
    if (password === ADMIN_PASSWORD) {
        isAdminLoggedIn = true;
        document.getElementById('adminLoginForm').style.display = 'none';
        document.getElementById('adminPanel').style.display = 'flex';
        
        // Show admin columns in header
        document.querySelectorAll('.admin-only').forEach(el => {
            el.style.display = '';
            console.log('Showing admin column:', el);
        });
        
        updateEmailCount();
        
        // Refresh table to show Send buttons
        if (window.currentResults && window.currentResults.length > 0) {
            console.log('Refreshing table with', window.currentResults.length, 'results');
            displayResults(window.currentResults);
        } else {
            console.log('No current results to refresh - will show Actions after next search');
        }
        
        alert('✅ Đăng nhập Admin thành công!\n\nNếu đã có kết quả tìm kiếm, vui lòng click "Search" lại để hiển thị nút gửi email.');
    } else {
        alert('❌ Mật khẩu không đúng!');
    }
}

function adminLogout() {
    isAdminLoggedIn = false;
    document.getElementById('adminLoginForm').style.display = 'flex';
    document.getElementById('adminPanel').style.display = 'none';
    document.getElementById('adminPassword').value = '';
    
    // Hide admin columns
    document.querySelectorAll('.admin-only').forEach(el => {
        el.style.display = 'none';
    });
    
    // Refresh table to hide Send buttons
    if (window.currentResults && window.currentResults.length > 0) {
        displayResults(window.currentResults);
    }
}

function updateEmailCount() {
    const count = studentsData.filter(s => s.email && s.email.trim()).length;
    document.getElementById('emailCount').textContent = count;
    document.getElementById('sendAllBtn').disabled = count === 0;
    console.log('Email count updated:', count, 'out of', studentsData.length, 'students');
}

async function sendAllEmails() {
    if (!isAdminLoggedIn) {
        alert('❌ Vui lòng đăng nhập Admin!');
        return;
    }
    
    const studentsWithEmail = studentsData.filter(s => s.email);
    if (studentsWithEmail.length === 0) {
        alert('❌ Không có học sinh nào có email!');
        return;
    }
    
    if (!confirm(`Bạn có chắc muốn gửi email cho ${studentsWithEmail.length} học sinh?`)) {
        return;
    }
    
    try {
        const response = await fetch('http://localhost:8000/send-email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'send_all' })
        });
        
        const result = await response.json();
        if (result.status === 'success') {
            alert(`✅ ${result.message}\n\n📊 Tiến độ sẽ hiển thị trong terminal Python.`);
        } else {
            alert(`❌ Lỗi: ${result.message}`);
        }
    } catch (error) {
        alert(`⚠️ Không thể kết nối đến server.\n\nVui lòng chạy lệnh:\npython email_server.py\n\nHoặc chạy trực tiếp:\npython send_student_awards.py`);
        console.error(error);
    }
}

async function sendSingleEmail(candidate) {
    if (!isAdminLoggedIn) {
        alert('❌ Vui lòng đăng nhập Admin!');
        return;
    }
    
    const student = window.currentResults.find(s => String(s.candidate) === String(candidate));
    if (!student) {
        alert('❌ Không tìm thấy học sinh!');
        return;
    }
    
    if (!student.email) {
        alert('❌ Học sinh này không có email!');
        return;
    }
    
    if (!confirm(`Gửi email cho:\n${student.fullName}\nEmail: ${student.email}`)) {
        return;
    }
    
    try {
        const response = await fetch('http://localhost:8000/send-email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'send_single', sbd: candidate })
        });
        
        const result = await response.json();
        if (result.status === 'success') {
            alert(`✅ Đang gửi email cho ${student.fullName}...\n\n📊 Kiểm tra terminal để xem tiến độ.`);
        } else {
            alert(`❌ Lỗi: ${result.message}`);
        }
    } catch (error) {
        alert(`⚠️ Không thể kết nối đến server.\n\nVui lòng chạy lệnh:\npython email_server.py`);
        console.error(error);
    }
}

// Auto-load Excel file from server
async function autoLoadExcel() {
    try {
        // Try to load from server endpoint
        const response = await fetch('http://localhost:8000/get-data');
        
        if (!response.ok) {
            throw new Error('Server not responding');
        }
        
        const result = await response.json();
        
        if (result.status === 'success' && result.data) {
            studentsData = result.data;
            
            console.log(`✅ Đã load ${studentsData.length} học sinh từ server`);
            
            // Show all students initially
            window.currentResults = studentsData;
            displayResults(studentsData);
            updateStudentCount(studentsData.length);
            updateEmailCount();
        } else {
            console.log('💡 Chưa có dữ liệu - vui lòng upload file DATA KQ.xlsx');
        }
        
    } catch (error) {
        console.log('💡 Server chưa chạy hoặc chưa có dữ liệu - vui lòng upload file Excel');
        console.log('💡 Để tự động load dữ liệu: chạy "python email_server.py"');
    }
}

// Generate QR codes for uploaded DATA file
async function generateQRCodes() {
    console.log('🔲 Generate QR button clicked');
    
    // Disable button and show loading
    const btn = document.getElementById('generateQRBtn');
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '⏳ Đang tạo QR codes...';
    
    try {
        console.log('📤 Sending request to server...');
        
        // Call generate QR endpoint
        const response = await fetch('http://localhost:8000/process-and-send', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'generate_qr' })
        });
        
        console.log('📥 Response received:', response.status);
        
        const result = await response.json();
        console.log('📊 Result:', result);
        
        if (result.status === 'success') {
            alert(`✅ Tạo QR thành công!\n\n📊 Đã tạo QR cho ${result.count || 'tất cả'} học sinh.\n\n📁 File DS_KQ_WITH_QR.xlsx đã được tạo.\n\n🔄 Đang load dữ liệu...`);
            
            // Hide generate button, show download button
            btn.style.display = 'none';
            document.getElementById('downloadBtn').style.display = 'block';
            
            // Load the new file with QR
            await autoLoadExcel();
            
        } else {
            throw new Error(result.message || 'Failed to generate QR');
        }
        
    } catch (error) {
        console.error('❌ Error:', error);
        alert(`❌ Lỗi: ${error.message}\n\n💡 Đảm bảo:\n1. Server đang chạy: python email_server.py\n2. File DATA KQ.xlsx đã được upload\n\nHoặc chạy thủ công:\npython create_qr_for_all_students.py`);
        
        // Restore button
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
}

// Initialize on page load
window.onload = function() {
    initializeDateSelects();
    autoLoadExcel();
};

