// Fungsi untuk mengubah format waktu
function formatTime(timeStr) {
    if (!timeStr) return "-";
    
    // Jika waktu sudah dalam format HH:MM
    if (timeStr.includes(":")) {
        return timeStr;
    }
    
    // Jika waktu dalam format desimal (misal 7.15)
    if (timeStr.includes(".")) {
        const parts = timeStr.split(".");
        const hours = parts[0].padStart(2, '0');
        const minutes = Math.round(parseFloat("0." + parts[1]) * 60).toString().padStart(2, '0');
        return `${hours}:${minutes}`;
    }
    
    return timeStr;
}

// Fungsi untuk menghitung durasi kerja
function calculateDuration(timeIn, timeOut) {
    if (!timeIn || !timeOut || timeIn === "-" || timeOut === "-") return "-";
    
    try {
        // Konversi waktu ke menit
        const [inHours, inMinutes] = timeIn.split(":").map(Number);
        const [outHours, outMinutes] = timeOut.split(":").map(Number);
        
        const totalInMinutes = inHours * 60 + inMinutes;
        const totalOutMinutes = outHours * 60 + outMinutes;
        
        // Jika waktu keluar lebih kecil dari waktu masuk, tambah 24 jam
        const durationMinutes = totalOutMinutes >= totalInMinutes 
            ? totalOutMinutes - totalInMinutes 
            : (totalOutMinutes + 24*60) - totalInMinutes;
        
        const hours = Math.floor(durationMinutes / 60);
        const minutes = durationMinutes % 60;
        
        return `${hours} jam ${minutes} menit`;
    } catch (e) {
        return "-";
    }
}

// Fungsi untuk menampilkan data absensi
function displayAttendanceData(filteredData = attendanceData) {
    const tbody = document.getElementById('attendanceData');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="5" class="empty-state">
                    <div>📊</div>
                    <p>Tidak ada data absensi yang sesuai dengan filter</p>
                </td>
            </tr>
        `;
        return;
    }
    
    tbody.innerHTML = '';
    
    filteredData.forEach(record => {
        const row = document.createElement('tr');
        const formattedTimeIn = formatTime(record.timeIn);
        const formattedTimeOut = formatTime(record.timeOut);
        const duration = calculateDuration(formattedTimeIn, formattedTimeOut);
        
        row.innerHTML = `
            <td><strong>${record.name}</strong></td>
            <td class="date-cell">${formatDate(record.date)}</td>
            <td class="time-cell">${formattedTimeIn}</td>
            <td class="time-cell">${formattedTimeOut}</td>
            <td class="time-cell">${duration}</td>
        `;
        tbody.appendChild(row);
    });
}

// Fungsi untuk menampilkan data lembur
function displayOvertimeData(filteredData = overtimeData) {
    const tbody = document.getElementById('overtimeData');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="6" class="empty-state">
                    <div>⏰</div>
                    <p>Tidak ada data lembur yang sesuai dengan filter</p>
                </td>
            </tr>
        `;
        return;
    }
    
    tbody.innerHTML = '';
    
    filteredData.forEach(record => {
        const row = document.createElement('tr');
        const formattedTimeIn = formatTime(record.timeIn);
        const formattedTimeOut = formatTime(record.timeOut);
        const overtimePay = record.hours * 10000;
        
        row.innerHTML = `
            <td><strong>${record.name}</strong></td>
            <td class="date-cell">${formatDate(record.date)}</td>
            <td class="time-cell">${formattedTimeIn}</td>
            <td class="time-cell">${formattedTimeOut}</td>
            <td class="time-cell overtime-hours">${record.hours} jam</td>
            <td class="time-cell overtime-hours">Rp ${overtimePay.toLocaleString('id-ID')}</td>
        `;
        tbody.appendChild(row);
    });
}

// Fungsi untuk menampilkan data lembur K3
function displayK3OvertimeData() {
    const tbody = document.getElementById('k3OvertimeData');
    tbody.innerHTML = '';
    
    k3OvertimeData.forEach(record => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong>${record.name}</strong></td>
            <td class="time-cell">${record.totalHours} jam</td>
            <td class="time-cell">${record.roundedHours} jam</td>
            <td class="time-cell overtime-hours">Rp ${record.totalPayment.toLocaleString('id-ID')}</td>
        `;
        tbody.appendChild(row);
    });
}

// Fungsi untuk memformat tanggal
function formatDate(dateStr) {
    const date = new Date(dateStr);
    return date.toLocaleDateString('id-ID', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

// Fungsi untuk mengisi dropdown filter karyawan
function populateEmployeeFilters() {
    const employeeFilter = document.getElementById('employeeFilter');
    const overtimeEmployeeFilter = document.getElementById('overtimeEmployeeFilter');
    
    // Dapatkan daftar karyawan unik dari data absensi
    const employees = [...new Set(attendanceData.map(record => record.name))].sort();
    
    // Kosongkan dropdown terlebih dahulu
    employeeFilter.innerHTML = '<option value="all">Semua Karyawan</option>';
    overtimeEmployeeFilter.innerHTML = '<option value="all">Semua Karyawan</option>';
    
    // Tambahkan setiap karyawan ke dropdown
    employees.forEach(employee => {
        const option1 = document.createElement('option');
        option1.value = employee;
        option1.textContent = employee;
        employeeFilter.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = employee;
        option2.textContent = employee;
        overtimeEmployeeFilter.appendChild(option2);
    });
}

// Fungsi untuk menghitung ringkasan
function calculateSummary() {
    // Hitung total karyawan
    const employees = [...new Set(attendanceData.map(record => record.name))];
    document.getElementById('totalEmployees').textContent = employees.length;
    
    // Hitung total kehadiran
    document.getElementById('totalAttendance').textContent = attendanceData.length;
    
    // Hitung total jam lembur
    const totalOvertimeHours = overtimeData.reduce((sum, record) => sum + record.hours, 0);
    document.getElementById('totalOvertime').textContent = totalOvertimeHours + " jam";
    
    // Buat ringkasan per karyawan
    const summaryTbody = document.getElementById('summaryData');
    summaryTbody.innerHTML = '';
    
    employees.forEach(employee => {
        const employeeRecords = attendanceData.filter(record => record.name === employee);
        const workDays = employeeRecords.length;
        
        // Hitung rata-rata jam kerja
        let totalWorkMinutes = 0;
        let validRecords = 0;
        
        employeeRecords.forEach(record => {
            const timeIn = formatTime(record.timeIn);
            const timeOut = formatTime(record.timeOut);
            
            if (timeIn !== "-" && timeOut !== "-") {
                const [inHours, inMinutes] = timeIn.split(":").map(Number);
                const [outHours, outMinutes] = timeOut.split(":").map(Number);
                
                const totalInMinutes = inHours * 60 + inMinutes;
                const totalOutMinutes = outHours * 60 + outMinutes;
                const durationMinutes = totalOutMinutes >= totalInMinutes 
                    ? totalOutMinutes - totalInMinutes 
                    : (totalOutMinutes + 24*60) - totalInMinutes;
                
                totalWorkMinutes += durationMinutes;
                validRecords++;
            }
        });
        
        const avgWorkHours = validRecords > 0 
            ? `${Math.floor(totalWorkMinutes / validRecords / 60)} jam ${Math.floor((totalWorkMinutes / validRecords) % 60)} menit`
            : "-";
        
        // Hitung total jam lembur
        const employeeOvertime = overtimeData.filter(record => 
            record.name.toLowerCase().includes(employee.toLowerCase().split(' ')[1] || employee.toLowerCase())
        );
        const totalOvertime = employeeOvertime.reduce((sum, record) => sum + record.hours, 0);
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong>${employee}</strong></td>
            <td class="time-cell">${workDays} hari</td>
            <td class="time-cell">${avgWorkHours}</td>
            <td class="time-cell overtime-hours">${totalOvertime} jam</td>
        `;
        summaryTbody.appendChild(row);
    });
}

// Fungsi untuk memfilter data absensi
function filterAttendanceData() {
    const employeeFilter = document.getElementById('employeeFilter').value;
    const dateFilter = document.getElementById('dateFilter').value;
    
    let filteredData = attendanceData;
    
    // Filter berdasarkan karyawan
    if (employeeFilter !== 'all') {
        filteredData = filteredData.filter(record => record.name === employeeFilter);
    }
    
    // Filter berdasarkan tanggal
    if (dateFilter) {
        filteredData = filteredData.filter(record => record.date === dateFilter);
    }
    
    displayAttendanceData(filteredData);
}

// Fungsi untuk memfilter data lembur
function filterOvertimeData() {
    const employeeFilter = document.getElementById('overtimeEmployeeFilter').value;
    
    let filteredData = overtimeData;
    
    // Filter berdasarkan karyawan
    if (employeeFilter !== 'all') {
        filteredData = filteredData.filter(record => 
            record.name.toLowerCase().includes(employeeFilter.toLowerCase().split(' ')[1] || employeeFilter.toLowerCase())
        );
    }
    
    displayOvertimeData(filteredData);
}

// Fungsi untuk reset filter
function resetFilters() {
    document.getElementById('employeeFilter').value = 'all';
    document.getElementById('dateFilter').value = '';
    document.getElementById('overtimeEmployeeFilter').value = 'all';
    
    displayAttendanceData();
    displayOvertimeData();
}

// Fungsi untuk inisialisasi tab
function initializeTabs() {
    const tabs = document.querySelectorAll('.tab');
    const tabContents = document.querySelectorAll('.tab-content');
    
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabId = tab.getAttribute('data-tab');
            
            // Hapus class active dari semua tab dan konten
            tabs.forEach(t => t.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));
            
            // Tambah class active ke tab dan konten yang diklik
            tab.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
}

// Fungsi untuk inisialisasi event listeners
function initializeEventListeners() {
    // Filter event listeners
    document.getElementById('employeeFilter').addEventListener('change', filterAttendanceData);
    document.getElementById('dateFilter').addEventListener('change', filterAttendanceData);
    document.getElementById('overtimeEmployeeFilter').addEventListener('change', filterOvertimeData);
    document.getElementById('resetFilters').addEventListener('click', resetFilters);
}

// Fungsi inisialisasi aplikasi
function initializeApp() {
    initializeTabs();
    populateEmployeeFilters();
    initializeEventListeners();
    
    // Tampilkan data awal
    displayAttendanceData();
    displayOvertimeData();
    displayK3OvertimeData();
    calculateSummary();
}

// Jalankan aplikasi ketika DOM siap
document.addEventListener('DOMContentLoaded', initializeApp);
