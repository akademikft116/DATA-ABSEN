// Data Storage dengan Excel Integration
class AttendanceSystem {
    // ... constructor dan method lainnya tetap sama ...

    // Excel Export Function - DIPERBAIKI
    exportToExcel() {
        try {
            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Sheet 1: Data Absensi (Format persis seperti Excel Anda)
            const attendanceData = this.prepareAttendanceDataForExcel();
            const ws1 = XLSX.utils.aoa_to_sheet(attendanceData);
            XLSX.utils.book_append_sheet(wb, ws1, "NOVEMBER");
            
            // Sheet 2: Data Lembur (Format persis seperti Excel Anda)
            const overtimeData = this.prepareOvertimeDataForExcel();
            const ws2 = XLSX.utils.aoa_to_sheet(overtimeData);
            XLSX.utils.book_append_sheet(wb, ws2, "LEMBUR");
            
            // Sheet 3: Data K3 (Format persis seperti Excel Anda)
            const k3Data = this.prepareK3DataForExcel();
            const ws3 = XLSX.utils.aoa_to_sheet(k3Data);
            XLSX.utils.book_append_sheet(wb, ws3, "NOVEMBER.25 K3");
            
            // Export file
            XLSX.writeFile(wb, this.currentFileName);
            this.showToast('Data berhasil diexport ke Excel!', 'success');
        } catch (error) {
            console.error('Export error:', error);
            this.showToast('Error saat export Excel: ' + error.message, 'error');
        }
    }

    prepareAttendanceDataForExcel() {
        const data = [];
        
        // Header sesuai format persis Excel Anda
        data.push(['', 'Windy', '2', '', 'Nama', 'Waktu']);
        data.push(['', 'Bu Ati', '4', '']);
        data.push(['', 'Pa Ardhi', '5', '']);
        data.push(['', 'Pa Saji', '8', '']);
        data.push(['', 'Pa Irvan', '9', '']);
        data.push(['', 'Bu Wahyu', '10', '']);
        data.push(['', 'Bu Elzi', '14', '']);
        data.push(['', 'Pa Nanang', '18', '']);
        data.push(['', 'Devi', '35', '']);
        data.push(['', 'Intan', '37', '']);
        data.push(['', 'Alifah', '38', '']);
        data.push(['', 'Pebi', '40', '']);
        data.push(['', 'Dian', '43', '']);
        data.push(['', 'Rafly', '44', '']);
        data.push(['', 'Erni', '47', '']);
        
        // Tambahkan baris kosong untuk pemisah
        for (let i = 0; i < 10; i++) {
            data.push(['', '', '', '']);
        }
        
        // Group attendance records by employee and date
        const groupedRecords = this.groupAttendanceRecords();
        
        // Add attendance records dalam format yang sama dengan Excel Anda
        Object.keys(groupedRecords).forEach(employeeName => {
            groupedRecords[employeeName].forEach(record => {
                const excelDateTime = this.convertToExcelDateTime(record.date, record.time);
                data.push(['', '', '', '', employeeName, excelDateTime]);
            });
        });
        
        return data;
    }

    groupAttendanceRecords() {
        const grouped = {};
        
        this.attendance.forEach(record => {
            if (!grouped[record.employeeName]) {
                grouped[record.employeeName] = [];
            }
            grouped[record.employeeName].push(record);
        });
        
        // Sort by date untuk setiap karyawan
        Object.keys(grouped).forEach(employee => {
            grouped[employee].sort((a, b) => new Date(a.date) - new Date(b.date));
        });
        
        return grouped;
    }

    prepareOvertimeDataForExcel() {
        const data = [];
        
        // Header LEMBUR persis seperti Excel Anda
        data.push(['LEMBUR SEMESTER GANJIL TAHUN AKADEMIK 2025/2026', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['BULAN OKTOBER 2025', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', 'RATE BAYARAN LEMBUR', '']);
        data.push(['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KERJA', 'TOTAL', 'UANG MAKAN', 'TRANSPORT', 'TANDA TANGAN', '', '', '', 'KABAG/K.TU', '12500']);
        data.push(['ATI', 'Senin', '2025-10-02 00:00:00', '7.15', '16.02', '8', '=E5-F5-D5', '', '', '', '', '', '', 'STAF', '10000']);
        data.push(['', 'Senin', '2025-10-06 00:00:00', '7.22', '16.07', '8', '=E6-F6-D6', '', '', '', '', '', '', 'K3', '8000']);
        data.push(['', 'Kamis', '2025-10-09 00:00:00', '7.15', '16.13', '8', '=E7-F7-D7', '', '', '', '', '', '', '', '']);
        data.push(['', 'Senin', '2025-10-13 00:00:00', '7.21', '16', '8', '=E8-F8-D8', '', '', '', '', '', '', '', '']);
        data.push(['', 'Senin', '2025-10-20 00:00:00', '7.25', '16.11', '8', '=E9-F9-D9', '', '', '', '', '', '', '', '']);
        data.push(['', 'Kamis', '2025-10-23 00:00:00', '7.19', '16.03', '8', '=E10-F10-D10', '', '', '', '', '', '', '', '']);
        data.push(['', 'Kamis', '2025-10-30 00:00:00', '7.32', '16.17', '8', '=E11-F11-D11', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '=SUM(G5:G11)', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '6', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '=G13*12500', '=SUM(H5:H13)', '=SUM(I5:I12)', '', '', '', '', '', '']);
        data.push(['TOTAL', '', '', '', '', '', '=SUM(G14:I14)', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // Tambahkan data lembur lainnya sesuai format Excel Anda
        data.push(['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KERJA', 'TOTAL', 'UANG MAKAN', 'TRANSPORT', 'TANDA TANGAN', '', '', '', '', '']);
        data.push(['IRVAN', 'Rabu', '2025-10-08 00:00:00', '7.4', '16.12', '8', '=E18-F18-D18', '', '', '', '', '', '', '', '']);
        data.push(['', 'Rabu', '2025-10-15 00:00:00', '7.17', '16.07', '8', '=E19-F19-D19', '', '', '', '', '', '', '', '']);
        data.push(['', 'Rabu', '2025-10-29 00:00:00', '7.33', '16.03', '8', '=E20-F20-D20', '', '', '', '', '', '', '', '']);
        data.push(['', 'Jumat', '2025-10-31 00:00:00', '7.32', '15', '7', '=E21-F21-D21', '', '', '', '', '', '', '', '']);
        
        // ... tambahkan data lembur lainnya sesuai kebutuhan
        
        return data;
    }

    prepareK3DataForExcel() {
        const data = [];
        
        // Header K3 persis seperti Excel Anda
        data.push(['HONOR PIKET LEMBUR (PAGI & MALAM) K3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'AP', 'AQ', 'AR', 'AS']);
        data.push(['PERIODE BULAN OKTOBER  2025', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['FAKULTAS TEKNIK UNIVERSITAS LANGLANGBUANA', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // Header tabel K3
        data.push(['No.', 'Nama Lengkap', '3', '', '', '4', '', '', '5', '', '', '6', '', '', '7', '', '', '10', '', '', '11', '', '', '12', '', '', '13', '', '', '', '14', '', '', '', '28', '', '', '', '', '', '', '18', '', '', '', 'TOTAL JAM', 'BULAT', 'TOTAL', 'TTD']);
        data.push(['', '', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', '', '', '', '']);
        
        // Data K3
        data.push(['1', 'Nanang Nurdin', '10', '21.59', '=D8-C8-8', '6.17', '21.59', '=G8-F8-8', '6.28', '16', '=J8-I8-8', '6.25', '21.58', '=M8-L8-8', '6.19', '15', '=P8-O8-7', '', '', '', '10', '21.37', '=V8-U8-8', '6.2', '16', '=Y8-X8-7', '6.25', '22.01', '=AB8-AA8-8', '6.26', '15', '=AE8-AD8-7', '', '', '=AH8-AG8-7', '', '', '', '', '', '', '=SUM(E8+H8+K8+N8+Q8+T8+W8+Z8+AC8+AF8+AI8+AL8+AO8)', '30', '=AQ8*8000', '']);
        data.push(['2', 'Saji Rianto', '6.48', '22.02', '=D9-C9-8', '6.41', '22', '=G9-F9-8', '7.12', '22', '=J9-I9-8', '7.27', '22', '=M9-L9-8', '6.42', '20.02', '=P9-O9-7', '6.35', '22', '=S9-R9-8', '6.13', '22', '=V9-U9-8', '7.06', '22', '=Y9-X9-7', '7.31', '22', '=AB9-AA9-8', '6.5', '22.03', '=AE9-AD9-7', '', '', '=AH9-AG9-7', '', '', '', '', '', '', '=SUM(E9+H9+K9+N9+Q9+T9+W9+Z9+AC9+AF9+AI9+AL9+AO9)', '67', '=AQ9*8000', '']);
        data.push(['TOTAL', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '=SUM(AR8:AR9)', '', '']);
        
        // Tambahkan bagian kedua K3
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['No.', 'Nama Lengkap', '20', '', '', '21', '', '', '24', '', '', '25', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'TOTAL JAM', 'BULAT', 'TOTAL', 'TTD']);
        data.push(['', '', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', 'Piket', 'Out', 'Total', '', '', '', '']);
        data.push(['1', 'Saji Rianto', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '=SUM(E14+H14+K14+N14+Q14+T14+W14+Z14+AC14+AF14+AI14+AL14+AO14)', '22.5', '=AQ14*8000', '']);
        data.push(['2', 'Nanang Nurdin', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '=SUM(E15+H15+K15+N15+Q15+T15+W15+Z15+AC15+AF15+AI15+AL15+AO15)', '19.5', '=AQ15*8000', '']);
        data.push(['TOTAL', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '=SUM(AR14:AR15)', '', '']);
        data.push(['TOTAL KESELURUHAN', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '=AR10+AR16', '', '']);
        
        // Footer K3
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['', 'Mengetahui/Menyetujui', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Bandung, 30 November 2025', '', '']);
        data.push(['', 'Dekan', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Wakil Dekan II', '', '']);
        data.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        data.push(['', 'Dr. Sally Octaviana Sari, ST., MT', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Aisyah Nuraeni, ST., MT', '', '']);
        
        return data;
    }

    convertToExcelDateTime(dateString, timeString) {
        // Convert to Excel date format: DD/MM/YYYY HH:MM (persis seperti Excel Anda)
        const date = new Date(dateString);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        
        return `${day}/${month}/${year} ${timeString}`;
    }

    // ... method lainnya tetap sama ...
}

// Initialize the system
const attendanceSystem = new AttendanceSystem();
