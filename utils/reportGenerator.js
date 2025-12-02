// Report Generator - Generate Excel reports

import { formatCurrency, formatDate } from './dataFormatter.js';

// Generate and download Excel report
export async function generateReport(data, filename, sheetName = 'Data') {
    try {
        // Prepare data for export
        const exportData = prepareExportData(data);
        
        // Create worksheet
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        
        // Set column widths
        const wscols = getColumnWidths(exportData);
        worksheet['!cols'] = wscols;
        
        // Create workbook
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        // Generate and trigger download
        XLSX.writeFile(workbook, filename);
        
        return true;
    } catch (error) {
        console.error('Error generating report:', error);
        throw error;
    }
}

// Prepare data for export
function prepareExportData(data) {
    if (data.length === 0) return [];
    
    // Check if it's processed data (has gaji fields)
    const isProcessedData = data[0].gajiPokok !== undefined;
    
    if (isProcessedData) {
        // Processed data format
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Total Hari Kerja': item.totalHari,
            'Total Jam Kerja': item.totalJam.toFixed(2),
            'Jam Normal': item.jamNormal.toFixed(2),
            'Jam Lembur': item.jamLembur.toFixed(2),
            'Gaji Pokok': item.gajiPokok,
            'Uang Lembur': item.uangLembur,
            'Potongan Pajak': item.pajak,
            'Gaji Bersih': item.gajiBersih,
            'Keterangan': 'Data terhitung otomatis'
        }));
    } else {
        // Original data format
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi (jam)': item.durasi ? item.durasi.toFixed(2) : '',
            'Keterangan': item.jamKeluar ? 'Lengkap' : 'Masuk saja'
        }));
    }
}

// Get appropriate column widths
function getColumnWidths(data) {
    if (data.length === 0) return [];
    
    const firstRow = data[0];
    const columns = Object.keys(firstRow);
    
    return columns.map(col => {
        // Determine width based on column content
        const maxLength = Math.max(
            col.length,
            ...data.map(row => (row[col] ? row[col].toString().length : 0))
        );
        
        return { wch: Math.min(Math.max(maxLength + 2, 10), 50) };
    });
}

// Download file (wrapper for XLSX)
export function downloadFile(data, filename, sheetName) {
    return generateReport(data, filename, sheetName);
}

// Generate summary report
export function generateSummaryReport(processedData, originalData, config) {
    const summaryData = {
        metadata: {
            generated: new Date().toISOString(),
            totalEmployees: new Set(processedData.map(item => item.nama)).size,
            totalDays: processedData.reduce((sum, item) => sum + item.totalHari, 0),
            period: getDateRange(originalData),
            config: config
        },
        summary: {
            totalGajiPokok: processedData.reduce((sum, item) => sum + item.gajiPokok, 0),
            totalUangLembur: processedData.reduce((sum, item) => sum + item.uangLembur, 0),
            totalPajak: processedData.reduce((sum, item) => sum + item.pajak, 0),
            totalGajiBersih: processedData.reduce((sum, item) => sum + item.gajiBersih, 0)
        },
        employees: processedData.map(item => ({
            nama: item.nama,
            totalHari: item.totalHari,
            totalJam: item.totalJam,
            jamLembur: item.jamLembur,
            gajiPokok: item.gajiPokok,
            uangLembur: item.uangLembur,
            pajak: item.pajak,
            gajiBersih: item.gajiBersih
        }))
    };
    
    return summaryData;
}

// Get date range from original data
function getDateRange(data) {
    if (data.length === 0) return 'Tidak ada data';
    
    const dates = data
        .map(item => new Date(item.tanggal))
        .filter(date => !isNaN(date.getTime()))
        .sort((a, b) => a - b);
    
    if (dates.length === 0) return 'Tanggal tidak valid';
    
    const start = dates[0];
    const end = dates[dates.length - 1];
    
    return `${formatDate(start)} - ${formatDate(end)}`;
}
