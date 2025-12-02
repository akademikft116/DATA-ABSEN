// Data Formatter - Format data for display

// Format currency (Rupiah)
export function formatCurrency(amount) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(amount);
}

// Format date
export function formatDate(dateString) {
    if (!dateString) return '-';
    
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) {
            // Try parsing different formats
            if (typeof dateString === 'string') {
                if (dateString.includes('/')) {
                    // DD/MM/YYYY
                    const [day, month, year] = dateString.split('/');
                    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
                }
            }
            return dateString;
        }
        
        return date.toLocaleDateString('id-ID', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    } catch (error) {
        return dateString;
    }
}

// Format time
export function formatTime(timeString) {
    if (!timeString) return '-';
    
    try {
        if (typeof timeString === 'string') {
            // Handle different time formats
            if (timeString.includes(':')) {
                const [hours, minutes] = timeString.split(':');
                return `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
            } else if (timeString.includes('.')) {
                // Decimal hours (e.g., 8.5 = 08:30)
                const hours = Math.floor(timeString);
                const minutes = Math.round((timeString - hours) * 60);
                return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
            }
        }
        return timeString;
    } catch (error) {
        return timeString;
    }
}

// Calculate hours between two time strings
export function calculateHours(timeIn, timeOut) {
    if (!timeIn || !timeOut) return 0;
    
    try {
        // Parse time strings
        const parseTime = (timeStr) => {
            if (!timeStr) return null;
            
            if (typeof timeStr === 'string') {
                if (timeStr.includes(':')) {
                    const [hours, minutes] = timeStr.split(':').map(Number);
                    return { hours, minutes: minutes || 0 };
                } else if (timeStr.includes('.')) {
                    const decimal = parseFloat(timeStr);
                    const hours = Math.floor(decimal);
                    const minutes = Math.round((decimal - hours) * 60);
                    return { hours, minutes };
                } else {
                    const decimal = parseFloat(timeStr);
                    if (!isNaN(decimal)) {
                        const hours = Math.floor(decimal);
                        const minutes = Math.round((decimal - hours) * 60);
                        return { hours, minutes };
                    }
                }
            }
            return null;
        };
        
        const inTime = parseTime(timeIn);
        const outTime = parseTime(timeOut);
        
        if (!inTime || !outTime) return 0;
        
        // Calculate total minutes
        let totalMinutes = (outTime.hours * 60 + outTime.minutes) - 
                          (inTime.hours * 60 + inTime.minutes);
        
        // Handle overnight shifts
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        // Convert to hours
        return totalMinutes / 60;
        
    } catch (error) {
        console.error('Error calculating hours:', error);
        return 0;
    }
}

// Format duration
export function formatDuration(hours) {
    const wholeHours = Math.floor(hours);
    const minutes = Math.round((hours - wholeHours) * 60);
    
    if (wholeHours === 0) {
        return `${minutes} menit`;
    } else if (minutes === 0) {
        return `${wholeHours} jam`;
    } else {
        return `${wholeHours} jam ${minutes} menit`;
    }
}

// Sanitize string for Excel
export function sanitizeForExcel(str) {
    if (typeof str !== 'string') return str;
    
    // Remove characters that might break Excel
    return str.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
}
