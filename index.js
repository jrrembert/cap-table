// Enhanced Cap Table Template with Google Sheets Compatible Formulas
import XLSX from 'xlsx';

/**
 * Creates an enhanced capitalization (cap) table workbook with detailed financial and equity information.
 * 
 * @returns {Object} A workbook containing three worksheets:
 *  - 'Cap Table': Comprehensive breakdown of share ownership, vesting, and investment details
 *  - 'Share Classes': Summary of different share classes and their characteristics
 *  - 'Vesting Summary': Detailed tracking of individual shareholder vesting schedules
 * 
 * @description
 * Generates a multi-sheet Excel workbook that provides a comprehensive view of a company's equity structure.
 * The workbook includes:
 *  - Detailed shareholder information
 *  - Share class breakdowns
 *  - Vesting schedules for founders, employees, and option pool
 *  - Calculated ownership percentages and investment amounts
 * 
 * @example
 * const capTableWorkbook = createEnhancedCapTable();
 * XLSX.writeFile(capTableWorkbook, 'enhanced_cap_table.xlsx');
 */
function createEnhancedCapTable() {
    if (typeof XLSX === 'undefined') {
        throw new Error('XLSX library not loaded');
    }

    const wb = XLSX.utils.book_new();
    
    // Helper function to calculate vesting details for a shareholder
    function calculateVestingDetails(shares, startDate, schedule) {
        if (!startDate) return { vested: 0, nextDate: null, nextAmount: 0 };
        
        const start = new Date(startDate);
        if (isNaN(start.getTime())) return { vested: 0, nextDate: null, nextAmount: 0 };

        return {
            vested: get_vested_shares(shares, start, schedule),
            nextDate: get_next_vesting_date(start, schedule),
            nextAmount: get_next_vesting_amount(shares, schedule)
        };
    }

    // Pre-calculate vesting for each shareholder
    const shareholders = [
        { name: 'Founder 1', shares: 1000000, startDate: '1/1/2025', schedule: '4 year / 1 year cliff' },
        { name: 'Founder 2', shares: 1000000, startDate: '1/1/2025', schedule: '4 year / 1 year cliff' },
        // ... other shareholders
    ];

    const vestingCalculations = shareholders.reduce((acc, sh) => {
        acc[sh.name] = calculateVestingDetails(sh.shares, sh.startDate, sh.schedule);
        return acc;
    }, {});

    // Create main cap table worksheet with pre-calculated values
    const mainTableData = [
        ['Share Class', 'Shareholder Type', 'Name', 'Number of Shares', 'Share Price', 'Investment Amount', 'Ownership %', 'Vesting Start', 'Vesting Schedule', 'Vested Shares', 'Unvested Shares'],
        
        // Common Stock - Founders with pre-calculated vesting
        ['Common Stock', 'Founders', 'Founder 1', 1000000, 0.0001, '=D2*E2', '=D2/SUBTOTAL(9,D2:D999)', '1/1/2025', '4 year / 1 year cliff', 
         vestingCalculations['Founder 1'].vested,
         `=D2-J2`],
        ['Common Stock', 'Founders', 'Founder 2', 1000000, 0.0001, '=D3*E3', '=D3/SUBTOTAL(9,D2:D999)', '1/1/2025', '4 year / 1 year cliff',
         vestingCalculations['Founder 2'].vested,
         '=D3-J3'],
        
        // Preferred Stock - Series A
        ['Series A Preferred', 'Investors', 'VC Fund 1', 500000, 1.00, '=D4*E4', '=D4/SUBTOTAL(9,D2:D999)', '', '', '', ''],
        ['Series A Preferred', 'Investors', 'VC Fund 2', 250000, 1.00, '=D5*E5', '=D5/SUBTOTAL(9,D2:D999)', '', '', '', ''],
        
        // Option Pool - Common Stock
        ['Common Stock', 'Option Pool', 'Employee 1', 50000, 0.50, '', '=D6/SUBTOTAL(9,D2:D999)', '3/1/2025', '4 year / 1 year cliff', '=D6*0.25', '=D6-J6'],
        ['Common Stock', 'Option Pool', 'Reserved', 450000, '', '', '=D7/SUBTOTAL(9,D2:D999)', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', ''],
        
        // Totals
        ['Totals', '', '', '=SUBTOTAL(9,D2:D999)', '', '=SUBTOTAL(9,F2:F999)', '=SUBTOTAL(9,G2:G999)', '', '', '=SUBTOTAL(9,J2:J999)', '=SUBTOTAL(9,K2:K999)']
    ];
    
    const ws_main = XLSX.utils.aoa_to_sheet(mainTableData, {cellDates: true});
    
    // Create share classes summary worksheet
    const shareClassesData = [
        ['Share Class', 'Total Shares', 'Liquidation Preference', 'Conversion Ratio', 'Anti-dilution', 'Voting Rights'],
        ['Common Stock', '=SUMIF(\'Cap Table\'!A2:A999, "Common Stock", \'Cap Table\'!D2:D999)', 'None', '1:1', '-', '1 vote per share'],
        ['Series A Preferred', '=SUMIF(\'Cap Table\'!A2:A999, "Series A Preferred", \'Cap Table\'!D2:D999)', '1x', '1:1', 'Broad-based weighted average', '1 vote per share']
    ];
    
    const ws_classes = XLSX.utils.aoa_to_sheet(shareClassesData);
    
    // Create vesting summary worksheet with actual vesting calculations
    const vestingData = [
        ['Name', 'Total Shares', 'Vesting Start', 'Schedule Type', 'Vested to Date', 'Unvested', 'Next Vesting Date', 'Next Vesting Amount'],
        ['=\'Cap Table\'!C2', '=\'Cap Table\'!D2', '=\'Cap Table\'!H2', '=\'Cap Table\'!I2', 
         vestingCalculations['Founder 1'].vested,
         '=B2-E2',
         (vestingCalculations['Founder 1'].nextDate
           ? vestingCalculations['Founder 1'].nextDate.toISOString()
           : ""),
         vestingCalculations['Founder 1'].nextAmount],
        ['=\'Cap Table\'!C3', '=\'Cap Table\'!D3', '=\'Cap Table\'!H3', '=\'Cap Table\'!I3', 
         vestingCalculations['Founder 2'].vested,
         '=B3-E3',
         vestingCalculations['Founder 2'].nextDate.toISOString(),
         vestingCalculations['Founder 2'].nextAmount]
    ];
    
    const ws_vesting = XLSX.utils.aoa_to_sheet(vestingData);
    
    // Set column widths and formats
    ['!cols', '!rows'].forEach(prop => {
        [ws_main, ws_classes, ws_vesting].forEach(ws => {
            ws[prop] = ws[prop] || [];
        });
    });

    // Add formatting information
    const formats = {
        percentage: '0.00%',
        currency: '$#,##0.00',
        number: '#,##0',
        date: 'mm/dd/yyyy'
    };

    // Format cells in main worksheet
    for (let i = 2; i < mainTableData.length; i++) {
        // Format share price as currency
        if (ws_main[XLSX.utils.encode_cell({r: i, c: 4})]) {
            ws_main[XLSX.utils.encode_cell({r: i, c: 4})].z = formats.currency;
        }
        // Format investment amount as currency
        if (ws_main[XLSX.utils.encode_cell({r: i, c: 5})]) {
            ws_main[XLSX.utils.encode_cell({r: i, c: 5})].z = formats.currency;
        }
        // Format ownership as percentage
        if (ws_main[XLSX.utils.encode_cell({r: i, c: 6})]) {
            ws_main[XLSX.utils.encode_cell({r: i, c: 6})].z = formats.percentage;
        }
        // Format shares as numbers
        if (ws_main[XLSX.utils.encode_cell({r: i, c: 3})]) {
            ws_main[XLSX.utils.encode_cell({r: i, c: 3})].z = formats.number;
        }
    }
    
    // Add worksheets to workbook
    XLSX.utils.book_append_sheet(wb, ws_main, 'Cap Table');
    XLSX.utils.book_append_sheet(wb, ws_classes, 'Share Classes');
    XLSX.utils.book_append_sheet(wb, ws_vesting, 'Vesting Summary');
    
    return wb;
}

/**
 * Calculates the number of vested shares based on total shares, vesting start date, and vesting schedule.
 * 
 * @param {number} totalShares - The total number of shares subject to vesting.
 * @param {Date} startDate - The date when the vesting schedule begins.
 * @param {string} schedule - The vesting schedule type (e.g., 'quarterly', 'cliff').
 * @returns {number} The number of shares that have vested, currently a placeholder returning 25% of total shares.
 * @description
 * This is a placeholder implementation that returns a fixed 25% of total shares.
 * In a production environment, this function would dynamically calculate vested shares
 * based on the current date, start date, and specific vesting schedule.
 */
function get_vested_shares(totalShares, startDate, schedule) {
    if (!totalShares || !startDate || !schedule) {
        throw new Error('Missing required parameters');
    }
        
    const start = new Date(startDate);
    
    if (isNaN(start.getTime())) {
        throw new Error('Invalid start date');
    }
    
    const now = new Date();
    const monthsElapsed = (now.getFullYear() - start.getFullYear()) * 12 + now.getMonth() - start.getMonth();
    
    // Parse schedule (e.g., "4 year / 1 year cliff")
    const [duration, cliff] = parseSchedule(schedule);
    
    if (monthsElapsed < cliff) {
        return 0;
    }
    
    const vestedPercentage = Math.min(1, monthsElapsed / duration);
    return Math.floor(totalShares * vestedPercentage);
}

/**
 * Parses a vesting schedule string into duration and cliff periods (in months)
 * 
 * @param {string} schedule - Schedule string (e.g., "4 year / 1 year cliff")
 * @returns {[number, number]} Array containing [total duration in months, cliff period in months]
 */
function parseSchedule(schedule) {
    const defaultValues = [48, 12]; // 4 years total, 1 year cliff
    
    if (!schedule || typeof schedule !== 'string') {
        return defaultValues;
    }

    try {
        const parts = schedule.toLowerCase().split('/');
        const durationMatch = parts[0].match(/(\d+)\s*year/);
        const cliffMatch = parts[1]?.match(/(\d+)\s*year/);
        
        const duration = (durationMatch ? parseInt(durationMatch[1]) : 4) * 12;
        const cliff = (cliffMatch ? parseInt(cliffMatch[1]) : 1) * 12;
        
        return [duration, cliff];
    } catch (error) {
        console.warn(`Failed to parse schedule "${schedule}", using defaults`, error);
        return defaultValues;
    }
}

/**
 * Calculates the next vesting date based on a given start date and vesting schedule.
 * 
 * @param {Date} startDate - The initial date when vesting begins.
 * @param {string} [schedule='quarterly'] - The vesting schedule type (e.g., 'quarterly', 'annual').
 * @returns {Date} The next date when additional shares will vest.
 * @description
 * This is a placeholder implementation that returns a date 90 days after the start date.
 * In a production environment, this would be replaced with a more sophisticated calculation
 * that considers the specific vesting schedule and current date.
 */
function get_next_vesting_date(startDate, schedule = 'quarterly') {
    if (!startDate || !(startDate instanceof Date)) {
        throw new Error('Invalid start date');
    }

    const now = new Date();
    const start = new Date(startDate);
    
    if (isNaN(start.getTime())) {
        throw new Error('Invalid start date');
    }

    // If vesting hasn't started yet, return start date
    if (start > now) {
        return start;
    }

    const [duration, cliff] = parseSchedule(schedule);
    const monthsElapsed = (now.getFullYear() - start.getFullYear()) * 12 + 
                         (now.getMonth() - start.getMonth());

    // If we haven't reached the cliff yet, return cliff date
    if (monthsElapsed < cliff) {
        const cliffDate = new Date(start);
        cliffDate.setMonth(cliffDate.getMonth() + cliff);
        return cliffDate;
    }

    // If fully vested, return null
    if (monthsElapsed >= duration) {
        return null;
    }

    // Calculate next quarterly vesting date
    const quartersPassed = Math.floor(monthsElapsed / 3);
    const nextVestingDate = new Date(start);
    nextVestingDate.setMonth(start.getMonth() + (quartersPassed + 1) * 3);
    
    return nextVestingDate;
}

/**
 * Calculates the next vesting amount for a given number of total shares.
 * 
 * @param {number} totalShares - The total number of shares subject to vesting.
 * @param {string} [schedule='quarterly'] - The vesting schedule type (default is quarterly).
 * @returns {number} The amount of shares that will vest in the next vesting period.
 * @description
 * This function provides a simplified calculation of the next vesting amount, 
 * assuming a standard quarterly vesting schedule where shares vest in equal 
 * increments over time. For demonstration purposes, it returns 1/16th of the 
 * total shares, representing a typical 4-year vesting period with quarterly 
 * vest intervals.
 * 
 * @example
 * // Returns 625 (assuming 10,000 total shares)
 * const nextVest = NEXT_VESTING_AMOUNT(10000);
 * 
 * @note This is a placeholder implementation and should be replaced with 
 * more sophisticated vesting calculation logic in a production environment.
 */
function get_next_vesting_amount(totalShares, schedule = 'quarterly') {
    if (!totalShares || Number.isNaN(totalShares)) {
        throw new Error('Invalid total shares amount');
    }

    const [duration] = parseSchedule(schedule);
    const quartersTotal = Math.floor(duration / 3);
    if (quartersTotal <= 0) {
        return 0;
    }
    
    // Standard quarterly vesting amount (after cliff)
    const quarterlyAmount = Math.floor(totalShares / quartersTotal);
    
    return quarterlyAmount;
}
}

// Export to file
const workbook = createEnhancedCapTable();
XLSX.writeFile(workbook, 'enhanced_cap_table.xlsx');