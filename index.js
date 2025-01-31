// Enhanced Cap Table Template with Vesting and Share Classes
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
    const wb = XLSX.utils.book_new();
    
    // Create main cap table worksheet
    const mainTableData = [
        ['Share Class', 'Shareholder Type', 'Name', 'Number of Shares', 'Share Price', 'Investment Amount', 'Ownership %', 'Vesting Start', 'Vesting Schedule', 'Vested Shares', 'Unvested Shares'],
        
        // Common Stock - Founders
        ['Common Stock', 'Founders', 'Founder 1', 1000000, '$0.0001', '=$D3*E3', '=D3/SUBTOTAL(9,D3:D999)', '2025-01-01', '4 year / 1 year cliff', '=VESTED_SHARES(D3,H3,I3)', '=D3-J3'],
        ['Common Stock', 'Founders', 'Founder 2', 1000000, '$0.0001', '=$D4*E4', '=D4/SUBTOTAL(9,D3:D999)', '2025-01-01', '4 year / 1 year cliff', '=VESTED_SHARES(D4,H4,I4)', '=D4-J4'],
        ['', '', '', '', '', '', '', '', '', '', ''],
        
        // Preferred Stock - Series A
        ['Series A Preferred', 'Investors', 'VC Fund 1', 500000, '$1.00', '=$D6*E6', '=D6/SUBTOTAL(9,D3:D999)', '-', '-', '-', '-'],
        ['Series A Preferred', 'Investors', 'VC Fund 2', 250000, '$1.00', '=$D7*E7', '=D7/SUBTOTAL(9,D3:D999)', '-', '-', '-', '-'],
        ['', '', '', '', '', '', '', '', '', '', ''],
        
        // Option Pool - Common Stock
        ['Common Stock', 'Option Pool', 'Employee 1', 50000, '$0.50', '-', '=D9/SUBTOTAL(9,D3:D999)', '2025-03-01', '4 year / 1 year cliff', '=VESTED_SHARES(D9,H9,I9)', '=D9-J9'],
        ['Common Stock', 'Option Pool', 'Reserved', 450000, '-', '-', '=D10/SUBTOTAL(9,D3:D999)', '-', '-', '-', '-'],
        ['', '', '', '', '', '', '', '', '', '', ''],
        
        // Totals
        ['Totals', '', '', '=SUBTOTAL(9,D3:D999)', '', '=SUBTOTAL(9,F3:F999)', '=SUBTOTAL(9,G3:G999)', '', '', '=SUBTOTAL(9,J3:J999)', '=SUBTOTAL(9,K3:K999)']
    ];
    
    const ws_main = XLSX.utils.aoa_to_sheet(mainTableData);
    
    // Create share classes summary worksheet
    const shareClassesData = [
        ['Share Class', 'Total Shares', 'Liquidation Preference', 'Conversion Ratio', 'Anti-dilution', 'Voting Rights'],
        ['Common Stock', '=SUMIF(\'Cap Table\'!A3:A999,"Common Stock",\'Cap Table\'!D3:D999)', 'None', '1:1', '-', '1 vote per share'],
        ['Series A Preferred', '=SUMIF(\'Cap Table\'!A3:A999,"Series A Preferred",\'Cap Table\'!D3:D999)', '1x', '1:1', 'Broad-based weighted average', '1 vote per share']
    ];
    
    const ws_classes = XLSX.utils.aoa_to_sheet(shareClassesData);
    
    // Create vesting summary worksheet
    const vestingData = [
        ['Name', 'Total Shares', 'Vesting Start', 'Schedule Type', 'Vested to Date', 'Unvested', 'Next Vesting Date', 'Next Vesting Amount'],
        ['=\'Cap Table\'!C3', '=\'Cap Table\'!D3', '=\'Cap Table\'!H3', '=\'Cap Table\'!I3', '=\'Cap Table\'!J3', '=\'Cap Table\'!K3', '=NEXT_VESTING_DATE(H2,I2)', '=NEXT_VESTING_AMOUNT(D2,I2)'],
        ['=\'Cap Table\'!C4', '=\'Cap Table\'!D4', '=\'Cap Table\'!H4', '=\'Cap Table\'!I4', '=\'Cap Table\'!J4', '=\'Cap Table\'!K4', '=NEXT_VESTING_DATE(H3,I3)', '=NEXT_VESTING_AMOUNT(D3,I3)']
    ];
    
    const ws_vesting = XLSX.utils.aoa_to_sheet(vestingData);
    
    // Set column widths for main worksheet
    ws_main['!cols'] = [
        {wch: 15}, // Share Class
        {wch: 15}, // Shareholder Type
        {wch: 20}, // Name
        {wch: 15}, // Number of Shares
        {wch: 12}, // Share Price
        {wch: 15}, // Investment Amount
        {wch: 12}, // Ownership %
        {wch: 12}, // Vesting Start
        {wch: 20}, // Vesting Schedule
        {wch: 12}, // Vested Shares
        {wch: 12}  // Unvested Shares
    ];
    
    // Set column widths for share classes worksheet
    ws_classes['!cols'] = [
        {wch: 15}, // Share Class
        {wch: 12}, // Total Shares
        {wch: 20}, // Liquidation Preference
        {wch: 15}, // Conversion Ratio
        {wch: 25}, // Anti-dilution
        {wch: 15}  // Voting Rights
    ];
    
    // Set column widths for vesting worksheet
    ws_vesting['!cols'] = [
        {wch: 20}, // Name
        {wch: 12}, // Total Shares
        {wch: 12}, // Vesting Start
        {wch: 20}, // Schedule Type
        {wch: 12}, // Vested to Date
        {wch: 12}, // Unvested
        {wch: 15}, // Next Vesting Date
        {wch: 15}  // Next Vesting Amount
    ];
    
    // Add worksheets to workbook
    XLSX.utils.book_append_sheet(wb, ws_main, 'Cap Table');
    XLSX.utils.book_append_sheet(wb, ws_classes, 'Share Classes');
    XLSX.utils.book_append_sheet(wb, ws_vesting, 'Vesting Summary');
    
    // Format percentage and currency cells
    const formatRanges = [ws_main, ws_classes, ws_vesting];
    formatRanges.forEach(ws => {
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            // Format ownership percentages
            const pct_cell = ws[XLSX.utils.encode_cell({r: R, c: 6})];
            if (pct_cell) pct_cell.z = '0.00%';
            
            // Format currency amounts
            const curr_cell = ws[XLSX.utils.encode_cell({r: R, c: 5})];
            if (curr_cell) curr_cell.z = '$#,##0.00';
        }
    });
    
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
function VESTED_SHARES(totalShares, startDate, schedule) {
    // This would normally calculate vested shares based on the current date
    // For demonstration, returning 25% of total shares
    return totalShares * 0.25;
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
function NEXT_VESTING_DATE(startDate, schedule) {
    // This would normally calculate the next vesting date
    // For demonstration, returning a placeholder date
    return new Date(startDate.getTime() + (90 * 24 * 60 * 60 * 1000));
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
function NEXT_VESTING_AMOUNT(totalShares, schedule) {
    // This would normally calculate the next vesting amount
    // For demonstration, returning 1/16th of total shares (quarterly vesting)
    return totalShares / 16;
}

// Export to file
const workbook = createEnhancedCapTable();
XLSX.writeFile(workbook, 'enhanced_cap_table.xlsx');