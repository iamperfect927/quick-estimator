import ExcelJS from 'exceljs';
import { CalculationInput } from '@/app/types/estimation';

export async function generateEstimateExcel(data: CalculationInput): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Solar Cost Estimate');

  // Set grid lines visibility setting
  sheet.views = [{ showGridLines: true }];

  // Column definitions & widths
  sheet.columns = [
    { key: 'desc', width: 70 },
    { key: 'qty', width: 25 },
    { key: 'rate', width: 35 },
    { key: 'total', width: 35 }
  ];

  // Helper colors
  const emeraldGreen = '059669'; // Primary Accent
  const slateDark = '0F172A'; // Dark Headers
  const slateLight = 'F8FAFC'; // Alternating row
  const borderGray = 'E2E8F0'; // Borders
  
  // Thin border definition with strict literal typing
  const thinBorder = {
    top: { style: 'thin' as const, color: { argb: 'FF' + borderGray } },
    bottom: { style: 'thin' as const, color: { argb: 'FF' + borderGray } },
    left: { style: 'thin' as const, color: { argb: 'FF' + borderGray } },
    right: { style: 'thin' as const, color: { argb: 'FF' + borderGray } }
  };

  // 1. BRAND LOGO BLOCK (Rows 1-2 merged)
  sheet.mergeCells('A1:D2');
  const logoCell = sheet.getCell('A1');
  logoCell.value = 'BLUE FRAMES SOLAR ESTIMATE';
  logoCell.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  logoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + emeraldGreen } };
  logoCell.alignment = { vertical: 'middle', horizontal: 'center' };
  
  sheet.getRow(1).height = 20;
  sheet.getRow(2).height = 20;

  // 2. CUSTOMER & PROJECT BANNER (Row 3)
  sheet.mergeCells('A3:D3');
  const subtitleCell = sheet.getCell('A3');
  subtitleCell.value = data.customerName
    ? `${data.customerName.toUpperCase()} — CONFIDENTIAL SYSTEM ESTIMATE`
    : 'CONFIDENTIAL SYSTEM ESTIMATE';
  subtitleCell.font = { name: 'Segoe UI', size: 9, bold: true, italic: true, color: { argb: 'FF94A3B8' } };
  subtitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + slateDark } };
  subtitleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getRow(3).height = 20;

  // 3. METRICS BANNER (if ingestion data available)
  if (data.metrics) {
    // Row 4: metric labels row
    const metricLabelRow = sheet.addRow([
      '📊 Peak Load (kW)',
      'Day Consumption (kWh)',
      'Night Consumption (kWh)',
      `${data.metrics.apartmentCount ?? '—'} Apartments`
    ]);
    metricLabelRow.height = 18;
    metricLabelRow.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 8, bold: true, color: { argb: 'FF64748B' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    });

    // Row 5: metric values
    const metricValueRow = sheet.addRow([
      data.metrics.peakKW?.toFixed(2) ?? '—',
      data.metrics.dayConsumptionKWh?.toFixed(2) ?? '—',
      data.metrics.nightConsumptionKWh?.toFixed(2) ?? '—',
      `${data.metrics.totalDeviceCount ?? '—'} Devices`
    ]);
    metricValueRow.height = 22;
    metricValueRow.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF10B981' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      cell.border = { bottom: { style: 'medium' as const, color: { argb: 'FF10B981' } } };
    });
  }

  // Add empty separator row
  sheet.addRow([]);

  // Function to style section headers
  const addSectionHeader = (title: string) => {
    const headerRow = sheet.addRow([title]);
    sheet.mergeCells(`A${headerRow.number}:D${headerRow.number}`);
    const cell = headerRow.getCell(1);
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF' + emeraldGreen } };
    headerRow.height = 24;
    
    // Add border below
    cell.border = { bottom: { style: 'medium' as const, color: { argb: 'FF' + emeraldGreen } } };
    
    sheet.addRow([]); // space
  };

  // Function to add table headers
  const addTableHeadersMaterial = () => {
    const tableHeaderRow = sheet.addRow([
      'Component Description',
      'Quantity',
      'Unit Price',
      'Calculated Cost'
    ]);
    
    tableHeaderRow.height = 24;
    tableHeaderRow.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + slateDark } };
      cell.alignment = { 
        vertical: 'middle', 
        horizontal: colNum === 1 ? 'left' : colNum === 2 ? 'center' : 'right' 
      };
      cell.border = thinBorder;
    });
  };

  const addTableHeadersLabor = () => {
    const tableHeaderRow = sheet.addRow([
      'Description',
      '—',
      '—',
      'Calculated Cost'
    ]);
    
    tableHeaderRow.height = 24;
    tableHeaderRow.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + slateDark } };
      cell.alignment = { 
        vertical: 'middle', 
        horizontal: colNum === 1 ? 'left' : colNum === 2 ? 'center' : colNum === 3 ? 'center' : 'right' 
      };
      cell.border = thinBorder;
    });
  };

  // --- SECTION 1: MATERIALS ---
  addSectionHeader('1. MATERIALS & DEPLOYED COMPONENT INVENTORY');
  addTableHeadersMaterial();

  let materialSubtotal = 0;
  data.materials.forEach((m, idx) => {
    const total = m.quantity * m.unitPrice;
    materialSubtotal += total;

    const row = sheet.addRow([m.name, m.quantity, m.unitPrice, total]);
    row.height = 20;
    
    const isEven = idx % 2 === 0;

    row.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF334155' } };
      cell.border = thinBorder;
      if (isEven) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + slateLight } };
      }
      
      if (colNum === 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      } else if (colNum === 2) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.numFmt = '#,##0';
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.numFmt = '#,##0" XAF"';
      }
    });
  });

  // Materials Subtotal Row
  const matSubtotalRow = sheet.addRow(['Materials Subtotal', '', '', materialSubtotal]);
  matSubtotalRow.height = 22;
  sheet.mergeCells(`A${matSubtotalRow.number}:C${matSubtotalRow.number}`);
  matSubtotalRow.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF334155' } };
  matSubtotalRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'right' };
  matSubtotalRow.getCell(4).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + emeraldGreen } };
  matSubtotalRow.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };
  matSubtotalRow.getCell(4).numFmt = '#,##0" XAF"';
  matSubtotalRow.getCell(4).border = {
    top: { style: 'thin' as const, color: { argb: 'FF' + emeraldGreen } },
    bottom: { style: 'thin' as const, color: { argb: 'FF' + emeraldGreen } }
  };

  // spacing
  sheet.addRow([]);
  sheet.addRow([]);

  // --- SECTION 2: LABOR ---
  addSectionHeader('2. INSTALLATION & COMMISSIONING LABOR');
  addTableHeadersLabor();

  let laborSubtotal = 0;
  data.labor.forEach((l, idx) => {
    // Labor cost is a flat fee (30% of materials) — hourlyRate holds the total amount, hours=1 is sentinel
    const total = l.hourlyRate; // flat cost
    laborSubtotal += total;

    const row = sheet.addRow([l.description, '—', '—', total]);
    row.height = 20;

    const isEven = idx % 2 === 0;

    row.eachCell((cell, colNum) => {
      cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF334155' } };
      cell.border = thinBorder;
      if (isEven) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + slateLight } };
      }

      if (colNum === 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      } else if (colNum === 2 || colNum === 3) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { ...cell.font, italic: true, color: { argb: 'FF94A3B8' } };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.numFmt = '#,##0" XAF"';
      }
    });
  });

  // Labor Subtotal Row
  const labSubtotalRow = sheet.addRow(['Technical Labor Subtotal', '', '', laborSubtotal]);
  labSubtotalRow.height = 22;
  sheet.mergeCells(`A${labSubtotalRow.number}:C${labSubtotalRow.number}`);
  labSubtotalRow.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF334155' } };
  labSubtotalRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'right' };
  labSubtotalRow.getCell(4).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + emeraldGreen } };
  labSubtotalRow.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };
  labSubtotalRow.getCell(4).numFmt = '#,##0" XAF"';
  labSubtotalRow.getCell(4).border = {
    top: { style: 'thin' as const, color: { argb: 'FF' + emeraldGreen } },
    bottom: { style: 'thin' as const, color: { argb: 'FF' + emeraldGreen } }
  };

  // spacing
  sheet.addRow([]);
  sheet.addRow([]);

  // --- SECTION 3: GRAND SUMMARY ---
  addSectionHeader('3. PROJECT ESTIMATION OVERALL SUMMARY');

  const baseSubtotal = materialSubtotal + laborSubtotal;
  const marginMultiplier = data.marginPercentage / 100;
  const marginMarkup = baseSubtotal * marginMultiplier;
  const grandTotal = baseSubtotal + marginMarkup;

  // Raw Subtotal
  const subtotalSummaryRow = sheet.addRow(['Subtotal', '', '', baseSubtotal]);
  subtotalSummaryRow.height = 22;
  sheet.mergeCells(`A${subtotalSummaryRow.number}:C${subtotalSummaryRow.number}`);
  subtotalSummaryRow.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF475569' } };
  subtotalSummaryRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'right' };
  subtotalSummaryRow.getCell(4).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF334155' } };
  subtotalSummaryRow.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };
  subtotalSummaryRow.getCell(4).numFmt = '#,##0" XAF"';
  subtotalSummaryRow.getCell(4).border = thinBorder;

  // Margin Markup Row
  const marginSummaryRow = sheet.addRow([`Overhead Margin & Risk Markup (${data.marginPercentage}%)`, '', '', marginMarkup]);
  marginSummaryRow.height = 22;
  sheet.mergeCells(`A${marginSummaryRow.number}:C${marginSummaryRow.number}`);
  marginSummaryRow.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF475569' } };
  marginSummaryRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'right' };
  marginSummaryRow.getCell(4).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF334155' } };
  marginSummaryRow.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };
  marginSummaryRow.getCell(4).numFmt = '#,##0" XAF"';
  marginSummaryRow.getCell(4).border = thinBorder;

  // Grand Total Row
  const grandTotalSummaryRow = sheet.addRow(['TOTAL', '', '', grandTotal]);
  grandTotalSummaryRow.height = 30;
  sheet.mergeCells(`A${grandTotalSummaryRow.number}:C${grandTotalSummaryRow.number}`);
  
  grandTotalSummaryRow.eachCell((cell, colNum) => {
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + emeraldGreen } };
    
    if (colNum === 1) {
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (colNum === 4) {
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
      cell.numFmt = '#,##0" XAF"';
      cell.border = {
        top: { style: 'thin' as const, color: { argb: 'FFFFFFFF' } },
        bottom: { style: 'double' as const, color: { argb: 'FFFFFFFF' } }
      };
    }
  });

  return workbook.xlsx.writeBuffer() as unknown as Promise<Buffer>;
}
