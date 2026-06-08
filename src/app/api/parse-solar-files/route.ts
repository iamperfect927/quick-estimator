import { NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

// Helper function to extract cell value, bypassing styled text / objects
const getCellValueRaw = (cell: any): any => {
  if (!cell || cell.value === null || cell.value === undefined) return null;
  if (typeof cell.value === 'string' || typeof cell.value === 'number' || typeof cell.value === 'boolean') {
    return cell.value;
  }
  if (typeof cell.value === 'object') {
    if ('richText' in cell.value && Array.isArray(cell.value.richText)) {
      return cell.value.richText.map((t: any) => t.text || '').join('');
    }
    if ('result' in cell.value) {
      return cell.value.result;
    }
    if ('text' in cell.value) {
      return cell.value.text;
    }
  }
  return String(cell.value);
};

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const fieldStudyFile = formData.get('fieldStudy') as File | null;
    const priceListFile = formData.get('priceList') as File | null;

    if (!fieldStudyFile) {
      return NextResponse.json({ error: 'Missing field study file' }, { status: 400 });
    }

    // 1. READ FIELD STUDY WORKBOOK
    const fieldStudyBuffer = Buffer.from(await fieldStudyFile.arrayBuffer());
    const fieldWorkbook = new ExcelJS.Workbook();
    await fieldWorkbook.xlsx.load(fieldStudyBuffer as any);


    let customerName = 'Valued Solar Client';
    let devices: any[] = [];
    let totalPeakPowerW = 0;
    let totalDayConsumptionWh = 0;
    let totalNightConsumptionWh = 0;
    let profilesParsed = 0;
    const apartmentsData: any[] = [];

    for (const sheet of fieldWorkbook.worksheets) {
      // Matches sheets with names like "01_Comsumption_profile", "02_Comsumption_profile", etc.
      // Note the spelling of "Comsumption" to match the user's specification.
      const isMatch = /^\d{2}_Comsumption_profile$/i.test(sheet.name);
      if (!isMatch) continue;

      profilesParsed++;
      let apartmentPeakPowerW = 0;
      let apartmentDayConsumptionWh = 0;
      let apartmentNightConsumptionWh = 0;
      let apartmentDeviceCount = 0;

      // Extract Customer Name from cell A1: "CUSTOMER'S NAME: <Name>"
      const a1Value = sheet.getCell('A1').value;
      if (a1Value && typeof a1Value === 'string') {
        const custMatch = a1Value.match(/CUSTOMER'S NAME:\s*(.*)/i);
        if (custMatch && custMatch[1]) {
          customerName = custMatch[1].trim();
        }
      }

      // Tabular parsing starts at row 7 (header is row 6)
      let rowIdx = 7;
      while (true) {
        const row = sheet.getRow(rowIdx);
        const deviceNameVal = row.getCell(1).value;
        if (!deviceNameVal || String(deviceNameVal).trim() === '') {
          break; // Empty cell marks end of table
        }

        const deviceName = String(deviceNameVal).trim();
        const powerWatts = parseFloat(row.getCell(2).value as any) || 0;
        const voltageV = parseFloat(row.getCell(3).value as any) || 220;
        const quantity = parseInt(row.getCell(4).value as any) || 0;
        const runtimeDayHours = parseFloat(row.getCell(5).value as any) || 0;
        const runtimeNightHours = parseFloat(row.getCell(6).value as any) || 0;

        // G = B * D * E
        const consumptionDayWh = parseFloat(row.getCell(7).value as any) || (powerWatts * quantity * runtimeDayHours);
        // H = B * D * F
        const consumptionNightWh = parseFloat(row.getCell(8).value as any) || (powerWatts * quantity * runtimeNightHours);

        devices.push({
          deviceName,
          powerWatts,
          voltageV,
          quantity,
          runtimeDayHours,
          runtimeNightHours,
          consumptionDayWh,
          consumptionNightWh,
          apartment: sheet.name
        });

        totalPeakPowerW += (powerWatts * quantity);
        totalDayConsumptionWh += consumptionDayWh;
        totalNightConsumptionWh += consumptionNightWh;

        apartmentPeakPowerW += (powerWatts * quantity);
        apartmentDayConsumptionWh += consumptionDayWh;
        apartmentNightConsumptionWh += consumptionNightWh;
        apartmentDeviceCount++;

        rowIdx++;
      }

      apartmentsData.push({
        name: sheet.name,
        peakKW: apartmentPeakPowerW / 1000,
        dayConsumptionKWh: apartmentDayConsumptionWh / 1000,
        nightConsumptionKWh: apartmentNightConsumptionWh / 1000,
        deviceCount: apartmentDeviceCount
      });
    }

    if (profilesParsed === 0) {
      return NextResponse.json({ error: 'No worksheets matching sequential profile regex found' }, { status: 400 });
    }


    // 2. MATHEMATICAL AGGREGATIONS
    const peakKW = totalPeakPowerW / 1000;
    const dayConsumptionKWh = totalDayConsumptionWh / 1000;
    const nightConsumptionKWh = totalNightConsumptionWh / 1000;

    // 3. PARSE PRICE LIST SPREADSHEET (MANDATORY)
    if (!priceListFile || priceListFile.size === 0) {
      return NextResponse.json({ error: 'Missing price list file. Please upload a valid pricing sheet.' }, { status: 400 });
    }

    let inverters: any[] = [];
    let batteries: any[] = [];
    let panels: any[] = [];
    let cables: any[] = [];

    let sheetNames: string[] = [];
    let debugInfo: any = {};

    try {
      const priceListBuffer = Buffer.from(await priceListFile.arrayBuffer());
      const priceWorkbook = new ExcelJS.Workbook();
      await priceWorkbook.xlsx.load(priceListBuffer as any);
      sheetNames = priceWorkbook.worksheets.map((s: any) => s.name);

      for (const sheet of priceWorkbook.worksheets) {
        const name = sheet.name.toLowerCase();
        
        let foundHeaders = false;
        let headerRowIndex = 1;
        let headerMapping: { [key: string]: number } = {};

        // Helper parser for dynamic table extraction using synonym/alias matching
        const parseTable = (fieldMappings: { [key: string]: string[] }) => {
          const parsedRows: any[] = [];
          headerMapping = {};
          
          // Scan top 15 rows for headers
          foundHeaders = false;
          headerRowIndex = 1;
          for (let r = 1; r <= 15; r++) {
            const row = sheet.getRow(r);
            row.eachCell({ includeEmpty: false }, (cell, colIdx) => {
              const text = String(getCellValueRaw(cell) || '').toLowerCase().trim();
              for (const [field, aliases] of Object.entries(fieldMappings)) {
                if (aliases.some(alias => text.includes(alias))) {
                  headerMapping[field] = colIdx;
                }
              }
            });
            
            // Require at least 2 distinct mapped columns (model/spec + price)
            const colIndices = new Set(Object.values(headerMapping));
            const hasModelOrSpec = !!headerMapping['model'] || !!headerMapping['spec'];
            const hasPrice = !!headerMapping['priceXAF'] || !!headerMapping['priceXAFPerMeter'];
            if (hasModelOrSpec && hasPrice && colIndices.size >= 2) {
              foundHeaders = true;
              headerRowIndex = r;
              break;
            }
          }

          if (foundHeaders) {
            const lastRow = sheet.rowCount;
            let consecutiveEmpty = 0;

            for (let r = headerRowIndex + 1; r <= lastRow; r++) {
              const row = sheet.getRow(r);

              // Use S/N (column 1) as the primary end-of-data sentinel
              const snVal = getCellValueRaw(row.getCell(1));
              const keyColIdx = headerMapping['model'] || headerMapping['spec'] || 2;
              const keyColVal = getCellValueRaw(row.getCell(keyColIdx));

              // If both S/N and key column are empty, increment empty counter and skip
              if ((!snVal || String(snVal).trim() === '') && (!keyColVal || String(keyColVal).trim() === '')) {
                consecutiveEmpty++;
                if (consecutiveEmpty >= 2) break; // Two consecutive blank rows = end of table
                continue;
              }
              consecutiveEmpty = 0;

              const entry: any = {};
              let hasAnyMappedValue = false;
              for (const field of Object.keys(fieldMappings)) {
                const colIdx = headerMapping[field];
                const val = colIdx ? getCellValueRaw(row.getCell(colIdx)) : null;
                entry[field] = val;
                if (val !== null && val !== undefined && String(val).trim() !== '') hasAnyMappedValue = true;
              }
              if (hasAnyMappedValue) parsedRows.push(entry);
            }
          }
          return parsedRows;
        };


        if (name.includes('solar inverters') || name.includes('inverter')) {
          const list = parseTable({
            brand: ['brand', 'manuf', 'make'],
            model: ['model', 'sku', 'desc', 'inverter', 'type'],
            powerKW: ['power', 'capacity', 'kw', 'rating'],
            priceXAF: ['price', 'cost', 'xaf', 'rate'],
            amperageA: ['amperage', 'amp', 'current', 'maxa']
          });
          debugInfo[sheet.name] = { foundHeaders, headerRowIndex, headerMapping, rowsParsed: list.length };
          if (list.length > 0) {
            inverters = list.map(item => {
              const powerKW = parseFloat(item.powerKW) || 5;
              const derivedAmperage = powerKW * 5;
              return {
                brand: String(item.brand || 'Generic'),
                model: String(item.model || 'Unknown Model'),
                powerKW,
                priceXAF: parseInt(item.priceXAF) || 500000,
                amperageA: parseFloat(item.amperageA) || derivedAmperage || 30
              };
            });
          }
        } else if (name.includes('solar batteries') || name.includes('batter')) {
          const list = parseTable({
            brand: ['brand', 'manuf', 'make'],
            model: ['model', 'sku', 'desc', 'battery', 'type'],
            type: ['type', 'chemistry'],
            powerKWh: ['power', 'capacity', 'kwh', 'rating'],
            priceXAF: ['price', 'cost', 'xaf', 'rate']
          });
          debugInfo[sheet.name] = { foundHeaders, headerRowIndex, headerMapping, rowsParsed: list.length };
          if (list.length > 0) {
            batteries = list.map(item => {
              const model = String(item.model || 'Unknown Model');
              const rawType = String(item.type || '').toLowerCase();
              const isLithium = rawType.includes('lith') || rawType.includes('lfp') || model.toLowerCase().includes('lith') || model.toLowerCase().includes('lfp');
              return {
                brand: String(item.brand || 'Generic'),
                model,
                type: isLithium ? 'Lithium' : 'Gel',
                powerKWh: parseFloat(item.powerKWh) || 2.4,
                priceXAF: parseInt(item.priceXAF) || 800000
              };
            });
          }
        } else if (name.includes('solar panels') || name.includes('panel') || name.includes('pv') || name.includes('module')) {
          const list = parseTable({
            brand: ['brand', 'manuf', 'make'],
            model: ['model', 'sku', 'desc', 'panel', 'pv', 'module', 'type'],
            powerW: ['power', 'capacity', 'w', 'watt', 'rating'],
            efficiencyPercent: ['efficiency', 'eff', '%'],
            priceXAF: ['price', 'cost', 'xaf', 'rate']
          });
          debugInfo[sheet.name] = { foundHeaders, headerRowIndex, headerMapping, rowsParsed: list.length };
          if (list.length > 0) {
            panels = list.map(item => {
              const powerW = parseFloat(item.powerW) || 550;
              const derivedEff = 15 + (powerW / 100);
              return {
                brand: String(item.brand || 'Generic'),
                model: String(item.model || 'Unknown Model'),
                powerW,
                efficiencyPercent: parseFloat(item.efficiencyPercent) || derivedEff || 21.0,
                priceXAF: parseInt(item.priceXAF) || 90000
              };
            });
          }
        } else if (name.includes('cables') || name.includes('cable') || name.includes('wire')) {
          const list = parseTable({
            spec: ['specification', 'spec', 'wire', 'cable', 'desc', 'type'],
            maxAmperage: ['amperage', 'amp', 'current', 'maxa'],
            priceXAFPerMeter: ['price', 'cost', 'xaf', 'meter', 'rate']
          });
          debugInfo[sheet.name] = { foundHeaders, headerRowIndex, headerMapping, rowsParsed: list.length };
          if (list.length > 0) {
            cables = list.map(item => {
              const spec = String(item.spec || 'DC Wire');
              let maxAmperage = 50;
              const mmMatch = spec.match(/(\d+)\s*mm²/i);
              if (mmMatch) {
                const mmVal = parseInt(mmMatch[1]);
                if (mmVal <= 4) maxAmperage = 30;
                else if (mmVal <= 6) maxAmperage = 55;
                else if (mmVal <= 10) maxAmperage = 80;
                else if (mmVal <= 16) maxAmperage = 110;
                else if (mmVal <= 25) maxAmperage = 150;
                else if (mmVal <= 35) maxAmperage = 200;
              }
              return {
                spec,
                maxAmperage: parseFloat(item.maxAmperage) || maxAmperage,
                priceXAFPerMeter: parseInt(item.priceXAFPerMeter) || 1500
              };
            });
          }
        }
      }
    } catch (err: any) {
      return NextResponse.json({ error: `Failed parsing uploaded price list: ${err.message}` }, { status: 400 });
    }


    // If a category was empty in the price list, insert a "—" placeholder item so the flow continues gracefully
    if (inverters.length === 0) {
      inverters = [{ brand: '—', model: '—', powerKW: 0, priceXAF: 0, amperageA: 0 }];
    }
    if (batteries.length === 0) {
      batteries = [{ brand: '—', model: '—', type: 'Unknown', powerKWh: 0, priceXAF: 0 }];
    }
    if (panels.length === 0) {
      panels = [{ brand: '—', model: '—', powerW: 0, efficiencyPercent: 0, priceXAF: 0 }];
    }
    if (cables.length === 0) {
      cables = [{ spec: '—', maxAmperage: 0, priceXAFPerMeter: 0 }];
    }


    const boq: any[] = [];
    const warnings: string[] = [];

    // 5. SYSTEM MATCHING: SOLAR INVERTERS
    let matchedInverter = inverters
      .filter(i => i.powerKW >= peakKW)
      .sort((a, b) => a.priceXAF - b.priceXAF)[0];

    let inverterQty = 1;
    if (!matchedInverter) {
      // Exception Capacity Overflow
      const maxInverter = inverters.reduce((max, i) => i.powerKW > max.powerKW ? i : max, inverters[0]);
      inverterQty = Math.ceil(peakKW / maxInverter.powerKW);
      matchedInverter = maxInverter;
      warnings.push(`CAPACITY_OVERFLOW_WARNING: Peak load of ${peakKW.toFixed(2)} kW exceeds the maximum inverter rating of ${maxInverter.powerKW} kW. Switched to ${inverterQty} x ${maxInverter.brand} ${maxInverter.model} inverters.`);
    }

    boq.push({
      category: 'Inverter',
      brand: matchedInverter.brand,
      model: matchedInverter.model,
      rating: `${matchedInverter.powerKW} kW`,
      unitPriceXAF: matchedInverter.priceXAF,
      quantity: inverterQty,
      totalPriceXAF: matchedInverter.priceXAF * inverterQty
    });

    // 6. SYSTEM MATCHING: SOLAR BATTERIES
    const targetBatteryCapacityKWh = nightConsumptionKWh * 1.2; // 20% safety margin buffer
    
    // Sort batteries preferring Lithium, then Gel, then sort by price
    const eligibleBatteries = batteries.filter(b => b.powerKWh >= targetBatteryCapacityKWh);
    
    let matchedBattery = eligibleBatteries
      .sort((a, b) => {
        // Lithium priority
        const aIsLith = a.type.toLowerCase().includes('lith');
        const bIsLith = b.type.toLowerCase().includes('lith');
        if (aIsLith && !bIsLith) return -1;
        if (!aIsLith && bIsLith) return 1;
        return a.priceXAF - b.priceXAF;
      })[0];

    let batteryQty = 1;
    if (!matchedBattery) {
      // Capacity overflow
      const maxBattery = batteries
        .sort((a, b) => {
          const aIsLith = a.type.toLowerCase().includes('lith');
          const bIsLith = b.type.toLowerCase().includes('lith');
          if (aIsLith && !bIsLith) return -1;
          if (!aIsLith && bIsLith) return 1;
          return b.powerKWh - a.powerKWh;
        })[0];
      batteryQty = Math.ceil(targetBatteryCapacityKWh / maxBattery.powerKWh);
      matchedBattery = maxBattery;
      warnings.push(`CAPACITY_OVERFLOW_WARNING: Night load demand of ${targetBatteryCapacityKWh.toFixed(2)} kWh (incl. 20% safety buffer) exceeds largest battery unit (${maxBattery.powerKWh} kWh). Configured ${batteryQty} x ${maxBattery.brand} ${maxBattery.model} in parallel.`);
    }

    boq.push({
      category: 'Battery',
      brand: matchedBattery.brand,
      model: `${matchedBattery.model} (${matchedBattery.type})`,
      rating: `${matchedBattery.powerKWh} kWh`,
      unitPriceXAF: matchedBattery.priceXAF,
      quantity: batteryQty,
      totalPriceXAF: matchedBattery.priceXAF * batteryQty
    });

    // 7. SYSTEM MATCHING: SOLAR PANELS (Efficiency footprint optimization)
    const bestPanel = panels.sort((a, b) => b.efficiencyPercent - a.efficiencyPercent)[0];
    
    // Standard daily sun hours production rule: Wattage * 5 peak sun hours = Daily Wh
    const dailyProductionKWh = (bestPanel.powerW * 5.0) / 1000;
    let panelQty = Math.ceil(dayConsumptionKWh / dailyProductionKWh);

    // Limit physical constraint warning
    if (panelQty > 80) {
      warnings.push(`PHYSICAL_ROOF_LIMIT_WARNING: The required panel array count (${panelQty} panels) is exceptionally high. Please verify if the installation site has sufficient square footage.`);
    }

    boq.push({
      category: 'Panel',
      brand: bestPanel.brand,
      model: bestPanel.model,
      rating: `${bestPanel.powerW} W`,
      unitPriceXAF: bestPanel.priceXAF,
      quantity: panelQty,
      totalPriceXAF: bestPanel.priceXAF * panelQty
    });

    // 8. SYSTEM MATCHING: CABLES
    const totalAmperage = matchedInverter.amperageA * inverterQty;
    let matchedCable = cables.filter(c => c.maxAmperage >= totalAmperage).sort((a, b) => a.priceXAFPerMeter - b.priceXAFPerMeter)[0];
    
    if (!matchedCable) {
      matchedCable = cables.sort((a, b) => b.maxAmperage - a.maxAmperage)[0];
    }

    // cabling requirements scale with number of panels and inverters
    const cableQty = 50 + (panelQty * 2);

    boq.push({
      category: 'Cable',
      brand: 'Standard DC Gauge',
      model: matchedCable.spec,
      rating: `Max ${matchedCable.maxAmperage}A`,
      unitPriceXAF: matchedCable.priceXAFPerMeter,
      quantity: cableQty,
      totalPriceXAF: matchedCable.priceXAFPerMeter * cableQty
    });

    const materials = boq.map(item => ({
      name: `${item.category} - ${item.brand} ${item.model} (${item.rating})`,
      unitPrice: item.unitPriceXAF,
      quantity: item.quantity
    }));

    // Labor = 30% of total equipment cost
    const materialsSubtotal = materials.reduce((sum, m) => sum + m.unitPrice * m.quantity, 0);
    const laborCost = Math.round(materialsSubtotal * 0.30);

    const labor = [
      {
        description: 'Installation, Wiring & Commissioning Labor (30% of Equipment Cost)',
        hours: 1,
        hourlyRate: laborCost
      }
    ];

    return NextResponse.json({
      success: true,
      customerName,
      metrics: {
        peakKW,
        dayConsumptionKWh,
        nightConsumptionKWh,
        apartmentCount: profilesParsed,
        totalDeviceCount: devices.length
      },
      materials,
      labor,
      marginPercentage: 0,
      apartments: apartmentsData,
      warnings
    });
  } catch (err: any) {
    console.error('Ingestion Engine Parse Error:', err);
    return NextResponse.json({ error: `File Processing Failed: ${err.message}` }, { status: 500 });
  }
}
