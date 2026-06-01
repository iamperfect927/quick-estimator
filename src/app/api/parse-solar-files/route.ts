import { NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const fieldStudyFile = formData.get('fieldStudy') as File | null;
    const priceListFile = formData.get('priceList') as File | null;

    if (!fieldStudyFile) {
      return NextResponse.json({ error: 'Missing field study file' }, { status: 400 });
    }

    // 1. READ FIELD STUDY WORKBOOK
    // const fieldStudyBuffer = Buffer.from(await fieldStudyFile.arrayBuffer());
    // const fieldWorkbook = new ExcelJS.Workbook();
    // await fieldWorkbook.xlsx.load(fieldStudyBuffer);

    const fieldStudyArrayBuffer = await fieldStudyFile.arrayBuffer();
    const fieldWorkbook = new ExcelJS.Workbook();
    await fieldWorkbook.xlsx.load(fieldStudyArrayBuffer as any);

    let customerName = 'Valued Solar Client';
    let devices: any[] = [];
    let totalPeakPowerW = 0;
    let totalDayConsumptionWh = 0;
    let totalNightConsumptionWh = 0;
    let profilesParsed = 0;

    for (const sheet of fieldWorkbook.worksheets) {
      // Matches sheets with names like "01_Comsumption_profile", "02_Comsumption_profile", etc.
      // Note the spelling of "Comsumption" to match the user's specification.
      const isMatch = /^\d{2}_Comsumption_profile$/i.test(sheet.name);
      if (!isMatch) continue;

      profilesParsed++;

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

        rowIdx++;
      }
    }

    if (profilesParsed === 0) {
      return NextResponse.json({ error: 'No worksheets matching sequential profile regex found' }, { status: 400 });
    }

    // 2. MATHEMATICAL AGGREGATIONS
    const peakKW = totalPeakPowerW / 1000;
    const dayConsumptionKWh = totalDayConsumptionWh / 1000;
    const nightConsumptionKWh = totalNightConsumptionWh / 1000;

    // 3. DEFINE SEED HARDWARE CATALOG
    let inverters = [
      { brand: 'Felicity', model: 'Hybrid Inverter 30kW', powerKW: 30, priceXAF: 1500000, amperageA: 150 },
  { brand: 'Felicity', model: 'Hybrid Inverter 25kW', powerKW: 25, priceXAF: 1300000, amperageA: 125 },
  { brand: 'Felicity', model: 'Hybrid Inverter 15kW', powerKW: 15, priceXAF: 900000, amperageA: 75 },
  { brand: 'Cworth',   model: 'Hybrid Inverter 10kW', powerKW: 10, priceXAF: 840000, amperageA: 50 },
  { brand: 'Cworth',   model: 'Hybrid Inverter 5kW',  powerKW: 5,  priceXAF: 260000, amperageA: 25 },
  { brand: 'Cworth',   model: 'Hybrid Inverter 3kW',  powerKW: 3,  priceXAF: 140000, amperageA: 15 }
    ];

    let batteries = [
  { brand: 'Felicity',    model: 'Premium Lithium 30kWh', type: 'Lithium', powerKWh: 30, priceXAF: 2200000 },
  { brand: 'Felicity',    model: 'Premium Lithium 25kWh', type: 'Lithium', powerKWh: 25, priceXAF: 1900000 },
  { brand: 'Felicity',    model: 'Premium Lithium 15kWh', type: 'Lithium', powerKWh: 15, priceXAF: 1000000 },
  { brand: 'Cworth',      model: 'Eco Gel 10kWh',         type: 'Gel',     powerKWh: 10, priceXAF: 940000  },
  { brand: 'Cworth',      model: 'Eco Gel 5kWh',          type: 'Gel',     powerKWh: 5,  priceXAF: 360000  },
  { brand: 'Cworth',      model: 'Eco Gel 3kWh',          type: 'Gel',     powerKWh: 3,  priceXAF: 240000  }
];

    let panels = [
  { brand: '', model: 'Solar PV Panel 650W', powerW: 650, efficiencyPercent: 21.5, priceXAF: 70000 },
  { brand: '', model: 'Solar PV Panel 625W', powerW: 625, efficiencyPercent: 21.0, priceXAF: 65000 },
  { brand: '', model: 'Solar PV Panel 600W', powerW: 600, efficiencyPercent: 20.5, priceXAF: 60000 },
  { brand: '', model: 'Solar PV Panel 580W', powerW: 580, efficiencyPercent: 20.2, priceXAF: 58000 },
  { brand: '', model: 'Solar PV Panel 500W', powerW: 500, efficiencyPercent: 19.5, priceXAF: 55000 },
  { brand: '', model: 'Solar PV Panel 400W', powerW: 400, efficiencyPercent: 18.0, priceXAF: 45000 }
];

    let cables = [
      { spec: '4mm² DC Cable (Red/Black)', maxAmperage: 30, priceXAFPerMeter: 1200 },
      { spec: '6mm² DC Cable (Red/Black)', maxAmperage: 55, priceXAFPerMeter: 1800 },
      { spec: '10mm² DC Cable (Red/Black)', maxAmperage: 80, priceXAFPerMeter: 2800 },
      { spec: '16mm² AC/DC Copper Wire', maxAmperage: 110, priceXAFPerMeter: 4500 },
      { spec: '25mm² Heavy-Duty Power Cable', maxAmperage: 150, priceXAFPerMeter: 7500 },
      { spec: '35mm² Heavy-Duty Power Cable', maxAmperage: 200, priceXAFPerMeter: 11000 }
    ];

    // 4. OPTIONALLY PARSE INGESTED PRICE LIST SPREADSHEET IF PROVIDED
    if (priceListFile && priceListFile.size > 0) {
      try {
        const priceListBuffer = Buffer.from(await priceListFile.arrayBuffer());
        const priceWorkbook = new ExcelJS.Workbook();
        await priceWorkbook.xlsx.load(priceListBuffer);

        for (const sheet of priceWorkbook.worksheets) {
          const name = sheet.name.toLowerCase();
          
          // Helper parser for dynamic table extraction
          const parseTable = (requiredCols: string[]) => {
            const parsedRows: any[] = [];
            let headerMapping: { [key: string]: number } = {};
            
            // Scan top 5 rows for headers
            let foundHeaders = false;
            let headerRowIndex = 1;
            for (let r = 1; r <= 5; r++) {
              const row = sheet.getRow(r);
              row.eachCell((cell, colIdx) => {
                const text = String(cell.value || '').toLowerCase().trim();
                requiredCols.forEach(colName => {
                  if (text.includes(colName)) {
                    headerMapping[colName] = colIdx;
                  }
                });
              });
              
              if (Object.keys(headerMapping).length >= requiredCols.length - 1) {
                foundHeaders = true;
                headerRowIndex = r;
                break;
              }
            }

            if (foundHeaders) {
              for (let r = headerRowIndex + 1; r <= 100; r++) {
                const row = sheet.getRow(r);
                const firstColVal = row.getCell(headerMapping[requiredCols[0]] || 1).value;
                if (!firstColVal || String(firstColVal).trim() === '') break;
                
                const entry: any = {};
                requiredCols.forEach(colName => {
                  const colIdx = headerMapping[colName];
                  entry[colName] = colIdx ? row.getCell(colIdx).value : null;
                });
                parsedRows.push(entry);
              }
            }
            return parsedRows;
          };

          if (name.includes('inverter')) {
            const list = parseTable(['brand', 'model', 'power', 'price', 'amperage']);
            if (list.length > 0) {
              inverters = list.map(item => ({
                brand: String(item.brand || 'Generic'),
                model: String(item.model || 'Unknown Model'),
                powerKW: parseFloat(item.power) || 5,
                priceXAF: parseInt(item.price) || 500000,
                amperageA: parseFloat(item.amperage) || 30
              }));
            }
          } else if (name.includes('batter')) {
            const list = parseTable(['brand', 'model', 'type', 'power', 'price']);
            if (list.length > 0) {
              batteries = list.map(item => ({
                brand: String(item.brand || 'Generic'),
                model: String(item.model || 'Unknown Model'),
                type: String(item.type || 'Lithium'),
                powerKWh: parseFloat(item.power) || 2.4,
                priceXAF: parseInt(item.price) || 800000
              }));
            }
          } else if (name.includes('panel') || name.includes('pv') || name.includes('module')) {
            const list = parseTable(['brand', 'model', 'power', 'efficiency', 'price']);
            if (list.length > 0) {
              panels = list.map(item => ({
                brand: String(item.brand || 'Generic'),
                model: String(item.model || 'Unknown Model'),
                powerW: parseFloat(item.power) || 550,
                efficiencyPercent: parseFloat(item.efficiency) || 21.0,
                priceXAF: parseInt(item.price) || 90000
              }));
            }
          } else if (name.includes('cable') || name.includes('wire')) {
            const list = parseTable(['specification', 'amperage', 'price']);
            if (list.length > 0) {
              cables = list.map(item => ({
                spec: String(item.specification || 'DC Wire'),
                maxAmperage: parseFloat(item.amperage) || 50,
                priceXAFPerMeter: parseInt(item.price) || 1500
              }));
            }
          }
        }
      } catch (err) {
        console.warn('Price list parsing warned fallback triggered:', err);
      }
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
      marginPercentage: 15,
      warnings
    });
  } catch (err: any) {
    console.error('Ingestion Engine Parse Error:', err);
    return NextResponse.json({ error: `File Processing Failed: ${err.message}` }, { status: 500 });
  }
}
