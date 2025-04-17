const ExcelJS = require('exceljs');


async function generateExcel(jsonData) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Device Summary');

  sheet.columns = [
    { width: 40 }, { width: 18 }, { width: 18 },
    { width: 20 }, { width: 20 }, { width: 15 },
    { width: 20 }, { width: 20 }, { width: 15 },
    { width: 20 }, { width: 20 }, { width: 15 },
    { width: 20 }, { width: 20 }, { width: 15 },
    { width: 20 }, { width: 20 }, { width: 15 },
    { width: 20 }, { width: 20 }, { width: 15 }
  ];
  const date = new Date()
  const day = String(date.getDate()-1).padStart(2,'0')
  const month = String(date.getMonth()+1).padStart(2,'0')
  const year = String(date.getFullYear())
  const today = `${day}/${month}/${year}`
  console.log(today);
  // Row 1: Report Title
  sheet.getCell('A1').value = `Report ${today}`;
  sheet.getCell('A1').font = { bold: true, size: 14 };
  sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.getCell('A2').value = 'DEVICE ID';
  sheet.mergeCells('B2:C2');
  sheet.getCell('B2').value = 'Uptime Info';
  sheet.mergeCells('D2:F2');
  sheet.getCell('D2').value = 'CPU';
  sheet.mergeCells('G2:I2');
  sheet.getCell('G2').value = 'GPU';
  sheet.mergeCells('J2:L2');
  sheet.getCell('J2').value = 'Panel';
  sheet.mergeCells('M2:O2');
  sheet.getCell('M2').value = 'Vision';
  sheet.mergeCells('P2:R2');
  sheet.getCell('P2').value = 'Motor Ampere';
  sheet.mergeCells('S2:U2');
  sheet.getCell('S2').value = 'Air Pressure';

  // Style header row
  ['A2', 'B2', 'D2', 'G2', 'J2', 'M2', 'P2', 'S2'].forEach((cell) => {
    sheet.getCell(cell).font = { bold: true };
    sheet.getCell(cell).fill = { pattern: 'solid' };
    sheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Row 3: Sub-Headers
  sheet.addRow([
    '',
    'Device Running', 'Sorting Running',
    'Frequency (max val)', 'Total Duration', 'Threshold',
    'Frequency (max val)', 'Total Duration', 'Threshold',
    'Frequency (max val)', 'Total Duration', 'Threshold',
    'Frequency (max val)', 'Total Duration', 'Threshold',
    'Frequency (max val)', 'Total Duration', 'Threshold',
    'Frequency (min val)', 'Total Duration', 'Threshold'
  ]);

  // Style sub-header row (row 3)
  sheet.getRow(3).eachCell((cell) => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
  });
  // const rawData = await fs.readFile(path.join(__dirname, '../output.txt'), 'utf8');
  // const jsonData = JSON.parse(rawData);

  function setCellWithHighlight(sheet, cellRef, rawValue) {
    const cell = sheet.getCell(cellRef);
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid cell reference format: ${cellRef}`);
  
    let finalValue = rawValue;
    let highlight = false;
  
    // Check for format like "1 (4.5678)"
    if (typeof rawValue === 'string' && rawValue.includes('(')) {
      const numberMatch = rawValue.match(/^(\d+)\s*\(([\d.]+)\)$/);
      if (numberMatch) {
        const prefix = parseInt(numberMatch[1]);
        const num = parseFloat(numberMatch[2]);         
        const rounded = num.toFixed(2);                
  
        if (prefix === 0) {
          finalValue = '-';
        } else {
          finalValue = `${prefix} (${rounded})`;
          if (prefix > 1 || parseFloat(rounded) > 1) {
            highlight = true;
            if (highlight) {
              // console.log('highlight:',cellRef, rawValue);
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } 
              };
            }
          }
          
        }
      } else {
        finalValue = '-';
      }
    }
  
    cell.value = finalValue;
    cell.alignment = { horizontal: 'center' };
  }
  
  function setCellTime(sheet,cellRef, rawValue){
    const cell = sheet.getCell(cellRef);
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid cell reference format: ${cellRef}`);
  
    const isZeroTime = rawValue === '00:00:00';
    const finalValue = isZeroTime ? '-' : rawValue;

    cell.value = finalValue;
    cell.alignment = { horizontal: 'center' };

    if (!isZeroTime) {
      
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' }
      };
    }
    // console.log(`Set ${cellRef} to ${finalValue} with highlight: ${!isZeroTime}`);
  }


  let currentRow = 4;
  jsonData.forEach(entry => {
    sheet.getCell(`A${currentRow}`).value = `${entry.orgName}`;
    sheet.getCell(`A${currentRow}`).alignment = { horizontal: 'left' };
    sheet.getCell(`A${currentRow}`).font = { bold: true };
    currentRow ++;
    const machineData = entry.data
    // console.log(machineData);
    machineData.forEach(deviceEntry=>{
    
    const deviceId = deviceEntry.deviceId;
    // console.log(deviceId);
    const deviceRunning = deviceEntry.data.device_running;
    const sortingRunning = deviceEntry.data.sorting_running;
    // console.log(deviceRunning, sortingRunning);

    sheet.getCell(`A${currentRow}`).value = deviceId;
    sheet.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };

    sheet.getCell(`B${currentRow}`).value = deviceRunning;
    sheet.getCell(`B${currentRow}`).alignment = { horizontal: 'center' };

    sheet.getCell(`C${currentRow}`).value = sortingRunning;
    sheet.getCell(`C${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `D${currentRow}`,deviceEntry.data.cpufreq)
    setCellTime(sheet, `E${currentRow}`, deviceEntry.data.cpu_time)
    sheet.getCell(`F${currentRow}`).value = deviceEntry.data.maxCPUTemperatureThreshold
    sheet.getCell(`F${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `G${currentRow}`, deviceEntry.data.gpufreq)
    setCellTime(sheet, `H${currentRow}`, deviceEntry.data.gpu_time)
    sheet.getCell(`I${currentRow}`).value = deviceEntry.data.maxGPUTemperatureThreshold
    sheet.getCell(`I${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `J${currentRow}`, deviceEntry.data.panelfreq)
    setCellTime(sheet, `K${currentRow}`, deviceEntry.data.panel_time)
    sheet.getCell(`L${currentRow}`).value = deviceEntry.data.upperPanelThreshold
    sheet.getCell(`L${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `M${currentRow}`, deviceEntry.data.visionfreq)
    setCellTime(sheet, `N${currentRow}`, deviceEntry.data.vision_time)
    sheet.getCell(`O${currentRow}`).value = deviceEntry.data.upperVisionPanelThreshold
    sheet.getCell(`O${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `P${currentRow}`, deviceEntry.data.motorfreq);
    setCellTime(sheet, `Q${currentRow}`, deviceEntry.data.ampere_time)
    sheet.getCell(`R${currentRow}`).value = deviceEntry.data.motorAmpereThreshold
    sheet.getCell(`R${currentRow}`).alignment = { horizontal: 'center' }

    setCellWithHighlight(sheet, `S${currentRow}`, deviceEntry.data.pressurefreq);
    setCellTime(sheet, `T${currentRow}`, deviceEntry.data.air_pressure_time)
    sheet.getCell(`U${currentRow}`).value = deviceEntry.data.airPressureLowerValueThreshold
    sheet.getCell(`U${currentRow}`).alignment = { horizontal: 'center' }


    currentRow++;
    })
    currentRow++;
  });

  const borderPairs = ['A', 'C', 'F', 'I', 'L', 'O', 'R'];
  for (let row = 2; row <= currentRow; row++) {
    for (const col of borderPairs) {
      sheet.getCell(`${col}${row}`).border = {
        right: { style: 'thin' }
      };
    }
  }

  sheet.views = [
    {
      state: 'frozen',
      xSplit: 1,
      ySplit: 3
    }
  ];

  // Save the file
  await workbook.xlsx.writeFile('Abnormality_Report.xlsx');
  console.log('Excel generated!');
}

module.exports = { generateExcel };


