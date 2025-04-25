import * as XLSX from 'sheetjs-style';

export function exportToExcel(data, headers, filename, highlightedCells, visibleHeaders) {
  const workbook = XLSX.utils.book_new();
  const ws = {};

  // Write headers
  visibleHeaders.forEach((header, colIndex) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
    ws[cellRef] = {
      v: header,
      t: 's',
      s: {
        font: { bold: true },
        fill: { patternType: 'solid', fgColor: { rgb: 'DDDDDD' } }
      }
    };
  });

  // Debug output
  console.log("Highlighted cells:", [...highlightedCells]);

  // Write data rows
  data.forEach((row, rowIndex) => {
    visibleHeaders.forEach((header, colIndex) => {
      const originalColIndex = headers.indexOf(header);
      const value = row[originalColIndex];
      const cellRef = XLSX.utils.encode_cell({ r: rowIndex + 1, c: colIndex });

      const isHighlighted = highlightedCells.has(`${rowIndex}-${originalColIndex}`);

      console.log(`(${rowIndex}, ${originalColIndex}) -> highlight:`, isHighlighted);

      const baseCell = {
        v: value,
        t: typeof value === 'number' ? 'n' : 's'
      };

      if (isHighlighted) {
        baseCell.s = {
          font: { bold: true },
          fill: {
            patternType: 'solid',
            fgColor: { rgb: 'FFFF00' }
          }
        };
      }

      ws[cellRef] = baseCell;
    });
  });

  const totalRows = data.length + 1;
  const totalCols = visibleHeaders.length;
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: totalRows - 1, c: totalCols - 1 }
  });

  XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
  XLSX.writeFile(workbook, `${filename}.xlsx`);
}

export function filterExportData(data, headers, highlightedCells, showOnlyHighlighted, highlightThreshold) {
  return data.filter((row, rowIndex) => {
    const highlightCount = headers.reduce((count, _, colIndex) => (
      highlightedCells.has(`${rowIndex}-${colIndex}`) ? count + 1 : count
    ), 0);

    if (showOnlyHighlighted && highlightThreshold === '') {
      return highlightCount > 0;
    } else if (highlightThreshold !== '') {
      return highlightCount === parseInt(highlightThreshold);
    }
    return true;
  });
}
