import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

function exportToExcel(data, headers, filename, highlightedCells, visibleHeaders) {
  const sheetData = [visibleHeaders];
  const workbook = XLSX.utils.book_new();
  const worksheet = {};

  data.forEach((row, rowIndex) => {
    const sheetRow = [];
    visibleHeaders.forEach((header, colIndex) => {
      const originalColIndex = headers.indexOf(header);
      const cellValue = row[originalColIndex];
      const cellKey = XLSX.utils.encode_cell({ r: sheetData.length, c: colIndex });
      worksheet[cellKey] = {
        v: cellValue,
        s: highlightedCells.has(`${rowIndex}-${originalColIndex}`)
          ? { fill: { fgColor: { rgb: 'FFFF00' } } }
          : {}
      };
      sheetRow.push(cellValue);
    });
    sheetData.push(sheetRow);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  worksheet['!ref'] = ws['!ref'];
  Object.assign(ws, worksheet);

  XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
  XLSX.writeFile(workbook, `${filename}.xlsx`);
}

function filterExportData(data, headers, highlightedCells, showOnlyHighlighted, highlightThreshold) {
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

function handleExport({ data, headers, exportFileName, highlightedCells, visibleColumns, showOnlyHighlighted, highlightThreshold }) {
  const filtered = filterExportData(data, headers, highlightedCells, showOnlyHighlighted, highlightThreshold);
  exportToExcel(filtered, headers, exportFileName, highlightedCells, visibleColumns);
}

export default handleExport;
export { exportToExcel, filterExportData };