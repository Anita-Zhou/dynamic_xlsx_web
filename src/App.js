import React, { useState } from 'react';
import * as XLSX from 'sheetjs-style';
import './App.css';

function exportToExcel(data, headers, filename, highlightedCells, visibleHeaders) {
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

      // Log each cell check for debugging
      console.log(`(${rowIndex}, ${originalColIndex}) -> highlight:`, highlightedCells.has(`${rowIndex}-${originalColIndex}`));

      const baseCell = {
        v: value,
        t: typeof value === 'number' ? 'n' : 's'
      };

      if (isHighlighted) {
        baseCell.s = {
          font: { bold: true },
          fill: {
            patternType: 'solid',
            fgColor: { rgb: 'FFFF00' } // Pure yellow fill
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

function processDefaultMode({
  selectedColumn,
  headers,
  data,
  condition,
  highlightedCells,
  highlightedConditions,
  setHighlightedCells,
  setHighlightedConditions
}) {
  const colIndex = headers.indexOf(selectedColumn);
  if (colIndex === -1) return;

  const newHighlightedCells = new Set(highlightedCells);
  const newConditions = { ...highlightedConditions, [colIndex]: condition };

  for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
    newHighlightedCells.delete(`${rowIndex}-${colIndex}`);
  }

  data.forEach((row, rowIndex) => {
    try {
      if (eval(`${row[colIndex]} ${condition}`)) {
        newHighlightedCells.add(`${rowIndex}-${colIndex}`);
      }
    } catch (e) {
      console.warn("Condition failed on row", rowIndex, e);
    }
  });

  setHighlightedCells(newHighlightedCells);
  setHighlightedConditions(newConditions);
}

function processDataByMode(mode, handlers) {
  switch (mode) {
    case 'default':
      processDefaultMode(handlers);
      break;
    case 'keywords':
    case 'orders':
    case 'top100':
    case 'unlimited':
    case 'merge':
    default:
      break;
  }
}

function App() {
  const [rawData, setRawData] = useState([]);
  const [titleRowIndex, setTitleRowIndex] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [data, setData] = useState([]);
  const [highlightedCells, setHighlightedCells] = useState(new Set());
  const [highlightedConditions, setHighlightedConditions] = useState({});
  const [selectedColumn, setSelectedColumn] = useState('');
  const [condition, setCondition] = useState('');
  const [showOnlyHighlighted, setShowOnlyHighlighted] = useState(false);
  const [highlightThreshold, setHighlightThreshold] = useState('');
  const [mode, setMode] = useState('default');
  const [visibleColumns, setVisibleColumns] = useState([]);
  const [exportFileName, setExportFileName] = useState('exported_table');

  const pastelColors = ['#fce4ec', '#e3f2fd', '#e8f5e9', '#fff3e0', '#ede7f6', '#f3e5f5', '#e0f2f1', '#f1f8e9'];

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const arrayBuffer = evt.target.result;
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      setRawData(raw);
      setHeaders([]);
      setData([]);
      setTitleRowIndex(null);
      setHighlightedCells(new Set());
      setHighlightedConditions({});
      setVisibleColumns([]);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleHighlight = () => {
    processDataByMode(mode, {
      selectedColumn,
      headers,
      data,
      condition,
      highlightedCells,
      highlightedConditions,
      setHighlightedCells,
      setHighlightedConditions
    });
  };

  const handleExportClick = () => {
    const filtered = filterExportData(data, headers, highlightedCells, showOnlyHighlighted, highlightThreshold);
    exportToExcel(filtered, headers, exportFileName, highlightedCells, visibleColumns);
  };

  const toggleColumnVisibility = (column) => {
    setVisibleColumns(prev =>
      prev.includes(column) ? prev.filter(c => c !== column) : [...prev, column]
    );
  };

  const toggleSelectAllColumns = (selectAll) => {
    setVisibleColumns(selectAll ? [...headers] : []);
  };

  const shouldRenderRow = (row, rowIndex) => {
    const highlightCount = headers.reduce((count, _, colIndex) => (
      highlightedCells.has(`${rowIndex}-${colIndex}`) ? count + 1 : count
    ), 0);

    if (showOnlyHighlighted && highlightThreshold === '') {
      return highlightCount > 0;
    } else if (highlightThreshold !== '') {
      return highlightCount === parseInt(highlightThreshold);
    }
    return true;
  };

  const renderRow = (row, rowIndex) => {
    if (!shouldRenderRow(row, rowIndex)) return null;

    return (
      <tr key={rowIndex}>
        {headers.map((header, colIndex) => (
          visibleColumns.includes(header) && (
            <td
              key={colIndex}
              style={{
                backgroundColor: highlightedCells.has(`${rowIndex}-${colIndex}`)
                  ? pastelColors[colIndex % pastelColors.length]
                  : 'transparent'
              }}
            >
              {row[colIndex]}
            </td>
          )
        ))}
      </tr>
    );
  };

  return (
    <div className="container">
      <h1>Excel Highlighter</h1>
      <div className="mode-selector">
        <label htmlFor="mode">Select Processing Mode: </label>
        <select id="mode" value={mode} onChange={(e) => setMode(e.target.value)}>
          <option value="default">默认模式 (Default Processing Mode)</option>
          <option value="keywords">Sorftime 反查关键词—表格处理</option>
          <option value="orders">Sorftime 反查出单词—表格处理</option>
          <option value="top100">Sorftime Top100产品—表格处理</option>
          <option value="unlimited">Sorftime 不限产品—表格处理</option>
          <option value="merge">集合表格 (Combine Tables)</option>
        </select>
      </div>

      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />

      {rawData.length > 0 && titleRowIndex === null && (
        <div>
          <h3>Select the row to use as the column header:</h3>
          <table style={{ borderCollapse: 'collapse' }}>
            <tbody>
              {rawData.slice(0, 10).map((row, i) => (
                <tr
                  key={i}
                  onClick={() => {
                    setTitleRowIndex(i);
                    setHeaders(rawData[i]);
                    setData(rawData.slice(i + 1));
                    setVisibleColumns(rawData[i]);
                  }}
                  style={{ cursor: 'pointer', backgroundColor: '#f5f5f5', border: '1px solid #ddd' }}
                >
                  {row.map((cell, j) => (
                    <td key={j} style={{ border: '1px solid #ccc', padding: '4px 8px' }}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {headers.length > 0 && (
        <>
          <div className="column-filter">
            <strong>显示列：</strong>
            <label style={{ marginRight: '10px' }}>
              <input
                type="checkbox"
                checked={visibleColumns.length === headers.length}
                onChange={(e) => toggleSelectAllColumns(e.target.checked)}
              /> 全选
            </label>
            {headers.map((header, i) => (
              <label key={i} style={{ marginRight: '10px' }}>
                <input
                  type="checkbox"
                  checked={visibleColumns.includes(header)}
                  onChange={() => toggleColumnVisibility(header)}
                /> {header}
              </label>
            ))}
          </div>

          <div className="controls">
            <select onChange={(e) => setSelectedColumn(e.target.value)} value={selectedColumn}>
              <option value="">Select Column</option>
              {visibleColumns.map((header, i) => (
                <option key={i} value={header}>{header}</option>
              ))}
            </select>
            <input
              type="text"
              placeholder="e.g., > 100"
              value={condition}
              onChange={(e) => setCondition(e.target.value)}
            />
            <button onClick={handleHighlight}>Highlight</button>
            <button onClick={() => setShowOnlyHighlighted(!showOnlyHighlighted)}>
              {showOnlyHighlighted ? 'Show All Rows' : 'Show Highlighted Only'}
            </button>
            <input
              type="number"
              min="1"
              placeholder="Highlight count = x"
              value={highlightThreshold}
              onChange={(e) => setHighlightThreshold(e.target.value)}
            />
            <div style={{ marginTop: '10px' }}>
              <input
                type="text"
                value={exportFileName}
                onChange={(e) => setExportFileName(e.target.value)}
                placeholder="Enter filename"
              />
              <button onClick={handleExportClick}>Export to Excel</button>
            </div>
          </div>
        </>
      )}

      {headers.length > 0 && (
        <table>
          <thead>
            <tr>
              {headers.map((h, i) => (
                visibleColumns.includes(h) && <th key={i}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>{data.map((row, i) => renderRow(row, i))}</tbody>
        </table>
      )}
    </div>
  );
}

export default App;
