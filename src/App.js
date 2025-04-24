import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

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
  const [mode, setMode] = useState('default');

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
    };
    reader.readAsArrayBuffer(file);
  };

  const handleHighlight = () => {
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
  };

  const renderRow = (row, rowIndex) => {
    const isRowHighlighted = row.some((_, colIndex) => highlightedCells.has(`${rowIndex}-${colIndex}`));
    if (showOnlyHighlighted && !isRowHighlighted) return null;

    return (
      <tr key={rowIndex}>
        {row.map((cell, colIndex) => (
          <td
            key={colIndex}
            style={{
              backgroundColor: highlightedCells.has(`${rowIndex}-${colIndex}`)
                ? 'lightyellow'
                : 'transparent'
            }}
          >
            {cell}
          </td>
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
                  }}
                  style={{ cursor: 'pointer', backgroundColor: '#f5f5f5', border: '1px solid #ddd' }}
                >
                  {row.map((cell, j) => (
                    <td key={j} style={{ border: '1px solid #ccc', padding: '4px 8px' }}>
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {headers.length > 0 && mode === 'default' && (
        <div className="controls">
          <select onChange={(e) => setSelectedColumn(e.target.value)} value={selectedColumn}>
            <option value="">Select Column</option>
            {headers.map((header, i) => (
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
        </div>
      )}

      {headers.length > 0 && (
        <table>
          <thead>
            <tr>{headers.map((h, i) => <th key={i}>{h}</th>)}</tr>
          </thead>
          <tbody>{data.map((row, i) => renderRow(row, i))}</tbody>
        </table>
      )}
    </div>
  );
}

export default App;
