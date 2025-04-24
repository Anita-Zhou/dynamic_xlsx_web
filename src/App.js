import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

function App() {
  const [rawData, setRawData] = useState([]);
  const [titleRowIndex, setTitleRowIndex] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [data, setData] = useState([]);
  const [highlightedRows, setHighlightedRows] = useState([]);
  const [selectedColumn, setSelectedColumn] = useState('');
  const [condition, setCondition] = useState('');
  const [showOnlyHighlighted, setShowOnlyHighlighted] = useState(false);

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
      setHighlightedRows([]);
    };
    reader.readAsArrayBuffer(file); // ✅ safer & modern
  };

  const handleHighlight = () => {
    const index = headers.indexOf(selectedColumn);
    const matches = data.map((row, i) => {
      try {
        return eval(`${row[index]} ${condition}`) ? i : null;
      } catch {
        return null;
      }
    }).filter(i => i !== null);
    setHighlightedRows(matches);
  };

  const renderRow = (row, rowIndex) => {
    const isHighlighted = highlightedRows.includes(rowIndex);
    if (showOnlyHighlighted && !isHighlighted) return null;
    return (
      <tr key={rowIndex}>
        {row.map((cell, colIndex) => (
          <td
            key={colIndex}
            style={{
              backgroundColor:
                isHighlighted && headers[colIndex] === selectedColumn
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
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />

      {/* Step 1: Pick a title row from preview */}
      {rawData.length > 0 && titleRowIndex === null && (
        <div>
          <h3>在下方预览中，点击标题行:</h3>
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

      {/* Step 2: Column and condition filtering UI */}
      {headers.length > 0 && (
        <div className="controls">
          <select onChange={(e) => setSelectedColumn(e.target.value)} value={selectedColumn}>
            <option value="">你要处理的列</option>
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
          <button onClick={handleHighlight}>染色</button>
          <button onClick={() => setShowOnlyHighlighted(!showOnlyHighlighted)}>
            {showOnlyHighlighted ? '显示所有行' : '只显示染色行'}
          </button>
        </div>
      )}

      {/* Step 3: Show table */}
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
