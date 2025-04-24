import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';


function App() {
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [highlightedRows, setHighlightedRows] = useState([]);
  const [selectedColumn, setSelectedColumn] = useState('');
  const [condition, setCondition] = useState('');
  const [showOnlyHighlighted, setShowOnlyHighlighted] = useState(false);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
  
      const headerRowIndex = parseInt(prompt("Which row number is your title row? (e.g., 1 or 2)")) - 1;
      if (isNaN(headerRowIndex) || headerRowIndex < 0 || headerRowIndex >= raw.length) {
        alert("Invalid row number.");
        return;
      }
  
      setHeaders(raw[headerRowIndex]);
      setData(raw.slice(headerRowIndex + 1));
      setHighlightedRows([]);
    };
    reader.readAsBinaryString(file);
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
      <table>
        <thead>
          <tr>{headers.map((h, i) => <th key={i}>{h}</th>)}</tr>
        </thead>
        <tbody>{data.map((row, i) => renderRow(row, i))}</tbody>
      </table>
    </div>
  );
}

export default App;
