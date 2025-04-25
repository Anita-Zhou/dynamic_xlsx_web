export function processDefaultMode({
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



// Add placeholders for other modes:
export function processKeywordsMode(params) {
    console.log('processKeywordsMode not implemented yet', params);
  }
  
  export function processOrdersMode(params) {
    console.log('processOrdersMode not implemented yet', params);
  }
  
  export function processTop100Mode(params) {
    console.log('processTop100Mode not implemented yet', params);
  }
  
  export function processUnlimitedMode(params) {
    console.log('processUnlimitedMode not implemented yet', params);
  }
  
  export function processMergeMode(params) {
    console.log('processMergeMode not implemented yet', params);
  }