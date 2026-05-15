export function parseHtmlTable(html: string) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const table = doc.querySelector('table');
  if (!table) return null;

  const rowsHtml = Array.from(table.querySelectorAll('tr'));
  
  const grid: string[][] = [];
  const spans: any[][] = [];
  
  // Initialize dynamic grid
  for (let r = 0; r < rowsHtml.length; r++) {
    grid[r] = [];
    spans[r] = [];
  }

  for (let r = 0; r < rowsHtml.length; r++) {
    const tr = rowsHtml[r];
    const cells = Array.from(tr.querySelectorAll('td, th')) as HTMLTableCellElement[];
    
    let cIndex = 0;
    
    for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        
        // Find empty slot (skip occupied by previous rowspans)
        while (cIndex < grid[r].length && grid[r][cIndex] !== undefined) {
             cIndex++;
        }
        
        const colSpan = parseInt(cell.getAttribute('colspan') || '1', 10) || 1;
        const rowSpan = parseInt(cell.getAttribute('rowspan') || '1', 10) || 1;
        const text = cell.innerText.trim();

        // Fill current slot and merged slots
        for (let ri = 0; ri < rowSpan; ri++) {
            if (!grid[r + ri]) {
                grid[r + ri] = [];
                spans[r + ri] = [];
            }
            for (let ci = 0; ci < colSpan; ci++) {
                grid[r + ri][cIndex + ci] = text; // Just store text
                
                if (ri === 0 && ci === 0) {
                    spans[r + ri][cIndex + ci] = { rowspan: rowSpan, colspan: colSpan };
                } else {
                    spans[r + ri][cIndex + ci] = 1;
                }
            }
        }
        
        cIndex += colSpan;
    }
  }

  // Ensure all rows have same length
  let maxCols = 0;
  for (let r = 0; r < grid.length; r++) {
      if (grid[r].length > maxCols) maxCols = grid[r].length;
  }
  for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < maxCols; c++) {
          if (grid[r][c] === undefined) {
              grid[r][c] = '';
              spans[r][c] = 1;
          }
      }
  }

  return { tableData: grid, tableSpans: spans };
}
