Office.onReady(() => {
  document.getElementById('dataSource').onchange = toggleSourceUI;
  document.getElementById('runBtn').onclick = runAgent;
  document.getElementById('refreshTablesBtn').onclick = loadTables;
  loadTables();
});

function toggleSourceUI() {
  const v = document.getElementById('dataSource').value;
  document.getElementById('tableChooser').style.display = v === 'namedTables' ? 'block' : 'none';
  document.getElementById('manualRangeBox').style.display = v === 'manualRange' ? 'block' : 'none';
}

async function loadTables() {
  try {
    await Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load('items/name');
      await context.sync();
      const sel = document.getElementById('tableList');
      sel.innerHTML = '';
      if (tables.items.length === 0) {
        const opt = document.createElement('option'); opt.text = '(no tables)'; sel.add(opt);
      } else {
        tables.items.forEach(t => {
          const opt = document.createElement('option'); opt.value = t.name; opt.text = t.name; sel.add(opt);
        });
      }
    });
  } catch (e) {
    console.error(e);
  }
}

function setStatus(s) {
  document.getElementById('status').textContent = 'Status: ' + s;
}

async function runAgent() {
  setStatus('Reading data from workbook...');
  const prompt = document.getElementById('prompt').value.trim();
  if (!prompt) { setStatus('Enter prompt'); return; }

  let tableData = null;
  let originalRangeAddress = null;
  try {
    await Excel.run(async (context) => {
      const ds = document.getElementById('dataSource').value;
      if (ds === 'selection') {
        const range = context.workbook.getSelectedRange();
        range.load('values, address');
        await context.sync();
        tableData = { address: range.address, values: range.values };
        originalRangeAddress = range.address;
      } else if (ds === 'namedTables') {
        const tname = document.getElementById('tableList').value;
        const t = context.workbook.tables.getItem(tname);
        const r = t.getDataBodyRange();
        r.load('values, address');
        await context.sync();
        tableData = { address: r.address, values: r.values };
        originalRangeAddress = r.address;
      } else {
        const manual = document.getElementById('pasteRangeInput').value.trim();
        if (!manual) throw new Error('Manual range empty');
        const r = context.workbook.worksheets.getActiveWorksheet().getRange(manual);
        r.load('values, address');
        await context.sync();
        tableData = { address: r.address, values: r.values };
        originalRangeAddress = r.address;
      }
    });
  } catch (err) {
    console.error(err);
    setStatus('Error reading workbook: ' + (err.message || err));
    return;
  }

  setStatus('Sending to serverless LLM endpoint...');
  try {
    const resp = await fetch('https://YOUR_VERCEL_URL/api/ask', {   // <-- REPLACE with your deployed Vercel project URL
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ prompt, tableData })
    });
    const json = await resp.json();
    if (!json.success) throw new Error(json.error || JSON.stringify(json));
    const answer = json.answer;
    setStatus('Received response from LLM; parsing...');
    const parsed = parseLLMOutput(answer); // returns array-of-arrays OR single-string

    // write back into workbook
    await Excel.run(async (context) => {
      let targetSheet = null;
      const writeToNew = document.getElementById('newSheet').checked;
      const pasteRangeText = document.getElementById('pasteRange').value.trim();

      if (writeToNew) {
        targetSheet = context.workbook.worksheets.add('AI_Result_' + Date.now());
      } else {
        if (pasteRangeText) {
          // user gave explicit address; find sheet and range
          // Excel API accepts A1-style without sheet name on getRange - if includes sheet, we need to split
          const hasSheet = pasteRangeText.includes('!');
          if (hasSheet) {
            const parts = pasteRangeText.split('!');
            const sheetName = parts[0];
            const rangeAddr = parts[1];
            targetSheet = context.workbook.worksheets.getItem(sheetName);
            const r = targetSheet.getRange(rangeAddr);
            if (Array.isArray(parsed)) {
              const rows = parsed.length, cols = parsed[0].length;
              const t = r.getResizedRange(rows-1, cols-1);
              t.values = parsed;
            } else {
              r.values = [[parsed]];
            }
            await context.sync();
            setStatus('Done: written to ' + pasteRangeText);
            return;
          } else {
            targetSheet = context.workbook.worksheets.getActiveWorksheet();
          }
        } else {
          // overwrite original selection/table
          // try to get the original address's sheet
          if (originalRangeAddress && originalRangeAddress.includes('!')) {
            const parts = originalRangeAddress.split('!');
            const sheetName = parts[0];
            targetSheet = context.workbook.worksheets.getItem(sheetName);
            const r = targetSheet.getRange(parts[1]);
            if (Array.isArray(parsed)) {
              const rows = parsed.length, cols = parsed[0].length;
              const t = r.getResizedRange(rows-1, cols-1);
              t.values = parsed;
            } else {
              r.values = [[parsed]];
            }
            await context.sync();
            setStatus('Done: overwritten ' + originalRangeAddress);
            return;
          } else {
            targetSheet = context.workbook.worksheets.getActiveWorksheet();
          }
        }
      }

      // If we get here: targetSheet is set and we need to write starting at A1 or provided range within same sheet
      let startRange = 'A1';
      if (!writeToNew && document.getElementById('pasteRange').value.trim()) {
        startRange = document.getElementById('pasteRange').value.trim();
        // If includes sheet, strip it (we are on this targetSheet)
        if (startRange.includes('!')) {
          startRange = startRange.split('!')[1];
        }
      }
      const start = targetSheet.getRange(startRange);
      if (Array.isArray(parsed)) {
        const rows = parsed.length, cols = parsed[0].length;
        const t = start.getResizedRange(rows-1, cols-1);
        t.values = parsed;
      } else {
        start.values = [[parsed]];
      }
      await context.sync();
      setStatus('Done: written output (sheet: ' + targetSheet.name + ', start: ' + startRange + ')');
    });

  } catch (err) {
    console.error(err);
    setStatus('Error from server/LLM: ' + (err.message || JSON.stringify(err)));
  }
}

// Try JSON parse, then CSV parse, fallback to raw text
function parseLLMOutput(text) {
  if (!text || typeof text !== 'string') return [[String(text)]];
  text = text.trim();
  // JSON array-of-arrays
  try {
    const j = JSON.parse(text);
    if (Array.isArray(j) && Array.isArray(j[0])) return j;
    // If JSON array-of-objects, convert to array-of-arrays with headers
    if (Array.isArray(j) && typeof j[0] === 'object') {
      const keys = Object.keys(j[0]);
      const rows = [keys];
      for (const obj of j) rows.push(keys.map(k => obj[k]));
      return rows;
    }
  } catch(e) {}

  // CSV guess: lines separated by \n and comma separated (handle quoted)
  if (text.indexOf('\n') !== -1) {
    const rows = text.split(/\r?\n/).map(line => parseCSVLine(line));
    // if every row is single empty string, fallback
    if (rows.length > 0 && rows.some(r => r.length > 1 || (r.length ===1 && r[0]!=='') )) return rows;
  }

  // else return as single cell
  return text;
}

// Basic CSV line parser handling quoted fields
function parseCSVLine(line) {
  const out = [];
  let cur = '';
  let inQuotes = false;
  for (let i=0;i<line.length;i++) {
    const ch = line[i];
    if (ch === '"' ) {
      if (inQuotes && line[i+1] === '"') { cur += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out.map(s => s.trim());
}
