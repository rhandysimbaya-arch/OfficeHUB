/* Argument guide: describes the current argument and offers tappable choices (TRUE/FALSE, 0/1, match types, units, formats...). */
const ARGH = {};
(function () {
  const A = (fn, arg, d, o) => { fn.split('/').forEach(f => ARGH[f + '.' + arg] = { d, o: o || [] }); };
  const TF = [['FALSE', 'Exact match (same as 0)'], ['TRUE', 'Approximate match (same as 1) – first column must be sorted A→Z']];
  A('VLOOKUP', 'lookup_value', 'The value to find in the FIRST column of the table.');
  A('VLOOKUP', 'table_array', 'The block of cells to search, e.g. A2:D100. The lookup column must be its first column.');
  A('VLOOKUP', 'col_index_num', 'Which column of the table to return: 1 = first column, 2 = second, ...', [['2', 'Second column'], ['3', 'Third column']]);
  A('VLOOKUP', 'range_lookup', 'Choose FALSE (or 0) for an exact match – recommended. TRUE (or 1) finds the closest smaller value.', TF.concat([['0', 'Exact match'], ['1', 'Approximate match']]));
  A('HLOOKUP', 'row_index_num', 'Which row of the table to return: 1 = first row, 2 = second, ...', [['2', 'Second row']]);
  A('HLOOKUP', 'range_lookup', 'Choose FALSE (or 0) for an exact match – recommended. TRUE (or 1) finds the closest smaller value.', TF.concat([['0', 'Exact match'], ['1', 'Approximate match']]));
  A('XLOOKUP', 'if_not_found', 'Text or value to show when nothing matches (instead of #N/A).', [['"Not found"', 'Show the text Not found'], ['""', 'Leave blank'], ['0', 'Show 0']]);
  A('XLOOKUP/XMATCH', 'match_mode', 'How to match the lookup value.', [['0', 'Exact match (default)'], ['-1', 'Exact, or next smaller item'], ['1', 'Exact, or next larger item'], ['2', 'Wildcard match (* ? ~)']]);
  A('XLOOKUP/XMATCH', 'search_mode', 'Which direction / method to search.', [['1', 'First to last (default)'], ['-1', 'Last to first'], ['2', 'Binary search, ascending data'], ['-2', 'Binary search, descending data']]);
  A('MATCH', 'match_type', 'Choose 0 for an exact match. 1 finds the largest value ≤ lookup (data sorted ascending). -1 finds the smallest value ≥ lookup (data sorted descending).', [['0', 'Exact match'], ['1', 'Largest value ≤ lookup (ascending data)'], ['-1', 'Smallest value ≥ lookup (descending data)']]);
  A('LOOKUP', 'result_vector', 'Range to return the result from (same size as lookup_vector).');
  A('INDEX', 'row_num', 'Row position inside the array (1 = first row). Use 0 for the whole column.', [['1', 'First row'], ['0', 'All rows']]);
  A('INDEX', 'column_num', 'Column position inside the array (1 = first column). Use 0 for the whole row.', [['1', 'First column'], ['0', 'All columns']]);
  A('IF', 'logical_test', 'A condition that is TRUE or FALSE, e.g. A1>10 or B2="Yes".', [['=', ''], ['>', ''], ['<', ''], ['<>', '']].slice(0, 0));
  A('IF', 'value_if_false', 'Result when the test is FALSE. Leave out to return FALSE.', [['""', 'Leave blank'], ['0', 'Zero']]);
  A('IFERROR', 'value_if_error', 'What to show when the formula gives an error.', [['""', 'Leave blank'], ['0', 'Zero'], ['"N/A"', 'Text N/A']]);
  A('IFNA', 'value_if_na', 'What to show when the result is #N/A.', [['""', 'Leave blank'], ['"Not found"', 'Text Not found']]);
  A('SUMIF/COUNTIF/AVERAGEIF', 'criteria', 'What to match. Use text in quotes, numbers, or comparisons like ">=50".', [['">0"', 'Greater than 0'], ['">=50"', 'At least 50'], ['"<>"', 'Not blank'], ['""', 'Blank'], ['"*text*"', 'Contains text']]);
  A('SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS', 'criteria1', 'What to match in the first criteria range, e.g. "North" or ">=50".', [['">0"', 'Greater than 0'], ['"<>"', 'Not blank'], ['"*text*"', 'Contains text']]);
  A('SUMIF/AVERAGEIF', 'sum_range', 'Cells to add up when criteria match. Leave out to add the first range itself.');
  A('FILTER', 'include', 'A condition per row, e.g. B2:B50>100 (TRUE rows are kept).');
  A('FILTER', 'if_empty', 'What to show when nothing matches.', [['"No data"', 'Text No data'], ['""', 'Blank']]);
  A('SORT', 'sort_index', 'Which column (or row) to sort by: 1 = first.', [['1', 'First column'], ['2', 'Second column']]);
  A('SORT/SORTBY', 'sort_order', 'Sort direction.', [['1', 'Ascending (A→Z, small→large)'], ['-1', 'Descending (Z→A, large→small)']]);
  A('SORTBY', 'sort_order1', 'Sort direction.', [['1', 'Ascending'], ['-1', 'Descending']]);
  A('SORT/UNIQUE', 'by_col', 'FALSE compares rows (default). TRUE compares columns.', [['FALSE', 'By row (default)'], ['TRUE', 'By column']]);
  A('UNIQUE', 'exactly_once', 'FALSE returns each distinct value. TRUE returns only values that occur once.', [['FALSE', 'All distinct values'], ['TRUE', 'Only values that appear once']]);
  A('RANK', 'order', '0 or omitted ranks largest = 1. Any other number (1) ranks smallest = 1.', [['0', 'Descending (largest = 1)'], ['1', 'Ascending (smallest = 1)']]);
  A('WEEKDAY', 'return_type', 'How days are numbered.', [['1', 'Sunday = 1 ... Saturday = 7'], ['2', 'Monday = 1 ... Sunday = 7'], ['3', 'Monday = 0 ... Sunday = 6']]);
  A('DATEDIF', 'unit', 'What to measure between the two dates (in quotes).', [['"Y"', 'Complete years'], ['"M"', 'Complete months'], ['"D"', 'Days'], ['"YM"', 'Months ignoring years'], ['"MD"', 'Days ignoring months and years'], ['"YD"', 'Days ignoring years']]);
  A('PMT/FV/PV/NPER/RATE', 'type', 'When payments are due.', [['0', 'End of each period (default)'], ['1', 'Beginning of each period']]);
  A('PMT/FV/PV/NPER/RATE', 'fv', 'Future value left after the last payment. Usually 0.', [['0', 'Zero (fully paid)']]);
  A('PMT/FV/NPER/RATE', 'pv', 'Present value – the amount today (loan = positive).');
  A('PMT/NPER/RATE/FV/PV', 'rate', 'Interest rate PER PERIOD, e.g. 12%/12 for monthly payments on a 12% yearly rate.');
  A('PMT/NPER/RATE/FV/PV', 'nper', 'Total number of payment periods, e.g. 36 months.');
  A('IRR', 'guess', 'Your estimate of the rate. Usually left out (10%).', [['0.1', '10%']]);
  A('TEXTJOIN', 'delimiter', 'Text put between items, in quotes.', [['", "', 'Comma and space'], ['" "', 'Space'], ['"-"', 'Hyphen'], ['CHAR(10)', 'New line']]);
  A('TEXTJOIN', 'ignore_empty', 'TRUE skips empty cells. FALSE includes them.', [['TRUE', 'Skip empty cells'], ['FALSE', 'Include empty cells']]);
  A('TEXT', 'format_text', 'The display format in quotes.', [['"0.00"', '2 decimals'], ['"#,##0"', 'Thousands separator'], ['"0%"', 'Percent'], ['"mm/dd/yyyy"', 'Date 11/28/2026'], ['"mmm d, yyyy"', 'Date Nov 28, 2026'], ['"mmmm"', 'Month name'], ['"dddd"', 'Weekday name'], ['"hh:mm AM/PM"', 'Time']]);
  A('ROUND/ROUNDUP/ROUNDDOWN', 'num_digits', 'Digits to keep: 2 = two decimals, 0 = whole number, -1 = nearest 10, -2 = nearest 100.', [['0', 'Whole number'], ['1', '1 decimal'], ['2', '2 decimals'], ['-1', 'Nearest 10'], ['-2', 'Nearest 100']]);
  A('ADDRESS', 'abs_num', 'Reference style.', [['1', 'Absolute $A$1 (default)'], ['2', 'Row absolute A$1'], ['3', 'Column absolute $A1'], ['4', 'Relative A1']]);
  A('ADDRESS', 'a1', 'TRUE = A1 style, FALSE = R1C1 style.', [['TRUE', 'A1 style'], ['FALSE', 'R1C1 style']]);
  A('INDIRECT', 'ref_text', 'Text that holds a cell reference, e.g. "A1" or B1.');
  A('SUBTOTAL', 'function_num', '1-11 include manually hidden rows, 101-111 ignore them.', [['9', 'SUM'], ['109', 'SUM (ignore hidden rows)'], ['1', 'AVERAGE'], ['101', 'AVERAGE (ignore hidden)'], ['2', 'COUNT'], ['3', 'COUNTA'], ['4', 'MAX'], ['5', 'MIN']]);
  A('QUARTILE', 'quart', 'Which quartile.', [['0', 'Minimum'], ['1', '1st quartile'], ['2', 'Median'], ['3', '3rd quartile'], ['4', 'Maximum']]);
  A('PERCENTILE', 'k', 'A value from 0 to 1, e.g. 0.9 = 90th percentile.', [['0.25', '25th'], ['0.5', '50th (median)'], ['0.9', '90th']]);
  A('LARGE/SMALL', 'k', 'Position: 1 = largest (or smallest), 2 = second, ...', [['1', '1st'], ['2', '2nd'], ['3', '3rd']]);
  A('SUBSTITUTE', 'instance_num', 'Which occurrence to replace. Leave out to replace all.', [['1', 'First occurrence only']]);
  A('FIND/SEARCH', 'start_num', 'Character position to start from. Default 1.', [['1', 'From the start']]);
  A('LEFT/RIGHT', 'num_chars', 'How many characters to return. Default 1.', [['1', '1 character'], ['2', '2 characters'], ['3', '3 characters']]);
  A('MID', 'start_num', 'Position of the first character to take (1 = first).', [['1', 'First character']]);
  A('EDATE/EOMONTH', 'months', 'Months to move: positive = future, negative = past, 0 = same month (EOMONTH).', [['0', 'This month'], ['1', 'Next month'], ['-1', 'Previous month']]);
  A('NETWORKDAYS/WORKDAY', 'holidays', 'Optional range of dates to treat as days off.');
  A('OFFSET', 'height', 'Number of rows in the returned range. Leave out to keep the original height.');
  A('OFFSET', 'width', 'Number of columns in the returned range. Leave out to keep the original width.');
  A('SEQUENCE', 'start', 'First number. Default 1.', [['1', 'Start at 1']]);
  A('SEQUENCE', 'step', 'Increase between numbers. Default 1.', [['1', 'Add 1'], ['-1', 'Subtract 1']]);
  A('CONCAT/CONCATENATE', 'text1', 'Text or cell to join.', []);
  A('SWITCH', 'default', 'Result when nothing matches (last argument, no value pair).', [['"Other"', 'Text Other'], ['""', 'Blank']]);
  A('COUNTA', 'value1', 'Cells or values to count (counts anything not empty).');
  A('CHOOSE', 'index_num', 'Which item to return: 1 = first, 2 = second ...', [['1', '1'], ['2', '2'], ['3', '3']]);
  A('MROUND', 'multiple', 'Round to the nearest multiple of this number.', [['5', 'Nearest 5'], ['10', 'Nearest 10'], ['0.05', 'Nearest 0.05']]);
  A('CEILING/FLOOR', 'significance', 'Multiple to round to.', [['1', '1'], ['5', '5'], ['10', '10']]);
  A('TAKE/DROP', 'rows', 'Rows to keep (TAKE) or remove (DROP). Negative counts from the end.', [['1', '1 row'], ['-1', 'Last row']]);
  A('TEXTSPLIT/TEXTBEFORE/TEXTAFTER', 'col_delimiter', 'Text that separates the pieces, in quotes.', [['","', 'Comma'], ['" "', 'Space'], ['"-"', 'Hyphen']]);
  A('TEXTBEFORE/TEXTAFTER', 'delimiter', 'Text to look for, in quotes.', [['","', 'Comma'], ['" "', 'Space'], ['"-"', 'Hyphen']]);
})();
function argBounds(v, p) { let d = 0, st = 0, inS = false; for (let i = 0; i < p; i++) { const c = v[i]; if (c === '"') { inS = !inS; continue; } if (inS) continue; if (c === '(' || c === '{') { d++; if (d === 1 || true) st = i + 1; } else if (c === ')' || c === '}') { d--; } else if ((c === ',' || c === ';') && d >= 1) { st = i + 1; } }
  let en = v.length; d = 0; inS = false; for (let i = p; i < v.length; i++) { const c = v[i]; if (c === '"') { inS = !inS; continue; } if (inS) continue; if (c === '(' || c === '{') d++; else if (c === ')' || c === '}') { if (d === 0) { en = i; break; } d--; } else if ((c === ',' || c === ';') && d === 0) { en = i; break; } }
  return [st, en]; }
function pickArg(val) { const f = $('fx'), v = f.value, p = f.selectionStart == null ? v.length : f.selectionStart; const [s, e] = argBounds(v, p); f.value = v.slice(0, s) + val + v.slice(e); const pos = s + val.length; f.focus(); f.setSelectionRange(pos, pos); guide(); }
const _guide0 = guide;
guide = function () {
  _guide0(); const g = $('fguide'), f = $('fx'), v = f.value; if (g.style.display === 'none' || v[0] !== '=') return;
  const p = f.selectionStart == null ? v.length : f.selectionStart, cx = ctxAt(v, p); if (!cx.top || cx.inStr && false) return;
  const n = cx.top.n, sig = SIG[n]; if (sig === undefined) return; const args = sig.split(', ').filter(Boolean); let idx = cx.top.a; const variadic = args.length && args[args.length - 1] === '...'; if (variadic && idx >= args.length - 1) idx = args.length - 2 >= 0 ? args.length - 2 : 0; if (idx >= args.length) return;
  let raw = args[idx]; const opt = /^\[/.test(raw); const name = raw.replace(/[\[\]]/g, '').replace(/\s*\(.*$/, '').replace(/\d+$/, m => '');
  const H = ARGH[n + '.' + name] || ARGH[n + '.' + name + '1'] || ARGH[n + '.' + raw.replace(/[\[\]]/g, '')];
  let h = '<div class="ah"><span class="an">' + esc2(name) + (opt ? ' <i>(optional)</i>' : ' <i>(required)</i>') + '</span>'; if (H && H.d) h += ' – ' + esc2(H.d); h += '</div>';
  if (H && H.o && H.o.length) h += '<div class="aopts">' + H.o.map(o => `<span class="ch ao" onpointerdown="event.preventDefault();pickArg(${JSON.stringify(o[0]).replace(/"/g, '&quot;')})"><b>${esc2(o[0])}</b>${o[1] ? ' <small>' + esc2(o[1]) + '</small>' : ''}</span>`).join('') + '</div>';
  else if (!H) { if (!opt) return; }
  g.insertAdjacentHTML('beforeend', h); g.style.display = 'block';
};
(function () { const st = document.createElement('style'); st.textContent = '#fguide{max-height:150px!important}#fguide .ah{margin-top:4px;font-size:12px;color:#444;line-height:1.35}#fguide .ah .an{font-weight:600;color:#222}#fguide .ah i{font-weight:400;color:#888}#fguide .aopts{display:flex;flex-wrap:wrap;gap:4px;margin-top:4px}#fguide .ao{display:inline-flex;align-items:baseline;gap:5px;background:#fff;border:1px solid #d6cb82;border-radius:12px;padding:3px 9px;font-size:12px;cursor:pointer}#fguide .ao b{font-family:Consolas,monospace;color:#1a4d8f;background:none;padding:0}#fguide .ao small{color:#666;font-size:11px}'; document.head.appendChild(st); })();
