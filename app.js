let rawRows = [];
let charts = {};

const INR = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  maximumFractionDigits: 2
});

const compactINR = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  notation: 'compact',
  maximumFractionDigits: 1
});

const el = id => document.getElementById(id);

/* =========================================================
   FORMATTING
   ========================================================= */

function money(v) {
  return INR.format(Number(v) || 0);
}

function compactMoney(v) {
  return compactINR.format(Number(v) || 0);
}

/* =========================================================
   AMOUNT PARSER
   Handles:
   360
   360.00
   ₹360.00
   ₹1,25,000.00
   1,25,000
   ========================================================= */

function parseAmount(v) {
  if (typeof v === 'number' && Number.isFinite(v)) {
    return v;
  }

  const s = String(v ?? '')
    .trim()
    .replace(/₹/g, '')
    .replace(/,/g, '')
    .replace(/\s/g, '');

  if (!s) return 0;

  const n = Number(s);

  return Number.isFinite(n) ? n : 0;
}

/* =========================================================
   DATE PARSER
   PRIMARY FORMAT:
   DD/MM/YYYY

   Examples:
   01/09/2026 → 1 September 2026
   03/09/2026 → 3 September 2026
   09/09/2026 → 9 September 2026
   30/09/2026 → 30 September 2026
   ========================================================= */

function parseDate(v) {

  // -----------------------------------------
  // Excel / JavaScript Date object
  // -----------------------------------------
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(
      v.getFullYear(),
      v.getMonth(),
      v.getDate()
    );
  }

  // -----------------------------------------
  // Excel serial date
  // -----------------------------------------
  if (
    typeof v === 'number' &&
    window.XLSX &&
    XLSX.SSF
  ) {
    try {
      const o = XLSX.SSF.parse_date_code(v);

      if (o) {
        return new Date(
          o.y,
          o.m - 1,
          o.d
        );
      }
    } catch (err) {
      console.warn('Excel date parsing error:', err);
    }
  }

  const s = String(v ?? '').trim();

  if (!s) return null;

  // -----------------------------------------
  // IMPORTANT:
  // Read DD/MM/YYYY FIRST.
  // Do NOT allow new Date(s) to guess.
  // -----------------------------------------

  let m = s.match(
    /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/
  );

  if (m) {

    const day = Number(m[1]);
    const month = Number(m[2]);
    const year = Number(m[3]);

    const d = new Date(
      year,
      month - 1,
      day
    );

    // Validate
    if (
      d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day
    ) {
      return d;
    }

    return null;
  }

  // -----------------------------------------
  // ISO fallback: YYYY-MM-DD
  // -----------------------------------------

  m = s.match(
    /^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/
  );

  if (m) {

    const year = Number(m[1]);
    const month = Number(m[2]);
    const day = Number(m[3]);

    const d = new Date(
      year,
      month - 1,
      day
    );

    if (
      d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day
    ) {
      return d;
    }

    return null;
  }

  return null;
}

/* =========================================================
   DATE HELPERS
   IMPORTANT:
   Do not use toISOString() here.
   It converts local time to UTC and can shift dates.
   ========================================================= */

function dateKey(d) {

  if (!(d instanceof Date) || isNaN(d.getTime())) {
    return '';
  }

  return [
    d.getFullYear(),
    String(d.getMonth() + 1).padStart(2, '0'),
    String(d.getDate()).padStart(2, '0')
  ].join('-');
}

function fmtDate(d) {

  if (!(d instanceof Date) || isNaN(d.getTime())) {
    return '—';
  }

  return d.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });
}

function monthKey(d) {

  return `${d.getFullYear()}-${String(
    d.getMonth() + 1
  ).padStart(2, '0')}`;
}

function monthLabel(key) {

  const [y, m] = key.split('-').map(Number);

  return new Date(
    y,
    m - 1,
    1
  ).toLocaleDateString('en-IN', {
    month: 'long',
    year: 'numeric'
  });
}

/* =========================================================
   ESCAPE HTML
   ========================================================= */

function escapeHtml(s) {

  return String(s ?? '').replace(
    /[&<>"']/g,
    c => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    }[c])
  );
}

/* =========================================================
   HEADER FINDER
   Makes the Excel reader more tolerant of spaces/case.
   ========================================================= */

function cleanHeader(value) {

  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function findColumn(row, possibleNames) {

  const keys = Object.keys(row);

  for (const name of possibleNames) {

    const wanted = cleanHeader(name);

    const found = keys.find(
      key => cleanHeader(key) === wanted
    );

    if (found !== undefined) {
      return found;
    }
  }

  return null;
}

/* =========================================================
   NORMALIZE EXCEL ROWS
   Expected Excel columns:

   Date
   Expense Name
   Expense Type
   Expense Sub Class
   Amount
   ========================================================= */

function normalize(rows) {

  return rows.map(r => {

    const dateCol = findColumn(r, [
      'Date',
      'DATE',
      'date'
    ]);

    const nameCol = findColumn(r, [
      'Expense Name',
      'Expense Description',
      'Description',
      'Expense Description '
    ]);

    const typeCol = findColumn(r, [
      'Expense Type',
      'Category',
      'Type'
    ]);

    const subClassCol = findColumn(r, [
      'Expense Sub Class',
      'Expense Subclass',
      'Sub Class',
      'Subclass',
      'Subcategory',
      'Expense Sub Class '
    ]);

    const amountCol = findColumn(r, [
      'Amount',
      'Amount (₹)',
      'Amount (INR)',
      'amount',
      'AMOUNT'
    ]);

    const d = parseDate(
      dateCol ? r[dateCol] : ''
    );

    const amount = parseAmount(
      amountCol ? r[amountCol] : ''
    );

    const description = String(
      nameCol ? r[nameCol] : ''
    ).trim();

    const type = String(
      typeCol ? r[typeCol] : 'Uncategorised'
    ).trim() || 'Uncategorised';

    const subClass = String(
      subClassCol ? r[subClassCol] : ''
    ).trim();

    return {
      date: d,
      description,
      type,
      subClass,
      amount,

      payment: String(
        r[
          findColumn(r, [
            'Payment Method',
            'Payment',
            'Mode'
          ])
        ] ?? 'Not specified'
      ).trim() || 'Not specified',

      notes: String(
        r[
          findColumn(r, [
            'Notes',
            'Note'
          ])
        ] ?? ''
      ).trim()
    };

  }).filter(
    x =>
      x.date &&
      x.description &&
      x.amount > 0
  );
}

/* =========================================================
   LOAD WORKBOOK
   ========================================================= */

function loadWorkbook(
  data,
  name = 'Excel file'
) {

  try {

    const wb = XLSX.read(data, {
      type: 'array',
      cellDates: true
    });

    // Prefer a sheet containing "expense"
    // otherwise use first sheet
    let chosen =
      wb.SheetNames.find(
        s => /expense/i.test(s)
      ) ||
      wb.SheetNames[0];

    if (!chosen) {
      throw new Error(
        'No worksheet found.'
      );
    }

    const ws = wb.Sheets[chosen];

    const rows =
      XLSX.utils.sheet_to_json(
        ws,
        {
          defval: '',
          raw: true
        }
      );

    rawRows = normalize(rows);

    rawRows.sort(
      (a, b) => a.date - b.date
    );

    el('fileStatus').textContent =
      `${name} · ${rawRows.length} expenses`;

    el('subhead').textContent =
      `${chosen} · ${rawRows.length} valid expense rows loaded locally in your browser.`;

    populateFilters();
    render();

  } catch (err) {

    console.error(
      'Workbook loading error:',
      err
    );

    el('fileStatus').textContent =
      'Could not read Excel file';

    el('subhead').textContent =
      'Please check the Excel format and try again.';

    alert(
      'I could not read this Excel file. Please check the column names and date format.'
    );
  }
}

/* =========================================================
   FILTERS
   ========================================================= */

function populateFilters() {

  const cats = [
    ...new Set(
      rawRows.map(r => r.type)
    )
  ].sort();

  const pays = [
    ...new Set(
      rawRows.map(r => r.payment)
    )
  ].sort();

  const months = [
    ...new Set(
      rawRows.map(r => monthKey(r.date))
    )
  ]
    .sort()
    .reverse();

  el('categorySelect').innerHTML =
    '<option value="ALL">All categories</option>' +
    cats
      .map(
        x =>
          `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`
      )
      .join('');

  el('paymentSelect').innerHTML =
    '<option value="ALL">All payment methods</option>' +
    pays
      .map(
        x =>
          `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`
      )
      .join('');

  el('monthSelect').innerHTML =
    months.length
      ? months
          .map(
            x =>
              `<option value="${x}">${monthLabel(x)}</option>`
          )
          .join('')
      : '<option value="ALL">No data</option>';
}

/* =========================================================
   FILTERED DATA
   ========================================================= */

function filtered() {

  const month =
    el('monthSelect').value;

  const cat =
    el('categorySelect').value;

  const pay =
    el('paymentSelect').value;

  const q =
    el('searchInput').value
      .toLowerCase()
      .trim();

  return rawRows.filter(r => {

    const searchable =
      `${r.description} ${r.type} ${r.subClass} ${r.payment} ${r.notes}`
        .toLowerCase();

    return (
      (month === 'ALL' ||
        monthKey(r.date) === month) &&

      (cat === 'ALL' ||
        r.type === cat) &&

      (pay === 'ALL' ||
        r.payment === pay) &&

      (!q ||
        searchable.includes(q))
    );
  });
}

/* =========================================================
   MAIN RENDER
   ========================================================= */

function render() {

  const rows = filtered();

  const month =
    el('monthSelect').value;

  const total =
    rows.reduce(
      (s, r) => s + r.amount,
      0
    );

  const max =
    rows.reduce(
      (a, b) =>
        b.amount > a.amount
          ? b
          : a,
      { amount: 0 }
    );

  const activeDays =
    new Set(
      rows.map(
        r => dateKey(r.date)
      )
    ).size;

  el('totalSpend').textContent =
    money(total);

  el('txnCount').textContent =
    rows.length.toLocaleString('en-IN');

  el('dailyAvg').textContent =
    money(
      activeDays
        ? total / activeDays
        : 0
    );

  el('largestExpense').textContent =
    money(max.amount);

  el('largestLabel').textContent =
    max.description || '—';

  el('periodLabel').textContent =
    month === 'ALL'
      ? 'All loaded data'
      : monthLabel(month);

  renderCharts(rows);
  renderTables(rows);
}

/* =========================================================
   CHARTS
   ========================================================= */

function renderCharts(rows) {

  Object.values(charts)
    .forEach(c => c?.destroy());

  const byDay = new Map();
  const byCat = new Map();

  rows.forEach(r => {

    const dk = dateKey(r.date);

    byDay.set(
      dk,
      (byDay.get(dk) || 0) +
        r.amount
    );

    byCat.set(
      r.type,
      (byCat.get(r.type) || 0) +
        r.amount
    );
  });

  /* -----------------------------------------
     DAILY SPENDING
     ----------------------------------------- */

  const days =
    [...byDay.keys()].sort();

  const dailyCanvas =
    el('dailyChart');

  if (dailyCanvas) {

    charts.daily =
      new Chart(
        dailyCanvas,
        {
          type: 'line',

          data: {
            labels: days.map(d =>
              new Date(
                `${d}T00:00:00`
              ).toLocaleDateString(
                'en-IN',
                {
                  day: '2-digit',
                  month: 'short'
                }
              )
            ),

            datasets: [
              {
                data:
                  days.map(
                    d => byDay.get(d)
                  ),

                borderWidth: 2,
                fill: true,
                tension: 0.28,
                pointRadius: 2
              }
            ]
          },

          options: baseChart({
            scales: {
              y: {
                ticks: {
                  callback: v =>
                    compactMoney(v)
                }
              },

              x: {
                grid: {
                  display: false
                }
              }
            }
          })
        }
      );
  }

  /* -----------------------------------------
     CATEGORY DOUGHNUT
     ----------------------------------------- */

  const cats =
    [...byCat.entries()]
      .sort(
        (a, b) => b[1] - a[1]
      );

  const categoryCanvas =
    el('categoryChart');

  if (categoryCanvas) {

    charts.category =
      new Chart(
        categoryCanvas,
        {
          type: 'doughnut',

          data: {
            labels:
              cats.map(x => x[0]),

            datasets: [
              {
                data:
                  cats.map(x => x[1]),

                borderWidth: 2
              }
            ]
          },

          options: {
            responsive: true,
            maintainAspectRatio: false,

            plugins: {

              legend: {
                position: 'right',

                labels: {
                  color: '#c9d1df',
                  boxWidth: 12,
                  font: {
                    size: 11
                  }
                }
              },

              tooltip: {
                callbacks: {
                  label: ctx =>
                    ` ${ctx.label}: ${money(ctx.raw)}`
                }
              }
            }
          }
        }
      );
  }

  /* -----------------------------------------
     TOP CATEGORY BAR
     ----------------------------------------- */

  const top =
    cats.slice(0, 8);

  const barCanvas =
    el('barChart');

  if (barCanvas) {

    charts.bar =
      new Chart(
        barCanvas,
        {
          type: 'bar',

          data: {
            labels:
              top.map(x => x[0]),

            datasets: [
              {
                data:
                  top.map(x => x[1]),

                borderRadius: 7
              }
            ]
          },

          options: baseChart({
            indexAxis: 'y',

            scales: {
              x: {
                ticks: {
                  callback: v =>
                    compactMoney(v)
                }
              },

              y: {
                grid: {
                  display: false
                }
              }
            }
          })
        }
      );
  }
}

/* =========================================================
   CHART BASE OPTIONS
   ========================================================= */

function baseChart(extra = {}) {

  return {

    responsive: true,

    maintainAspectRatio: false,

    plugins: {

      legend: {
        display: false
      },

      tooltip: {
        callbacks: {
          label: ctx =>
            ` ${money(ctx.raw)}`
        }
      }
    },

    scales: {

      x: {
        ticks: {
          color: '#8f9bb0'
        },

        grid: {
          color: '#1c2435'
        }
      },

      y: {
        ticks: {
          color: '#8f9bb0'
        },

        grid: {
          color: '#1c2435'
        }
      }
    },

    ...extra
  };
}

/* =========================================================
   TABLES
   ========================================================= */

function renderTables(rows) {

  /* -----------------------------------------
     TOP EXPENSES
     ----------------------------------------- */

  const top =
    [...rows]
      .sort(
        (a, b) =>
          b.amount - a.amount
      )
      .slice(0, 10);

  const topBody =
    el('topExpensesBody');

  if (topBody) {

    topBody.innerHTML =
      top.length

        ? top
            .map(
              r => `
                <tr>
                  <td>${fmtDate(r.date)}</td>

                  <td>
                    ${escapeHtml(
                      r.description
                    )}
                  </td>

                  <td>
                    <span class="pill">
                      ${escapeHtml(
                        r.type
                      )}
                    </span>
                  </td>

                  <td class="right amount">
                    ${money(r.amount)}
                  </td>
                </tr>
              `
            )
            .join('')

        : `
          <tr>
            <td
              colspan="4"
              class="empty"
            >
              No matching expenses.
            </td>
          </tr>
        `;
  }

  /* -----------------------------------------
     TRANSACTION LEDGER
     ----------------------------------------- */

  const recent =
    [...rows]
      .sort(
        (a, b) =>
          b.date - a.date ||
          b.amount - a.amount
      )
      .slice(0, 100);

  const logSummary =
    el('logSummary');

  if (logSummary) {

    logSummary.textContent =
      `${rows.length.toLocaleString('en-IN')} rows · showing ${recent.length}`;
  }

  const logBody =
    el('logBody');

  if (logBody) {

    logBody.innerHTML =

      recent.length

        ? recent
            .map(
              r => `
                <tr>

                  <td>
                    ${fmtDate(r.date)}
                  </td>

                  <td>
                    ${escapeHtml(
                      r.description
                    )}
                  </td>

                  <td>
                    ${escapeHtml(
                      r.type
                    )}
                  </td>

                  <td>
                    ${escapeHtml(
                      r.payment
                    )}
                  </td>

                  <td class="right amount">
                    ${money(r.amount)}
                  </td>

                  <td>
                    ${escapeHtml(
                      r.notes
                    )}
                  </td>

                </tr>
              `
            )
            .join('')

        : `
          <tr>
            <td
              colspan="6"
              class="empty"
            >
              No matching expenses.
            </td>
          </tr>
        `;
  }
}

/* =========================================================
   FILE INPUT
   ========================================================= */

const fileInput =
  el('fileInput');

if (fileInput) {

  fileInput.addEventListener(
    'change',
    async e => {

      const f =
        e.target.files?.[0];

      if (f) {

        loadWorkbook(
          await f.arrayBuffer(),
          f.name
        );
      }
    }
  );
}

/* =========================================================
   FILTER EVENTS
   ========================================================= */

[
  'monthSelect',
  'categorySelect',
  'paymentSelect'
].forEach(id => {

  const node = el(id);

  if (node) {

    node.addEventListener(
      'change',
      render
    );
  }
});

/* =========================================================
   SEARCH
   ========================================================= */

const searchInput =
  el('searchInput');

if (searchInput) {

  searchInput.addEventListener(
    'input',
    render
  );
}

/* =========================================================
   CLEAR BUTTON
   ========================================================= */

const clearBtn =
  el('clearBtn');

if (clearBtn) {

  clearBtn.addEventListener(
    'click',
    () => {

      rawRows = [];

      if (fileInput) {
        fileInput.value = '';
      }

      if (el('fileStatus')) {
        el('fileStatus').textContent =
          'No file loaded';
      }

      if (el('subhead')) {
        el('subhead').textContent =
          'Load your Excel file to turn your expense log into a live dashboard.';
      }

      if (el('monthSelect')) {
        el('monthSelect').innerHTML =
          '<option value="ALL">No data</option>';
      }

      if (el('categorySelect')) {
        el('categorySelect').innerHTML =
          '<option value="ALL">All categories</option>';
      }

      if (el('paymentSelect')) {
        el('paymentSelect').innerHTML =
          '<option value="ALL">All payment methods</option>';
      }

      [
        'totalSpend',
        'txnCount',
        'dailyAvg',
        'largestExpense'
      ].forEach(id => {

        if (!el(id)) return;

        el(id).textContent =
          id === 'txnCount'
            ? '0'
            : '₹0';
      });

      if (el('largestLabel')) {
        el('largestLabel').textContent =
          '—';
      }

      Object.values(charts)
        .forEach(c => c?.destroy());

      charts = {};

      renderTables([]);
    }
  );
}

/* =========================================================
   DRAG & DROP
   ========================================================= */

const dz =
  el('dropzone');

if (dz) {

  ['dragenter', 'dragover']
    .forEach(ev => {

      dz.addEventListener(
        ev,
        e => {

          e.preventDefault();

          dz.classList.add('drag');
        }
      );
    });

  ['dragleave', 'drop']
    .forEach(ev => {

      dz.addEventListener(
        ev,
        e => {

          e.preventDefault();

          dz.classList.remove('drag');
        }
      );
    });

  dz.addEventListener(
    'drop',
    async e => {

      const f =
        e.dataTransfer.files?.[0];

      if (f) {

        if (fileInput) {
          fileInput.value = '';
        }

        loadWorkbook(
          await f.arrayBuffer(),
          f.name
        );
      }
    }
  );
}

/* =========================================================
   INITIAL STATE
   ========================================================= */

renderTables([]);
