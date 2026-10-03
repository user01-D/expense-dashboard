let rawRows = [];
let charts = {};


/* =========================================================
   NUMBER FORMATTING
   ========================================================= */

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


function money(value) {

  return INR.format(
    Number(value) || 0
  );

}


function compactMoney(value) {

  return compactINR.format(
    Number(value) || 0
  );

}


/* =========================================================
   AMOUNT PARSER
   Handles:

   360
   360.00
   ₹360.00
   ₹1,25,000.00
   ========================================================= */

function parseAmount(value) {

  if (
    typeof value === 'number' &&
    Number.isFinite(value)
  ) {

    return value;

  }


  const cleaned =
    String(value ?? '')
      .trim()
      .replace(/₹/g, '')
      .replace(/,/g, '')
      .replace(/\s/g, '');


  if (!cleaned) {
    return 0;
  }


  const number =
    Number(cleaned);


  return Number.isFinite(number)
    ? number
    : 0;

}


/* =========================================================
   DATE PARSER

   YOUR FORMAT:
   DD/MM/YYYY

   01/09/2026
   means:
   1 September 2026

   NOT:
   January 9
   ========================================================= */

function parseDate(value) {


  /* Excel Date object */

  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );

  }


  /* Excel serial number */

  if (
    typeof value === 'number' &&
    window.XLSX?.SSF
  ) {

    try {

      const parsed =
        XLSX.SSF.parse_date_code(value);


      if (parsed) {

        return new Date(
          parsed.y,
          parsed.m - 1,
          parsed.d
        );

      }

    }

    catch (error) {

      console.warn(
        'Excel date error:',
        error
      );

    }

  }


  const text =
    String(value ?? '')
      .trim();


  if (!text) {
    return null;
  }


  /*
    IMPORTANT:

    Parse DD/MM/YYYY BEFORE
    JavaScript gets to guess.
  */

  let match =
    text.match(
      /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/
    );


  if (match) {

    const day =
      Number(match[1]);

    const month =
      Number(match[2]);

    const year =
      Number(match[3]);


    const date =
      new Date(
        year,
        month - 1,
        day
      );


    /*
      Validate date.
    */

    if (

      date.getFullYear() === year &&

      date.getMonth() === month - 1 &&

      date.getDate() === day

    ) {

      return date;

    }


    return null;

  }


  /*
    ISO fallback
    YYYY-MM-DD
  */

  match =
    text.match(
      /^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/
    );


  if (match) {

    const year =
      Number(match[1]);

    const month =
      Number(match[2]);

    const day =
      Number(match[3]);


    const date =
      new Date(
        year,
        month - 1,
        day
      );


    if (

      date.getFullYear() === year &&
      date.getMonth() === month - 1 &&
      date.getDate() === day

    ) {

      return date;

    }

  }


  return null;

}


/* =========================================================
   DATE KEY

   DON'T USE toISOString()

   because UTC conversion can shift
   the date backward in India.
   ========================================================= */

function dateKey(date) {

  return [

    date.getFullYear(),

    String(
      date.getMonth() + 1
    ).padStart(2, '0'),

    String(
      date.getDate()
    ).padStart(2, '0')

  ].join('-');

}


/* =========================================================
   DATE LABEL
   ========================================================= */

function fmtDate(date) {

  if (
    !(date instanceof Date) ||
    Number.isNaN(date.getTime())
  ) {

    return '—';

  }


  return date.toLocaleDateString(
    'en-IN',
    {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }
  );

}


/* =========================================================
   MONTH HELPERS
   ========================================================= */

function monthKey(date) {

  return `${

    date.getFullYear()

  }-${

    String(
      date.getMonth() + 1
    ).padStart(2, '0')

  }`;

}


function monthLabel(key) {

  const [
    year,
    month
  ] =
    key
      .split('-')
      .map(Number);


  return new Date(
    year,
    month - 1,
    1
  ).toLocaleDateString(
    'en-IN',
    {
      month: 'long',
      year: 'numeric'
    }
  );

}


/* =========================================================
   HTML ESCAPE
   ========================================================= */

function escapeHtml(value) {

  return String(
    value ?? ''
  ).replace(
    /[&<>"']/g,
    character => ({

      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'

    }[character])
  );

}


/* =========================================================
   NORMALIZE HEADER NAMES
   ========================================================= */

function normalizeHeader(header) {

  return String(
    header ?? ''
  )
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

}


/* =========================================================
   FIND EXCEL COLUMN
   ========================================================= */

function findColumn(row, possibleNames) {

  const wanted =
    possibleNames.map(
      normalizeHeader
    );


  return Object.keys(row)
    .find(
      key =>
        wanted.includes(
          normalizeHeader(key)
        )
    ) || null;

}


/* =========================================================
   NORMALIZE EXCEL DATA

   Expected columns:

   Date
   Expense Name
   Expense Type
   Expense Sub Class
   Amount
   ========================================================= */

function normalize(rows) {

  return rows

    .map(row => {


      const dateColumn =
        findColumn(
          row,
          [
            'Date',
            'DATE',
            'date'
          ]
        );


      const nameColumn =
        findColumn(
          row,
          [
            'Expense Name',
            'Expense Description',
            'Description'
          ]
        );


      const typeColumn =
        findColumn(
          row,
          [
            'Expense Type',
            'Category',
            'Type'
          ]
        );


      const subClassColumn =
        findColumn(
          row,
          [
            'Expense Sub Class',
            'Expense Subclass',
            'Sub Class',
            'Subclass',
            'Subcategory'
          ]
        );


      const amountColumn =
        findColumn(
          row,
          [
            'Amount',
            'Amount (₹)',
            'Amount (INR)',
            'amount',
            'AMOUNT'
          ]
        );


      const date =
        parseDate(
          dateColumn
            ? row[dateColumn]
            : ''
        );


      const description =
        String(
          nameColumn
            ? row[nameColumn]
            : ''
        ).trim();


      const type =
        String(
          typeColumn
            ? row[typeColumn]
            : 'Uncategorised'
        ).trim()
        || 'Uncategorised';


      const subClass =
        String(
          subClassColumn
            ? row[subClassColumn]
            : 'Uncategorised'
        ).trim()
        || 'Uncategorised';


      const amount =
        parseAmount(
          amountColumn
            ? row[amountColumn]
            : ''
        );


      return {

        date,

        description,

        type,

        subClass,

        amount

      };

    })

    .filter(
      row =>
        row.date &&
        row.description &&
        row.amount > 0
    );

}


/* =========================================================
   LOAD WORKBOOK
   ========================================================= */

function loadWorkbook(
  data,
  filename = 'Excel file'
) {

  try {


    const workbook =
      XLSX.read(
        data,
        {
          type: 'array',
          cellDates: true
        }
      );


    /*
      Prefer a sheet whose name
      contains "expense".
    */

    const sheetName =
      workbook.SheetNames.find(
        sheet =>
          /expense/i.test(sheet)
      )
      ||
      workbook.SheetNames[0];


    if (!sheetName) {

      throw new Error(
        'No worksheet found.'
      );

    }


    const worksheet =
      workbook.Sheets[sheetName];


    const rows =
      XLSX.utils.sheet_to_json(
        worksheet,
        {
          defval: '',
          raw: true
        }
      );


    rawRows =
      normalize(rows);


    rawRows.sort(
      (a,b) =>
        a.date - b.date
    );


    el(
      'fileStatus'
    ).textContent =
      `${filename} · ${rawRows.length} expenses`;


    el(
      'subhead'
    ).textContent =
      `${sheetName} · Visual dashboard built from valid expense rows.`;


    populateFilters();

    render();


  }

  catch (error) {

    console.error(
      error
    );


    el(
      'fileStatus'
    ).textContent =
      'Could not read file';


    el(
      'subhead'
    ).textContent =
      'Please check the Excel format and try again.';


    alert(
      'Could not read this Excel file. Use columns: Date, Expense Name, Expense Type, Expense Sub Class, Amount.'
    );

  }

}


/* =========================================================
   FILTER DROPDOWNS
   ========================================================= */

function populateFilters() {


  const categories =
    [
      ...new Set(
        rawRows.map(
          row => row.type
        )
      )
    ]
      .sort();


  const months =
    [
      ...new Set(
        rawRows.map(
          row =>
            monthKey(row.date)
        )
      )
    ]
      .sort()
      .reverse();


  el(
    'categorySelect'
  ).innerHTML =

    '<option value="ALL">All categories</option>' +

    categories
      .map(
        category =>
          `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`
      )
      .join('');


  el(
    'monthSelect'
  ).innerHTML =

    months.length

      ? months
          .map(
            month =>
              `<option value="${month}">${monthLabel(month)}</option>`
          )
          .join('')

      : '<option value="ALL">No data</option>';

}


/* =========================================================
   FILTER DATA
   ========================================================= */

function filtered() {


  const selectedMonth =
    el(
      'monthSelect'
    ).value;


  const selectedCategory =
    el(
      'categorySelect'
    ).value;


  const search =
    el(
      'searchInput'
    )
      .value
      .toLowerCase()
      .trim();


  return rawRows.filter(
    row => {


      const searchText =

        `${row.description} ${row.type} ${row.subClass}`

          .toLowerCase();


      return (

        (
          selectedMonth === 'ALL' ||

          monthKey(row.date) ===
          selectedMonth
        )

        &&

        (
          selectedCategory === 'ALL' ||

          row.type ===
          selectedCategory
        )

        &&

        (
          !search ||

          searchText.includes(
            search
          )
        )

      );

    }
  );

}


/* =========================================================
   MAIN RENDER
   ========================================================= */

function render() {


  const rows =
    filtered();


  const total =
    rows.reduce(
      (sum,row) =>
        sum + row.amount,
      0
    );


  const activeDates =
    new Set(
      rows.map(
        row =>
          dateKey(row.date)
      )
    );


  const activeDays =
    activeDates.size;


  const averageTransaction =
    rows.length
      ? total / rows.length
      : 0;


  const largest =
    rows.reduce(
      (max,row) =>
        row.amount >
        max.amount
          ? row
          : max,
      {
        amount: 0
      }
    );


  /*
    Daily totals
  */

  const dailyTotals =
    new Map();


  rows.forEach(
    row => {

      const key =
        dateKey(row.date);


      dailyTotals.set(
        key,
        (
          dailyTotals.get(key)
          || 0
        )
        + row.amount
      );

    }
  );


  const highestDay =
    [
      ...dailyTotals.entries()
    ]
      .sort(
        (a,b) =>
          b[1] - a[1]
      )[0];


  /*
    KPI values
  */

  el(
    'totalSpend'
  ).textContent =
    money(total);


  el(
    'txnCount'
  ).textContent =
    rows.length.toLocaleString(
      'en-IN'
    );


  el(
    'dailyAvg'
  ).textContent =
    money(
      activeDays
        ? total / activeDays
        : 0
    );


  el(
    'largestExpense'
  ).textContent =
    money(
      largest.amount
    );


  el(
    'largestLabel'
  ).textContent =
    largest.description ||
    '—';


  /*
    Quick numbers
  */

  el(
    'activeDays'
  ).textContent =
    activeDays.toLocaleString(
      'en-IN'
    );


  el(
    'avgTransaction'
  ).textContent =
    money(
      averageTransaction
    );


  if (highestDay) {

    const highestDate =
      new Date(
        `${highestDay[0]}T00:00:00`
      );


    el(
      'highestDay'
    ).textContent =
      fmtDate(highestDate);


    el(
      'highestDayAmount'
    ).textContent =
      money(highestDay[1]);

  }

  else {

    el(
      'highestDay'
    ).textContent =
      '—';


    el(
      'highestDayAmount'
    ).textContent =
      '₹0';

  }


  /*
    Period
  */

  const selectedMonth =
    el(
      'monthSelect'
    ).value;


  el(
    'periodLabel'
  ).textContent =

    selectedMonth === 'ALL'

      ? 'All loaded data'

      : monthLabel(
          selectedMonth
        );


  /*
    Visuals
  */

  renderCharts(rows);

  renderSnapshot(rows);

  renderTable(rows);

}


/* =========================================================
   DESTROY OLD CHARTS
   ========================================================= */

function destroyCharts() {

  Object.values(charts)
    .forEach(
      chart =>
        chart?.destroy()
    );


  charts = {};

}


/* =========================================================
   CHART COLORS
   Let Chart.js generate colors automatically.
   ========================================================= */

function baseChart(extra = {}) {

  return {

    responsive: true,

    maintainAspectRatio: false,

    interaction: {

      intersect: false,

      mode: 'index'

    },


    plugins: {

      legend: {

        display: false

      },


      tooltip: {

        callbacks: {

          label:
            context =>
              ` ${money(context.raw)}`

        }

      }

    },


    scales: {

      x: {

        ticks: {

          color: '#aeb8ca'

        },

        grid: {

          color:
            'rgba(150,170,200,.10)'

        }

      },


      y: {

        ticks: {

          color: '#aeb8ca',

          callback:
            value =>
              compactMoney(value)

        },

        grid: {

          color:
            'rgba(150,170,200,.10)'

        }

      }

    },


    ...extra

  };

}


/* =========================================================
   RENDER CHARTS
   ========================================================= */

function renderCharts(rows) {


  destroyCharts();


  const byDay =
    new Map();


  const byCategory =
    new Map();


  const bySubClass =
    new Map();


  rows.forEach(
    row => {


      const day =
        dateKey(
          row.date
        );


      byDay.set(
        day,
        (
          byDay.get(day)
          || 0
        )
        + row.amount
      );


      byCategory.set(
        row.type,
        (
          byCategory.get(row.type)
          || 0
        )
        + row.amount
      );


      bySubClass.set(
        row.subClass,
        (
          bySubClass.get(row.subClass)
          || 0
        )
        + row.amount
      );

    }
  );


  /*
    DAILY LINE CHART
  */

  const days =
    [
      ...byDay.keys()
    ].sort();


  charts.daily =
    new Chart(
      el('dailyChart'),
      {

        type: 'line',

        data: {

          labels:

            days.map(
              day =>

                new Date(
                  `${day}T00:00:00`
                )
                  .toLocaleDateString(
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
                  day =>
                    byDay.get(day)
                ),

              borderWidth: 2.5,

              fill: true,

              tension: .3,

              pointRadius: 2,

              pointHoverRadius: 5

            }

          ]

        },


        options:
          baseChart()

      }

    );


  /*
    CATEGORY PIE / DOUGHNUT
  */

  const categories =
    [
      ...byCategory.entries()
    ]
      .sort(
        (a,b) =>
          b[1] - a[1]
      );


  charts.category =
    new Chart(
      el('categoryChart'),
      {

        type: 'doughnut',

        data: {

          labels:
            categories.map(
              item => item[0]
            ),

          datasets: [

            {

              data:
                categories.map(
                  item => item[1]
                ),

              borderWidth: 3,

              hoverOffset: 6

            }

          ]

        },


        options: {

          responsive: true,

          maintainAspectRatio:
            false,

          cutout:
            '66%',


          plugins: {

            legend: {

              position:
                'right',

              labels: {

                color:
                  '#aeb8ca',

                boxWidth:
                  12,

                padding:
                  14,

                font: {

                  size:
                    11

                }

              }

            },


            tooltip: {

              callbacks: {

                label:
                  context =>
                    ` ${context.label}: ${money(context.raw)}`

              }

            }

          }

        }

      }

    );


  /*
    TOP CATEGORIES
  */

  const topCategories =
    categories.slice(
      0,
      8
    );


  charts.bar =
    new Chart(
      el('barChart'),
      {

        type:
          'bar',

        data: {

          labels:
            topCategories.map(
              item =>
                item[0]
            ),

          datasets: [

            {

              data:
                topCategories.map(
                  item =>
                    item[1]
                ),

              borderRadius:
                8,

              borderSkipped:
                false

            }

          ]

        },


        options:

          baseChart({

            indexAxis:
              'y',

            scales: {

              x: {

                ticks: {

                  color:
                    '#aeb8ca',

                  callback:
                    value =>
                      compactMoney(value)

                },

                grid: {

                  color:
                    'rgba(150,170,200,.10)'

                }

              },


              y: {

                ticks: {

                  color:
                    '#aeb8ca'

                },

                grid: {

                  display:
                    false

                }

              }

            }

          })

      }

    );


  /*
    SUB CLASS
  */

  const topSubClasses =

    [
      ...bySubClass.entries()
    ]
      .sort(
        (a,b) =>
          b[1] - a[1]
      )
      .slice(
        0,
        8
      );


  charts.subClass =
    new Chart(
      el('subClassChart'),
      {

        type:
          'bar',

        data: {

          labels:
            topSubClasses.map(
              item =>
                item[0]
            ),

          datasets: [

            {

              data:
                topSubClasses.map(
                  item =>
                    item[1]
                ),

              borderRadius:
                8,

              borderSkipped:
                false

            }

          ]

        },


        options:

          baseChart({

            indexAxis:
              'y',

            scales: {

              x: {

                ticks: {

                  color:
                    '#aeb8ca',

                  callback:
                    value =>
                      compactMoney(value)

                },

                grid: {

                  color:
                    'rgba(150,170,200,.10)'

                }

              },


              y: {

                ticks: {

                  color:
                    '#aeb8ca',

                  font: {

                    size:
                      10

                  }

                },

                grid: {

                  display:
                    false

                }

              }

            }

          })

      }

    );

}


/* =========================================================
   CATEGORY SNAPSHOT
   ========================================================= */

function renderSnapshot(rows) {


  const byCategory =
    new Map();


  rows.forEach(
    row => {

      byCategory.set(

        row.type,

        (
          byCategory.get(row.type)
          || 0
        )
        + row.amount

      );

    }
  );


  const categories =

    [
      ...byCategory.entries()
    ]
      .sort(
        (a,b) =>
          b[1] - a[1]
      )
      .slice(
        0,
        5
      );


  const total =
    rows.reduce(
      (sum,row) =>
        sum + row.amount,
      0
    );


  const max =
    categories[0]?.[1]
    || 1;


  if (!categories.length) {

    el(
      'categorySnapshot'
    ).innerHTML =

      '<div class="empty-state">No matching data.</div>';

    return;

  }


  el(
    'categorySnapshot'
  ).innerHTML =

    categories
      .map(
        ([name,value]) => {


          const barWidth =
            (
              value /
              max
            )
            * 100;


          const percentage =
            total

              ? (
                  value /
                  total *
                  100
                ).toFixed(1)

              : '0.0';


          return `

            <div class="snapshot-row">

              <div>

                <div class="snapshot-top">

                  <span>
                    ${escapeHtml(name)}
                  </span>

                  <strong>
                    ${money(value)}
                  </strong>

                </div>

                <div class="snapshot-bar">

                  <i
                    style="width:${barWidth}%"
                  ></i>

                </div>

              </div>

              <div class="snapshot-amount">

                ${percentage}%

              </div>

            </div>

          `;

        }
      )
      .join('');

}


/* =========================================================
   DETAIL TABLE
   Hidden until user opens it.
   ========================================================= */

function renderTable(rows) {


  const recent =

    [
      ...rows
    ]

      .sort(
        (a,b) =>

          b.date - a.date ||

          b.amount - a.amount
      )

      .slice(
        0,
        100
      );


  el(
    'logSummary'
  ).textContent =

    `${rows.length.toLocaleString('en-IN')} rows · showing ${recent.length}`;


  el(
    'logBody'
  ).innerHTML =

    recent.length

      ? recent
          .map(
            row => `

              <tr>

                <td>
                  ${fmtDate(row.date)}
                </td>

                <td>
                  ${escapeHtml(
                    row.description
                  )}
                </td>

                <td>

                  <span class="pill">

                    ${escapeHtml(
                      row.type
                    )}

                  </span>

                </td>

                <td>
                  ${escapeHtml(
                    row.subClass
                  )}
                </td>

                <td class="right amount">

                  ${money(
                    row.amount
                  )}

                </td>

              </tr>

            `
          )
          .join('')

      : `

        <tr>

          <td
            colspan="5"
            class="empty-state"
          >

            No matching expenses.

          </td>

        </tr>

      `;

}


/* =========================================================
   FILE INPUT
   ========================================================= */

el(
  'fileInput'
).addEventListener(
  'change',
  async event => {


    const file =
      event.target
        .files?.[0];


    if (!file) {
      return;
    }


    loadWorkbook(
      await file.arrayBuffer(),
      file.name
    );

  }
);


/* =========================================================
   FILTER EVENTS
   ========================================================= */

[
  'monthSelect',
  'categorySelect'
]
  .forEach(
    id => {

      el(id)
        .addEventListener(
          'change',
          render
        );

    }
  );


el(
  'searchInput'
)
  .addEventListener(
    'input',
    render
  );


/* =========================================================
   RESET
   ========================================================= */

el(
  'clearBtn'
).addEventListener(
  'click',
  () => {


    rawRows = [];


    el(
      'fileInput'
    ).value = '';


    el(
      'fileStatus'
    ).textContent =
      'No file loaded';


    el(
      'subhead'
    ).textContent =
      'Load your Excel file to build your visual dashboard.';


    el(
      'periodLabel'
    ).textContent =
      'No data';


    el(
      'monthSelect'
    ).innerHTML =

      '<option value="ALL">All periods</option>';


    el(
      'categorySelect'
    ).innerHTML =

      '<option value="ALL">All categories</option>';


    el(
      'searchInput'
    ).value = '';


    el(
      'totalSpend'
    ).textContent =
      '₹0';


    el(
      'dailyAvg'
    ).textContent =
      '₹0';


    el(
      'largestExpense'
    ).textContent =
      '₹0';


    el(
      'highestDayAmount'
    ).textContent =
      '₹0';


    el(
      'avgTransaction'
    ).textContent =
      '₹0';


    el(
      'txnCount'
    ).textContent =
      '0';


    el(
      'activeDays'
    ).textContent =
      '0';


    el(
      'largestLabel'
    ).textContent =
      '—';


    el(
      'highestDay'
    ).textContent =
      '—';


    destroyCharts();


    renderTable([]);

    renderSnapshot([]);

  }
);


/* =========================================================
   VIEW / HIDE DETAILS
   ========================================================= */

el(
  'toggleDetailsBtn'
)
  .addEventListener(
    'click',
    () => {


      const panel =
        el(
          'detailsPanel'
        );


      const button =
        el(
          'toggleDetailsBtn'
        );


      const isOpen =
        !panel.classList.contains(
          'hidden'
        );


      const newState =
        !isOpen;


      panel.classList.toggle(
        'hidden',
        !newState
      );


      button.classList.toggle(
        'open',
        newState
      );


      button.setAttribute(
        'aria-expanded',
        String(newState)
      );


      button.querySelector(
        '.toggle-icon'
      ).textContent =

        newState
          ? '−'
          : '+';


      button.querySelector(
        'span:nth-child(2)'
      ).textContent =

        newState

          ? 'Hide detailed transactions'

          : 'View detailed transactions';


      if (newState) {

        panel.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });

      }

    }
  );


/* =========================================================
   INITIAL STATE
   ========================================================= */

renderTable([]);

renderSnapshot([]);
