<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Marking Deadline Calculator</title>
<style>
  :root {
    --ink: #1F2A37;
    --paper: #FAF9F6;
    --line: #D8D2C4;
    --accent: #7A2E2E;
    --accent-soft: #F1E4E0;
    --muted: #6B7280;
  }

  * { box-sizing: border-box; }

  body {
    margin: 0;
    background: var(--paper);
    color: var(--ink);
    font-family: 'Georgia', 'Iowan Old Style', serif;
    display: flex;
    justify-content: center;
    padding: 48px 20px;
    min-height: 100vh;
  }

  main {
    width: 100%;
    max-width: 560px;
  }

  header {
    margin-bottom: 32px;
    border-bottom: 1px solid var(--line);
    padding-bottom: 20px;
  }

  h1 {
    font-size: 1.55rem;
    font-weight: 600;
    margin: 0 0 6px 0;
    letter-spacing: 0.2px;
  }

  .subtitle {
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.88rem;
    color: var(--muted);
    margin: 0;
  }

  .rule-note {
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.78rem;
    color: var(--muted);
    background: var(--accent-soft);
    border-left: 3px solid var(--accent);
    padding: 8px 12px;
    margin-top: 14px;
    line-height: 1.4;
  }

  .card {
    background: #fff;
    border: 1px solid var(--line);
    padding: 28px;
    margin-bottom: 28px;
  }

  label {
    display: block;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.82rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--muted);
    margin-bottom: 8px;
  }

  input[type="date"] {
    width: 100%;
    padding: 12px 14px;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 1rem;
    border: 1px solid var(--line);
    background: var(--paper);
    color: var(--ink);
  }

  input[type="date"]:focus {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }

  button {
    margin-top: 16px;
    width: 100%;
    padding: 13px 14px;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.92rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    background: var(--accent);
    color: #fff;
    border: none;
    cursor: pointer;
  }

  button:hover { background: #5E2323; }
  button:focus-visible { outline: 2px solid var(--ink); outline-offset: 2px; }

  .result {
    margin-top: 20px;
    padding: 18px;
    background: var(--accent-soft);
    border: 1px dashed var(--accent);
    display: none;
  }

  .result.visible { display: block; }

  .result .label {
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--accent);
    margin-bottom: 4px;
  }

  .result .date {
    font-size: 1.3rem;
    font-weight: 600;
  }

  .result .day-of-week {
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.85rem;
    color: var(--muted);
    margin-top: 2px;
  }

  h2 {
    font-size: 1rem;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--muted);
    margin: 0 0 14px 0;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.9rem;
  }

  th, td {
    text-align: left;
    padding: 10px 8px;
    border-bottom: 1px solid var(--line);
  }

  th {
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--muted);
    font-weight: 600;
  }

  .remove-btn {
    background: none;
    border: none;
    color: var(--muted);
    font-size: 1.1rem;
    cursor: pointer;
    padding: 0;
    width: auto;
    margin: 0;
    line-height: 1;
  }

  .remove-btn:hover { color: var(--accent); }

  .empty-state {
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.85rem;
    color: var(--muted);
    font-style: italic;
  }
</style>
</head>
<body>
<main>
  <header>
    <h1>Marking Deadline Calculator</h1>
    <p class="subtitle">Enter a submission date to work out when marks are due.</p>
    <div class="rule-note">Marking deadlines are calculated as 20 calendar days after the submission date (weekends included).</div>
  </header>

  <div class="card">
    <label for="submission-date">Submission date</label>
    <input type="date" id="submission-date">
    <button id="calculate-btn">Calculate deadline</button>

    <div class="result" id="result">
      <div class="label">Marking deadline</div>
      <div class="date" id="result-date"></div>
      <div class="day-of-week" id="result-day"></div>
    </div>
  </div>

  <div class="card">
    <h2>Logged submissions</h2>
    <div id="table-wrap">
      <p class="empty-state" id="empty-state">No submissions logged yet.</p>
    </div>
  </div>
</main>

<script>
  const dateInput = document.getElementById('submission-date');
  const calcBtn = document.getElementById('calculate-btn');
  const resultBox = document.getElementById('result');
  const resultDate = document.getElementById('result-date');
  const resultDay = document.getElementById('result-day');
  const tableWrap = document.getElementById('table-wrap');
  const emptyState = document.getElementById('empty-state');

  const DAYS_TO_ADD = 20;
  let log = [];

  function formatDate(d) {
    return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  }

  function formatDayOfWeek(d) {
    return d.toLocaleDateString('en-GB', { weekday: 'long' });
  }

  function calcDeadline(submissionDateStr) {
    const parts = submissionDateStr.split('-').map(Number);
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    d.setDate(d.getDate() + DAYS_TO_ADD);
    return d;
  }

  function renderTable() {
    if (log.length === 0) {
      tableWrap.innerHTML = '<p class="empty-state" id="empty-state">No submissions logged yet.</p>';
      return;
    }

    let rows = log.map((entry, i) => `
      <tr>
        <td>${entry.submission}</td>
        <td>${entry.deadline}</td>
        <td><button class="remove-btn" data-index="${i}" title="Remove">&times;</button></td>
      </tr>
    `).join('');

    tableWrap.innerHTML = `
      <table>
        <thead>
          <tr><th>Submitted</th><th>Marks due</th><th></th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;

    tableWrap.querySelectorAll('.remove-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        log.splice(Number(btn.dataset.index), 1);
        renderTable();
      });
    });
  }

  calcBtn.addEventListener('click', () => {
    if (!dateInput.value) {
      dateInput.focus();
      return;
    }

    const parts = dateInput.value.split('-').map(Number);
    const submissionDate = new Date(parts[0], parts[1] - 1, parts[2]);
    const deadline = calcDeadline(dateInput.value);

    resultDate.textContent = formatDate(deadline);
    resultDay.textContent = formatDayOfWeek(deadline);
    resultBox.classList.add('visible');

    log.unshift({
      submission: formatDate(submissionDate),
      deadline: formatDate(deadline)
    });
    renderTable();
  });

  dateInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') calcBtn.click();
  });
</script>
</body>
</html>
