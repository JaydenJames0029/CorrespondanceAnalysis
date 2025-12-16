const state = {
  rows: [],
  charts: {},
};

const bucketOrder = [
  "Approved",
  "Commented & to be Resubmitted",
  "Rejected & to be Resubmitted",
  "Under Review",
  "Ready for use",
  "Completed",
  "Not Accepted",
  "Cancelled",
  "Obsolete",
  "Unknown",
];

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-input");
  const clearBtn = document.getElementById("clear-btn");
  fileInput.addEventListener("change", onFilesSelected);
  clearBtn.addEventListener("click", resetFilters);
  document
    .querySelectorAll(".filter select, .filter input")
    .forEach((el) => {
      const handler =
        el.type === "search" ? debounce(renderAll, 150) : renderAll;
      el.addEventListener(el.type === "search" ? "input" : "change", handler);
    });
  renderAll(); // initial empty state
});

async function onFilesSelected(evt) {
  const files = Array.from(evt.target.files || []);
  if (!files.length) return;

  const allRows = [];
  for (const file of files) {
    const workbook = await readWorkbook(file);
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      json.forEach((row, idx) => {
        const normalized = normalizeRow(row, {
          file: file.name,
          sheet: sheetName,
          idx,
        });
        if (
          normalized.originatingCompany ||
          normalized.documentNumber ||
          normalized.correspondenceNumber ||
          normalized.title
        ) {
          allRows.push(normalized);
        }
      });
    });
  }

  state.rows = allRows;
  document.getElementById("file-meta").textContent = `${files.length} file(s) loaded • ${allRows.length} rows`;
  populateFilters(allRows);
  renderAll();
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      resolve(workbook);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeRow(row, meta) {
  const get = (key) => row[key] ?? row[key?.trim()] ?? "";
  const discipline = clean(get("Discipline"));
  const recipientsRaw = clean(
    get("Recipient Companies") || get("Recipients") || get("Recipient Company")
  );
  const recipients = recipientsRaw
    .split(/;|,/)
    .map((s) => s.trim())
    .filter(Boolean);
  const docNumber = clean(
    get("Document Number") ||
      get("Drawing Number") ||
      get("Name") ||
      get("Title")
  );
  const title = clean(
    get("Title") || get("Correspondence Title") || get("Name") || ""
  );
  const origin = clean(get("Originating Company"));
  const verdict = clean(get("Final Review Verdict"));
  const status = clean(get("Correspondence Status"));
  const issueText = clean(get("Issue Reason Text"));
  const issueReason = clean(get("Issue Reason"));
  const dueDate = parseDate(
    get("Calculated Response Due Date") || get("Response Due Date")
  );
  const dateIssued = parseDate(get("Date Issued") || get("Actual Submission Date"));
  const finalReviewDate = parseDate(get("Final Review Date"));
  const completedDate = parseDate(get("Completed Date"));
  const correspondenceCreated = parseDate(
    get("Correspondence Created") || get("Created")
  );
  const revision = clean(
    get("Project Document Revision") ||
      get("Revision") ||
      get("Current Revision")
  );
  const corrNum = clean(get("Correspondence Number"));
  const itemType = clean(get("Correspondence Type") || get("Item Type"));
  const workPackage = clean(get("Work Package"));
  const path = clean(get("Path"));
  const bucket = deriveBucket({ status, verdict });
  const bestDate =
    dateIssued || finalReviewDate || completedDate || correspondenceCreated || dueDate;
  return {
    id: `${meta.file}-${meta.sheet}-${meta.idx}`,
    sourceFile: meta.file,
    sheet: meta.sheet,
    documentNumber: docNumber,
    correspondenceNumber: corrNum,
    title,
    discipline,
    disciplineCode: toDisciplineCode(discipline),
    originatingCompany: origin,
    recipients,
    recipientRaw: recipientsRaw,
    status,
    verdict,
    bucket,
    issueReasonText: issueText,
    issueReason,
    correspondenceType: itemType,
    workPackage,
    revision,
    dateIssued,
    finalReviewDate,
    dueDate,
    completedDate,
    correspondenceCreated,
    path,
    bestDate,
  };
}

function deriveBucket(row) {
  const verdict = (row.verdict || "").toLowerCase();
  const status = (row.status || "").toLowerCase();
  if (verdict.includes("commented")) return "Commented & to be Resubmitted";
  if (verdict.includes("rejected") && verdict.includes("resubmission"))
    return "Rejected & to be Resubmitted";
  if (verdict.includes("approved")) return "Approved";
  if (status === "under review") return "Under Review";
  if (status === "ready for use") return "Ready for use";
  if (status === "completed") return row.verdict ? row.verdict : "Completed";
  if (status === "not accepted") return "Not Accepted";
  if (status === "cancelled") return "Cancelled";
  if (status === "obsolete") return "Obsolete";
  return row.verdict || row.status || "Unknown";
}

function toDisciplineCode(text) {
  const map = {
    Architectural: "AR",
    "Civil & Structural": "CS",
    Electrical: "EL",
    Mechanical: "ME",
    Piping: "PI",
    "Project Management": "PM",
    Process: "PR",
    Instrumentation: "IC",
    "Building Services": "BS",
    "Construction Management": "CO",
    QAQC: "QA",
    Safety: "SF",
  };
  if (!text) return "";
  if (map[text]) return map[text];
  return text
    .split(/\s+/)
    .map((w) => w[0])
    .join("")
    .slice(0, 2)
    .toUpperCase();
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const epoch = Math.round((value - 25569) * 86400 * 1000);
    return new Date(epoch);
  }
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : new Date(parsed);
}

function formatDate(date) {
  if (!date) return "";
  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "2-digit",
  });
}

function clean(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function populateFilters(rows) {
  const docRows = rows.filter((r) => r.documentNumber);
  fillSelect(
    "origin-filter",
    uniq(rows.map((r) => r.originatingCompany))
  );
  fillSelect(
    "recipient-filter",
    uniq(
      rows.flatMap((r) =>
        r.recipients && r.recipients.length
          ? r.recipients
          : r.recipientRaw
          ? [r.recipientRaw]
          : []
      )
    )
  );
  fillSelect(
    "discipline-filter",
    uniq(rows.map((r) => r.disciplineCode || r.discipline))
  );
  fillSelect("status-filter", uniq(rows.map((r) => r.status)));
  fillSelect("verdict-filter", uniq(rows.map((r) => r.verdict)));
  fillSelect("issue-filter", uniq(rows.map((r) => r.issueReason)));
  fillSelect("reason-text-filter", uniq(rows.map((r) => r.issueReasonText)));
  fillSelect("revision-filter", uniq(docRows.map((r) => r.revision).filter(Boolean)));
  fillSelect("type-filter", uniq(rows.map((r) => r.correspondenceType)));
}

function fillSelect(id, values) {
  const select = document.getElementById(id);
  if (!select) return;
  select.innerHTML = "";
  const sorted = values
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  sorted.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    select.appendChild(opt);
  });
}

function selectedValues(id) {
  const el = document.getElementById(id);
  if (!el) return [];
  return Array.from(el.selectedOptions || []).map((o) => o.value);
}

function readFilters() {
  return {
    origins: selectedValues("origin-filter"),
    recipients: selectedValues("recipient-filter"),
    disciplines: selectedValues("discipline-filter"),
    statuses: selectedValues("status-filter"),
    verdicts: selectedValues("verdict-filter"),
    issues: selectedValues("issue-filter"),
    issueTexts: selectedValues("reason-text-filter"),
    revisions: selectedValues("revision-filter"),
    types: selectedValues("type-filter"),
    search: (document.getElementById("search-input")?.value || "")
      .trim()
      .toLowerCase(),
  };
}

function applyFilters(rows) {
  const f = readFilters();
  const recipientFilters = f.recipients.map((r) => r.toLowerCase());
  const disciplineFilters = f.disciplines.map((d) => d.toLowerCase());
  const statusFilters = f.statuses.map((s) => s.toLowerCase());
  const verdictFilters = f.verdicts.map((v) => v.toLowerCase());
  const issueFilters = f.issues.map((v) => v.toLowerCase());
  const issueTextFilters = f.issueTexts.map((v) => v.toLowerCase());
  const revisionFilters = f.revisions.map((v) => v.toLowerCase());
  const typeFilters = f.types.map((v) => v.toLowerCase());

  return rows.filter((r) => {
    if (f.search) {
      const blob = [
        r.documentNumber,
        r.title,
        r.correspondenceNumber,
      ]
        .join(" ")
        .toLowerCase();
      if (!blob.includes(f.search)) return false;
    }
    if (
      f.origins.length &&
      !f.origins.some(
        (o) => (r.originatingCompany || "").toLowerCase() === o.toLowerCase()
      )
    )
      return false;
    if (
      recipientFilters.length &&
      !((r.recipients && r.recipients.length
          ? r.recipients
          : r.recipientRaw
          ? [r.recipientRaw]
          : []
        ).some((rec) => recipientFilters.includes(rec.toLowerCase())))
    )
      return false;
    if (
      disciplineFilters.length &&
      !(disciplineFilters.includes((r.disciplineCode || r.discipline || "").toLowerCase()))
    )
      return false;
    if (
      statusFilters.length &&
      !statusFilters.includes((r.status || "").toLowerCase())
    )
      return false;
    if (
      verdictFilters.length &&
      !verdictFilters.includes((r.verdict || "").toLowerCase())
    )
      return false;
    if (
      issueFilters.length &&
      !issueFilters.includes((r.issueReason || "").toLowerCase())
    )
      return false;
    if (
      issueTextFilters.length &&
      !issueTextFilters.includes((r.issueReasonText || "").toLowerCase())
    )
      return false;
    if (
      revisionFilters.length &&
      !revisionFilters.includes((r.revision || "").toLowerCase())
    )
      return false;
    if (
      typeFilters.length &&
      !typeFilters.includes((r.correspondenceType || "").toLowerCase())
    )
      return false;
    return true;
  });
}

function latestPerDocument(rows) {
  const map = new Map();
  rows.forEach((r) => {
    const key = r.documentNumber || r.correspondenceNumber || r.title;
    if (!key) return;
    const date =
      r.dateIssued ||
      r.finalReviewDate ||
      r.completedDate ||
      r.correspondenceCreated ||
      r.dueDate ||
      r.bestDate;
    const current = map.get(key);
    if (!current) {
      map.set(key, { ...r, latestDate: date });
      return;
    }
    if (date && (!current.latestDate || date > current.latestDate)) {
      map.set(key, { ...r, latestDate: date });
    }
  });
  return Array.from(map.values());
}

function renderAll() {
  const rows = state.rows || [];
  const filteredAll = applyFilters(rows);
  const docRows = rows.filter((r) => r.documentNumber);
  const filteredDocs = applyFilters(docRows);
  const latestDocs = latestPerDocument(filteredDocs);

  renderStats(rows, latestDocs, filteredDocs);
  renderDisciplineTables(latestDocs);
  renderRevisionTable(filteredDocs);
  renderRecordsTable(filteredAll);
  renderCharts(latestDocs);
}

function renderStats(allRows, latestDocs, filteredDocs) {
  document.getElementById("rows-loaded").textContent = allRows.length.toLocaleString();
  document.getElementById("rows-loaded-sub").textContent =
    "documents + correspondences";
  document.getElementById("doc-count").textContent =
    latestDocs.length.toLocaleString();
  document.getElementById("doc-count-sub").textContent =
    "latest revision per document";
  const underReview = latestDocs.filter(
    (r) => (r.bucket || deriveBucket(r)) === "Under Review"
  );
  document.getElementById("under-review").textContent =
    underReview.length.toLocaleString();
  document.getElementById("under-review-sub").textContent =
    "open reviews (latest per document)";
  const today = new Date();
  const overdue = filteredDocs.filter((r) => {
    const bucket = r.bucket || deriveBucket(r);
    if (bucket !== "Under Review") return false;
    const due = r.dueDate;
    return due && due < today;
  });
  document.getElementById("overdue-count").textContent =
    overdue.length.toLocaleString();
  document.getElementById("overdue-sub").textContent =
    "response due date passed";
}

function renderDisciplineTables(latestDocs) {
  const totals = {};
  latestDocs.forEach((r) => {
    const disc = r.disciplineCode || r.discipline || "Unspecified";
    if (!totals[disc]) totals[disc] = { total: 0, issued: 0 };
    totals[disc].total += 1;
    if (r.dateIssued || r.finalReviewDate || r.bestDate) {
      totals[disc].issued += 1;
    }
  });
  const discKeys = Object.keys(totals).sort();
  const totalsTable = document.getElementById("discipline-totals");
  totalsTable.innerHTML = "";
  const totalsHead = `<thead><tr><th>Discipline</th><th>Total drawings/documents</th><th>Total issued</th></tr></thead>`;
  const totalsBody = discKeys
    .map(
      (d) =>
        `<tr><td>${d}</td><td>${totals[d].total}</td><td>${totals[d].issued}</td></tr>`
    )
    .join("");
  const totalsFooter = `<tr><th>Grand total</th><th>${discKeys.reduce(
    (s, d) => s + totals[d].total,
    0
  )}</th><th>${discKeys.reduce((s, d) => s + totals[d].issued, 0)}</th></tr>`;
  totalsTable.innerHTML = `${totalsHead}<tbody>${totalsBody}${totalsFooter}</tbody>`;

  const pivot = {};
  latestDocs.forEach((r) => {
    const disc = r.disciplineCode || r.discipline || "Unspecified";
    const bucket = r.bucket || deriveBucket(r);
    if (!pivot[bucket]) pivot[bucket] = {};
    pivot[bucket][disc] = (pivot[bucket][disc] || 0) + 1;
  });
  const pivotTable = document.getElementById("status-pivot");
  const disciplines = Array.from(
    new Set(
      latestDocs.map((r) => r.disciplineCode || r.discipline || "Unspecified")
    )
  ).sort();
  const header = `<thead><tr><th>Status</th>${disciplines
    .map((d) => `<th>${d}</th>`)
    .join("")}<th>Grand total</th></tr></thead>`;
  const body = bucketOrder
    .filter((b) => pivot[b])
    .map((bucket) => {
      const rowTotal = disciplines.reduce(
        (sum, d) => sum + (pivot[bucket][d] || 0),
        0
      );
      return `<tr><td>${bucket}</td>${disciplines
        .map((d) => `<td>${pivot[bucket][d] || 0}</td>`)
        .join("")}<td>${rowTotal}</td></tr>`;
    })
    .join("");
  const grandTotals = disciplines.reduce((acc, d) => {
    acc[d] = bucketOrder.reduce(
      (sum, b) => sum + (pivot[b]?.[d] || 0),
      0
    );
    return acc;
  }, {});
  const grandTotalRow = `<tr><th>Grand total</th>${disciplines
    .map((d) => `<th>${grandTotals[d]}</th>`)
    .join("")}<th>${bucketOrder.reduce(
    (sum, b) =>
      sum +
      disciplines.reduce((inner, d) => inner + (pivot[b]?.[d] || 0), 0),
    0
  )}</th></tr>`;
  pivotTable.innerHTML = `${header}<tbody>${body}${grandTotalRow}</tbody>`;
}

function renderRevisionTable(docRows) {
  const targetStatuses = new Set([
    "Under Review",
    "Commented & to be Resubmitted",
    "Rejected & to be Resubmitted",
  ]);
  const subset = docRows.filter((r) => targetStatuses.has(r.bucket));
  const grouped = new Map();
  subset.forEach((r) => {
    if (!r.documentNumber) return;
    const key = r.documentNumber;
    const existing = grouped.get(key) || {
      base: r,
      baseDate: null,
      revisions: new Map(),
      latest: null,
    };
    const rev = r.revision || "N/A";
    const date = r.dateIssued || r.finalReviewDate || r.bestDate;
    const currentRev = existing.revisions.get(rev);
    if (!currentRev || (date && (!currentRev.date || date > currentRev.date))) {
      existing.revisions.set(rev, { date, status: r.bucket, corr: r.correspondenceNumber });
    }
    if (!existing.latest || (date && (!existing.latest.date || date > existing.latest.date))) {
      existing.latest = { date, rev, bucket: r.bucket };
    }
    if (!existing.baseDate || (date && date > existing.baseDate)) {
      existing.base = r;
      existing.baseDate = date;
    }
    grouped.set(key, existing);
  });

  const allRevs = new Set();
  grouped.forEach((g) => g.revisions.forEach((_, rev) => allRevs.add(rev)));
  const revColumns = Array.from(allRevs).sort(sortRevision);
  const table = document.getElementById("revision-table");
  if (!grouped.size) {
    table.innerHTML = "<tbody><tr><td>No rows in the selected filters.</td></tr></tbody>";
    return;
  }
  const head =
    "<thead><tr>" +
    ["Document", "Title", "Discipline", "Originating company", "Recipients", "Status", "Due date", "Current rev"]
      .map((h) => `<th>${h}</th>`)
      .join("") +
    revColumns.map((r) => `<th>${r}</th>`).join("") +
    "</tr></thead>";

  const rows = Array.from(grouped.values())
    .sort((a, b) => (a.base.documentNumber || "").localeCompare(b.base.documentNumber || ""))
    .map((g) => {
      const base = g.base;
      const latest = g.latest || {};
      const recipients = base.recipients.length
        ? base.recipients.join(", ")
        : base.recipientRaw;
      const dueCell =
        base.bucket === "Under Review" && base.dueDate && base.dueDate < new Date()
          ? `<span class="danger">${formatDate(base.dueDate)}</span>`
          : formatDate(base.dueDate);
      const revCells = revColumns
        .map((rev) => {
          const info = g.revisions.get(rev);
          if (!info) return "<td></td>";
          const text = info.date ? formatDate(info.date) : "";
          return `<td>${text}</td>`;
        })
        .join("");
      return `<tr>
        <td>${base.documentNumber}</td>
        <td>${base.title}</td>
        <td><span class="pill">${base.disciplineCode || base.discipline || ""}</span></td>
        <td>${base.originatingCompany}</td>
        <td>${recipients || ""}</td>
        <td>${base.bucket}</td>
        <td>${dueCell}</td>
        <td>${latest.rev || ""}</td>
        ${revCells}
      </tr>`;
    })
    .join("");
  table.innerHTML = `${head}<tbody>${rows}</tbody>`;
}

function renderRecordsTable(rows) {
  const table = document.getElementById("records-table");
  if (!rows.length) {
    table.innerHTML = "<tbody><tr><td>No rows in the selected filters.</td></tr></tbody>";
    return;
  }
  const head =
    "<thead><tr>" +
    [
      "Document number",
      "Correspondence no.",
      "Title",
      "Discipline",
      "Originating company",
      "Recipients",
      "Status",
      "Final verdict",
      "Revision",
      "Date issued",
      "Due date",
      "Final review date",
      "Work package",
      "Type",
      "Source",
    ]
      .map((h) => `<th>${h}</th>`)
      .join("") +
    "</tr></thead>";
  const body = rows
    .map((r) => {
      const recipients = r.recipients.length
        ? r.recipients.join(", ")
        : r.recipientRaw;
      return `<tr>
        <td>${r.documentNumber || ""}</td>
        <td>${r.correspondenceNumber || ""}</td>
        <td>${r.title || ""}</td>
        <td>${r.disciplineCode || r.discipline || ""}</td>
        <td>${r.originatingCompany || ""}</td>
        <td>${recipients || ""}</td>
        <td>${r.status || ""}</td>
        <td>${r.verdict || ""}</td>
        <td>${r.revision || ""}</td>
        <td>${formatDate(r.dateIssued)}</td>
        <td>${formatDate(r.dueDate)}</td>
        <td>${formatDate(r.finalReviewDate)}</td>
        <td>${r.workPackage || ""}</td>
        <td>${r.correspondenceType || ""}</td>
        <td>${r.sourceFile}</td>
      </tr>`;
    })
    .join("");
  table.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderCharts(latestDocs) {
  renderChart(
    "status-chart",
    "pie",
    buildStatusDataset(latestDocs),
    "Status distribution"
  );
  renderChart(
    "discipline-chart",
    "doughnut",
    buildDisciplineDataset(latestDocs),
    "Discipline mix"
  );
}

function buildStatusDataset(latestDocs) {
  const counts = {};
  latestDocs.forEach((r) => {
    const bucket = r.bucket || deriveBucket(r);
    counts[bucket] = (counts[bucket] || 0) + 1;
  });
  const labels = Object.keys(counts);
  return {
    labels,
    datasets: [
      {
        label: "Status",
        data: labels.map((l) => counts[l]),
        backgroundColor: labels.map(colorForStatus),
      },
    ],
  };
}

function buildDisciplineDataset(latestDocs) {
  const counts = {};
  latestDocs.forEach((r) => {
    const disc = r.disciplineCode || r.discipline || "Unspecified";
    counts[disc] = (counts[disc] || 0) + 1;
  });
  const labels = Object.keys(counts);
  return {
    labels,
    datasets: [
      {
        label: "Discipline",
        data: labels.map((l) => counts[l]),
        backgroundColor: labels.map((_, idx) => palette(idx)),
      },
    ],
  };
}

function renderChart(id, type, data, label) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  if (state.charts[id]) {
    state.charts[id].data = data;
    state.charts[id].update();
    return;
  }
  state.charts[id] = new Chart(ctx, {
    type,
    data,
    options: {
      plugins: {
        legend: {
          position: "bottom",
          labels: { color: "#e2e8f0" },
        },
        title: {
          display: false,
          text: label,
          color: "#e2e8f0",
        },
      },
    },
  });
}

function colorForStatus(status) {
  const map = {
    Approved: "#63f3c3",
    "Commented & to be Resubmitted": "#f7c948",
    "Rejected & to be Resubmitted": "#f57777",
    "Under Review": "#6a8bff",
    "Ready for use": "#63f3c3",
    Completed: "#4dd4b0",
    "Not Accepted": "#f57777",
    Cancelled: "#9ca3af",
    Obsolete: "#64748b",
    Unknown: "#cbd5e1",
  };
  return map[status] || palette(status.length);
}

function palette(i) {
  const colors = [
    "#63f3c3",
    "#6a8bff",
    "#f7c948",
    "#f57777",
    "#ff9bd0",
    "#8be9fd",
    "#c792ea",
    "#94a3b8",
  ];
  return colors[i % colors.length];
}

function resetFilters() {
  document
    .querySelectorAll(".filter select")
    .forEach((select) => {
      Array.from(select.options).forEach((opt) => (opt.selected = false));
    });
  const search = document.getElementById("search-input");
  if (search) search.value = "";
  renderAll();
}

function sortRevision(a, b) {
  const aNum = Number(a);
  const bNum = Number(b);
  const aIsNum = !Number.isNaN(aNum);
  const bIsNum = !Number.isNaN(bNum);
  if (aIsNum && bIsNum) return aNum - bNum;
  return a.localeCompare(b, undefined, { numeric: true });
}

function uniq(arr) {
  return Array.from(
    new Set(arr.filter((v) => v !== null && v !== undefined && v !== ""))
  );
}

function debounce(fn, wait = 200) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}
