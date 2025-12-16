const state = {
  rows: [],
  charts: {},
  view: {},
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
  const exportBtn = document.getElementById("export-btn");
  fileInput.addEventListener("change", onFilesSelected);
  clearBtn.addEventListener("click", resetFilters);
  exportBtn.addEventListener("click", exportToExcel);
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
  const names = files.map((f) => f.name).join(", ");
  document.getElementById("file-meta").textContent = `${files.length} file(s): ${names} • ${allRows.length} rows`;
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
  const finalReviewComments = clean(get("Final Review Comments"));
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
    finalReviewComments,
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

function applyFilters(rows, opts = {}) {
  const ignoreStatus = !!opts.ignoreStatus;
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
    if (!ignoreStatus) {
      if (
        statusFilters.length &&
        !statusFilters.includes((r.status || "").toLowerCase())
      )
        return false;
    }
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

function computeDisciplineData(latestDocs) {
  const totals = {};
  latestDocs.forEach((r) => {
    const disc = r.disciplineCode || r.discipline || "Unspecified";
    if (!totals[disc]) totals[disc] = { total: 0, issued: 0 };
    totals[disc].total += 1;
    if (r.dateIssued || r.finalReviewDate || r.bestDate) {
      totals[disc].issued += 1;
    }
  });

  const pivot = {};
  latestDocs.forEach((r) => {
    const disc = r.disciplineCode || r.discipline || "Unspecified";
    const bucket = r.bucket || deriveBucket(r);
    if (!pivot[bucket]) pivot[bucket] = {};
    pivot[bucket][disc] = (pivot[bucket][disc] || 0) + 1;
  });

  const disciplines = Object.keys(totals).sort();
  return { totals, pivot, disciplines };
}

function computeRevisionData(docRows) {
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
  const rows = Array.from(grouped.values())
    .sort((a, b) => (a.base.documentNumber || "").localeCompare(b.base.documentNumber || ""))
    .map((g) => {
      const base = g.base;
      const latest = g.latest || {};
      const recipients = base.recipients.length
        ? base.recipients.join(", ")
        : base.recipientRaw;
      const revDates = {};
      revColumns.forEach((rev) => {
        const info = g.revisions.get(rev);
        revDates[rev] = info?.date || null;
      });
      return {
        documentNumber: base.documentNumber,
        title: base.title,
        discipline: base.disciplineCode || base.discipline || "",
        originatingCompany: base.originatingCompany,
        recipients,
        status: base.bucket,
        dueDate: base.dueDate,
        correspondenceNumber: base.correspondenceNumber,
        currentRev: latest.rev || "",
        docRevision: base.revision || "",
        issueReasonText: base.issueReasonText || "",
        issueReason: base.issueReason || "",
        completedDate: base.completedDate || "",
        finalReviewDate: base.finalReviewDate || "",
        workPackage: base.workPackage || "",
        comments: base.finalReviewComments || "",
        revisions: revDates,
      };
    });

  return { columns: revColumns, rows };
}

function computeLetteredOpen(targetDocs, historyDocs) {
  const targetStatuses = new Set([
    "Under Review",
    "Commented & to be Resubmitted",
    "Rejected & to be Resubmitted",
  ]);
  // doc keys we need to surface (based on target statuses)
  const neededKeys = new Set(
    targetDocs
      .filter((r) => targetStatuses.has(r.bucket || deriveBucket(r)))
      .map((r) => r.documentNumber || `corr:${r.correspondenceNumber}`)
      .filter(Boolean)
  );
  if (!neededKeys.size) return { columns: [], rows: [] };

  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let maxLetters = 0;
  const rows = [];

  neededKeys.forEach((key) => {
    const history = historyDocs.filter((r) => {
      const k = r.documentNumber || `corr:${r.correspondenceNumber}`;
      return k === key;
    });
    if (!history.length) return;

    // sort by best available date; missing dates go last
    const sorted = history
      .map((r) => {
        const date =
          r.dateIssued ||
          r.finalReviewDate ||
          r.completedDate ||
          r.correspondenceCreated ||
          r.dueDate ||
          r.bestDate;
        return { row: r, date };
      })
      .sort((a, b) => {
        if (!a.date && !b.date) return 0;
        if (!a.date) return 1;
        if (!b.date) return -1;
        return a.date - b.date;
      });

    // assign letters strictly in chronological order (no gaps)
    const letterMap = new Map();
    sorted.forEach((entry, idx) => {
      const letter = letters[idx] || `R${idx}`;
      letterMap.set(letter, entry);
    });
    maxLetters = Math.max(maxLetters, letterMap.size);

    const targetRows = targetDocs.filter((r) => {
      const k = r.documentNumber || `corr:${r.correspondenceNumber}`;
      return k === key && targetStatuses.has(r.bucket || deriveBucket(r));
    });
    // pick latest target for status/metadata
    const base =
      targetRows
        .map((r) => ({
          r,
          date:
            r.dateIssued ||
            r.finalReviewDate ||
            r.completedDate ||
            r.correspondenceCreated ||
            r.dueDate ||
            r.bestDate,
        }))
        .sort((a, b) => {
          if (!a.date && !b.date) return 0;
          if (!a.date) return 1;
          if (!b.date) return -1;
          return a.date - b.date;
        })
        .slice(-1)[0]?.r || targetRows[0] || history.slice(-1)[0].row;

    const revDates = {};
    letterMap.forEach((entry, letter) => {
      revDates[letter] = entry.date || null;
    });

    const flags = inferPhaseFlags(base.issueReasonText || base.issueReason);
    const recips = base.recipients.length ? base.recipients.join(", ") : base.recipientRaw;
    rows.push({
      index: rows.length + 1,
      discipline: base.disciplineCode || base.discipline || "",
      ifcFlag: flags.ifc ? "X" : "",
      documentNumber: base.documentNumber || "",
      title: base.title || "",
      currentRevision: base.revision || Array.from(letterMap.keys()).slice(-1)[0] || "",
      category: base.issueReasonText || base.issueReason || "",
      revDates,
      correspondenceNumber: base.correspondenceNumber || "",
      status: base.bucket || deriveBucket(base),
      dueDate: base.dueDate,
      completedDate: base.completedDate || base.finalReviewDate,
      remark: base.finalReviewComments || "",
      flags,
      recipients: recips || "",
    });
  });

  rows.sort((a, b) => (a.documentNumber || "").localeCompare(b.documentNumber || ""));
  rows.forEach((r, i) => (r.index = i + 1));

  const letterColumns = Array.from({ length: maxLetters }, (_, i) => letters[i] || `R${i}`);
  return { columns: letterColumns, rows };
}

function renderAll() {
  const rows = state.rows || [];
  const filteredAll = applyFilters(rows);
  const docRows = rows.filter((r) => r.documentNumber);
  const filteredDocs = applyFilters(docRows);
  const filteredDocsNoStatus = applyFilters(docRows, { ignoreStatus: true });
  const latestDocs = latestPerDocument(filteredDocs);

  const disciplineData = computeDisciplineData(latestDocs);
  const revisionData = computeRevisionData(filteredDocs);
  const letteredData = computeLetteredOpen(filteredDocs, filteredDocsNoStatus);
  const statusDataset = buildStatusDataset(latestDocs);
  const disciplineDataset = buildDisciplineDataset(latestDocs);

  state.view = {
    filteredAll,
    filteredDocs,
    latestDocs,
    disciplineData,
    revisionData,
    letteredData,
    statusDataset,
    disciplineDataset,
  };

  renderStats(rows, latestDocs, filteredDocs);
  renderDisciplineTables(disciplineData);
  renderRevisionTable(revisionData);
  renderLetteredOpenTable(letteredData);
  renderRecordsTable(filteredAll);
  renderCharts(statusDataset, disciplineDataset);
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
  const underPct = latestDocs.length
    ? ((underReview.length / latestDocs.length) * 100).toFixed(1)
    : "0.0";
  document.getElementById("under-review").textContent =
    underReview.length.toLocaleString();
  document.getElementById("under-review-sub").textContent =
    `open reviews (${underPct}% of docs)`;
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

function renderDisciplineTables(data) {
  const { totals, pivot, disciplines } = data;
  const discKeys = disciplines.length ? disciplines : Object.keys(totals).sort();
  const totalsTable = document.getElementById("discipline-totals");
  totalsTable.innerHTML = "";
  const totalsHead = `<thead><tr><th>Discipline</th><th>Total drawings/documents</th><th>Total issued</th></tr></thead>`;
  const totalsBody = discKeys
    .map(
      (d) =>
        `<tr><td>${d}</td><td>${totals[d]?.total || 0}</td><td>${totals[d]?.issued || 0}</td></tr>`
    )
    .join("");
  const totalsFooter = `<tr><th>Grand total</th><th>${discKeys.reduce(
    (s, d) => s + (totals[d]?.total || 0),
    0
  )}</th><th>${discKeys.reduce((s, d) => s + (totals[d]?.issued || 0), 0)}</th></tr>`;
  totalsTable.innerHTML = `${totalsHead}<tbody>${totalsBody}${totalsFooter}</tbody>`;

  const pivotTable = document.getElementById("status-pivot");
  const header = `<thead><tr><th>Status</th>${disciplines
    .map((d) => `<th>${d}</th>`)
    .join("")}<th>Grand total</th></tr></thead>`;
  const body = bucketOrder
    .filter((b) => pivot[b])
    .map((bucket) => {
      const rowTotal = disciplines.reduce(
        (sum, d) => sum + (pivot[bucket]?.[d] || 0),
        0
      );
      return `<tr><td>${bucket}</td>${disciplines
        .map((d) => `<td>${pivot[bucket]?.[d] || 0}</td>`)
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
    .map((d) => `<th>${grandTotals[d] || 0}</th>`)
    .join("")}<th>${bucketOrder.reduce(
    (sum, b) =>
      sum +
      disciplines.reduce((inner, d) => inner + (pivot[b]?.[d] || 0), 0),
    0
  )}</th></tr>`;
  pivotTable.innerHTML = `${header}<tbody>${body}${grandTotalRow}</tbody>`;
}

function renderRevisionTable(revisionData) {
  const table = document.getElementById("revision-table");
  if (!revisionData.rows.length) {
    table.innerHTML = "<tbody><tr><td>No rows in the selected filters.</td></tr></tbody>";
    return;
  }
  const revColumns = revisionData.columns;
  const head =
    "<thead><tr>" +
    ["Document", "Title", "Discipline", "Originating company", "Recipients", "Status", "Due date", "Current rev"]
      .map((h) => `<th>${h}</th>`)
      .join("") +
    revColumns.map((r) => `<th>${r}</th>`).join("") +
    "</tr></thead>";

  const rows = revisionData.rows
    .map((row) => {
      const dueCell =
        row.status === "Under Review" && row.dueDate && row.dueDate < new Date()
          ? `<span class="danger">${formatDate(row.dueDate)}</span>`
          : formatDate(row.dueDate);
      const revCells = revColumns
        .map((rev) => {
          const date = row.revisions[rev];
          return `<td>${date ? formatDate(date) : ""}</td>`;
        })
        .join("");
      return `<tr>
        <td>${row.documentNumber}</td>
        <td>${row.title}</td>
        <td><span class="pill">${row.discipline}</span></td>
        <td>${row.originatingCompany || ""}</td>
        <td>${row.recipients || ""}</td>
        <td>${row.status}</td>
        <td>${dueCell}</td>
        <td>${row.currentRev}</td>
        ${revCells}
      </tr>`;
    })
    .join("");
  table.innerHTML = `${head}<tbody>${rows}</tbody>`;
}

function renderOpenReviewsTable(letteredData) {
  const table = document.getElementById("open-reviews-table");
  if (!table) return;
  if (!letteredData.rows.length) {
    table.innerHTML = "<tbody><tr><td>No rows in the selected filters.</td></tr></tbody>";
    return;
  }
  const revColumns = letteredData.columns;
  const head =
    "<thead><tr>" +
    [
      "#",
      "Discipline",
      "IFC to be Issued",
      "Drawing number",
      "Description",
      "Current revision",
      "Category",
    ]
      .map((h) => `<th>${h}</th>`)
      .join("") +
    revColumns.map((r) => `<th>${r}</th>`).join("") +
    [
      "Correspondence no.",
      "Status",
      "Due date",
      "Completed date",
      "Remark",
      "BD",
      "30%",
      "60%",
      "90%",
      "IFC",
      "Impacted?",
    ]
      .map((h) => `<th>${h}</th>`)
      .join("") +
    "</tr></thead>";

  const rows = letteredData.rows.map((row) => {
    const flags = row.flags || inferPhaseFlags(row.category);
    const revCells = revColumns
      .map((rev) => {
        const date = row.revDates[rev];
        return `<td>${date ? formatDate(date) : ""}</td>`;
      })
      .join("");
    const ifcToBeIssued = flags.ifc ? "X" : "";
    return `<tr>
      <td>${row.index}</td>
      <td>${row.discipline}</td>
      <td>${ifcToBeIssued}</td>
      <td>${row.documentNumber || ""}</td>
      <td>${row.title || ""}</td>
      <td>${row.currentRevision || ""}</td>
      <td>${row.category || ""}</td>
      ${revCells}
      <td>${row.correspondenceNumber || ""}</td>
      <td>${row.status}</td>
      <td>${formatDate(row.dueDate)}</td>
      <td>${formatDate(row.completedDate)}</td>
      <td>${row.remark || ""}</td>
      <td>${flags.bd ? "1" : ""}</td>
      <td>${flags.thirty ? "1" : ""}</td>
      <td>${flags.sixty ? "1" : ""}</td>
      <td>${flags.ninety ? "1" : ""}</td>
      <td>${flags.ifc ? "1" : ""}</td>
      <td></td>
    </tr>`;
  }).join("");
  table.innerHTML = `${head}<tbody>${rows}</tbody>`;
}

function renderLetteredOpenTable(letteredData) {
  // For now reuse the same table; function exists for clarity/extendability.
  renderOpenReviewsTable(letteredData);
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

function renderCharts(statusDataset, disciplineDataset) {
  renderChart("status-chart", "pie", statusDataset, "Status distribution");
  renderChart(
    "discipline-chart",
    "doughnut",
    disciplineDataset,
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
  const total = labels.reduce((sum, l) => sum + counts[l], 0) || 1;
  const percentages = labels.map((l) => (counts[l] / total) * 100);
  const displayLabels = labels.map(
    (l, i) => `${l} (${percentages[i].toFixed(1)}%)`
  );
  return {
    labels: displayLabels,
    rawLabels: labels,
    percentages,
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

function exportToExcel() {
  if (typeof XLSX === "undefined") {
    alert("SheetJS (XLSX) is not loaded.");
    return;
  }
  const view = state.view || {};
  if (!view.filteredAll || !view.filteredAll.length) {
    alert("Upload a query and wait for data to load before exporting.");
    return;
  }

  const wb = XLSX.utils.book_new();
  const addSheet = (name, rows) => {
    const safeName = name.slice(0, 31);
    const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ Notice: "No data" }]);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  };

  const recordRow = (r) => {
    const recips = Array.isArray(r.recipients) ? r.recipients : [];
    return {
      "Document number": r.documentNumber || "",
      "Correspondence no.": r.correspondenceNumber || "",
      Title: r.title || "",
      Discipline: r.disciplineCode || r.discipline || "",
      "Originating company": r.originatingCompany || "",
      Recipients: recips.length ? recips.join(", ") : r.recipientRaw || "",
      Status: r.status || "",
      "Status bucket": r.bucket || deriveBucket(r),
      "Final verdict": r.verdict || "",
      Revision: r.revision || "",
      "Date issued": r.dateIssued || "",
      "Due date": r.dueDate || "",
      "Final review date": r.finalReviewDate || "",
      "Completed date": r.completedDate || "",
      "Created date": r.correspondenceCreated || "",
      "Issue reason": r.issueReason || "",
      "Issue reason text": r.issueReasonText || "",
      "Work package": r.workPackage || "",
      Type: r.correspondenceType || "",
      "Source file": r.sourceFile || "",
    };
  };

  const statsRows = [
    { Metric: "Rows loaded (all sheets)", Value: state.rows.length },
    { Metric: "Filtered rows (all)", Value: view.filteredAll.length },
    { Metric: "Documents with revisions (filtered)", Value: view.filteredDocs?.length || 0 },
    { Metric: "Unique documents (latest)", Value: view.latestDocs?.length || 0 },
  ];

  const disciplineRows = [];
  const totals = view.disciplineData?.totals || {};
  Object.keys(totals)
    .sort()
    .forEach((disc) => {
      disciplineRows.push({
        Discipline: disc,
        "Total drawings/documents": totals[disc]?.total || 0,
        "Total issued": totals[disc]?.issued || 0,
      });
    });
  disciplineRows.push({
    Discipline: "Grand total",
    "Total drawings/documents": disciplineRows.reduce((s, r) => s + (r["Total drawings/documents"] || 0), 0),
    "Total issued": disciplineRows.reduce((s, r) => s + (r["Total issued"] || 0), 0),
  });

  const pivotRows = [];
  const disciplines = view.disciplineData?.disciplines || [];
  const pivot = view.disciplineData?.pivot || {};
  bucketOrder
    .filter((b) => pivot[b])
    .forEach((bucket) => {
      const row = { Status: bucket };
      disciplines.forEach((d) => {
        row[d] = pivot[bucket]?.[d] || 0;
      });
      row["Row total"] = disciplines.reduce((sum, d) => sum + (pivot[bucket]?.[d] || 0), 0);
      pivotRows.push(row);
    });
  const grandRow = { Status: "Grand total" };
  disciplines.forEach((d) => {
    grandRow[d] = bucketOrder.reduce((sum, bucket) => sum + (pivot[bucket]?.[d] || 0), 0);
  });
  grandRow["Row total"] = bucketOrder.reduce(
    (sum, bucket) => sum + (pivot[bucket] ? Object.values(pivot[bucket]).reduce((s, v) => s + v, 0) : 0),
    0
  );
  pivotRows.push(grandRow);

  const revisionRows = [];
  const revData = view.revisionData || { columns: [], rows: [] };
  revData.rows.forEach((row) => {
    const obj = {
      "Document number": row.documentNumber,
      Title: row.title,
      Discipline: row.discipline,
      "Originating company": row.originatingCompany,
      Recipients: row.recipients,
      Status: row.status,
      "Due date": row.dueDate || "",
      "Current rev": row.currentRev,
    };
    revData.columns.forEach((rev) => {
      obj[`Rev ${rev}`] = row.revisions[rev] || "";
    });
    revisionRows.push(obj);
  });

  const openRows = [];
  const letterData = view.letteredData || { columns: [], rows: [] };
  letterData.rows.forEach((row) => {
    const flags = row.flags || inferPhaseFlags(row.category);
    const obj = {
      "#": row.index,
      Discipline: row.discipline,
      "IFC to be Issued": flags.ifc ? "X" : "",
      "Drawing number": row.documentNumber || "",
      Description: row.title || "",
      "Current revision": row.currentRevision || "",
      Category: row.category || "",
      "Correspondence no.": row.correspondenceNumber || "",
      Status: row.status,
      "Due date": row.dueDate || "",
      "Completed date": row.completedDate || "",
      Remark: row.remark || "",
      BD: flags.bd ? 1 : "",
      "30%": flags.thirty ? 1 : "",
      "60%": flags.sixty ? 1 : "",
      "90%": flags.ninety ? 1 : "",
      IFC: flags.ifc ? 1 : "",
      "Impacted?": "",
    };
    letterData.columns.forEach((rev) => {
      obj[`Rev ${rev}`] = row.revDates[rev] || "";
    });
    openRows.push(obj);
  });

  const chartRows = [];
  const statusDataset = view.statusDataset || { labels: [], rawLabels: [], percentages: [], datasets: [] };
  const disciplineDataset = view.disciplineDataset || { labels: [], datasets: [] };
  if (statusDataset.datasets[0]) {
    const labels = statusDataset.rawLabels && statusDataset.rawLabels.length ? statusDataset.rawLabels : statusDataset.labels;
    labels.forEach((label, idx) => {
      chartRows.push({
        Series: "Status distribution",
        Label: label,
        Value: statusDataset.datasets[0].data[idx] || 0,
        Percent: statusDataset.percentages[idx]?.toFixed(1) || "",
      });
    });
  }
  if (disciplineDataset.datasets[0]) {
    disciplineDataset.labels.forEach((label, idx) => {
      chartRows.push({
        Series: "Discipline mix",
        Label: label,
        Value: disciplineDataset.datasets[0].data[idx] || 0,
      });
    });
  }

  addSheet("Stats", statsRows);
  addSheet("Discipline Summary", disciplineRows);
  addSheet("Status Pivot", pivotRows);
  addSheet("Revision Timeline", revisionRows);
  addSheet("Open Reviews", openRows);
  addSheet("Latest Documents", view.latestDocs.map(recordRow));
  addSheet("Filtered Rows", view.filteredAll.map(recordRow));
  addSheet("Chart Data", chartRows);

  XLSX.writeFile(wb, "correspondence_analysis.xlsx");
}

function inferPhaseFlags(issueText) {
  const t = (issueText || "").toLowerCase();
  return {
    thirty: t.includes("30%"),
    sixty: t.includes("60%"),
    ninety: t.includes("90%"),
    ifc: t.includes("construction") || t.includes("ifc"),
    bd: t.includes("bd") || t.includes("basic design"),
  };
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
