const ADMIN_PASS = "admin123";

// Exact exam definitions matching Class_6_to_9_Student_Marks_Sheet.xlsx
const CORE_SUBJECTS = ["Tamil", "English", "Maths", "Science", "Social Science"];
const FULL_SUBJECTS = [
  "Tamil",
  "English",
  "Maths",
  "Science",
  "Social Science",
  "Value Education",
  "PT",
  "Computer Science",
  "General Knowledge"
];

const EXAMS = [
  { key: "mid1", label: "1st Mid term", short: "1st Mid", maxTotal: 500, subjects: CORE_SUBJECTS },
  { key: "quarterly", label: "Quarterly", short: "Quarterly", maxTotal: 900, subjects: FULL_SUBJECTS },
  { key: "mid2", label: "2nd Mid term", short: "2nd Mid", maxTotal: 500, subjects: CORE_SUBJECTS },
  { key: "halfYearly", label: "Half Yearly", short: "Half Yearly", maxTotal: 900, subjects: FULL_SUBJECTS },
  { key: "mid3", label: "3rd Mid term", short: "3rd Mid", maxTotal: 500, subjects: CORE_SUBJECTS },
  { key: "annually", label: "Annually", short: "Annually", maxTotal: 900, subjects: FULL_SUBJECTS }
];

let workbook = null;
let fileHandle = null;
let sheetData = {}; // { "Class 6": [studentObj, ...], "Class 7": [...], ... }

// Unified Scope State
let selectedClass = "ALL";
let selectedSection = "ALL";
let selectedExam = "ALL";     // "ALL" or exam key ("mid1", "quarterly", "mid2", "halfYearly", "mid3", "annually")
let selectedStatus = "ALL";   // "ALL" or "FAILED"
let isAdmin = false;
let hasUnsavedChanges = false;
let activeHistoryStudent = null;
let activeEditExamKey = "mid1";
let tempEditMarks = {};

// Chart Instances
let subjectBarChart = null;
let strengthDoughnutChart = null;
let gradeDistChart = null;
let studentHistoryChart = null;

// DOM Elements
const openExcelBtn = document.getElementById("openExcelBtn");
const emptyLoadBtn = document.getElementById("emptyLoadBtn");
const exportFailedTopBtn = document.getElementById("exportFailedTopBtn");
const exportFailedRecordsBtn = document.getElementById("exportFailedRecordsBtn");
const saveExcelBtn = document.getElementById("saveExcelBtn");
const adminToggleBtn = document.getElementById("adminToggleBtn");
const fileStatus = document.getElementById("fileStatus");

const emptyState = document.getElementById("emptyState");
const appLayout = document.getElementById("appLayout");

const sidebarClassList = document.getElementById("sidebarClassList");
const sidebarSectionList = document.getElementById("sidebarSectionList");
const sidebarExamList = document.getElementById("sidebarExamList");
const sidebarStatusList = document.getElementById("sidebarStatusList");
const currentScopeLabel = document.getElementById("currentScopeLabel");

const navDashboardBtn = document.getElementById("navDashboardBtn");
const navRecordsBtn = document.getElementById("navRecordsBtn");
const dashboardView = document.getElementById("dashboardView");
const recordsView = document.getElementById("recordsView");

const searchInput = document.getElementById("searchInput");
const recordsTableHead = document.getElementById("recordsTableHead");
const studentTableBody = document.getElementById("studentTableBody");

// Search Match Banner Elements
const quickSearchBanner = document.getElementById("quickSearchBanner");
const quickMatchTitle = document.getElementById("quickMatchTitle");
const quickMatchSummary = document.getElementById("quickMatchSummary");
const quickViewHistoryBtn = document.getElementById("quickViewHistoryBtn");
const quickExportPdfBtn = document.getElementById("quickExportPdfBtn");

// History Modal Elements
const historyModal = document.getElementById("historyModal");
const closeHistoryBtn = document.getElementById("closeHistoryBtn");
const downloadReportCardPdfBtn = document.getElementById("downloadReportCardPdfBtn");

// Admin & Edit Modal Elements
const loginModal = document.getElementById("loginModal");
const adminPassword = document.getElementById("adminPassword");
const loginError = document.getElementById("loginError");
const editModal = document.getElementById("editModal");
const studentForm = document.getElementById("studentForm");

// Event Listeners
openExcelBtn.addEventListener("click", connectExcelFile);
emptyLoadBtn.addEventListener("click", connectExcelFile);
exportFailedTopBtn.addEventListener("click", exportFailedStudentsExcel);
exportFailedRecordsBtn.addEventListener("click", exportFailedStudentsExcel);
saveExcelBtn.addEventListener("click", saveToExcelFile);

navDashboardBtn.addEventListener("click", () => switchMainView("dashboard"));
navRecordsBtn.addEventListener("click", () => switchMainView("records"));
searchInput.addEventListener("input", renderRecordsTable);

closeHistoryBtn.addEventListener("click", () => historyModal.classList.add("hidden"));
downloadReportCardPdfBtn.addEventListener("click", () => {
  if (activeHistoryStudent) exportStudentReportCardPDF(activeHistoryStudent);
});

adminToggleBtn.addEventListener("click", () => {
  if (isAdmin) {
    setAdminMode(false);
  } else {
    adminPassword.value = "";
    loginError.classList.add("hidden");
    loginModal.classList.remove("hidden");
    adminPassword.focus();
  }
});

document.getElementById("cancelLoginBtn").addEventListener("click", () => loginModal.classList.add("hidden"));
document.getElementById("confirmLoginBtn").addEventListener("click", verifyAdminLogin);
adminPassword.addEventListener("keydown", (e) => {
  if (e.key === "Enter") verifyAdminLogin();
});

document.getElementById("addStudentBtn").addEventListener("click", () => openStudentModal(null, null));
document.getElementById("cancelEditBtn").addEventListener("click", () => editModal.classList.add("hidden"));
document.getElementById("cancelEditTopBtn").addEventListener("click", () => editModal.classList.add("hidden"));
studentForm.addEventListener("submit", handleSaveStudentForm);

// =========================================================
// 1. PARSE MULTI-EXAM EXCEL WORKBOOK (2-ROW HEADERS)
// =========================================================
async function connectExcelFile() {
  try {
    let file;
    if ("showOpenFilePicker" in window) {
      [fileHandle] = await window.showOpenFilePicker({
        types: [{
          description: "Excel Files",
          accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] }
        }],
        multiple: false
      });
      file = await fileHandle.getFile();
    } else {
      file = await pickFileFallback();
    }

    const arrayBuffer = await file.arrayBuffer();
    workbook = XLSX.read(arrayBuffer, { type: "array" });

    sheetData = {};
    workbook.SheetNames.forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      sheetData[sheetName] = parseClassSheetMatrix(matrix, sheetName);
    });

    selectedClass = "ALL";
    selectedSection = "ALL";
    selectedExam = "ALL";
    selectedStatus = "ALL";
    hasUnsavedChanges = false;
    updateSaveButtonState();

    fileStatus.textContent = `Connected: ${file.name}`;
    fileStatus.className = "badge badge-ok";

    emptyState.classList.add("hidden");
    appLayout.classList.remove("hidden");
    exportFailedTopBtn.classList.remove("hidden");

    refreshAll();
  } catch (err) {
    if (err.name !== "AbortError") {
      alert("Could not load Excel file: " + err.message);
    }
  }
}

function parseClassSheetMatrix(matrix, className) {
  // Row 0: Merged Exam Group Headers ("1st Mid term", "Quarterly", etc.)
  // Row 1: Column Headers ("Student Name", "Roll No", "Section", subjects...)
  // Row 2+: Student Rows
  if (!matrix || matrix.length < 3) return [];

  const row0 = matrix[0];
  const row1 = matrix[1];

  // Map column indices to { examKey, subject }
  const colMap = [];
  let activeExamKey = null;

  for (let c = 3; c < row1.length; c++) {
    const groupHeader = String(row0[c] || "").trim();
    if (groupHeader) {
      const foundExam = EXAMS.find((e) => e.label.toLowerCase() === groupHeader.toLowerCase());
      activeExamKey = foundExam ? foundExam.key : null;
    }
    const subHeader = String(row1[c] || "").trim();
    if (activeExamKey && FULL_SUBJECTS.includes(subHeader)) {
      colMap.push({ colIndex: c, examKey: activeExamKey, subject: subHeader });
    }
  }

  const students = [];
  for (let r = 2; r < matrix.length; r++) {
    const row = matrix[r];
    const name = String(row[0] || "").trim();
    if (!name) continue;

    const examsObj = {};
    EXAMS.forEach((ex) => {
      examsObj[ex.key] = {};
      ex.subjects.forEach((sub) => {
        examsObj[ex.key][sub] = 0;
      });
    });

    colMap.forEach(({ colIndex, examKey, subject }) => {
      examsObj[examKey][subject] = Number(row[colIndex]) || 0;
    });

    students.push({
      _idx: students.length,
      _class: className,
      name,
      rollNo: Number(row[1]) || 0,
      section: String(row[2] || "").trim().toUpperCase(),
      exams: examsObj
    });
  }
  return students;
}

function pickFileFallback() {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx";
    input.onchange = () => resolve(input.files[0]);
    input.click();
  });
}

// =========================================================
// 2. SCORE & PASS/FAIL HELPERS
// =========================================================
function getExamStats(student, examKey) {
  const examDef = EXAMS.find((e) => e.key === examKey);
  const marksMap = student.exams[examKey] || {};
  let total = 0;
  const failedSubjects = [];

  examDef.subjects.forEach((sub) => {
    const m = Number(marksMap[sub]) || 0;
    total += m;
    if (m < 35) failedSubjects.push(`${sub} (${m})`);
  });

  return {
    total,
    maxTotal: examDef.maxTotal,
    pct: (total / examDef.maxTotal) * 100,
    failedSubjects,
    isPass: failedSubjects.length === 0
  };
}

function getCumulativeStats(student) {
  let grandTotal = 0;
  const grandMax = 4200; // 500 + 900 + 500 + 900 + 500 + 900
  const allFailures = [];

  EXAMS.forEach((ex) => {
    const st = getExamStats(student, ex.key);
    grandTotal += st.total;
    if (!st.isPass) {
      allFailures.push(`${ex.short}: ${st.failedSubjects.join(", ")}`);
    }
  });

  return {
    total: grandTotal,
    maxTotal: grandMax,
    pct: (grandTotal / grandMax) * 100,
    failedDetails: allFailures,
    isPass: allFailures.length === 0
  };
}

function getActiveScopeStats(student) {
  return selectedExam === "ALL"
    ? getCumulativeStats(student)
    : getExamStats(student, selectedExam);
}

function isStudentFailedInScope(student) {
  return !getActiveScopeStats(student).isPass;
}

// =========================================================
// 3. LEFT VERTICAL SIDEBAR & NAVIGATION
// =========================================================
function switchMainView(view) {
  if (view === "dashboard") {
    navDashboardBtn.classList.add("active");
    navRecordsBtn.classList.remove("active");
    dashboardView.classList.remove("hidden");
    recordsView.classList.add("hidden");
    renderDashboard();
  } else {
    navRecordsBtn.classList.add("active");
    navDashboardBtn.classList.remove("active");
    recordsView.classList.remove("hidden");
    dashboardView.classList.add("hidden");
    renderRecordsTable();
  }
}

function buildVerticalSidebar() {
  const allRows = Object.values(sheetData).flat();

  // A. Class Scope
  sidebarClassList.innerHTML = "";
  const allBtn = document.createElement("button");
  allBtn.className = `side-item ${selectedClass === "ALL" ? "active" : ""}`;
  allBtn.innerHTML = `<span>All Classes</span><span class="side-count">${allRows.length}</span>`;
  allBtn.onclick = () => {
    selectedClass = "ALL";
    selectedSection = "ALL";
    refreshAll();
  };
  sidebarClassList.appendChild(allBtn);

  Object.keys(sheetData).forEach((clsName) => {
    const btn = document.createElement("button");
    btn.className = `side-item ${selectedClass === clsName ? "active" : ""}`;
    btn.innerHTML = `<span>${clsName}</span><span class="side-count">${sheetData[clsName].length}</span>`;
    btn.onclick = () => {
      selectedClass = clsName;
      selectedSection = "ALL";
      refreshAll();
    };
    sidebarClassList.appendChild(btn);
  });

  // B. Sections
  sidebarSectionList.innerHTML = "";
  const classPool = selectedClass === "ALL" ? allRows : (sheetData[selectedClass] || []);
  const sectionCounts = {};
  classPool.forEach((s) => {
    sectionCounts[s.section] = (sectionCounts[s.section] || 0) + 1;
  });

  const allSecBtn = document.createElement("button");
  allSecBtn.className = `side-item ${selectedSection === "ALL" ? "active" : ""}`;
  allSecBtn.innerHTML = `<span>All Sections</span><span class="side-count">${classPool.length}</span>`;
  allSecBtn.onclick = () => {
    selectedSection = "ALL";
    refreshAll();
  };
  sidebarSectionList.appendChild(allSecBtn);

  Object.keys(sectionCounts).sort().forEach((sec) => {
    const secBtn = document.createElement("button");
    secBtn.className = `side-item ${selectedSection === sec ? "active" : ""}`;
    secBtn.innerHTML = `<span>Section ${sec}</span><span class="side-count">${sectionCounts[sec]}</span>`;
    secBtn.onclick = () => {
      selectedSection = sec;
      refreshAll();
    };
    sidebarSectionList.appendChild(secBtn);
  });

  // C. Exam Term Scope
  sidebarExamList.innerHTML = "";
  const allExamsBtn = document.createElement("button");
  allExamsBtn.className = `side-item ${selectedExam === "ALL" ? "active" : ""}`;
  allExamsBtn.innerHTML = `<span>All 6 Exams</span><span class="side-count">4200</span>`;
  allExamsBtn.onclick = () => {
    selectedExam = "ALL";
    refreshAll();
  };
  sidebarExamList.appendChild(allExamsBtn);

  EXAMS.forEach((ex) => {
    const exBtn = document.createElement("button");
    exBtn.className = `side-item ${selectedExam === ex.key ? "active" : ""}`;
    exBtn.innerHTML = `<span>${ex.label}</span><span class="side-count">${ex.maxTotal}</span>`;
    exBtn.onclick = () => {
      selectedExam = ex.key;
      refreshAll();
    };
    sidebarExamList.appendChild(exBtn);
  });

  // D. Result Status
  sidebarStatusList.innerHTML = "";
  const classSecPool = getClassAndSectionStudents();
  const failedPoolCount = classSecPool.filter(isStudentFailedInScope).length;

  const allStatusBtn = document.createElement("button");
  allStatusBtn.className = `side-item ${selectedStatus === "ALL" ? "active" : ""}`;
  allStatusBtn.innerHTML = `<span>All Students</span><span class="side-count">${classSecPool.length}</span>`;
  allStatusBtn.onclick = () => {
    selectedStatus = "ALL";
    refreshAll();
  };
  sidebarStatusList.appendChild(allStatusBtn);

  const failStatusBtn = document.createElement("button");
  failStatusBtn.className = `side-item fail-item ${selectedStatus === "FAILED" ? "active" : ""}`;
  failStatusBtn.innerHTML = `<span>⚠️ Failed Students</span><span class="side-count">${failedPoolCount}</span>`;
  failStatusBtn.onclick = () => {
    selectedStatus = "FAILED";
    refreshAll();
  };
  sidebarStatusList.appendChild(failStatusBtn);

  // Breadcrumb
  const clsText = selectedClass === "ALL" ? "All Classes" : selectedClass;
  const secText = selectedSection === "ALL" ? "All Sections" : `Section ${selectedSection}`;
  const exText = selectedExam === "ALL" ? "All 6 Exams" : EXAMS.find((e) => e.key === selectedExam).label;
  const stText = selectedStatus === "ALL" ? "All Students" : "Failed Only";
  currentScopeLabel.textContent = `${clsText} • ${secText} • ${exText} • ${stText}`;
}

function getClassAndSectionStudents() {
  const base = selectedClass === "ALL"
    ? Object.values(sheetData).flat()
    : (sheetData[selectedClass] || []);
  return selectedSection === "ALL" ? base : base.filter((s) => s.section === selectedSection);
}

function getScopedStudents() {
  const list = getClassAndSectionStudents();
  return selectedStatus === "FAILED" ? list.filter(isStudentFailedInScope) : list;
}

function refreshAll() {
  buildVerticalSidebar();
  renderDashboard();
  renderRecordsTable();
}

// =========================================================
// 4. ANALYTICS DASHBOARD
// =========================================================
function renderDashboard() {
  const students = getScopedStudents();
  const count = students.length;

  let grandTotal = 0;
  let passedCount = 0;
  let topper = null;

  const brackets = {
    "< 35%": 0,
    "35–49%": 0,
    "50–59%": 0,
    "60–74%": 0,
    "75–89%": 0,
    "90–100%": 0
  };

  students.forEach((st) => {
    const stats = getActiveScopeStats(st);
    grandTotal += stats.total;
    if (stats.isPass) passedCount++;

    if (!topper || stats.pct > topper.pct) {
      topper = {
        name: st.name,
        cls: st._class,
        sec: st.section,
        total: stats.total,
        max: stats.maxTotal,
        pct: stats.pct.toFixed(1)
      };
    }

    if (stats.pct < 35) brackets["< 35%"]++;
    else if (stats.pct < 50) brackets["35–49%"]++;
    else if (stats.pct < 60) brackets["50–59%"]++;
    else if (stats.pct < 75) brackets["60–74%"]++;
    else if (stats.pct < 90) brackets["75–89%"]++;
    else brackets["90–100%"]++;
  });

  const maxPerStudent = selectedExam === "ALL" ? 4200 : EXAMS.find((e) => e.key === selectedExam).maxTotal;
  const overallPct = count ? ((grandTotal / (count * maxPerStudent)) * 100).toFixed(2) : "0.00";
  const avgMarks = count ? (grandTotal / count).toFixed(1) : "0";
  const passRate = count ? ((passedCount / count) * 100).toFixed(1) : "0.0";

  document.getElementById("dashStrengthLabel").textContent =
    selectedStatus === "FAILED" ? "Failed Student Count" : "Total Strength";
  document.getElementById("dashStrength").textContent = count.toLocaleString();
  document.getElementById("dashStrengthSub").textContent =
    selectedClass === "ALL" ? `Across ${Object.keys(sheetData).length} classes` : `${selectedClass} (${selectedSection === "ALL" ? "All Sections" : "Sec " + selectedSection})`;

  document.getElementById("dashOverallPct").textContent = `${overallPct}%`;
  document.getElementById("dashAvgMarksSub").textContent = `Avg Total: ${avgMarks} / ${maxPerStudent}`;

  document.getElementById("dashPassRate").textContent = `${passRate}%`;
  document.getElementById("dashPassCountSub").textContent = `${passedCount} of ${count} passed all subjects`;

  document.getElementById("dashTopperLabel").textContent =
    selectedStatus === "FAILED" ? "Highest Among Failed" : "Top Performer";
  if (topper) {
    document.getElementById("dashTopperName").textContent = topper.name;
    document.getElementById("dashTopperScore").textContent =
      `${topper.total}/${topper.max} (${topper.pct}%) • ${topper.cls}-${topper.sec}`;
  } else {
    document.getElementById("dashTopperName").textContent = "—";
    document.getElementById("dashTopperScore").textContent = "No records";
  }

  // Build Metrics Breakdown (All 6 Exams comparison when selectedExam === "ALL", or Subject-wise when specific exam selected)
  const metricsBody = document.getElementById("subjectMetricsBody");
  metricsBody.innerHTML = "";
  const chartLabels = [];
  const chartAvgs = [];
  const chartPassRates = [];

  if (selectedExam === "ALL") {
    document.getElementById("subjectChartTitle").textContent = "Exam-Wise Average Percentage & Pass Rate Across All 6 Terms";
    document.getElementById("metricsTableTitle").textContent = "Exam-Wise Detailed Performance Metrics";
    document.getElementById("metricsColHeader").textContent = "Exam Term";

    EXAMS.forEach((ex) => {
      let sum = 0;
      let max = -Infinity;
      let min = Infinity;
      let pass = 0;
      let fail = 0;

      students.forEach((st) => {
        const es = getExamStats(st, ex.key);
        sum += es.total;
        if (es.total > max) max = es.total;
        if (es.total < min) min = es.total;
        if (es.isPass) pass++;
        else fail++;
      });

      const avgPct = count ? ((sum / (count * ex.maxTotal)) * 100).toFixed(1) : "0.0";
      const avgRaw = count ? (sum / count).toFixed(1) : "0.0";
      const pRate = count ? ((pass / count) * 100).toFixed(1) : "0.0";

      chartLabels.push(ex.label);
      chartAvgs.push(Number(avgPct));
      chartPassRates.push(Number(pRate));

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong class="peacock-text">${ex.label}</strong></td>
        <td>${avgRaw} / ${ex.maxTotal} (${avgPct}%)</td>
        <td>${count ? max : "—"}</td>
        <td>${count ? min : "—"}</td>
        <td>${pass}</td>
        <td class="${fail > 0 ? "mark-fail" : ""}">${fail}</td>
        <td><span class="pill-pass">${pRate}%</span></td>
      `;
      metricsBody.appendChild(tr);
    });
  } else {
    const examDef = EXAMS.find((e) => e.key === selectedExam);
    document.getElementById("subjectChartTitle").textContent = `${examDef.label} — Subject-Wise Average & Pass Rate (%)`;
    document.getElementById("metricsTableTitle").textContent = `${examDef.label} — Subject-Wise Metrics (${examDef.subjects.length} Subjects)`;
    document.getElementById("metricsColHeader").textContent = "Subject";

    examDef.subjects.forEach((sub) => {
      let sum = 0;
      let max = -Infinity;
      let min = Infinity;
      let pass = 0;
      let fail = 0;

      students.forEach((st) => {
        const val = Number(st.exams[selectedExam][sub]) || 0;
        sum += val;
        if (val > max) max = val;
        if (val < min) min = val;
        if (val >= 35) pass++;
        else fail++;
      });

      const avg = count ? (sum / count).toFixed(1) : "0.0";
      const pRate = count ? ((pass / count) * 100).toFixed(1) : "0.0";

      chartLabels.push(sub);
      chartAvgs.push(Number(avg));
      chartPassRates.push(Number(pRate));

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong class="peacock-text">${sub}</strong></td>
        <td>${avg} / 100</td>
        <td>${count ? max : "—"}</td>
        <td>${count ? min : "—"}</td>
        <td>${pass}</td>
        <td class="${fail > 0 ? "mark-fail" : ""}">${fail}</td>
        <td><span class="pill-pass">${pRate}%</span></td>
      `;
      metricsBody.appendChild(tr);
    });
  }

  updateSubjectChart(chartLabels, chartAvgs, chartPassRates);
  updateStrengthChart(students);
  updateGradeChart(brackets);
}

function updateSubjectChart(labels, avgs, passRates) {
  const ctx = document.getElementById("subjectBarChart").getContext("2d");
  if (subjectBarChart) subjectBarChart.destroy();

  subjectBarChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Average Score (%)",
          data: avgs,
          backgroundColor: "#02c39a",
          borderRadius: 6
        },
        {
          label: "Pass Rate (%)",
          data: passRates,
          backgroundColor: "#eab308",
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#ccfbf1" } } },
      scales: {
        x: { ticks: { color: "#99c2bd" }, grid: { color: "rgba(45,212,191,0.08)" } },
        y: { beginAtZero: true, max: 100, ticks: { color: "#99c2bd" }, grid: { color: "rgba(45,212,191,0.08)" } }
      }
    }
  });
}

function updateStrengthChart(students) {
  const ctx = document.getElementById("strengthDoughnutChart").getContext("2d");
  if (strengthDoughnutChart) strengthDoughnutChart.destroy();

  const titleEl = document.getElementById("strengthChartTitle");
  const counts = {};

  if (selectedClass === "ALL") {
    titleEl.textContent = selectedStatus === "FAILED" ? "Failed Students by Class" : "Student Strength by Class";
    students.forEach((s) => {
      counts[s._class] = (counts[s._class] || 0) + 1;
    });
  } else {
    titleEl.textContent = selectedStatus === "FAILED"
      ? `Failed Students by Section (${selectedClass})`
      : `Student Strength by Section (${selectedClass})`;
    students.forEach((s) => {
      const label = `Section ${s.section}`;
      counts[label] = (counts[label] || 0) + 1;
    });
  }

  strengthDoughnutChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: Object.keys(counts),
      datasets: [{
        data: Object.values(counts),
        backgroundColor: ["#02c39a", "#00a896", "#05668d", "#eab308", "#2dd4bf"],
        borderColor: "#082020",
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { color: "#ccfbf1" } } }
    }
  });
}

function updateGradeChart(brackets) {
  const ctx = document.getElementById("gradeDistChart").getContext("2d");
  if (gradeDistChart) gradeDistChart.destroy();

  gradeDistChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: Object.keys(brackets),
      datasets: [{
        label: "Students",
        data: Object.values(brackets),
        backgroundColor: ["#ef4444", "#f59e0b", "#14b8a6", "#00a896", "#02c39a", "#eab308"],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#99c2bd" }, grid: { color: "rgba(45,212,191,0.08)" } },
        y: { beginAtZero: true, ticks: { precision: 0, color: "#99c2bd" }, grid: { color: "rgba(45,212,191,0.08)" } }
      }
    }
  });
}

// =========================================================
// 5. STUDENT RECORDS TABLE & SEARCH HISTORY BANNER
// =========================================================
function renderRecordsTable() {
  const baseRows = getScopedStudents();
  const query = searchInput.value.trim().toLowerCase();

  const filtered = baseRows.filter((row) => {
    if (!query) return true;
    return (
      row.name.toLowerCase().includes(query) ||
      String(row.rollNo).toLowerCase().includes(query)
    );
  });

  // Show quick Student History & Report Card Banner when searching
  if (query && filtered.length > 0) {
    const topMatch = filtered[0];
    const cum = getCumulativeStats(topMatch);
    quickSearchBanner.classList.remove("hidden");
    quickMatchTitle.textContent = `${topMatch.name} — ${topMatch._class} (Section ${topMatch.section}) • Roll No: ${topMatch.rollNo}`;
    quickMatchSummary.textContent = `Cumulative Score: ${cum.total} / 4200 (${cum.pct.toFixed(2)}%) • Overall Status: ${cum.isPass ? "ALL CLEAR (PASS)" : "NEEDS IMPROVEMENT (FAILED)"}`;
    quickViewHistoryBtn.onclick = () => openStudentHistoryModal(topMatch);
    quickExportPdfBtn.onclick = () => exportStudentReportCardPDF(topMatch);
  } else {
    quickSearchBanner.classList.add("hidden");
  }

  document.getElementById("statCount").textContent = filtered.length;
  let pctSum = 0;
  let passCount = 0;

  // Dynamic Table Headers based on Selected Exam
  if (selectedExam === "ALL") {
    recordsTableHead.innerHTML = `
      <tr>
        <th>Class</th>
        <th>Roll No</th>
        <th>Name</th>
        <th>Sec</th>
        <th>1st Mid<br><small>(500)</small></th>
        <th>Quarterly<br><small>(900)</small></th>
        <th>2nd Mid<br><small>(500)</small></th>
        <th>Half Yearly<br><small>(900)</small></th>
        <th>3rd Mid<br><small>(500)</small></th>
        <th>Annually<br><small>(900)</small></th>
        <th>Grand Total<br><small>(4200)</small></th>
        <th>Overall %</th>
        <th>Status</th>
        <th>Report Card / Action</th>
      </tr>
    `;
  } else {
    const exDef = EXAMS.find((e) => e.key === selectedExam);
    const subHeaders = exDef.subjects.map((s) => `<th>${s}</th>`).join("");
    recordsTableHead.innerHTML = `
      <tr>
        <th>Class</th>
        <th>Roll No</th>
        <th>Name</th>
        <th>Sec</th>
        ${subHeaders}
        <th>Total<br><small>(${exDef.maxTotal})</small></th>
        <th>%</th>
        <th>Status</th>
        <th>Report Card / Action</th>
      </tr>
    `;
  }

  studentTableBody.innerHTML = "";

  filtered.forEach((student) => {
    const scopeStats = getActiveScopeStats(student);
    pctSum += scopeStats.pct;
    if (scopeStats.isPass) passCount++;

    const tr = document.createElement("tr");
    const editBtnHtml = isAdmin
      ? `<button class="btn btn-outline-peacock btn-sm" onclick="openStudentModal('${student._class}', ${student._idx})">Edit</button>`
      : "";

    if (selectedExam === "ALL") {
      const m1 = getExamStats(student, "mid1");
      const q = getExamStats(student, "quarterly");
      const m2 = getExamStats(student, "mid2");
      const hy = getExamStats(student, "halfYearly");
      const m3 = getExamStats(student, "mid3");
      const an = getExamStats(student, "annually");

      tr.innerHTML = `
        <td><span class="peacock-text">${student._class}</span></td>
        <td>${student.rollNo}</td>
        <td><strong>${student.name}</strong></td>
        <td>${student.section}</td>
        <td class="${!m1.isPass ? "mark-fail" : ""}">${m1.total}</td>
        <td class="${!q.isPass ? "mark-fail" : ""}">${q.total}</td>
        <td class="${!m2.isPass ? "mark-fail" : ""}">${m2.total}</td>
        <td class="${!hy.isPass ? "mark-fail" : ""}">${hy.total}</td>
        <td class="${!m3.isPass ? "mark-fail" : ""}">${m3.total}</td>
        <td class="${!an.isPass ? "mark-fail" : ""}">${an.total}</td>
        <td><strong>${scopeStats.total}</strong></td>
        <td><strong>${scopeStats.pct.toFixed(1)}%</strong></td>
        <td><span class="${scopeStats.isPass ? "pill-pass" : "pill-fail"}">${scopeStats.isPass ? "PASS" : "FAIL"}</span></td>
        <td class="action-cell">
          <button class="btn btn-peacock btn-sm" onclick="openStudentHistoryByRef('${student._class}', ${student._idx})">📜 History</button>
          ${editBtnHtml}
        </td>
      `;
    } else {
      const exDef = EXAMS.find((e) => e.key === selectedExam);
      const subCells = exDef.subjects
        .map((sub) => {
          const val = Number(student.exams[selectedExam][sub]) || 0;
          return `<td class="${val < 35 ? "mark-fail" : ""}">${val}</td>`;
        })
        .join("");

      tr.innerHTML = `
        <td><span class="peacock-text">${student._class}</span></td>
        <td>${student.rollNo}</td>
        <td><strong>${student.name}</strong></td>
        <td>${student.section}</td>
        ${subCells}
        <td><strong>${scopeStats.total}</strong></td>
        <td><strong>${scopeStats.pct.toFixed(1)}%</strong></td>
        <td><span class="${scopeStats.isPass ? "pill-pass" : "pill-fail"}">${scopeStats.isPass ? "PASS" : "FAIL"}</span></td>
        <td class="action-cell">
          <button class="btn btn-peacock btn-sm" onclick="openStudentHistoryByRef('${student._class}', ${student._idx})">📜 History</button>
          ${editBtnHtml}
        </td>
      `;
    }

    studentTableBody.appendChild(tr);
  });

  const avgPercent = filtered.length ? (pctSum / filtered.length).toFixed(1) : "0.0";
  const passRate = filtered.length ? ((passCount / filtered.length) * 100).toFixed(1) : "0.0";

  document.getElementById("statAvg").textContent = `${avgPercent}%`;
  document.getElementById("statPass").textContent = `${passRate}%`;
}

// =========================================================
// 6. COMPLETE STUDENT HISTORY MODAL & PDF REPORT CARD EXPORT
// =========================================================
window.openStudentHistoryByRef = function (clsName, idx) {
  const st = sheetData[clsName].find((s) => s._idx === idx);
  if (st) openStudentHistoryModal(st);
};

function openStudentHistoryModal(student) {
  activeHistoryStudent = student;
  const cum = getCumulativeStats(student);

  document.getElementById("histClassBadge").textContent = `${student._class} • Section ${student.section}`;
  document.getElementById("histStudentName").textContent = student.name;
  document.getElementById("histRollInfo").textContent =
    `Roll No: ${student.rollNo} • Cumulative Total: ${cum.total} / 4200 (${cum.pct.toFixed(2)}%) • Overall Result: ${cum.isPass ? "PASS" : "FAIL"}`;

  // Populate 6-Exam KPI Strip
  const kpiStrip = document.getElementById("histKpiStrip");
  kpiStrip.innerHTML = "";
  const examPcts = [];

  EXAMS.forEach((ex) => {
    const es = getExamStats(student, ex.key);
    examPcts.push(Number(es.pct.toFixed(1)));
    const box = document.createElement("div");
    box.className = "hist-kpi-box";
    box.innerHTML = `
      <span>${ex.label}</span>
      <strong>${es.total} / ${ex.maxTotal} (${es.pct.toFixed(1)}%)</strong>
      <small class="${es.isPass ? "peacock-text" : "mark-fail"}">${es.isPass ? "PASS" : "FAIL"}</small>
    `;
    kpiStrip.appendChild(box);
  });

  // Render Progression Line Chart
  const ctx = document.getElementById("studentHistoryChart").getContext("2d");
  if (studentHistoryChart) studentHistoryChart.destroy();

  studentHistoryChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: EXAMS.map((e) => e.label),
      datasets: [{
        label: "Exam Percentage (%)",
        data: examPcts,
        borderColor: "#02c39a",
        backgroundColor: "rgba(2, 195, 154, 0.18)",
        borderWidth: 3,
        pointBackgroundColor: "#eab308",
        pointRadius: 5,
        fill: true,
        tension: 0.3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#ccfbf1" }, grid: { color: "rgba(45,212,191,0.1)" } },
        y: { beginAtZero: true, max: 100, ticks: { color: "#ccfbf1" }, grid: { color: "rgba(45,212,191,0.1)" } }
      }
    }
  });

  // Render Complete 9-Subject x 6-Exam Matrix
  const matrixBody = document.getElementById("historyMatrixBody");
  matrixBody.innerHTML = "";

  FULL_SUBJECTS.forEach((sub) => {
    let subSum = 0;
    let subCount = 0;

    const cells = EXAMS.map((ex) => {
      if (!ex.subjects.includes(sub)) {
        return `<td style="color:#64748b;">—</td>`;
      }
      const val = Number(student.exams[ex.key][sub]) || 0;
      subSum += val;
      subCount++;
      return `<td class="${val < 35 ? "mark-fail" : ""}">${val}</td>`;
    }).join("");

    const avg = subCount ? (subSum / subCount).toFixed(1) : "—";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong class="peacock-text">${sub}</strong></td>
      ${cells}
      <td><strong>${avg}%</strong></td>
    `;
    matrixBody.appendChild(tr);
  });

  // Add Totals Row
  const totalCells = EXAMS.map((ex) => {
    const es = getExamStats(student, ex.key);
    return `<td><strong>${es.total}/${ex.maxTotal}</strong></td>`;
  }).join("");

  const totalTr = document.createElement("tr");
  totalTr.style.background = "rgba(2, 195, 154, 0.12)";
  totalTr.innerHTML = `
    <td><strong style="color:#fde047;">EXAM TOTAL</strong></td>
    ${totalCells}
    <td><strong style="color:#fde047;">${cum.pct.toFixed(1)}%</strong></td>
  `;
  matrixBody.appendChild(totalTr);

  historyModal.classList.remove("hidden");
}

function exportStudentReportCardPDF(student) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "mm", "a4");
  const cum = getCumulativeStats(student);

  // Header Banner (Peacock Green)
  doc.setFillColor(4, 57, 57);
  doc.rect(0, 0, 210, 34, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.text("ACADEMIC PROGRESS REPORT CARD", 14, 15);

  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(180, 245, 235);
  doc.text("Complete Multi-Term Examination Transcript (1st Mid Term to Annual Examination)", 14, 23);

  // Student Info Box
  doc.setDrawColor(0, 168, 150);
  doc.setFillColor(240, 253, 250);
  doc.roundedRect(14, 40, 182, 26, 2, 2, "FD");

  doc.setTextColor(15, 23, 42);
  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.text(`Student Name: ${student.name}`, 18, 49);
  doc.text(`Class & Section: ${student._class} - Sec ${student.section}`, 18, 57);

  doc.text(`Roll Number: ${student.rollNo}`, 125, 49);
  doc.text(`Cumulative Score: ${cum.total} / 4200 (${cum.pct.toFixed(2)}%)`, 125, 57);

  // Table 1: Subject-Wise Marks Across All 6 Exams
  const tableHead = [[
    "Subject",
    "1st Mid\n(100)",
    "Quarterly\n(100)",
    "2nd Mid\n(100)",
    "Half Yearly\n(100)",
    "3rd Mid\n(100)",
    "Annually\n(100)",
    "Avg %"
  ]];

  const tableBody = FULL_SUBJECTS.map((sub) => {
    let sum = 0;
    let cnt = 0;
    const row = [sub];
    EXAMS.forEach((ex) => {
      if (!ex.subjects.includes(sub)) {
        row.push("—");
      } else {
        const m = Number(student.exams[ex.key][sub]) || 0;
        sum += m;
        cnt++;
        row.push(m < 35 ? `${m} (F)` : String(m));
      }
    });
    row.push(cnt ? `${(sum / cnt).toFixed(1)}%` : "—");
    return row;
  });

  // Append Exam Totals & Percentages Rows
  const totalsRow = ["EXAM TOTAL"];
  const pctRow = ["PERCENTAGE"];
  const statusRow = ["RESULT STATUS"];

  EXAMS.forEach((ex) => {
    const es = getExamStats(student, ex.key);
    totalsRow.push(`${es.total}/${ex.maxTotal}`);
    pctRow.push(`${es.pct.toFixed(1)}%`);
    statusRow.push(es.isPass ? "PASS" : "FAIL");
  });
  totalsRow.push(`${cum.total}/4200`);
  pctRow.push(`${cum.pct.toFixed(2)}%`);
  statusRow.push(cum.isPass ? "PASS" : "FAIL");

  tableBody.push(totalsRow, pctRow, statusRow);

  doc.autoTable({
    startY: 72,
    head: tableHead,
    body: tableBody,
    theme: "grid",
    headStyles: {
      fillColor: [0, 128, 115],
      textColor: [255, 255, 255],
      fontStyle: "bold",
      halign: "center"
    },
    styles: {
      fontSize: 9,
      cellPadding: 3,
      halign: "center"
    },
    columnStyles: {
      0: { halign: "left", fontStyle: "bold", cellWidth: 42 }
    },
    didParseCell: function (data) {
      const txt = String(data.cell.raw || "");
      if (txt.includes("(F)") || txt === "FAIL") {
        data.cell.styles.textColor = [220, 38, 38];
        data.cell.styles.fontStyle = "bold";
      }
      if (data.row.index >= FULL_SUBJECTS.length) {
        data.cell.styles.fillColor = [230, 247, 245];
        data.cell.styles.fontStyle = "bold";
      }
    }
  });

  // Remarks & Signatures Footer
  const finalY = doc.lastAutoTable.finalY + 12;
  doc.setFontSize(10);
  doc.setTextColor(30, 41, 59);
  doc.setFont("helvetica", "bold");
  doc.text("Overall Academic Remarks:", 14, finalY);

  doc.setFont("helvetica", "normal");
  const remarkText = cum.isPass
    ? `Congratulations! ${student.name} has successfully passed all subjects across all 6 examinations with an aggregate of ${cum.pct.toFixed(2)}%.`
    : `Attention Required: ${student.name} needs remedial support in the subjects marked (F) above. Overall aggregate: ${cum.pct.toFixed(2)}%.`;
  doc.text(doc.splitTextToSize(remarkText, 180), 14, finalY + 6);

  const sigY = finalY + 28;
  doc.setDrawColor(100, 116, 139);
  doc.line(14, sigY, 64, sigY);
  doc.line(82, sigY, 132, sigY);
  doc.line(148, sigY, 196, sigY);

  doc.setFontSize(9);
  doc.text("Class Teacher Signature", 19, sigY + 5);
  doc.text("Principal Signature", 92, sigY + 5);
  doc.text("Parent / Guardian Signature", 151, sigY + 5);

  const safeName = student.name.replace(/\s+/g, "_");
  doc.save(`Report_Card_${student._class.replace(/\s+/g, "_")}_${student.rollNo}_${safeName}.pdf`);
}

// =========================================================
// 7. EXPORT FAILED STUDENTS TO NEW EXCEL FILE
// =========================================================
function exportFailedStudentsExcel() {
  if (!workbook) return;

  const scopedPool = getClassAndSectionStudents();
  const failedStudents = scopedPool.filter(isStudentFailedInScope);

  if (failedStudents.length === 0) {
    alert("No failed students found in the currently selected scope.");
    return;
  }

  const newWb = XLSX.utils.book_new();
  const headers = [
    "Class",
    "Roll No",
    "Student Name",
    "Section",
    "1st Mid (500)",
    "Quarterly (900)",
    "2nd Mid (500)",
    "Half Yearly (900)",
    "3rd Mid (500)",
    "Annually (900)",
    "Grand Total (4200)",
    "Overall %",
    "Failed Subjects"
  ];

  const formatRow = (st) => {
    const cum = getCumulativeStats(st);
    const scopeSt = getActiveScopeStats(st);
    const failText = selectedExam === "ALL"
      ? cum.failedDetails.join(" | ")
      : scopeSt.failedSubjects.join(", ");

    return {
      "Class": st._class,
      "Roll No": st.rollNo,
      "Student Name": st.name,
      "Section": st.section,
      "1st Mid (500)": getExamStats(st, "mid1").total,
      "Quarterly (900)": getExamStats(st, "quarterly").total,
      "2nd Mid (500)": getExamStats(st, "mid2").total,
      "Half Yearly (900)": getExamStats(st, "halfYearly").total,
      "3rd Mid (500)": getExamStats(st, "mid3").total,
      "Annually (900)": getExamStats(st, "annually").total,
      "Grand Total (4200)": cum.total,
      "Overall %": `${cum.pct.toFixed(2)}%`,
      "Failed Subjects": failText
    };
  };

  if (selectedClass === "ALL") {
    const allWs = XLSX.utils.json_to_sheet(failedStudents.map(formatRow), { header: headers });
    XLSX.utils.book_append_sheet(newWb, allWs, "All Failed Students");

    Object.keys(sheetData).forEach((clsName) => {
      const clsFailed = failedStudents.filter((s) => s._class === clsName).map(formatRow);
      if (clsFailed.length > 0) {
        const ws = XLSX.utils.json_to_sheet(clsFailed, { header: headers });
        XLSX.utils.book_append_sheet(newWb, ws, `${clsName} Failed`);
      }
    });
  } else {
    const ws = XLSX.utils.json_to_sheet(failedStudents.map(formatRow), { header: headers });
    const title = `${selectedClass}_${selectedSection === "ALL" ? "All" : "Sec_" + selectedSection}`.slice(0, 31);
    XLSX.utils.book_append_sheet(newWb, ws, title);
  }

  const clsPart = selectedClass === "ALL" ? "All_Classes" : selectedClass.replace(/\s+/g, "_");
  const secPart = selectedSection === "ALL" ? "All_Sections" : `Sec_${selectedSection}`;
  const exPart = selectedExam === "ALL" ? "All_Exams" : selectedExam;
  XLSX.writeFile(newWb, `Failed_Students_${clsPart}_${secPart}_${exPart}.xlsx`);
}

// =========================================================
// 8. ADMIN AUTHENTICATION, MULTI-EXAM EDIT & SAVE TO EXCEL
// =========================================================
function verifyAdminLogin() {
  if (adminPassword.value === ADMIN_PASS) {
    loginModal.classList.add("hidden");
    setAdminMode(true);
    switchMainView("records");
  } else {
    loginError.classList.remove("hidden");
  }
}

function setAdminMode(enabled) {
  isAdmin = enabled;
  adminToggleBtn.textContent = enabled ? "🔓 Exit Admin" : "🔒 Admin Login";
  document.getElementById("statMode").textContent = enabled ? "Admin (Edit Enabled)" : "View Only";
  document.getElementById("adminBanner").classList.toggle("hidden", !enabled);
  updateSaveButtonState();
  renderRecordsTable();
}

function updateSaveButtonState() {
  if (isAdmin && workbook) {
    saveExcelBtn.classList.remove("hidden");
    saveExcelBtn.textContent = hasUnsavedChanges ? "💾 Save to Excel *" : "💾 Save to Excel";
  } else {
    saveExcelBtn.classList.add("hidden");
  }
}

window.openStudentModal = function (clsName, rowIndex) {
  if (!isAdmin) return;
  const isNew = rowIndex === null;

  const classSelectWrap = document.getElementById("classSelectWrap");
  const fieldClass = document.getElementById("fieldClass");

  if (isNew) {
    document.getElementById("editModalTitle").textContent = "Add New Student (All 6 Exams)";
    document.getElementById("editRowIndex").value = "";
    document.getElementById("editRowClass").value = "";
    studentForm.reset();

    classSelectWrap.classList.remove("hidden");
    fieldClass.innerHTML = "";
    Object.keys(sheetData).forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      if (s === selectedClass) opt.selected = true;
      fieldClass.appendChild(opt);
    });
    if (selectedSection !== "ALL") document.getElementById("fieldSection").value = selectedSection;

    tempEditMarks = {};
    EXAMS.forEach((ex) => {
      tempEditMarks[ex.key] = {};
      ex.subjects.forEach((sub) => { tempEditMarks[ex.key][sub] = 0; });
    });
  } else {
    classSelectWrap.classList.add("hidden");
    document.getElementById("editModalTitle").textContent = `Edit Student (${clsName})`;
    document.getElementById("editRowIndex").value = rowIndex;
    document.getElementById("editRowClass").value = clsName;

    const student = sheetData[clsName].find((s) => s._idx === rowIndex);
    document.getElementById("fieldRoll").value = student.rollNo;
    document.getElementById("fieldName").value = student.name;
    document.getElementById("fieldSection").value = student.section;

    tempEditMarks = JSON.parse(JSON.stringify(student.exams));
  }

  activeEditExamKey = selectedExam !== "ALL" ? selectedExam : "mid1";
  renderEditExamTabs();
  renderEditExamFields();
  editModal.classList.remove("hidden");
};

function renderEditExamTabs() {
  const container = document.getElementById("editExamTabs");
  container.innerHTML = "";
  EXAMS.forEach((ex) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `exam-edit-pill ${activeEditExamKey === ex.key ? "active" : ""}`;
    btn.textContent = ex.label;
    btn.onclick = () => {
      captureCurrentEditExamInputs();
      activeEditExamKey = ex.key;
      renderEditExamTabs();
      renderEditExamFields();
    };
    container.appendChild(btn);
  });
}

function renderEditExamFields() {
  const container = document.getElementById("editExamFieldsContainer");
  container.innerHTML = "";
  const examDef = EXAMS.find((e) => e.key === activeEditExamKey);

  examDef.subjects.forEach((sub) => {
    const val = tempEditMarks[activeEditExamKey][sub] ?? 0;
    const label = document.createElement("label");
    label.innerHTML = `${sub} (0–100) <input type="number" data-sub="${sub}" min="0" max="100" value="${val}" required />`;
    container.appendChild(label);
  });
}

function captureCurrentEditExamInputs() {
  const inputs = document.querySelectorAll("#editExamFieldsContainer input[data-sub]");
  inputs.forEach((inp) => {
    const sub = inp.getAttribute("data-sub");
    tempEditMarks[activeEditExamKey][sub] = Number(inp.value) || 0;
  });
}

function handleSaveStudentForm(e) {
  e.preventDefault();
  captureCurrentEditExamInputs();

  const rowIndexVal = document.getElementById("editRowIndex").value;
  const targetClass = rowIndexVal === ""
    ? document.getElementById("fieldClass").value
    : document.getElementById("editRowClass").value;

  const updatedData = {
    name: document.getElementById("fieldName").value.trim(),
    section: document.getElementById("fieldSection").value.trim().toUpperCase(),
    rollNo: Number(document.getElementById("fieldRoll").value),
    exams: JSON.parse(JSON.stringify(tempEditMarks))
  };

  if (rowIndexVal === "") {
    const newIdx = sheetData[targetClass].length;
    sheetData[targetClass].push({ ...updatedData, _idx: newIdx, _class: targetClass });
  } else {
    const idx = Number(rowIndexVal);
    const target = sheetData[targetClass].find((s) => s._idx === idx);
    Object.assign(target, updatedData);
  }

  hasUnsavedChanges = true;
  updateSaveButtonState();
  refreshAll();
  editModal.classList.add("hidden");
}

// Write updated rows back into the exact 52-column 2-row merged structure of Class_6_to_9_Student_Marks_Sheet.xlsx
async function saveToExcelFile() {
  if (!workbook) return;

  Object.keys(sheetData).forEach((sheetName) => {
    const students = sheetData[sheetName];
    const row0 = [
      `${sheetName.toUpperCase()} — STUDENT INFO`, "", "",
      "1st Mid term", "", "", "", "",
      "Quarterly", "", "", "", "", "", "", "", "",
      "2nd Mid term", "", "", "", "",
      "Half Yearly", "", "", "", "", "", "", "", "",
      "3rd Mid term", "", "", "", "",
      "Annually", "", "", "", "", "", "", "", "",
      "EXAM TOTALS & OVERALL %", "", "", "", "", "", ""
    ];

    const row1 = [
      "Student Name", "Roll No", "Section",
      ...CORE_SUBJECTS,
      ...FULL_SUBJECTS,
      ...CORE_SUBJECTS,
      ...FULL_SUBJECTS,
      ...CORE_SUBJECTS,
      ...FULL_SUBJECTS,
      "1st Mid Term Total (500)",
      "Quarterly Total (900)",
      "2nd Mid Term Total (500)",
      "Half Yearly Total (900)",
      "3rd Mid Term Total (500)",
      "Annual Total (900)",
      "Overall Avg %"
    ];

    const aoa = [row0, row1];

    students.forEach((st, i) => {
      const rNum = i + 3; // 1-indexed Excel row
      const m1Vals = CORE_SUBJECTS.map((s) => Number(st.exams.mid1[s]) || 0);
      const qVals = FULL_SUBJECTS.map((s) => Number(st.exams.quarterly[s]) || 0);
      const m2Vals = CORE_SUBJECTS.map((s) => Number(st.exams.mid2[s]) || 0);
      const hyVals = FULL_SUBJECTS.map((s) => Number(st.exams.halfYearly[s]) || 0);
      const m3Vals = CORE_SUBJECTS.map((s) => Number(st.exams.mid3[s]) || 0);
      const anVals = FULL_SUBJECTS.map((s) => Number(st.exams.annually[s]) || 0);

      const m1Sum = m1Vals.reduce((a, b) => a + b, 0);
      const qSum = qVals.reduce((a, b) => a + b, 0);
      const m2Sum = m2Vals.reduce((a, b) => a + b, 0);
      const hySum = hyVals.reduce((a, b) => a + b, 0);
      const m3Sum = m3Vals.reduce((a, b) => a + b, 0);
      const anSum = anVals.reduce((a, b) => a + b, 0);
      const avgPct = Number(((m1Sum + qSum + m2Sum + hySum + m3Sum + anSum) / 4200).toFixed(4));

      aoa.push([
        st.name, st.rollNo, st.section,
        ...m1Vals,
        ...qVals,
        ...m2Vals,
        ...hyVals,
        ...m3Vals,
        ...anVals,
        { t: "n", v: m1Sum, f: `SUM(D${rNum}:H${rNum})` },
        { t: "n", v: qSum, f: `SUM(I${rNum}:Q${rNum})` },
        { t: "n", v: m2Sum, f: `SUM(R${rNum}:V${rNum})` },
        { t: "n", v: hySum, f: `SUM(W${rNum}:AE${rNum})` },
        { t: "n", v: m3Sum, f: `SUM(AF${rNum}:AJ${rNum})` },
        { t: "n", v: anSum, f: `SUM(AK${rNum}:AS${rNum})` },
        { t: "n", v: avgPct, f: `ROUND(SUM(AT${rNum}:AY${rNum})/4200, 4)` }
      ]);
    });

    const newWs = XLSX.utils.aoa_to_sheet(aoa);
    newWs["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
      { s: { r: 0, c: 3 }, e: { r: 0, c: 7 } },
      { s: { r: 0, c: 8 }, e: { r: 0, c: 16 } },
      { s: { r: 0, c: 17 }, e: { r: 0, c: 21 } },
      { s: { r: 0, c: 22 }, e: { r: 0, c: 30 } },
      { s: { r: 0, c: 31 }, e: { r: 0, c: 35 } },
      { s: { r: 0, c: 36 }, e: { r: 0, c: 44 } },
      { s: { r: 0, c: 45 }, e: { r: 0, c: 51 } }
    ];
    workbook.Sheets[sheetName] = newWs;
  });

  const wbOut = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

  try {
    if (fileHandle) {
      const writable = await fileHandle.createWritable();
      await writable.write(wbOut);
      await writable.close();
      hasUnsavedChanges = false;
      updateSaveButtonState();
      alert("✅ Saved directly to Class_6_to_9_Student_Marks_Sheet.xlsx!");
    } else {
      XLSX.writeFile(workbook, "Class_6_to_9_Student_Marks_Sheet.xlsx");
      hasUnsavedChanges = false;
      updateSaveButtonState();
    }
  } catch (err) {
    alert("Error saving Excel file (close it in Microsoft Excel first): " + err.message);
  }
}