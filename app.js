const fileInput = document.getElementById("fileInput");
const keywordFiltersPanel = document.getElementById("keywordFiltersPanel");
const keywordFilters = document.getElementById("keywordFilters");
const tablePanel = document.getElementById("tablePanel");
const resultTable = document.getElementById("resultTable");
const copyBtn = document.getElementById("copyBtn");
const exportBtn = document.getElementById("exportBtn");
const mainContainer = document.getElementById("mainContainer");
const currentSheetName = document.getElementById("currentSheetName");

const sheetTabsPanel = document.getElementById("sheetTabsPanel");
const sheetTabs = document.getElementById("sheetTabs");

const fieldModal = document.getElementById("fieldModal");
const fieldOptions = document.getElementById("fieldOptions");
const selectAllFields = document.getElementById("selectAllFields");
const confirmFieldsBtn = document.getElementById("confirmFieldsBtn");

let workbookRowsBySheet = {};
let processedSheetCache = {};
let sheetHeaderSelection = {};
let activeSheetName = "";

let allRows = [];
let allHeaders = [];
let originalRows = [];
let currentRows = [];
let headers = [];
let keywordFilterState = {};
let sortState = { key: "", direction: "asc" };

fileInput.addEventListener("click", () => {
  fileInput.value = "";
});

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const parsed = await parseWorkbook(file);
    if (!parsed.sheetNames.length) {
      alert("未检测到工作表。");
      return;
    }

    workbookRowsBySheet = parsed.rowsBySheet;
    processedSheetCache = {};
    sheetHeaderSelection = {};
    mainContainer.classList.remove("is-empty");

    renderSheetTabs(parsed.sheetNames);
    activateSheet(parsed.sheetNames[0]);
  } catch (error) {
    console.error(error);
    alert("文件解析失败，请确认文件格式正确。");
  }
});

copyBtn.addEventListener("click", async () => {
  if (!headers.length) {
    return;
  }
  const lines = [
    headers.join("\t"),
    ...currentRows.map((row) => headers.map((header) => row[header] ?? "").join("\t")),
  ];
  try {
    await navigator.clipboard.writeText(lines.join("\n"));
    copyBtn.textContent = "已复制";
    setTimeout(() => {
      copyBtn.textContent = "复制";
    }, 1200);
  } catch (error) {
    console.error(error);
    alert("复制失败，请检查浏览器权限。");
  }
});

exportBtn.addEventListener("click", () => {
  if (!headers.length) {
    return;
  }
  const exportData = currentRows.map((row) => {
    const out = {};
    headers.forEach((header) => {
      out[header] = row[header] ?? "";
    });
    return out;
  });
  const sheet = XLSX.utils.json_to_sheet(exportData, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "结果");
  const safeName = (activeSheetName || "结果").replace(/[\\/:*?"<>|]/g, "_");
  XLSX.writeFile(wb, `整理结果_${safeName}.xlsx`);
});

selectAllFields.addEventListener("change", () => {
  const checked = selectAllFields.checked;
  Array.from(fieldOptions.querySelectorAll("input[type='checkbox']")).forEach((input) => {
    input.checked = checked;
  });
});

confirmFieldsBtn.addEventListener("click", () => {
  const selectedHeaders = Array.from(fieldOptions.querySelectorAll("input[type='checkbox']"))
    .filter((input) => input.checked)
    .map((input) => input.value);

  if (!selectedHeaders.length) {
    alert("请至少选择一个字段。");
    return;
  }

  sheetHeaderSelection[activeSheetName] = selectedHeaders;
  applySelectedHeadersForActiveSheet(selectedHeaders);
  fieldModal.hidden = true;
});

function renderSheetTabs(sheetNames) {
  sheetTabs.innerHTML = "";
  sheetNames.forEach((name) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "sheet-tab";
    btn.textContent = name;
    btn.addEventListener("click", () => activateSheet(name));
    sheetTabs.appendChild(btn);
  });
  sheetTabsPanel.hidden = sheetNames.length <= 1;
}

function activateSheet(sheetName) {
  activeSheetName = sheetName;
  currentSheetName.textContent = sheetName ? `（${sheetName}）` : "";

  Array.from(sheetTabs.querySelectorAll(".sheet-tab")).forEach((btn) => {
    btn.classList.toggle("active", btn.textContent === sheetName);
  });

  const processed = getProcessedSheetData(sheetName);
  if (!processed) {
    alert("该工作表无可用数据。");
    return;
  }

  allHeaders = processed.headers;
  allRows = processed.rows;

  const savedHeaders = sheetHeaderSelection[sheetName];
  if (savedHeaders && savedHeaders.length) {
    const validHeaders = savedHeaders.filter((h) => allHeaders.includes(h));
    applySelectedHeadersForActiveSheet(validHeaders.length ? validHeaders : allHeaders);
  } else {
    openFieldModal(allHeaders);
  }
}

function getProcessedSheetData(sheetName) {
  if (processedSheetCache[sheetName]) {
    return processedSheetCache[sheetName];
  }

  const rows = workbookRowsBySheet[sheetName] || [];
  if (!rows.length) {
    return null;
  }

  const merged = mergeDuplicateColumns(rows);
  const dedupedRows = deduplicateRows(merged.rows, merged.headers);
  const result = { headers: merged.headers, rows: dedupedRows };
  processedSheetCache[sheetName] = result;
  return result;
}

function applySelectedHeadersForActiveSheet(selectedHeaders) {
  headers = selectedHeaders;
  originalRows = allRows
    .map((row) => {
      const projected = {};
      headers.forEach((header) => {
        projected[header] = row[header] ?? "";
      });
      return projected;
    })
    .filter((row) => headers.some((header) => String(row[header] ?? "").trim() !== ""));

  currentRows = [...originalRows];
  keywordFilterState = {};
  sortState = { key: "", direction: "asc" };

  renderKeywordFilters();
  renderTable(currentRows);

  keywordFiltersPanel.hidden = false;
  tablePanel.hidden = false;
}

function openFieldModal(listHeaders) {
  fieldOptions.innerHTML = "";
  listHeaders.forEach((header) => {
    const label = document.createElement("label");
    label.className = "field-item";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = header;
    input.checked = true;
    input.addEventListener("change", () => {
      const total = fieldOptions.querySelectorAll("input[type='checkbox']").length;
      const checkedCount = fieldOptions.querySelectorAll("input[type='checkbox']:checked").length;
      selectAllFields.checked = total > 0 && checkedCount === total;
    });

    const text = document.createElement("span");
    text.textContent = header;

    label.appendChild(input);
    label.appendChild(text);
    fieldOptions.appendChild(label);
  });

  selectAllFields.checked = true;
  fieldModal.hidden = false;
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetNames = workbook.SheetNames || [];
        const rowsBySheet = {};

        sheetNames.forEach((name) => {
          const sheet = workbook.Sheets[name];
          rowsBySheet[name] = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: false,
          });
        });

        resolve({ sheetNames, rowsBySheet });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeRows(rows, currentHeaders) {
  return rows.map((row) => {
    const normalized = {};
    currentHeaders.forEach((header) => {
      const value = row[header];
      normalized[header] = value === null || value === undefined ? "" : String(value).trim();
    });
    return normalized;
  });
}

function mergeDuplicateColumns(rows) {
  const firstRowHeaders = Object.keys(rows[0] || {});
  const groupedHeaders = new Map();
  let removedEmptyColumns = 0;

  firstRowHeaders.forEach((header) => {
    if (!header || header.trim() === "") {
      return;
    }
    if (/^__EMPTY(_\d+)?$/i.test(header)) {
      removedEmptyColumns += 1;
      return;
    }

    const baseHeader = header.replace(/_\d+$/, "");
    const indexMatch = header.match(/_(\d+)$/);
    const index = indexMatch ? Number(indexMatch[1]) : 0;

    if (!groupedHeaders.has(baseHeader)) {
      groupedHeaders.set(baseHeader, []);
    }
    groupedHeaders.get(baseHeader).push({ header, index });
  });

  const mergedHeaders = Array.from(groupedHeaders.keys());
  const expandedRows = [];

  rows.forEach((row) => {
    const availableIndexes = new Set([0]);

    mergedHeaders.forEach((baseHeader) => {
      groupedHeaders.get(baseHeader).forEach(({ index, header }) => {
        const value = row[header];
        if (value !== null && value !== undefined && String(value).trim() !== "") {
          availableIndexes.add(index);
        }
      });
    });

    const indexes = Array.from(availableIndexes).sort((a, b) => a - b);
    const rowGroup = [];

    indexes.forEach((index) => {
      const newRow = {};
      let hasAnyValue = false;

      mergedHeaders.forEach((baseHeader) => {
        const variants = groupedHeaders.get(baseHeader);
        const hasSplitVariants = variants.some((v) => v.index !== 0);
        const exactVariant = variants.find((v) => v.index === index);

        let sourceHeader = "";
        if (exactVariant) {
          sourceHeader = exactVariant.header;
        } else if (!hasSplitVariants) {
          sourceHeader = variants[0].header;
        }

        const value = sourceHeader ? row[sourceHeader] : "";
        const normalizedValue = value === null || value === undefined ? "" : String(value).trim();

        if (normalizedValue !== "") {
          hasAnyValue = true;
        }
        newRow[baseHeader] = normalizedValue;
      });

      if (hasAnyValue) {
        rowGroup.push(newRow);
      }
    });

    if (rowGroup.length) {
      expandedRows.push(...rowGroup);
    }
  });

  return {
    headers: mergedHeaders,
    rows: normalizeRows(expandedRows, mergedHeaders),
    mergedColumns: firstRowHeaders.length - mergedHeaders.length - removedEmptyColumns,
    removedEmptyColumns,
  };
}

function deduplicateRows(rows, currentHeaders) {
  const seen = new Set();
  const deduped = [];

  rows.forEach((row) => {
    const key = JSON.stringify(currentHeaders.map((header) => row[header] || ""));
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(row);
    }
  });

  return deduped;
}

function renderKeywordFilters() {
  keywordFilters.innerHTML = "";

  headers.forEach((header) => {
    const wrap = document.createElement("div");
    wrap.className = "filter-item";

    const label = document.createElement("label");
    label.textContent = header;

    const input = document.createElement("input");
    const suggestBox = document.createElement("div");
    suggestBox.className = "suggest-box";

    const uniqueValues = [...new Set(originalRows.map((row) => row[header]))];
    const sortedValues = uniqueValues
      .filter((v) => v !== "")
      .sort((a, b) => String(a).localeCompare(String(b), "zh-Hans-CN", { numeric: true, sensitivity: "base" }));

    input.type = "text";
    input.placeholder = "";

    const suggestValues = uniqueValues.includes("") ? [...sortedValues, "(空白)"] : sortedValues;

    input.addEventListener("input", () => {
      keywordFilterState[header] = input.value.trim();
      applyAllFilters();
      renderSuggestList(header, input, suggestBox, suggestValues);
    });

    input.addEventListener("focus", () => {
      renderSuggestList(header, input, suggestBox, suggestValues);
    });

    input.addEventListener("blur", () => {
      setTimeout(() => {
        suggestBox.classList.remove("open");
      }, 120);
    });

    wrap.appendChild(label);
    wrap.appendChild(input);
    wrap.appendChild(suggestBox);
    keywordFilters.appendChild(wrap);
  });
}

function renderSuggestList(header, input, suggestBox, values) {
  const keyword = input.value.trim().toLowerCase();
  const matched = values.filter((value) => String(value).toLowerCase().includes(keyword)).slice(0, 200);

  suggestBox.innerHTML = "";
  if (!matched.length) {
    suggestBox.classList.remove("open");
    return;
  }

  matched.forEach((value) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "suggest-item";
    item.textContent = value;
    item.addEventListener("mousedown", (e) => {
      e.preventDefault();
      input.value = value;
      keywordFilterState[header] = value;
      applyAllFilters();
      suggestBox.classList.remove("open");
    });
    suggestBox.appendChild(item);
  });

  suggestBox.classList.add("open");
}

function applyAllFilters() {
  const filtered = originalRows.filter((row) => {
    return headers.every((header) => {
      const keyword = keywordFilterState[header];
      if (!keyword) {
        return true;
      }
      if (keyword === "(空白)") {
        return row[header] === "";
      }
      return row[header].toLowerCase().includes(keyword.toLowerCase());
    });
  });

  currentRows = sortRows(filtered);
  renderTable(currentRows);
}

function renderTable(rows) {
  resultTable.innerHTML = "";

  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    th.className = "sortable-th";
    if (sortState.key === header) {
      th.classList.add(sortState.direction === "asc" ? "sorted-asc" : "sorted-desc");
    }
    th.addEventListener("click", () => {
      toggleSort(header);
    });
    headRow.appendChild(th);
  });
  head.appendChild(headRow);
  resultTable.appendChild(head);

  const body = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header];
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
  resultTable.appendChild(body);
}

function toggleSort(header) {
  if (sortState.key === header) {
    sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
  } else {
    sortState.key = header;
    sortState.direction = "asc";
  }
  applyAllFilters();
}

function sortRows(rows) {
  if (!sortState.key) {
    return [...rows];
  }
  const key = sortState.key;
  const dir = sortState.direction === "asc" ? 1 : -1;

  return [...rows].sort((a, b) => {
    const va = a[key] ?? "";
    const vb = b[key] ?? "";
    const na = Number(va);
    const nb = Number(vb);
    const bothNumber = va !== "" && vb !== "" && !Number.isNaN(na) && !Number.isNaN(nb);
    if (bothNumber) {
      return (na - nb) * dir;
    }
    return (
      String(va).localeCompare(String(vb), "zh-Hans-CN", {
        numeric: true,
        sensitivity: "base",
      }) * dir
    );
  });
}
