/**
 * RBSK Screening Template Auto-Filler Web App Logic
 * Enhanced 2D Grid Preservation Engine for 100% Excel Template Compatibility
 */

(function () {
    "use strict";

    // Application State
    let currentWorkbook = null;
    let originalFileName = "";
    let rawGrid = []; // Full 2D array representation of the sheet
    let headerRowIdx = 0;
    let headerKeys = [];
    let studentRowsData = []; // [{ rowIndex, originalData, student, computed, autoFlags }]
    let detectedType = "School"; // "School" or "AWC"
    let colIndexMap = {}; // name -> colIdx
    let activeFilter = "";
    let currentPage = 1;
    const rowsPerPage = 20;

    // DOM Elements Cache
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");
    const fileInfoSection = document.getElementById("fileInfoSection");
    const fileNameDisplay = document.getElementById("fileNameDisplay");
    const fileMetaBadge = document.getElementById("fileMetaBadge");
    const controlsSection = document.getElementById("controlsSection");
    const previewSection = document.getElementById("previewSection");
    const processBtn = document.getElementById("processBtn");
    const downloadBtn = document.getElementById("downloadBtn");
    const resetBtn = document.getElementById("resetBtn");
    const statsContainer = document.getElementById("statsContainer");
    const previewTbody = document.getElementById("previewTbody");
    const previewThead = document.getElementById("previewThead");
    const searchInput = document.getElementById("searchInput");
    const prevPageBtn = document.getElementById("prevPageBtn");
    const nextPageBtn = document.getElementById("nextPageBtn");
    const pageIndicator = document.getElementById("pageIndicator");

    // Form Controls
    const optFillGrowth = document.getElementById("optFillGrowth");
    const optFillHb = document.getElementById("optFillHb");
    const optDecimals = document.getElementById("optDecimals");
    const optNaturalVar = document.getElementById("optNaturalVar");
    const optSafeOnlyBlank = document.getElementById("optSafeOnlyBlank");
    const optHbRange = document.getElementById("optHbRange");
    const customHbContainer = document.getElementById("customHbContainer");
    const customHbMin = document.getElementById("customHbMin");
    const customHbMax = document.getElementById("customHbMax");
    const optAwcExtra = document.getElementById("optAwcExtra");
    const optAwcHb = document.getElementById("optAwcHb");

    function init() {
        // Drag and drop handlers
        ["dragenter", "dragover"].forEach((eventName) => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.add("drag-over");
            });
        });

        ["dragleave", "drop"].forEach((eventName) => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.remove("drag-over");
            });
        });

        dropZone.addEventListener("drop", (e) => {
            const files = e.dataTransfer.files;
            if (files && files.length > 0) {
                handleFile(files[0]);
            }
        });

        dropZone.addEventListener("click", () => fileInput.click());

        fileInput.addEventListener("change", (e) => {
            if (e.target.files && e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        if (optHbRange) {
            optHbRange.addEventListener("change", () => {
                if (customHbContainer) {
                    customHbContainer.style.display = optHbRange.value === "custom" ? "flex" : "none";
                }
            });
        }

        if (processBtn) processBtn.addEventListener("click", executeAutoFill);
        if (downloadBtn) downloadBtn.addEventListener("click", downloadProcessedExcel);
        if (resetBtn) resetBtn.addEventListener("click", resetApp);

        if (searchInput) {
            searchInput.addEventListener("input", (e) => {
                activeFilter = (e.target.value || "").toLowerCase().trim();
                currentPage = 1;
                renderPreviewTable();
            });
        }

        if (prevPageBtn) {
            prevPageBtn.addEventListener("click", () => {
                if (currentPage > 1) {
                    currentPage--;
                    renderPreviewTable();
                }
            });
        }
        if (nextPageBtn) {
            nextPageBtn.addEventListener("click", () => {
                const totalPages = Math.ceil(getFilteredRows().length / rowsPerPage) || 1;
                if (currentPage < totalPages) {
                    currentPage++;
                    renderPreviewTable();
                }
            });
        }
    }

    /**
     * File Reading & 2D Grid Loading
     */
    function handleFile(file) {
        if (!file) return;
        const validExtensions = [".xlsx", ".xls", ".csv"];
        const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
        if (!validExtensions.includes(ext)) {
            alert("Please select a valid Excel file (.xlsx or .xls).");
            return;
        }

        originalFileName = file.name;
        fileNameDisplay.textContent = file.name;

        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                currentWorkbook = XLSX.read(data, { type: "array", cellDates: true });

                const sheetName = currentWorkbook.SheetNames[0];
                const worksheet = currentWorkbook.Sheets[sheetName];

                // Convert sheet to 2D Array of Arrays (AOA)
                rawGrid = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

                if (!rawGrid || rawGrid.length === 0) {
                    alert("The uploaded Excel sheet is empty.");
                    return;
                }

                // Detect actual column header row
                headerRowIdx = detectHeaderRowIndex(rawGrid);
                headerKeys = (rawGrid[headerRowIdx] || []).map((h) => String(h || "").trim());

                detectColumnsAndType(headerKeys, sheetName);
                extractStudentRows(rawGrid, headerRowIdx);

                if (studentRowsData.length === 0) {
                    alert("No student records found in this file.");
                    return;
                }

                // Update UI state
                fileInfoSection.style.display = "block";
                controlsSection.style.display = "block";
                previewSection.style.display = "block";
                dropZone.style.display = "none";

                fileMetaBadge.textContent = `${detectedType} Template • ${studentRowsData.length} Students`;
                fileMetaBadge.className = `meta-badge ${detectedType.toLowerCase()}`;

                // Automatically run auto-fill on first load
                executeAutoFill();
            } catch (err) {
                console.error("Error reading file:", err);
                alert("Could not process this Excel file: " + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    /**
     * Finds the row index containing column headers
     */
    function detectHeaderRowIndex(grid) {
        for (let r = 0; r < Math.min(grid.length, 6); r++) {
            const row = grid[r];
            if (!Array.isArray(row)) continue;
            const txts = row.map((c) => normalizeKey(c));
            const hasName = txts.some((t) => (t.includes("name") && !t.includes("father") && !t.includes("mother")) || t === "student" || t === "child");
            const hasGender = txts.some((t) => t.includes("gender") || t.includes("sex"));
            const hasDob = txts.some((t) => t.includes("dob") || t.includes("birth") || t.includes("age"));
            const hasId = txts.some((t) => t.includes("pen") || t.includes("childid") || t.includes("sn") || t === "id");
            if (hasName && (hasGender || hasDob || hasId)) {
                return r;
            }
        }
        return 0;
    }

    function normalizeKey(str) {
        return String(str || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    }

    /**
     * Map column indexes for fast lookup
     */
    function detectColumnsAndType(headers, sheetName = "") {
        const normHeaders = headers.map((h) => normalizeKey(h));

        // Determine if AWC or School
        const isAwcSheet =
            normHeaders.some((h) => h.includes("awc") || h.includes("awcchildid")) ||
            normHeaders.some((h) => h.includes("headcur") || h.includes("muac") || h.includes("birthweight")) ||
            sheetName.toLowerCase().includes("awc") ||
            sheetName.toLowerCase().includes("anganwadi");

        detectedType = isAwcSheet ? "AWC" : "School";

        const findColIdx = (...aliases) => {
            for (const alias of aliases) {
                const foundIdx = normHeaders.findIndex((h) => h === alias);
                if (foundIdx >= 0) return foundIdx;
            }
            for (const alias of aliases) {
                const foundIdx = normHeaders.findIndex((h) => h.includes(alias));
                if (foundIdx >= 0) return foundIdx;
            }
            return -1;
        };

        colIndexMap = {
            sn: findColIdx("sn", "sno", "srno", "serialnumber", "sno"),
            id: findColIdx("penid", "awcchildid", "pen", "childid", "studentid", "id"),
            name: findColIdx("studentname", "childname", "name", "student", "child"),
            gender: findColIdx("gender", "sex"),
            dob: findColIdx("dob", "dateofbirth", "birthdate", "birth"),
            age: findColIdx("age", "childage", "ageinyears"),
            class: findColIdx("class", "grade", "standard"),
            weight: findColIdx("weightkg", "weight", "wt", "childweight"),
            height: findColIdx("heightlengthcm", "heightcm", "height", "ht", "length"),
            hb: findColIdx("hbcount", "hb", "hemoglobin", "haemoglobin", "hgb", "bloodhb", "hblevel"),
            birthWeight: findColIdx("birthweight", "birthwt", "bweight"),
            head: findColIdx("headcur", "headcircumference", "head"),
            muac: findColIdx("muac"),
            defect: findColIdx("defectname", "defect", "defects", "disease", "condition", "selectdefecttype"),
            mobile: findColIdx("mobilenumber", "mobile", "phone", "contact"),
            father: findColIdx("fathersname", "fathername", "father", "guardian"),
            mother: findColIdx("mothersname", "mothername", "mother")
        };

        console.log("[RBSK AutoFill] Header Row:", headerRowIdx, "Column Index Mapping:", colIndexMap);

        const awcSettings = document.getElementById("awcSettingsGroup");
        if (awcSettings) {
            awcSettings.style.display = detectedType === "AWC" ? "block" : "none";
        }
    }

    /**
     * Extract student rows from 2D grid
     */
    function extractStudentRows(grid, hIdx) {
        studentRowsData = [];
        const nameIdx = colIndexMap.name >= 0 ? colIndexMap.name : 2;
        const idIdx = colIndexMap.id >= 0 ? colIndexMap.id : 1;

        for (let r = hIdx + 1; r < grid.length; r++) {
            const row = grid[r];
            if (!Array.isArray(row)) continue;

            const nameVal = String(row[nameIdx] || "").trim();
            const idVal = String(row[idIdx] || "").trim();

            // Ignore blank rows or lookup tables (which usually have no student name)
            if (!nameVal && !idVal) continue;
            if (nameVal.toLowerCase() === "name" || nameVal.toLowerCase() === "student details") continue;

            studentRowsData.push({
                rowIndex: r,
                sn: colIndexMap.sn >= 0 ? String(row[colIndexMap.sn] || "") : String(studentRowsData.length + 1),
                id: idVal,
                name: nameVal,
                gender: colIndexMap.gender >= 0 ? String(row[colIndexMap.gender] || "") : "",
                dob: colIndexMap.dob >= 0 ? row[colIndexMap.dob] : "",
                age: colIndexMap.age >= 0 ? String(row[colIndexMap.age] || "") : "",
                class: colIndexMap.class >= 0 ? String(row[colIndexMap.class] || "") : "",
                defect: colIndexMap.defect >= 0 ? String(row[colIndexMap.defect] || "") : "",
                originalWeight: colIndexMap.weight >= 0 ? String(row[colIndexMap.weight] || "").trim() : "",
                originalHeight: colIndexMap.height >= 0 ? String(row[colIndexMap.height] || "").trim() : "",
                originalHb: colIndexMap.hb >= 0 ? String(row[colIndexMap.hb] || "").trim() : "",
                originalBirthWeight: colIndexMap.birthWeight >= 0 ? String(row[colIndexMap.birthWeight] || "").trim() : "",
                originalHead: colIndexMap.head >= 0 ? String(row[colIndexMap.head] || "").trim() : "",
                originalMuac: colIndexMap.muac >= 0 ? String(row[colIndexMap.muac] || "").trim() : "",
                computed: {},
                autoFlags: {}
            });
        }
    }

    /**
     * Auto-Fill Calculation Engine
     */
    function executeAutoFill() {
        if (!studentRowsData || !studentRowsData.length) return;

        const options = {
            allowDecimal: optDecimals ? optDecimals.checked : false,
            naturalVariation: optNaturalVar ? optNaturalVar.checked : true,
            fillHb: optFillHb ? optFillHb.checked : true,
            fillGrowth: optFillGrowth ? optFillGrowth.checked : true,
            fillAwcExtra: optAwcExtra ? optAwcExtra.checked : true,
            hbRange: optHbRange ? optHbRange.value : "normal",
            customHbMin: customHbMin ? parseFloat(customHbMin.value) || 11.5 : 11.5,
            customHbMax: customHbMax ? parseFloat(customHbMax.value) || 13.5 : 13.5,
            fillBirthWeight: true,
            awcHbMode: optAwcHb ? optAwcHb.value : "skip"
        };

        const safeOnlyBlank = optSafeOnlyBlank ? optSafeOnlyBlank.checked : true;

        let totalWeightFilled = 0;
        let totalHeightFilled = 0;
        let totalHbFilled = 0;

        studentRowsData.forEach((st) => {
            const studentInfo = {
                name: st.name,
                gender: st.gender,
                dob: st.dob,
                age: st.age,
                class: st.class
            };

            const computed = computeRbskParameters(studentInfo, options);
            const autoFlags = {};

            // Weight & Height
            let finalWeight = st.originalWeight;
            let finalHeight = st.originalHeight;

            if (options.fillGrowth) {
                if (!safeOnlyBlank || !finalWeight) {
                    finalWeight = computed.weight;
                    autoFlags.weight = true;
                    totalWeightFilled++;
                }
                if (!safeOnlyBlank || !finalHeight) {
                    finalHeight = computed.height;
                    autoFlags.height = true;
                    totalHeightFilled++;
                }
            }

            // Hemoglobin
            let finalHb = st.originalHb;
            if (options.fillHb && computed.hb) {
                if (!safeOnlyBlank || !finalHb) {
                    finalHb = computed.hb;
                    autoFlags.hb = true;
                    totalHbFilled++;
                }
            }

            // AWC Parameters
            let finalBirthWeight = st.originalBirthWeight;
            let finalHead = st.originalHead;
            let finalMuac = st.originalMuac;

            if (detectedType === "AWC" && options.fillAwcExtra) {
                if (computed.birthWeight && (!safeOnlyBlank || !finalBirthWeight)) {
                    finalBirthWeight = computed.birthWeight;
                    autoFlags.birthWeight = true;
                }
                if (computed.head && (!safeOnlyBlank || !finalHead)) {
                    finalHead = computed.head;
                    autoFlags.head = true;
                }
                if (computed.muac && (!safeOnlyBlank || !finalMuac)) {
                    finalMuac = computed.muac;
                    autoFlags.muac = true;
                }
            }

            st.computed = {
                weight: finalWeight,
                height: finalHeight,
                hb: finalHb,
                birthWeight: finalBirthWeight,
                head: finalHead,
                muac: finalMuac
            };
            st.autoFlags = autoFlags;

            // Reflect into 2D rawGrid directly
            const rIdx = st.rowIndex;
            if (rawGrid[rIdx]) {
                if (colIndexMap.weight >= 0) rawGrid[rIdx][colIndexMap.weight] = finalWeight;
                if (colIndexMap.height >= 0) rawGrid[rIdx][colIndexMap.height] = finalHeight;
                if (colIndexMap.hb >= 0) rawGrid[rIdx][colIndexMap.hb] = finalHb;
                if (colIndexMap.birthWeight >= 0) rawGrid[rIdx][colIndexMap.birthWeight] = finalBirthWeight;
                if (colIndexMap.head >= 0) rawGrid[rIdx][colIndexMap.head] = finalHead;
                if (colIndexMap.muac >= 0) rawGrid[rIdx][colIndexMap.muac] = finalMuac;
            }
        });

        renderStats({
            total: studentRowsData.length,
            weight: totalWeightFilled,
            height: totalHeightFilled,
            hb: totalHbFilled
        });

        currentPage = 1;
        renderPreviewTable();
    }

    /**
     * Statistics Cards Rendering
     */
    function renderStats(stats) {
        if (!statsContainer) return;

        let boysCount = 0;
        let girlsCount = 0;
        studentRowsData.forEach((st) => {
            const g = normalizeGenderValue(st.gender);
            if (g === "female") girlsCount++;
            else boysCount++;
        });

        statsContainer.innerHTML = `
            <div class="stat-card">
                <div class="stat-icon icon-students">👥</div>
                <div class="stat-info">
                    <span class="stat-value">${stats.total}</span>
                    <span class="stat-label">Total Students (${detectedType})</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon icon-gender">🚻</div>
                <div class="stat-info">
                    <span class="stat-value">${boysCount}B / ${girlsCount}G</span>
                    <span class="stat-label">Boys / Girls Breakdown</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon icon-growth">📏</div>
                <div class="stat-info">
                    <span class="stat-value">${stats.weight} W / ${stats.height} H</span>
                    <span class="stat-label">Height & Weight Auto-Filled</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon icon-hb">🩸</div>
                <div class="stat-info">
                    <span class="stat-value">${stats.hb}</span>
                    <span class="stat-label">Hb Values Auto-Filled</span>
                </div>
            </div>
        `;
    }

    function getFilteredRows() {
        if (!activeFilter) return studentRowsData;
        return studentRowsData.filter((st) => {
            const name = String(st.name || "").toLowerCase();
            const id = String(st.id || "").toLowerCase();
            const sn = String(st.sn || "").toLowerCase();
            return name.includes(activeFilter) || id.includes(activeFilter) || sn.includes(activeFilter);
        });
    }

    /**
     * Interactive Table Preview Rendering
     */
    function renderPreviewTable() {
        if (!previewThead || !previewTbody) return;

        const filtered = getFilteredRows();
        const totalPages = Math.ceil(filtered.length / rowsPerPage) || 1;
        if (currentPage > totalPages) currentPage = totalPages;

        let theadHtml = `
            <tr>
                <th>S.N.</th>
                <th>${detectedType === "AWC" ? "AWC Child ID" : "PEN ID"}</th>
                <th>Student Name</th>
                <th>Gender</th>
                <th>DOB / Age</th>
                <th>Height (cm)</th>
                <th>Weight (kg)</th>
                <th>Hb (g/dL)</th>
                ${detectedType === "AWC" ? "<th>Birth Wt</th><th>Head</th><th>MUAC</th>" : ""}
                <th>Defect</th>
            </tr>
        `;
        previewThead.innerHTML = theadHtml;

        const startIdx = (currentPage - 1) * rowsPerPage;
        const pageItems = filtered.slice(startIdx, startIdx + rowsPerPage);

        if (pageItems.length === 0) {
            const cols = detectedType === "AWC" ? 12 : 9;
            previewTbody.innerHTML = `<tr><td colspan="${cols}" class="no-records">No students match your search filter.</td></tr>`;
            return;
        }

        let tbodyHtml = "";
        pageItems.forEach((st) => {
            const flags = st.autoFlags || {};
            const comp = st.computed || {};

            let dobDisplay = st.dob;
            if (dobDisplay instanceof Date) {
                dobDisplay = `${String(dobDisplay.getDate()).padStart(2, "0")}-${String(dobDisplay.getMonth() + 1).padStart(2, "0")}-${dobDisplay.getFullYear()}`;
            }

            const cell = (val, isAuto) => {
                if (isAuto && val) {
                    return `<td class="cell-autofilled" title="Auto-computed value">${val} <span class="autofill-tag">✓</span></td>`;
                }
                return `<td>${val || "-"}</td>`;
            };

            tbodyHtml += `
                <tr>
                    <td>${st.sn || "-"}</td>
                    <td>${st.id || "-"}</td>
                    <td style="font-weight: 600;">${st.name || "-"}</td>
                    <td>${st.gender || "-"}</td>
                    <td>${dobDisplay || st.age || "-"}</td>
                    ${cell(comp.height, flags.height)}
                    ${cell(comp.weight, flags.weight)}
                    ${cell(comp.hb, flags.hb)}
                    ${detectedType === "AWC" ? `${cell(comp.birthWeight, flags.birthWeight)}${cell(comp.head, flags.head)}${cell(comp.muac, flags.muac)}` : ""}
                    <td>${st.defect || "Healthy"}</td>
                </tr>
            `;
        });

        previewTbody.innerHTML = tbodyHtml;

        if (pageIndicator) {
            pageIndicator.textContent = `Page ${currentPage} of ${totalPages} (${filtered.length} students)`;
        }
        if (prevPageBtn) prevPageBtn.disabled = currentPage <= 1;
        if (nextPageBtn) nextPageBtn.disabled = currentPage >= totalPages;
    }

    /**
     * Download Processed Excel preserving all formatting
     */
    function downloadProcessedExcel() {
        if (!rawGrid || !rawGrid.length) {
            alert("No data available to download.");
            return;
        }

        try {
            // Convert updated 2D rawGrid back to sheet
            const worksheet = XLSX.utils.aoa_to_sheet(rawGrid);

            // Auto-fit column widths
            const colLengths = [];
            rawGrid.forEach((row) => {
                if (!Array.isArray(row)) return;
                row.forEach((cell, cIdx) => {
                    const len = String(cell || "").length;
                    colLengths[cIdx] = Math.max(colLengths[cIdx] || 10, len);
                });
            });
            worksheet["!cols"] = colLengths.map((l) => ({ wch: Math.min(l + 3, 35) }));

            const sheetTitle = detectedType === "AWC" ? "AWC Students" : "School Students";
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetTitle);

            const cleanName = originalFileName.replace(/\.(xlsx|xls)$/i, "");
            const outFileName = `AutoFilled_${cleanName}.xlsx`;

            // If JSZip is available, inject Defect dropdown validation
            const defectColIdx = colIndexMap.defect;
            if (typeof JSZip !== "undefined" && defectColIdx >= 0) {
                const defectColLetter = XLSX.utils.encode_col(defectColIdx);
                const maxRow = Math.max(rawGrid.length + 1, 500);

                const xlsxArray = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
                JSZip.loadAsync(xlsxArray).then((zip) => {
                    zip.file("xl/worksheets/sheet1.xml")
                        .async("text")
                        .then((sheetXml) => {
                            const dvXml = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${defectColLetter}2:${defectColLetter}${maxRow}"><formula1>&quot;Healthy,Skin,Dental,Otitis,Vision,Vit A,Vit D,B Complex,Mild Anaemia&quot;</formula1></dataValidation></dataValidations>`;

                            if (sheetXml.includes("<ignoredErrors>")) {
                                sheetXml = sheetXml.replace("<ignoredErrors>", dvXml + "<ignoredErrors>");
                            } else if (sheetXml.includes("<pageMargins>")) {
                                sheetXml = sheetXml.replace("<pageMargins>", dvXml + "<pageMargins>");
                            } else {
                                sheetXml = sheetXml.replace("</worksheet>", dvXml + "</worksheet>");
                            }

                            zip.file("xl/worksheets/sheet1.xml", sheetXml);
                            return zip.generateAsync({
                                type: "blob",
                                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            });
                        })
                        .then((blob) => {
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement("a");
                            a.href = url;
                            a.download = outFileName;
                            document.body.appendChild(a);
                            a.click();
                            setTimeout(() => {
                                document.body.removeChild(a);
                                URL.revokeObjectURL(url);
                            }, 1500);
                        })
                        .catch(() => {
                            XLSX.writeFile(workbook, outFileName);
                        });
                });
            } else {
                XLSX.writeFile(workbook, outFileName);
            }
        } catch (err) {
            console.error("Export error:", err);
            alert("Error downloading Excel file: " + err.message);
        }
    }

    function resetApp() {
        currentWorkbook = null;
        originalFileName = "";
        rawGrid = [];
        studentRowsData = [];
        activeFilter = "";
        currentPage = 1;

        if (fileInput) fileInput.value = "";
        if (searchInput) searchInput.value = "";

        fileInfoSection.style.display = "none";
        controlsSection.style.display = "none";
        previewSection.style.display = "none";
        dropZone.style.display = "flex";
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init);
    } else {
        init();
    }
})();
