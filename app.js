/**
 * RBSK Screening Template Auto-Filler Web App Logic
 * Enhanced 2D Grid Preservation Engine for 100% Excel Template Compatibility
 * 
 * Features:
 * - Granular Radio Button Toggles for every field (Height, Weight, Hb, Vision, Defect, Head, MUAC, Birth Weight)
 * - Calibrated WHO 0-60m Medians to naturally avoid Stunting/Wasting on Portal
 * - Dynamic Excel Download with Dropdowns (Defect & Vision Left/Right) preserved via JSZip
 * - Text formatting preservation to avoid 6/6 becoming Date format
 * - Live Interactive In-Table Defect and Vision dropdowns
 */

(function () {
    "use strict";

    // Application State
    let currentWorkbook = null;
    let originalFileName = "";
    let rawGrid = []; // Full 2D array representation of the sheet
    let headerRowIdx = 0;
    let headerKeys = [];
    let studentRowsData = []; // [{ rowIndex, sn, id, name, gender, dob, age, class, ... }]
    let detectedType = "School"; // "School" or "AWC"
    let colIndexMap = {}; // field name -> colIdx
    let activeFilter = "";
    let currentPage = 1;
    const rowsPerPage = 20;

    // RBSK Disease / Condition Reference List
    const DEFECT_OPTIONS = [
        { val: "healthy", label: "Healthy", color: "rgb(107, 114, 128)" },
        { val: "skin", label: "Skin", color: "rgb(220, 38, 38)" },
        { val: "dental", label: "Dental", color: "rgb(220, 38, 38)" },
        { val: "otitis", label: "Otitis", color: "rgb(220, 38, 38)" },
        { val: "vision", label: "Vision", color: "rgb(220, 38, 38)" },
        { val: "vitaminA", label: "Vit A", color: "rgb(220, 38, 38)" },
        { val: "vitaminB", label: "B Complex", color: "rgb(220, 38, 38)" },
        { val: "vitaminD", label: "Vit D", color: "rgb(220, 38, 38)" },
        { val: "mildAnaemia", label: "Mild Anaemia", color: "rgb(220, 38, 38)" },
        { val: "moderateAnaemia", label: "Moderate Anaemia", color: "rgb(220, 38, 38)" },
        { val: "severeAnemia", label: "Severe Anemia", color: "rgb(220, 38, 38)" },
        { val: "reactiveAirway", label: "Reactive Airway Disease", color: "rgb(220, 38, 38)" },
        { val: "growingUpConcerns", label: "Growing up concerns", color: "rgb(220, 38, 38)" },
        { val: "painUrination", label: "Pain/Burning Urinating", color: "rgb(220, 38, 38)" },
        { val: "painMenstruation", label: "Pain during menstruation", color: "rgb(220, 38, 38)" }
    ];

    const DEFECT_LABELS = {
        healthy: "Healthy",
        skin: "Skin",
        dental: "Dental",
        otitis: "Otitis",
        vision: "Vision",
        vitaminA: "Vit A",
        vitaminB: "B Complex",
        vitaminD: "Vit D",
        mildAnaemia: "Mild Anaemia",
        moderateAnaemia: "Moderate Anaemia",
        severeAnemia: "Severe Anemia",
        reactiveAirway: "Reactive Airway Disease",
        growingUpConcerns: "Growing up concerns",
        painUrination: "Pain/Burning Urinating",
        painMenstruation: "Pain during menstruation"
    };

    // Vision Options
    const VISION_OPTIONS = [
        "6/6",
        "6/9",
        "6/12",
        "6/18",
        "6/24",
        "6/36",
        "6/60",
        "No vision"
    ];

    function normalizeDefectKey(raw) {
        const s = String(raw || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!s || s === "healthy" || s === "nil" || s === "none" || s === "normal" || s === "no" || s === "0" || s === "fit") {
            return "healthy";
        }
        if (s.includes("severeanemia") || s.includes("severeanaemia") || s.includes("severeanem")) return "severeAnemia";
        if (s.includes("moderateanaemia") || s.includes("moderateanemia") || s.includes("modanemia")) return "moderateAnaemia";
        if (s.includes("mildanaemia") || s.includes("mildanemia") || s.includes("anaemia") || s.includes("anemia")) return "mildAnaemia";
        if (s.includes("skin") || s.includes("scabies") || s.includes("leprosy")) return "skin";
        if (s.includes("dent") || s.includes("teeth") || s.includes("caries")) return "dental";
        if (s.includes("otit") || s.includes("ear")) return "otitis";
        if (s.includes("vis") || s.includes("eye") || s.includes("blind")) return "vision";
        if (s.includes("vita") || s.includes("vitamina") || s.includes("bitot")) return "vitaminA";
        if (s.includes("vitb") || s.includes("bcomp") || s.includes("complex")) return "vitaminB";
        if (s.includes("vitd") || s.includes("vitamind") || s.includes("rickets")) return "vitaminD";
        if (s.includes("reactiveairway") || s.includes("airway") || s.includes("asthma") || s.includes("rad")) return "reactiveAirway";
        if (s.includes("growingup") || s.includes("growingconcern")) return "growingUpConcerns";
        if (s.includes("urinat") || s.includes("burningurine")) return "painUrination";
        if (s.includes("menstruat") || s.includes("periodpain")) return "painMenstruation";
        return "healthy";
    }

    /**
     * Automatic Disease / Anaemia Determination based on Hb value:
     * - Severe Anemia: Hb < 8.0 g/dL
     * - Moderate Anaemia: Hb 8.0 to 10.99 g/dL
     * - Mild Anaemia: Hb 11.0 to 11.49 g/dL
     * - Healthy: Hb >= 11.5 g/dL (or retains pre-existing non-anemic condition)
     */
    function determineDefectFromHb(hbVal, fallbackDefect = "") {
        const num = parseFloat(hbVal);
        if (!isNaN(num) && num > 0) {
            if (num < 8.0) {
                return "severeAnemia";
            } else if (num >= 8.0 && num <= 10.99) {
                return "moderateAnaemia";
            } else if (num >= 11.0 && num <= 11.49) {
                return "mildAnaemia";
            }
        }
        if (fallbackDefect) {
            const norm = normalizeDefectKey(fallbackDefect);
            if (norm && !norm.toLowerCase().includes("anaemia") && !norm.toLowerCase().includes("anemia") && norm !== "healthy") {
                return norm;
            }
        }
        return "healthy";
    }

    function renderDiseaseSelectHtml(currentVal, rowIdx) {
        const isHealthy = currentVal === "healthy";
        const borderStyle = isHealthy ? "border: 1px solid #d1d5db;" : "border: 1px solid rgb(249, 115, 22);";
        const colorStyle = isHealthy ? "color: rgb(107, 114, 128);" : "color: rgb(234, 88, 12);";

        let html = `<select class="rbsk-disease-select" data-row-idx="${rowIdx}" style="${borderStyle} ${colorStyle}">`;

        DEFECT_OPTIONS.forEach((opt) => {
            const sel = opt.val === currentVal ? " selected" : "";
            html += `<option value="${opt.val}" style="color: ${opt.color};"${sel}>${opt.label}</option>`;
        });

        html += `</select>`;
        return html;
    }

    function renderVisionSelectHtml(currentVal, rowIdx, eyeSide) {
        const val = String(currentVal || "").trim();
        let html = `<select class="rbsk-vision-select" data-row-idx="${rowIdx}" data-eye="${eyeSide}" style="border: 1px solid #d1d5db; color: #1e293b;">`;
        html += `<option value=""${!val ? " selected" : ""}>-</option>`;
        VISION_OPTIONS.forEach((opt) => {
            const sel = opt === val ? " selected" : "";
            html += `<option value="${opt}"${sel}>${opt}</option>`;
        });
        html += `</select>`;
        return html;
    }

    function updateSelectStyle(sel, val) {
        if (!sel) return;
        if (val === "healthy") {
            sel.style.borderColor = "#d1d5db";
            sel.style.color = "rgb(107, 114, 128)";
        } else {
            sel.style.borderColor = "rgb(249, 115, 22)";
            sel.style.color = "rgb(234, 88, 12)";
        }
    }

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

    // Presets
    const presetStandardBtn = document.getElementById("presetStandardBtn");
    const presetAllBtn = document.getElementById("presetAllBtn");
    const presetNoneBtn = document.getElementById("presetNoneBtn");

    // Granular Radio Controls
    const radioHeightFill = document.getElementById("radioHeightFill");
    const radioHeightSkip = document.getElementById("radioHeightSkip");
    const optDecimals = document.getElementById("optDecimals");

    const radioWeightFill = document.getElementById("radioWeightFill");
    const radioWeightSkip = document.getElementById("radioWeightSkip");
    const optNaturalVar = document.getElementById("optNaturalVar");

    const radioHbFill = document.getElementById("radioHbFill");
    const radioHbSkip = document.getElementById("radioHbSkip");
    const optHbRange = document.getElementById("optHbRange");
    const customHbContainer = document.getElementById("customHbContainer");
    const customHbMin = document.getElementById("customHbMin");
    const customHbMax = document.getElementById("customHbMax");

    const radioVisionFill = document.getElementById("radioVisionFill");
    const radioVisionSkip = document.getElementById("radioVisionSkip");
    const optDefaultVision = document.getElementById("optDefaultVision");
    const optAwcVision = document.getElementById("optAwcVision");

    const radioDefectFill = document.getElementById("radioDefectFill");
    const radioDefectSkip = document.getElementById("radioDefectSkip");
    const optDefectMode = document.getElementById("optDefectMode");
    const specificDefectContainer = document.getElementById("specificDefectContainer");
    const optSpecificDefect = document.getElementById("optSpecificDefect");

    const radioSafeOnlyBlank = document.getElementById("radioSafeOnlyBlank");
    const radioSafeOverwrite = document.getElementById("radioSafeOverwrite");

    // AWC Radios
    const radioHeadFill = document.getElementById("radioHeadFill");
    const radioHeadSkip = document.getElementById("radioHeadSkip");
    const radioMuacFill = document.getElementById("radioMuacFill");
    const radioMuacSkip = document.getElementById("radioMuacSkip");
    const radioBirthWeightFill = document.getElementById("radioBirthWeightFill");
    const radioBirthWeightSkip = document.getElementById("radioBirthWeightSkip");

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

        // Hb range listener
        if (optHbRange) {
            optHbRange.addEventListener("change", () => {
                if (customHbContainer) {
                    customHbContainer.style.display = optHbRange.value === "custom" ? "flex" : "none";
                }
            });
        }

        // Defect mode listener
        if (optDefectMode) {
            optDefectMode.addEventListener("change", () => {
                if (specificDefectContainer) {
                    specificDefectContainer.style.display = optDefectMode.value === "specific" ? "block" : "none";
                }
            });
        }

        // Preset Toolbar Listeners
        if (presetStandardBtn) {
            presetStandardBtn.addEventListener("click", () => applyPreset("standard"));
        }
        if (presetAllBtn) {
            presetAllBtn.addEventListener("click", () => applyPreset("all"));
        }
        if (presetNoneBtn) {
            presetNoneBtn.addEventListener("click", () => applyPreset("none"));
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
     * Apply Preset configurations
     */
    function applyPreset(preset) {
        [presetStandardBtn, presetAllBtn, presetNoneBtn].forEach((b) => b && b.classList.remove("active"));

        if (preset === "standard") {
            if (presetStandardBtn) presetStandardBtn.classList.add("active");
            if (radioHeightFill) radioHeightFill.checked = true;
            if (radioWeightFill) radioWeightFill.checked = true;
            if (radioHbFill) radioHbFill.checked = true;
            if (optHbRange) optHbRange.value = "normal";
            if (customHbContainer) customHbContainer.style.display = "none";
            if (radioVisionFill) radioVisionFill.checked = true;
            if (optDefaultVision) optDefaultVision.value = "6/6";
            if (optAwcVision) optAwcVision.checked = false;
            if (radioDefectFill) radioDefectFill.checked = true;
            if (optDefectMode) optDefectMode.value = "auto_hb";
            if (specificDefectContainer) specificDefectContainer.style.display = "none";
            if (radioSafeOnlyBlank) radioSafeOnlyBlank.checked = true;
            if (radioHeadFill) radioHeadFill.checked = true;
            if (radioMuacFill) radioMuacFill.checked = true;
            if (radioBirthWeightFill) radioBirthWeightFill.checked = true;
        } else if (preset === "all") {
            if (presetAllBtn) presetAllBtn.classList.add("active");
            if (radioHeightFill) radioHeightFill.checked = true;
            if (radioWeightFill) radioWeightFill.checked = true;
            if (radioHbFill) radioHbFill.checked = true;
            if (radioVisionFill) radioVisionFill.checked = true;
            if (optAwcVision) optAwcVision.checked = true;
            if (radioDefectFill) radioDefectFill.checked = true;
            if (radioHeadFill) radioHeadFill.checked = true;
            if (radioMuacFill) radioMuacFill.checked = true;
            if (radioBirthWeightFill) radioBirthWeightFill.checked = true;
            if (radioSafeOverwrite) radioSafeOverwrite.checked = true;
        } else if (preset === "none") {
            if (presetNoneBtn) presetNoneBtn.classList.add("active");
            if (radioHeightSkip) radioHeightSkip.checked = true;
            if (radioWeightSkip) radioWeightSkip.checked = true;
            if (radioHbSkip) radioHbSkip.checked = true;
            if (radioVisionSkip) radioVisionSkip.checked = true;
            if (radioDefectSkip) radioDefectSkip.checked = true;
            if (radioHeadSkip) radioHeadSkip.checked = true;
            if (radioMuacSkip) radioMuacSkip.checked = true;
            if (radioBirthWeightSkip) radioBirthWeightSkip.checked = true;
        }
        executeAutoFill();
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
     * Map column indexes for fast lookup and detect template type
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
            sn: findColIdx("sn", "sno", "srno", "serialnumber"),
            id: findColIdx("penid", "awcchildid", "pen", "childid", "studentid", "id"),
            name: findColIdx("studentname", "childname", "name", "student", "child"),
            gender: findColIdx("gender", "sex"),
            dob: findColIdx("dob", "dateofbirth", "birthdate", "birth"),
            age: findColIdx("age", "childage", "ageinyears", "ageyears"),
            class: findColIdx("class", "grade", "standard"),
            weight: findColIdx("weightkg", "weight", "wt", "childweight"),
            height: findColIdx("heightlengthcm", "heightcm", "height", "ht", "length"),
            hb: findColIdx("hbcount", "hb", "hemoglobin", "haemoglobin", "hgb", "bloodhb", "hblevel"),
            visionLeft: findColIdx("visionleft", "leftvision", "visleft", "visionl", "bayeinaankh"),
            visionRight: findColIdx("visionright", "rightvision", "visright", "visionr", "dayeinaankh"),
            defect: findColIdx("defectname", "defect", "defects", "healthstatus", "healthdefect", "condition", "disease", "selectdefecttype"),
            birthWeight: findColIdx("birthweight", "birthwt", "bweight"),
            head: findColIdx("headcur", "headcircumference", "head"),
            muac: findColIdx("muac"),
            mobile: findColIdx("mobilenumber", "mobile", "phone", "contact"),
            father: findColIdx("fathersname", "fathername", "father", "guardian"),
            mother: findColIdx("mothersname", "mothername", "mother")
        };

        // Guarantee Defect column exists
        if (colIndexMap.defect < 0) {
            colIndexMap.defect = headerKeys.length;
            headerKeys.push("Defect / Health Status");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.defect] = "Defect / Health Status";
            }
        }

        // Guarantee Vision columns exist if missing
        if (colIndexMap.visionLeft < 0) {
            colIndexMap.visionLeft = headerKeys.length;
            headerKeys.push("Vision Left");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.visionLeft] = "Vision Left";
            }
        }
        if (colIndexMap.visionRight < 0) {
            colIndexMap.visionRight = headerKeys.length;
            headerKeys.push("Vision Right");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.visionRight] = "Vision Right";
            }
        }

        console.log("[RBSK AutoFill] Detected Type:", detectedType, "Column Index Mapping:", colIndexMap);

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

            // Ignore blank rows or lookup tables
            if (!nameVal && !idVal) continue;
            if (nameVal.toLowerCase() === "name" || nameVal.toLowerCase() === "student details") continue;

            const rawDefect = colIndexMap.defect >= 0 ? String(row[colIndexMap.defect] || "").trim() : "";
            const rawHb = colIndexMap.hb >= 0 ? String(row[colIndexMap.hb] || "").trim() : "";
            const initialDefectKey = determineDefectFromHb(rawHb, rawDefect);

            const rawVl = colIndexMap.visionLeft >= 0 ? String(row[colIndexMap.visionLeft] || "").trim() : "";
            const rawVr = colIndexMap.visionRight >= 0 ? String(row[colIndexMap.visionRight] || "").trim() : "";

            studentRowsData.push({
                rowIndex: r,
                sn: colIndexMap.sn >= 0 ? String(row[colIndexMap.sn] || "") : String(studentRowsData.length + 1),
                id: idVal,
                name: nameVal,
                gender: colIndexMap.gender >= 0 ? String(row[colIndexMap.gender] || "") : "",
                dob: colIndexMap.dob >= 0 ? row[colIndexMap.dob] : "",
                age: colIndexMap.age >= 0 ? String(row[colIndexMap.age] || "") : "",
                class: colIndexMap.class >= 0 ? String(row[colIndexMap.class] || "") : "",
                originalDefect: rawDefect,
                defect: DEFECT_LABELS[initialDefectKey] || "Healthy",
                currentDefectKey: initialDefectKey,
                manualDefect: null,
                originalVisionLeft: rawVl,
                originalVisionRight: rawVr,
                manualVisionLeft: null,
                manualVisionRight: null,
                originalWeight: colIndexMap.weight >= 0 ? String(row[colIndexMap.weight] || "").trim() : "",
                originalHeight: colIndexMap.height >= 0 ? String(row[colIndexMap.height] || "").trim() : "",
                originalHb: rawHb,
                originalBirthWeight: colIndexMap.birthWeight >= 0 ? String(row[colIndexMap.birthWeight] || "").trim() : "",
                originalHead: colIndexMap.head >= 0 ? String(row[colIndexMap.head] || "").trim() : "",
                originalMuac: colIndexMap.muac >= 0 ? String(row[colIndexMap.muac] || "").trim() : "",
                computed: {},
                autoFlags: {}
            });
        }
    }

    /**
     * Auto-Fill Calculation Engine based on Granular Radio Toggles
     */
    function executeAutoFill() {
        if (!studentRowsData || !studentRowsData.length) return;

        const fillHeight = radioHeightFill ? radioHeightFill.checked : true;
        const fillWeight = radioWeightFill ? radioWeightFill.checked : true;
        const fillHb = radioHbFill ? radioHbFill.checked : true;
        const fillVision = radioVisionFill ? radioVisionFill.checked : true;
        const fillDefect = radioDefectFill ? radioDefectFill.checked : true;
        const safeOnlyBlank = radioSafeOnlyBlank ? radioSafeOnlyBlank.checked : true;

        const defaultVisionVal = optDefaultVision ? optDefaultVision.value : "6/6";
        const awcVisionEnabled = optAwcVision ? optAwcVision.checked : false;

        const defectMode = optDefectMode ? optDefectMode.value : "auto_hb";
        const specificDefectVal = optSpecificDefect ? optSpecificDefect.value : "healthy";

        const fillHead = radioHeadFill ? radioHeadFill.checked : true;
        const fillMuac = radioMuacFill ? radioMuacFill.checked : true;
        const fillBirthWeight = radioBirthWeightFill ? radioBirthWeightFill.checked : true;

        const options = {
            fillHeight,
            fillWeight,
            fillHb,
            fillVision,
            defaultVision: defaultVisionVal,
            awcVision: awcVisionEnabled,
            fillBirthWeight,
            fillHead,
            fillMuac,
            allowDecimal: optDecimals ? optDecimals.checked : false,
            naturalVariation: optNaturalVar ? optNaturalVar.checked : true,
            hbRange: optHbRange ? optHbRange.value : "normal",
            customHbMin: customHbMin ? parseFloat(customHbMin.value) || 11.5 : 11.5,
            customHbMax: customHbMax ? parseFloat(customHbMax.value) || 13.5 : 13.5,
            isAwc: detectedType === "AWC"
        };

        let totalWeightFilled = 0;
        let totalHeightFilled = 0;
        let totalHbFilled = 0;
        let totalVisionFilled = 0;

        studentRowsData.forEach((st) => {
            const studentInfo = {
                name: st.name,
                gender: st.gender,
                dob: st.dob,
                age: st.age,
                class: st.class,
                isAwc: detectedType === "AWC"
            };

            const computed = computeRbskParameters(studentInfo, options);
            const autoFlags = {};

            // 1. Height & Weight
            let finalHeight = st.originalHeight;
            let finalWeight = st.originalWeight;

            if (fillHeight && computed.height) {
                if (!safeOnlyBlank || !finalHeight) {
                    finalHeight = computed.height;
                    autoFlags.height = true;
                    totalHeightFilled++;
                }
            }

            if (fillWeight && computed.weight) {
                if (!safeOnlyBlank || !finalWeight) {
                    finalWeight = computed.weight;
                    autoFlags.weight = true;
                    totalWeightFilled++;
                }
            }

            // 2. Hemoglobin (Hb)
            let finalHb = st.originalHb;
            if (fillHb && computed.hb) {
                if (!safeOnlyBlank || !finalHb) {
                    finalHb = computed.hb;
                    autoFlags.hb = true;
                    totalHbFilled++;
                }
            }

            // 3. Vision Left & Vision Right
            let finalVisionLeft = st.manualVisionLeft !== null ? st.manualVisionLeft : st.originalVisionLeft;
            let finalVisionRight = st.manualVisionRight !== null ? st.manualVisionRight : st.originalVisionRight;

            if (fillVision && computed.visionLeft) {
                if (!safeOnlyBlank || !finalVisionLeft) {
                    finalVisionLeft = computed.visionLeft;
                    autoFlags.visionLeft = true;
                    totalVisionFilled++;
                }
                if (!safeOnlyBlank || !finalVisionRight) {
                    finalVisionRight = computed.visionRight;
                    autoFlags.visionRight = true;
                }
            }

            // 4. AWC Parameters
            let finalBirthWeight = st.originalBirthWeight;
            let finalHead = st.originalHead;
            let finalMuac = st.originalMuac;

            if (detectedType === "AWC" || computed.isUnderFive) {
                if (fillHead && computed.head && (!safeOnlyBlank || !finalHead)) {
                    finalHead = computed.head;
                    autoFlags.head = true;
                }
                if (fillMuac && computed.muac && (!safeOnlyBlank || !finalMuac)) {
                    finalMuac = computed.muac;
                    autoFlags.muac = true;
                }
                if (fillBirthWeight && computed.birthWeight && (!safeOnlyBlank || !finalBirthWeight)) {
                    finalBirthWeight = computed.birthWeight;
                    autoFlags.birthWeight = true;
                }
            }

            // 5. Defect / Health Status
            let finalDefectKey = "healthy";
            if (st.manualDefect !== null) {
                finalDefectKey = st.manualDefect;
            } else if (fillDefect) {
                if (defectMode === "auto_hb") {
                    finalDefectKey = determineDefectFromHb(finalHb, st.originalDefect);
                } else if (defectMode === "healthy") {
                    finalDefectKey = "healthy";
                } else if (defectMode === "specific") {
                    finalDefectKey = specificDefectVal;
                }
            } else {
                finalDefectKey = normalizeDefectKey(st.originalDefect) || "healthy";
            }

            st.currentDefectKey = finalDefectKey;
            st.defect = DEFECT_LABELS[finalDefectKey] || "Healthy";

            st.computed = {
                weight: finalWeight,
                height: finalHeight,
                hb: finalHb,
                visionLeft: finalVisionLeft,
                visionRight: finalVisionRight,
                birthWeight: finalBirthWeight,
                head: finalHead,
                muac: finalMuac,
                defect: st.defect
            };
            st.autoFlags = autoFlags;

            // Reflect into 2D rawGrid directly
            const rIdx = st.rowIndex;
            if (rawGrid[rIdx]) {
                if (colIndexMap.weight >= 0) rawGrid[rIdx][colIndexMap.weight] = finalWeight;
                if (colIndexMap.height >= 0) rawGrid[rIdx][colIndexMap.height] = finalHeight;
                if (colIndexMap.hb >= 0) rawGrid[rIdx][colIndexMap.hb] = finalHb;
                if (colIndexMap.visionLeft >= 0) rawGrid[rIdx][colIndexMap.visionLeft] = finalVisionLeft;
                if (colIndexMap.visionRight >= 0) rawGrid[rIdx][colIndexMap.visionRight] = finalVisionRight;
                if (colIndexMap.birthWeight >= 0) rawGrid[rIdx][colIndexMap.birthWeight] = finalBirthWeight;
                if (colIndexMap.head >= 0) rawGrid[rIdx][colIndexMap.head] = finalHead;
                if (colIndexMap.muac >= 0) rawGrid[rIdx][colIndexMap.muac] = finalMuac;
                if (colIndexMap.defect >= 0) rawGrid[rIdx][colIndexMap.defect] = st.defect;
            }
        });

        renderStats({
            total: studentRowsData.length,
            weight: totalWeightFilled,
            height: totalHeightFilled,
            hb: totalHbFilled,
            vision: totalVisionFilled
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
                    <span class="stat-value">${stats.hb} Hb / ${stats.vision} Vis</span>
                    <span class="stat-label">Hb & Vision Auto-Filled</span>
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
            const defect = String(st.defect || "").toLowerCase();
            return name.includes(activeFilter) || id.includes(activeFilter) || sn.includes(activeFilter) || defect.includes(activeFilter);
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
                <th>Vision (L)</th>
                <th>Vision (R)</th>
                ${detectedType === "AWC" ? "<th>Birth Wt</th><th>Head</th><th>MUAC</th>" : ""}
                <th>Defect / Health Status</th>
            </tr>
        `;
        previewThead.innerHTML = theadHtml;

        const startIdx = (currentPage - 1) * rowsPerPage;
        const pageItems = filtered.slice(startIdx, startIdx + rowsPerPage);

        if (pageItems.length === 0) {
            const cols = detectedType === "AWC" ? 14 : 11;
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
                    <td>${renderVisionSelectHtml(comp.visionLeft, st.rowIndex, "left")}</td>
                    <td>${renderVisionSelectHtml(comp.visionRight, st.rowIndex, "right")}</td>
                    ${detectedType === "AWC" ? `${cell(comp.birthWeight, flags.birthWeight)}${cell(comp.head, flags.head)}${cell(comp.muac, flags.muac)}` : ""}
                    <td>${renderDiseaseSelectHtml(st.currentDefectKey || "healthy", st.rowIndex)}</td>
                </tr>
            `;
        });

        previewTbody.innerHTML = tbodyHtml;

        // Attach live listeners to Defect dropdowns
        previewTbody.querySelectorAll(".rbsk-disease-select").forEach((sel) => {
            sel.addEventListener("change", (e) => {
                const rIdx = parseInt(e.target.getAttribute("data-row-idx"), 10);
                const selectedVal = e.target.value;
                const st = studentRowsData.find((s) => s.rowIndex === rIdx);
                if (st) {
                    st.manualDefect = selectedVal;
                    st.currentDefectKey = selectedVal;
                    st.defect = DEFECT_LABELS[selectedVal] || "Healthy";
                    if (st.computed) st.computed.defect = st.defect;
                    if (colIndexMap.defect >= 0 && rawGrid[rIdx]) {
                        rawGrid[rIdx][colIndexMap.defect] = st.defect;
                    }
                    updateSelectStyle(sel, selectedVal);
                }
            });
        });

        // Attach live listeners to Vision dropdowns
        previewTbody.querySelectorAll(".rbsk-vision-select").forEach((sel) => {
            sel.addEventListener("change", (e) => {
                const rIdx = parseInt(e.target.getAttribute("data-row-idx"), 10);
                const eyeSide = e.target.getAttribute("data-eye");
                const selectedVal = e.target.value;
                const st = studentRowsData.find((s) => s.rowIndex === rIdx);
                if (st) {
                    if (eyeSide === "left") {
                        st.manualVisionLeft = selectedVal;
                        if (st.computed) st.computed.visionLeft = selectedVal;
                        if (colIndexMap.visionLeft >= 0 && rawGrid[rIdx]) {
                            rawGrid[rIdx][colIndexMap.visionLeft] = selectedVal;
                        }
                    } else {
                        st.manualVisionRight = selectedVal;
                        if (st.computed) st.computed.visionRight = selectedVal;
                        if (colIndexMap.visionRight >= 0 && rawGrid[rIdx]) {
                            rawGrid[rIdx][colIndexMap.visionRight] = selectedVal;
                        }
                    }
                }
            });
        });

        if (pageIndicator) {
            pageIndicator.textContent = `Page ${currentPage} of ${totalPages} (${filtered.length} students)`;
        }
        if (prevPageBtn) prevPageBtn.disabled = currentPage <= 1;
        if (nextPageBtn) nextPageBtn.disabled = currentPage >= totalPages;
    }

    /**
     * Download Processed Excel preserving dynamic format, dropdowns, and text columns
     */
    async function downloadProcessedExcel() {
        if (!rawGrid || !rawGrid.length) {
            alert("No data available to download.");
            return;
        }

        try {
            // Convert updated 2D rawGrid to sheet
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

            const vlColIdx = colIndexMap.visionLeft;
            const vrColIdx = colIndexMap.visionRight;
            const defectColIdx = colIndexMap.defect;

            // Ensure Vision cells have explicit String type 's' and Text format '@'
            const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
            for (let R = 1; R <= range.e.r; R++) {
                if (vlColIdx >= 0) {
                    const addrL = XLSX.utils.encode_cell({ r: R, c: vlColIdx });
                    if (!worksheet[addrL]) worksheet[addrL] = { t: "s", v: "", z: "@" };
                    else { worksheet[addrL].t = "s"; worksheet[addrL].z = "@"; }
                }
                if (vrColIdx >= 0) {
                    const addrR = XLSX.utils.encode_cell({ r: R, c: vrColIdx });
                    if (!worksheet[addrR]) worksheet[addrR] = { t: "s", v: "", z: "@" };
                    else { worksheet[addrR].t = "s"; worksheet[addrR].z = "@"; }
                }
            }

            const sheetTitle = detectedType === "AWC" ? "AWC Students" : "School Students";
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetTitle);

            // Build comprehensive Instructions guide sheet
            const screeningInstructions = [
                ["RBSK SCREENING AUTO-FILLED TEMPLATE - USER GUIDE & INSTRUCTIONS"],
                [""],
                ["Column Name", "Required?", "Accepted Values / Format", "Example", "Detailed Instructions / Rules"],
                ["Defect / Health Status", "Optional", "Healthy / Specific Defect Name", "Dental", "Healthy ya specific bimari chunein: Healthy, Skin, Dental, Otitis, Vision, Vit A, Vit D, B Complex, Mild Anaemia, Severe Anemia, Moderate Anaemia, Reactive Airway Disease, Growing up concerns, Pain/Burning Urinating, Pain during menstruation."],
                ["Weight (kg)", "Optional", "Number (kg)", "18.5", "Bachhe ka wazan kg me (jaise: 18, 22.5). WHO / RBSK reference ke anusaar auto-calculated."],
                ["Height (cm)", "Optional", "Number (cm)", "112", "Bachhe ki unchai cm me (jaise: 110, 125.4). WHO / RBSK reference ke anusaar auto-calculated."],
                ["Hb", "Optional", "Number (g/dL)", "12.0", "Hemoglobin value g/dL me (Anemia cases ke liye)."],
                ["Vision Left", "Optional", "6/6 / 6/9 / 6/12 / 6/18 / 6/24 / 6/36 / 6/60 / No vision", "6/6", "Left eye (bayein aankh) ka vision dropdown."],
                ["Vision Right", "Optional", "6/6 / 6/9 / 6/12 / 6/18 / 6/24 / 6/36 / 6/60 / No vision", "6/6", "Right eye (dayein aankh) ka vision dropdown."],
                ["Head Circumference", "Optional (AWC)", "Number (cm)", "48.5", "Sirf Anganwadi (AWC) bachho ke liye (WHO Normal Range)."],
                ["MUAC", "Optional (AWC)", "Number (cm)", "13.5", "Mid-Upper Arm Circumference cm me (WHO Green Range)."],
                ["Birth Weight", "Optional (AWC)", "Number (kg)", "2.8", "Janm ke samay wazan kg me (AWC DOB aadharit)."],
                ["DOB", "Reference", "DD-MM-YYYY", "15-08-2015", "Student matching ke liye portal se scraped DOB."],
                ["PEN ID / Child ID", "Reference", "Unique ID", "XXXXXXXXXXX", "Student ka unique ID."],
                [""],
                ["IMPORTANT NOTES / KHAAS BAATEIN:"],
                ["1. Is downloaded Excel file me Defect, Vision Left aur Vision Right ke live dropdowns active hain."],
                ["2. Aap Excel me bhi kisi bhi bachhe ka wazan, lambai, defect ya vision badal sakte hain, koi restriction nahi hai."],
                ["3. Badlav ke baad is file ko direct RBSK screening portal me feed kar sakte hain!"]
            ];

            const guideSheet = XLSX.utils.aoa_to_sheet(screeningInstructions);
            guideSheet["!cols"] = [
                { wch: 22 },
                { wch: 14 },
                { wch: 35 },
                { wch: 20 },
                { wch: 80 }
            ];
            XLSX.utils.book_append_sheet(workbook, guideSheet, "Instructions");

            const cleanName = originalFileName.replace(/\.(xlsx|xls)$/i, "");
            const outFileName = `AutoFilled_${cleanName}.xlsx`;

            // Inject Dropdowns (Defect, Vision Left, Vision Right) and Text Column Format using JSZip
            if (typeof JSZip !== "undefined") {
                const maxRow = Math.max(rawGrid.length + 50, 500);
                const xlsxArray = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
                const zip = await JSZip.loadAsync(xlsxArray);

                let stylesXml = zip.file("xl/styles.xml") ? await zip.file("xl/styles.xml").async("text") : "";
                let sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");

                // Ensure Text format (numFmtId="49") in styles.xml
                let textStyleIdx = 1;
                if (stylesXml) {
                    if (!stylesXml.includes('numFmtId="49"')) {
                        const cellXfsMatch = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
                        if (cellXfsMatch) {
                            const oldCount = parseInt(cellXfsMatch[1], 10);
                            textStyleIdx = oldCount;
                            const newCellXfs = `<cellXfs count="${oldCount + 1}">${cellXfsMatch[2]}<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`;
                            stylesXml = stylesXml.replace(cellXfsMatch[0], newCellXfs);
                            zip.file("xl/styles.xml", stylesXml);
                        }
                    }
                }

                // Apply text format to Vision columns so 6/6 is never parsed as a date
                if (vlColIdx >= 0 || vrColIdx >= 0) {
                    const minC = Math.min(vlColIdx >= 0 ? vlColIdx + 1 : 999, vrColIdx >= 0 ? vrColIdx + 1 : 999);
                    const maxC = Math.max(vlColIdx >= 0 ? vlColIdx + 1 : 0, vrColIdx >= 0 ? vrColIdx + 1 : 0);
                    const colTag = `<col min="${minC}" max="${maxC}" width="16" style="${textStyleIdx}" customWidth="1"/>`;
                    if (sheetXml.includes("<cols>")) {
                        sheetXml = sheetXml.replace("<cols>", `<cols>${colTag}`);
                    } else if (sheetXml.includes("<sheetData>")) {
                        sheetXml = sheetXml.replace("<sheetData>", `<cols>${colTag}</cols><sheetData>`);
                    }
                }

                // Build Data Validation rules
                const dvRules = [];

                if (defectColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(defectColIdx);
                    const dvDefectOptions = DEFECT_OPTIONS.map((o) => o.label).join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvDefectOptions}&quot;</formula1></dataValidation>`);
                }

                if (vlColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(vlColIdx);
                    const dvVisionOptions = VISION_OPTIONS.join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvVisionOptions}&quot;</formula1></dataValidation>`);
                }

                if (vrColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(vrColIdx);
                    const dvVisionOptions = VISION_OPTIONS.join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvVisionOptions}&quot;</formula1></dataValidation>`);
                }

                if (dvRules.length) {
                    const dvXml = `<dataValidations count="${dvRules.length}">${dvRules.join("")}</dataValidations>`;
                    if (sheetXml.includes("<ignoredErrors>")) {
                        sheetXml = sheetXml.replace("<ignoredErrors>", dvXml + "<ignoredErrors>");
                    } else if (sheetXml.includes("<pageMargins>")) {
                        sheetXml = sheetXml.replace("<pageMargins>", dvXml + "<pageMargins>");
                    } else {
                        sheetXml = sheetXml.replace("</worksheet>", dvXml + "</worksheet>");
                    }
                }

                zip.file("xl/worksheets/sheet1.xml", sheetXml);
                const blob = await zip.generateAsync({
                    type: "blob",
                    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                });

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
