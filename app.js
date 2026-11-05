/**
 * RBSK Screening Template Auto-Filler Web App Logic
 * Enhanced 2D Grid Preservation Engine for 100% Excel Template Compatibility
 * 
 * Features:
 * - Simple, One-Click Toggle Buttons (Height, Weight, Vision, Hb, Defect, Head, MUAC, Birth Weight)
 * - Complete 4D & RBSK Defect List (Healthy, Skin, Dental, Otitis, Vision, Vit A/B/D, SAM, Anaemia, RAD, etc.)
 * - Preserved Referred Facility Type Dropdowns in Excel
 * - Calibrated WHO 0-60m Medians to naturally prevent Stunting / Wasting on RBSK Portal
 * - Dynamic XML Injection via JSZip for Defect, Vision, and Referred Facility Dropdowns
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
    let lastAutoStats = null;

            // RBSK Comprehensive Disease / Condition Reference List (All 4Ds & Adolescent concerns)
    const RBSK_SCREENING_DEFECT_CATEGORIES = [
        {
            category: "1. Defects at Birth",
            items: [
                { val: "neuralTubeDefect", label: "1. Neural Tube Defect", facility: "DEIC" },
                { val: "downsSyndrome", label: "2. Down's Syndrome", facility: "DEIC" },
                { val: "cleftLipPalate", label: "3. Cleft Lip & Palate", facility: "DEIC" },
                { val: "clubFoot", label: "4. Talipes (Club Foot)", facility: "DEIC" },
                { val: "hipDysplasia", label: "5. Developmental Dysplasia of Hip", facility: "DEIC" },
                { val: "congenitalCataract", label: "6. Congenital Cataract", facility: "DEIC" },
                { val: "congenitalDeafness", label: "7. Congenital Deafness", facility: "DEIC" },
                { val: "congenitalHeartDisease", label: "8. Congenital Heart Disease", facility: "DEIC" },
                { val: "retinopathy", label: "9. Retinopathy of Prematurity", facility: "District Level Facility" },
                { val: "microcephaly", label: "42. Microcephaly", facility: "DEIC" },
                { val: "macrocephaly", label: "43. Macrocephaly", facility: "DEIC" },
                { val: "congenitalEar", label: "47. Congenital Ear Defects", facility: "DEIC" },
                { val: "neckFaceDefects", label: "48. Neck & Face Defects", facility: "DEIC" },
                { val: "congenitalEye", label: "50. Congenital Eye Defects", facility: "DEIC" }
            ]
        },
        {
            category: "2. Deficiencies",
            items: [
                { val: "severeAnemia", label: "10. Severe Anemia", facility: "DEIC" },
                { val: "vitaminA", label: "11. Vit A Deficiency (Bitot Spot)", facility: "Block Level Facility" },
                { val: "vitaminD", label: "12. Vit D Deficiency (Rickets)", facility: "Block Level Facility" },
                { val: "sam", label: "13. SAM / Severe Thinness / SUW", facility: "NRC" },
                { val: "goiter", label: "14. Goiter", facility: "Block Level Facility" },
                { val: "severeStunting", label: "41. Severe Stunting", facility: "NRC" },
                { val: "vitaminB", label: "44. Vit B Complex Deficiency", facility: "Block Level Facility" },
                { val: "obesity", label: "45. Obesity", facility: "Block Level Facility" },
                { val: "moderateAnaemia", label: "52. Moderate Anaemia", facility: "Block Level Facility" },
                { val: "mildAnaemia", label: "89. Mild Anaemia", facility: "Block Level Facility" },
                { val: "mam", label: "90. MAM", facility: "Block Level Facility" },
                { val: "vitaminD120", label: "120. Vitamin D Deficiency", facility: "Block Level Facility" }
            ]
        },
        {
            category: "3. Childhood Diseases",
            items: [
                { val: "skin", label: "15. Skin Conditions", facility: "Block Level Facility" },
                { val: "otitis", label: "16. Otitis Media", facility: "Block Level Facility" },
                { val: "rheumaticHeart", label: "17. Rheumatic Heart Disease", facility: "DEIC" },
                { val: "reactiveAirway", label: "18. Reactive Airway Disease", facility: "Block Level Facility" },
                { val: "dental", label: "19. Dental Conditions", facility: "Block Level Facility" },
                { val: "convulsiveDisorders", label: "20. Convulsive Disorders", facility: "DEIC" },
                { val: "childhoodLeprosy", label: "39. Childhood Leprosy", facility: "Block Level Facility" },
                { val: "tbPulmonary", label: "40. Tuberculosis Pulmonary", facility: "Block Level Facility" },
                { val: "stage2Hypertension", label: "46. Stage 2 Hypertension", facility: "District Level Facility" },
                { val: "refractiveProblems", label: "49. Refractive Problems", facility: "DEIC" },
                { val: "tbExtraPulmonary", label: "140. TB Extra Pulmonary", facility: "District Level Facility" },
                { val: "skinNotLeprosy", label: "150. Skin Conditions (Not Leprosy)", facility: "Block Level Facility" },
                { val: "bronchialAsthma", label: "180. Bronchial Asthma", facility: "Block Level Facility" },
                { val: "prehypertension", label: "460. Prehypertension", facility: "Block Level Facility" },
                { val: "stage1Hypertension", label: "461. Stage 1 Hypertension", facility: "Block Level Facility" }
            ]
        },
        {
            category: "4. Development Delay",
            items: [
                { val: "vision", label: "21. Vision Impairment", facility: "DEIC" },
                { val: "hearingImpairment", label: "22. Hearing Impairment", facility: "DEIC" },
                { val: "neuroMotor", label: "23. Neuro-Motor Impairment", facility: "DEIC" },
                { val: "motorDelay", label: "24. Motor Delay", facility: "DEIC" },
                { val: "cognitiveDelay", label: "25. Cognitive Delay", facility: "DEIC" },
                { val: "languageDelay", label: "26. Language Delay", facility: "DEIC" },
                { val: "autism", label: "27. Behavior Disorder (Autism)", facility: "DEIC" },
                { val: "learningDisorder", label: "28. Learning Disorder", facility: "DEIC" },
                { val: "adhd", label: "29. ADHD", facility: "DEIC" },
                { val: "convulsiveDevelopmental", label: "51. Convulsive Disorders (Dev Delay)", facility: "DEIC" }
            ]
        },
        {
            category: "5. Adolescent Health concerns",
            items: [
                { val: "growingUpConcerns", label: "31. Growing Up Concerns", facility: "AFHC" },
                { val: "substanceAbuse", label: "32. Substance Abuse", facility: "AFHC" },
                { val: "feelDepressed", label: "33. Feel Depressed", facility: "AFHC" },
                { val: "delayMenstruation", label: "34. Delay in Menstruation", facility: "AFHC" },
                { val: "irregularPeriods", label: "35. Irregular Periods", facility: "AFHC" },
                { val: "painUrination", label: "36. Pain/Burning Urinating", facility: "AFHC" },
                { val: "genitorUrinaryDischarge", label: "37. Genito-Urinary Discharge", facility: "AFHC" },
                { val: "painMenstruation", label: "38. Pain During Menstruation", facility: "AFHC" }
            ]
        }
    ];

    // Flat options list for processing and lookup
    const DEFECT_OPTIONS = [
        { val: "healthy", label: "Healthy", color: "#4b5563", facility: "" }
    ];
    const DEFECT_LABELS = { "healthy": "Healthy" };
    const DEFECT_FACILITIES = { "healthy": "" };

    RBSK_SCREENING_DEFECT_CATEGORIES.forEach((cat) => {
        cat.items.forEach((item) => {
            const entry = {
                val: item.val,
                label: item.label,
                color: "#dc2626",
                facility: item.facility || "Block Level Facility"
            };
            DEFECT_OPTIONS.push(entry);
            DEFECT_LABELS[item.val] = item.label;
            DEFECT_LABELS[item.label] = item.label;
            DEFECT_FACILITIES[item.val] = entry.facility;
            DEFECT_FACILITIES[item.label] = entry.facility;
        });
    });

    // Backward-compatible aliases for legacy keys
    DEFECT_LABELS["downSyndrome"] = "2. Down's Syndrome";
    DEFECT_LABELS["cleftLip"] = "3. Cleft Lip & Palate";
    DEFECT_LABELS["chd"] = "8. Congenital Heart Disease";
    DEFECT_LABELS["convulsive"] = "20. Convulsive Disorders";
    DEFECT_LABELS["convulsiveDevDelay"] = "51. Convulsive Disorders (Dev Delay)";

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

    // Referred Facility Options
    const FACILITY_OPTIONS = [
        "Block Level Facility",
        "District Level Facility",
        "State Level Facility",
        "DEIC",
        "NRC",
        "AFHC"
    ];

    function normalizeDefectKey(raw) {
        if (!raw) return "healthy";
        const s = String(raw).trim();
        const sClean = s.toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!sClean || sClean === "healthy" || sClean === "nil" || sClean === "none" || sClean === "normal" || sClean === "no" || sClean === "0" || sClean === "fit" || sClean === "swasth" || sClean === "theek") {
            return "healthy";
        }
        for (const opt of DEFECT_OPTIONS) {
            if (opt.val.toLowerCase() === sClean || opt.label.toLowerCase() === s.toLowerCase() || opt.label.toLowerCase().replace(/[^a-z0-9]/g, "") === sClean) {
                return opt.val;
            }
        }
        // Substring / keyword fallback
        for (const opt of DEFECT_OPTIONS) {
            const lLow = opt.label.toLowerCase();
            if (lLow.includes(s.toLowerCase()) || s.toLowerCase().includes(opt.val.toLowerCase())) {
                return opt.val;
            }
        }
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
        const isHealthy = currentVal === "healthy" || currentVal === "Healthy";
        const borderStyle = isHealthy ? "border: 1px solid #d1d5db;" : "border: 1px solid rgb(249, 115, 22);";
        const colorStyle = isHealthy ? "color: #4b5563;" : "color: #dc2626; font-weight: 700;";

        let html = `<select class="rbsk-disease-select" data-row-idx="${rowIdx}" style="${borderStyle} ${colorStyle}">`;
        html += `<option value="healthy" style="color: #4b5563;"${(currentVal === "healthy" || currentVal === "Healthy") ? " selected" : ""}>Healthy</option>`;

        RBSK_SCREENING_DEFECT_CATEGORIES.forEach((grp) => {
            html += `<optgroup label="${grp.category}">`;
            grp.items.forEach((opt) => {
                const sel = (opt.val === currentVal || opt.label === currentVal) ? " selected" : "";
                html += `<option value="${opt.val}" style="color: #dc2626;"${sel}>${opt.label}</option>`;
            });
            html += `</optgroup>`;
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

    function renderFacilitySelectHtml(currentVal, rowIdx) {
        const val = String(currentVal || "").trim();
        let html = `<select class="rbsk-facility-select" data-row-idx="${rowIdx}" style="border: 1px solid #d1d5db; color: #1e293b;">`;
        html += `<option value=""${!val ? " selected" : ""}>- (None)</option>`;
        FACILITY_OPTIONS.forEach((opt) => {
            const sel = opt.toLowerCase() === val.toLowerCase() ? " selected" : "";
            html += `<option value="${opt}"${sel}>${opt}</option>`;
        });
        html += `</select>`;
        return html;
    }

    function updateSelectStyle(sel, val) {
        if (!sel) return;
        if (val === "healthy") {
            sel.style.borderColor = "#d1d5db";
            sel.style.color = "#4b5563";
            sel.style.fontWeight = "600";
        } else {
            sel.style.borderColor = "rgb(249, 115, 22)";
            sel.style.color = "#dc2626";
            sel.style.fontWeight = "700";
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

    const schoolButtonsBar = document.getElementById("schoolButtonsBar");
    const awcButtonsBar = document.getElementById("awcButtonsBar");
    const activeModeBadge = document.getElementById("activeModeBadge");

    // Defect & Hb Controls
    const optDefectMode = document.getElementById("optDefectMode");
    const optSpecificDefect = document.getElementById("optSpecificDefect");
    const optSafeOnlyBlank = document.getElementById("optSafeOnlyBlank");
    const optDecimals = document.getElementById("optDecimals");
    const optHbRange = document.getElementById("optHbRange");
    const customHbRow = document.getElementById("customHbRow");
    const customHbMin = document.getElementById("customHbMin");
    const customHbMax = document.getElementById("customHbMax");

    function isButtonActive(btnId) {
        const btn = document.getElementById(btnId);
        return Boolean(btn && btn.classList.contains("active"));
    }

    function toggleButtonState(btn) {
        if (!btn) return;
        const isActive = btn.classList.contains("active");
        if (isActive) {
            btn.classList.remove("active");
            const icon = btn.querySelector(".toggle-icon");
            if (icon) icon.textContent = "✕";
        } else {
            btn.classList.add("active");
            const icon = btn.querySelector(".toggle-icon");
            if (icon) icon.textContent = "✓";
        }
        executeAutoFill();
    }

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

        // Toggle buttons listeners
        document.querySelectorAll(".btn-toggle").forEach((btn) => {
            btn.addEventListener("click", () => toggleButtonState(btn));
        });

        // Hb range listener
        if (optHbRange) {
            optHbRange.addEventListener("change", () => {
                if (customHbRow) {
                    customHbRow.style.display = optHbRange.value === "custom" ? "inline-flex" : "none";
                }
                executeAutoFill();
            });
        }

        // Defect mode listener
        if (optDefectMode) {
            optDefectMode.addEventListener("change", () => {
                if (optSpecificDefect) {
                    optSpecificDefect.style.display = optDefectMode.value === "specific" ? "inline-block" : "none";
                }
                executeAutoFill();
            });
        }

        if (optSpecificDefect) {
            optSpecificDefect.addEventListener("change", executeAutoFill);
        }

        if (optSafeOnlyBlank) optSafeOnlyBlank.addEventListener("change", executeAutoFill);
        if (optDecimals) optDecimals.addEventListener("change", executeAutoFill);
        if (customHbMin) customHbMin.addEventListener("input", executeAutoFill);
        if (customHbMax) customHbMax.addEventListener("input", executeAutoFill);

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

                if (activeModeBadge) {
                    activeModeBadge.textContent = `${detectedType} Mode`;
                    activeModeBadge.className = `meta-badge ${detectedType.toLowerCase()}`;
                }

                // Show School vs AWC button toolbar
                if (detectedType === "AWC") {
                    if (schoolButtonsBar) schoolButtonsBar.style.display = "none";
                    if (awcButtonsBar) awcButtonsBar.style.display = "flex";
                } else {
                    if (schoolButtonsBar) schoolButtonsBar.style.display = "flex";
                    if (awcButtonsBar) awcButtonsBar.style.display = "none";
                }

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
     * Map column indexes for fast lookup and ensure Defect, Vision, and Facility columns exist
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
            facilityType: findColIdx("referredfacilitytype", "facilitytype", "referredfacility", "facility"),
            facilityName: findColIdx("referredfacilityname", "facilityname"),
            birthWeight: findColIdx("birthweight", "birthwt", "bweight"),
            head: findColIdx("headcur", "headcircumference", "head"),
            muac: findColIdx("muac"),
            mobile: findColIdx("mobilenumber", "mobile", "phone", "contact"),
            father: findColIdx("fathersname", "fathername", "father", "guardian"),
            mother: findColIdx("mothersname", "mothername", "mother")
        };

        // Ensure Defect column exists
        if (colIndexMap.defect < 0) {
            colIndexMap.defect = headerKeys.length;
            headerKeys.push("Defect / Health Status");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.defect] = "Defect / Health Status";
            }
        }

        // Ensure Vision columns exist
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

        // Ensure Referred Facility columns exist
        if (colIndexMap.facilityType < 0) {
            colIndexMap.facilityType = headerKeys.length;
            headerKeys.push("Referred Facility Type");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.facilityType] = "Referred Facility Type";
            }
        }
        if (colIndexMap.facilityName < 0) {
            colIndexMap.facilityName = headerKeys.length;
            headerKeys.push("Referred Facility Name");
            if (rawGrid[headerRowIdx]) {
                rawGrid[headerRowIdx][colIndexMap.facilityName] = "Referred Facility Name";
            }
        }

        console.log("[RBSK AutoFill] Detected Type:", detectedType, "Column Mapping:", colIndexMap);
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
            const rawFacType = colIndexMap.facilityType >= 0 ? String(row[colIndexMap.facilityType] || "").trim() : "";
            const rawFacName = colIndexMap.facilityName >= 0 ? String(row[colIndexMap.facilityName] || "").trim() : "";

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
                originalFacilityType: rawFacType,
                manualFacilityType: null,
                originalFacilityName: rawFacName,
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
     * Auto-Fill Calculation Engine based on One-Click Toggle Buttons
     */
    function executeAutoFill() {
        if (!studentRowsData || !studentRowsData.length) return;

        const isAwc = detectedType === "AWC";

        // Read active buttons based on mode
        let fillHeight = false;
        let fillWeight = false;
        let fillVision = false;
        let fillHb = false;
        let fillHead = false;
        let fillMuac = false;
        let fillBirthWeight = false;

        if (isAwc) {
            fillHeight = isButtonActive("btnToggleAwcHeight");
            fillWeight = isButtonActive("btnToggleAwcWeight");
            fillHead = isButtonActive("btnToggleAwcHead");
            fillMuac = isButtonActive("btnToggleAwcMuac");
            fillBirthWeight = isButtonActive("btnToggleAwcBirthWeight");
            fillHb = isButtonActive("btnToggleAwcHb");
            fillVision = false; // Never auto-fill vision for AWC (RBSK rules)
        } else {
            fillHeight = isButtonActive("btnToggleHeight");
            fillWeight = isButtonActive("btnToggleWeight");
            fillVision = isButtonActive("btnToggleVision");
            fillHb = isButtonActive("btnToggleHb");
        }

        const safeOnlyBlank = optSafeOnlyBlank ? optSafeOnlyBlank.checked : true;
        const defectMode = optDefectMode ? optDefectMode.value : "auto_hb";
        const specificDefectVal = optSpecificDefect ? optSpecificDefect.value : "healthy";

        const options = {
            fillHeight,
            fillWeight,
            fillHb,
            fillVision,
            defaultVision: "6/6",
            awcVision: false,
            fillBirthWeight,
            fillHead,
            fillMuac,
            allowDecimal: optDecimals ? optDecimals.checked : false,
            naturalVariation: true,
            hbRange: optHbRange ? optHbRange.value : "normal",
            customHbMin: customHbMin ? parseFloat(customHbMin.value) || 11.5 : 11.5,
            customHbMax: customHbMax ? parseFloat(customHbMax.value) || 13.5 : 13.5,
            isAwc: isAwc
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
                isAwc: isAwc
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

            if (isAwc || computed.isUnderFive) {
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
            } else {
                if (defectMode === "auto_hb") {
                    finalDefectKey = determineDefectFromHb(finalHb, st.originalDefect);
                } else if (defectMode === "healthy") {
                    finalDefectKey = "healthy";
                } else if (defectMode === "specific") {
                    finalDefectKey = specificDefectVal;
                }
            }

            st.currentDefectKey = finalDefectKey;
            st.defect = DEFECT_LABELS[finalDefectKey] || "Healthy";

            // 6. Referred Facility
            let finalFacilityType = st.manualFacilityType !== null ? st.manualFacilityType : st.originalFacilityType;
            if (!finalFacilityType && finalDefectKey !== "healthy") {
                finalFacilityType = DEFECT_FACILITIES[finalDefectKey] || "Block Level Facility";
            } else if (finalDefectKey === "healthy" && !st.manualFacilityType) {
                finalFacilityType = "";
            }

            st.computed = {
                weight: finalWeight,
                height: finalHeight,
                hb: finalHb,
                visionLeft: finalVisionLeft,
                visionRight: finalVisionRight,
                birthWeight: finalBirthWeight,
                head: finalHead,
                muac: finalMuac,
                defect: st.defect,
                facilityType: finalFacilityType
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
                if (colIndexMap.facilityType >= 0) rawGrid[rIdx][colIndexMap.facilityType] = finalFacilityType;
            }
        });

        lastAutoStats = {
            total: studentRowsData.length,
            weight: totalWeightFilled,
            height: totalHeightFilled,
            hb: totalHbFilled,
            vision: totalVisionFilled
        };
        renderStats(lastAutoStats);

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
                <th style="width: 75px; text-align: center;">Action</th>
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
                <th>Referred Facility</th>
            </tr>
        `;
        previewThead.innerHTML = theadHtml;

        const startIdx = (currentPage - 1) * rowsPerPage;
        const pageItems = filtered.slice(startIdx, startIdx + rowsPerPage);

        if (pageItems.length === 0) {
            const cols = detectedType === "AWC" ? 16 : 13;
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
                    <td style="text-align: center;">
                        <button type="button" class="btn-delete-student" data-row-idx="${st.rowIndex}" title="Delete ${st.name || 'Student'}" style="background: #fee2e2; color: #dc2626; border: 1px solid #fca5a5; padding: 4px 8px; border-radius: 6px; cursor: pointer; font-size: 11px; font-weight: 700; display: inline-flex; align-items: center; gap: 3px; white-space: nowrap; transition: all 0.15s ease;">
                            <span>🗑️</span> Delete
                        </button>
                    </td>
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
                    <td>${renderFacilitySelectHtml(comp.facilityType || "", st.rowIndex)}</td>
                </tr>
            `;
        });

        previewTbody.innerHTML = tbodyHtml;

        // Attach live listeners to Delete buttons
        previewTbody.querySelectorAll(".btn-delete-student").forEach((btn) => {
            btn.addEventListener("click", (e) => {
                const targetBtn = e.target.closest(".btn-delete-student");
                if (!targetBtn) return;
                const rIdx = parseInt(targetBtn.getAttribute("data-row-idx"), 10);
                const stIndex = studentRowsData.findIndex((s) => s.rowIndex === rIdx);
                if (stIndex === -1) return;

                const st = studentRowsData[stIndex];
                const stName = st.name || "Student";
                if (confirm(`Kya aap student "${stName}" (S.N. ${st.sn}) ko delete karna chahte hain?`)) {
                    // 1. Remove from studentRowsData
                    studentRowsData.splice(stIndex, 1);

                    // 2. Remove row from rawGrid
                    if (rawGrid && rawGrid[rIdx]) {
                        rawGrid.splice(rIdx, 1);
                    }

                    // 3. Shift rowIndex of all remaining students below rIdx
                    studentRowsData.forEach((item) => {
                        if (item.rowIndex > rIdx) {
                            item.rowIndex--;
                        }
                    });

                    // 4. Re-number S.N. cleanly
                    studentRowsData.forEach((item, idx) => {
                        item.sn = idx + 1;
                        if (colIndexMap.sn >= 0 && rawGrid[item.rowIndex]) {
                            rawGrid[item.rowIndex][colIndexMap.sn] = idx + 1;
                        }
                    });

                    // 5. Update stats and file badge
                    if (fileMetaBadge) {
                        fileMetaBadge.textContent = `${detectedType} Template • ${studentRowsData.length} Students`;
                    }
                    if (lastAutoStats) {
                        lastAutoStats.total = studentRowsData.length;
                        renderStats(lastAutoStats);
                    }

                    // 6. Adjust page if page is now empty
                    const maxPages = Math.ceil(studentRowsData.length / rowsPerPage) || 1;
                    if (currentPage > maxPages) currentPage = maxPages;

                    // 7. Re-render table
                    renderPreviewTable();
                }
            });
        });

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

                    // Update facility automatically if non-healthy
                    if (selectedVal !== "healthy" && (!st.computed.facilityType || st.computed.facilityType === "")) {
                        let autoFac = DEFECT_FACILITIES[selectedVal] || "Block Level Facility";
                        st.manualFacilityType = autoFac;
                        if (st.computed) st.computed.facilityType = autoFac;
                        if (colIndexMap.facilityType >= 0 && rawGrid[rIdx]) {
                            rawGrid[rIdx][colIndexMap.facilityType] = autoFac;
                        }
                        const facSel = e.target.closest("tr").querySelector(".rbsk-facility-select");
                        if (facSel) facSel.value = autoFac;
                    }
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

        // Attach live listeners to Facility dropdowns
        previewTbody.querySelectorAll(".rbsk-facility-select").forEach((sel) => {
            sel.addEventListener("change", (e) => {
                const rIdx = parseInt(e.target.getAttribute("data-row-idx"), 10);
                const selectedVal = e.target.value;
                const st = studentRowsData.find((s) => s.rowIndex === rIdx);
                if (st) {
                    st.manualFacilityType = selectedVal;
                    if (st.computed) st.computed.facilityType = selectedVal;
                    if (colIndexMap.facilityType >= 0 && rawGrid[rIdx]) {
                        rawGrid[rIdx][colIndexMap.facilityType] = selectedVal;
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
            const facColIdx = colIndexMap.facilityType;

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
                ["Defect / Health Status", "Optional", "Healthy / Specific Defect Name", "Dental", "Healthy ya specific bimari chunein: Healthy, Skin, Dental, Otitis, Vision, Vit A, Vit D, B Complex, SAM, Mild Anaemia, Severe Anemia, Moderate Anaemia, Reactive Airway Disease, Growing up concerns, Pain/Burning Urinating, Pain during menstruation, Neural Tube Defect, Cleft Lip/Palate, Club Foot, Congenital Heart Disease, Down Syndrome, Convulsive Disorders."],
                ["Referred Facility Type", "Optional", "Block Level Facility / District Level Facility / State Level Facility / DEIC / NRC / AFHC", "Block Level Facility", "Bimari hone par facility type chunein."],
                ["Referred Facility Name", "Optional", "Hospital / CHC / DH ka naam", "Hospital Name", "Jis facility/hospital me refer karna hai uska naam."],
                ["Weight (kg)", "Optional", "Number (kg)", "18.5", "Bachhe ka wazan kg me (jaise: 18, 22.5). WHO / RBSK reference ke anusaar auto-calculated."],
                ["Height (cm)", "Optional", "Number (cm)", "112", "Bachhe ki unchai cm me (jaise: 110, 125.4). WHO / RBSK reference ke anusaar auto-calculated."],
                ["Hb", "Optional", "Number (g/dL)", "12.0", "Hemoglobin value g/dL me."],
                ["Vision Left", "Optional", "6/6 / 6/9 / 6/12 / 6/18 / 6/24 / 6/36 / 6/60 / No vision", "6/6", "Left eye (bayein aankh) ka vision dropdown."],
                ["Vision Right", "Optional", "6/6 / 6/9 / 6/12 / 6/18 / 6/24 / 6/36 / 6/60 / No vision", "6/6", "Right eye (dayein aankh) ka vision dropdown."],
                ["Head Circumference", "Optional (AWC)", "Number (cm)", "48.5", "Sirf Anganwadi (AWC) bachho ke liye (WHO Normal Range)."],
                ["MUAC", "Optional (AWC)", "Number (cm)", "13.5", "Mid-Upper Arm Circumference cm me (WHO Green Range)."],
                ["Birth Weight", "Optional (AWC)", "Number (kg)", "2.8", "Janm ke samay wazan kg me (AWC DOB aadharit)."],
                ["DOB", "Reference", "DD-MM-YYYY", "15-08-2015", "Student matching ke liye portal se scraped DOB."],
                ["PEN ID / Child ID", "Reference", "Unique ID", "XXXXXXXXXXX", "Student ka unique ID."],
                [""],
                ["IMPORTANT NOTES / KHAAS BAATEIN:"],
                ["1. Is downloaded Excel file me Defect, Vision Left, Vision Right aur Referred Facility Type ke live dropdowns active hain."],
                ["2. Aap Excel me bhi kisi bhi bachhe ka wazan, lambai, defect ya vision badal sakte hain, koi restriction nahi hai."],
                ["3. Badlav ke baad is file ko direct RBSK screening portal me feed kar sakte hain!"]
            ];

            const guideSheet = XLSX.utils.aoa_to_sheet(screeningInstructions);
            guideSheet["!cols"] = [
                { wch: 22 },
                { wch: 14 },
                { wch: 35 },
                { wch: 20 },
                { wch: 80 },
                { wch: 4 },
                { wch: 40 }
            ];
            // Populate Column G (c: 6) with full RBSK defect list for dropdown reference without 255-char limit
            guideSheet[XLSX.utils.encode_cell({ r: 0, c: 6 })] = { t: "s", v: "RBSK Defect Reference List" };
            DEFECT_OPTIONS.forEach((opt, idx) => {
                const addr = XLSX.utils.encode_cell({ r: idx + 1, c: 6 });
                guideSheet[addr] = { t: "s", v: opt.label };
            });
            const oldRange = XLSX.utils.decode_range(guideSheet["!ref"] || "A1:E20");
            guideSheet["!ref"] = XLSX.utils.encode_range({
                s: { r: 0, c: 0 },
                e: { r: Math.max(oldRange.e.r, DEFECT_OPTIONS.length + 2), c: Math.max(oldRange.e.c, 6) }
            });
            XLSX.utils.book_append_sheet(workbook, guideSheet, "Instructions");

            const cleanName = originalFileName.replace(/\.(xlsx|xls)$/i, "");
            const outFileName = `AutoFilled_${cleanName}.xlsx`;

            // Inject Dropdowns (Defect, Vision Left, Vision Right, Referred Facility) and Text Column Format using JSZip
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

                // 1. Defect Column
                if (defectColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(defectColIdx);
                    const formulaRef = `Instructions!$G$2:$G${DEFECT_OPTIONS.length + 1}`;
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>${formulaRef}</formula1></dataValidation>`);
                }

                // 2. Vision Left
                if (vlColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(vlColIdx);
                    const dvVisionOptions = VISION_OPTIONS.join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvVisionOptions}&quot;</formula1></dataValidation>`);
                }

                // 3. Vision Right
                if (vrColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(vrColIdx);
                    const dvVisionOptions = VISION_OPTIONS.join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvVisionOptions}&quot;</formula1></dataValidation>`);
                }

                // 4. Referred Facility Type
                if (facColIdx >= 0) {
                    const colLetter = XLSX.utils.encode_col(facColIdx);
                    const dvFacOptions = FACILITY_OPTIONS.join(",");
                    dvRules.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${colLetter}2:${colLetter}${maxRow}"><formula1>&quot;${dvFacOptions}&quot;</formula1></dataValidation>`);
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
