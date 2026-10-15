/**
 * RBSK & WHO Standard Growth Computation Engine
 * Designed for RBSK Screening Auto-Fill Tool
 */

const RBSK_GROWTH_BOYS = {
    // Months (2m to 11m) - WHO Standards (Whole Numbers, No BP, No Hb)
    "2m": { h: 56, w: 6, head: 39, muac: 13, bp: "", hb: "" },
    "3m": { h: 57, w: 6, head: 41, muac: 13, bp: "", hb: "" },
    "4m": { h: 62, w: 7, head: 42, muac: 13, bp: "", hb: "" },
    "5m": { h: 63, w: 8, head: 43, muac: 13, bp: "", hb: "" },
    "6m": { h: 66, w: 8, head: 43, muac: 13, bp: "", hb: "" },
    "7m": { h: 66, w: 8, head: 44, muac: 13, bp: "", hb: "" },
    "8m": { h: 68, w: 9, head: 45, muac: 13, bp: "", hb: "" },
    "9m": { h: 68, w: 9, head: 45, muac: 13, bp: "", hb: "" },
    "10m": { h: 70, w: 9, head: 45, muac: 13, bp: "", hb: "" },
    "11m": { h: 73, w: 9, head: 46, muac: 13, bp: "", hb: "" },

    // Years 1 to 5 - RBSK / Indian Reference Standards (AWC - Whole Numbers)
    1: { h: 76, w: 10, head: 46, muac: 13, bp: "", hb: "" },
    2: { h: 87, w: 12, head: 48, muac: 13, bp: "", hb: "" },
    3: { h: 96, w: 14, head: 49, muac: 13, bp: "", hb: "" },
    4: { h: 103, w: 16, head: 50, muac: 13, bp: "", hb: "" },
    5: { h: 110, w: 18, head: 51, muac: 13, bp: "", hb: "" },

    // Years 6 to 20 - School Students (Realistic Moderate Heights & Weights, Normal BP <120/<80, Hb 12-14)
    6: { h: 110, w: 18, head: null, muac: null, bp: "95/65", hb: 12 },
    7: { h: 115, w: 20, head: null, muac: null, bp: "98/65", hb: 12 },
    8: { h: 120, w: 23, head: null, muac: null, bp: "100/68", hb: 12 },
    9: { h: 125, w: 26, head: null, muac: null, bp: "102/68", hb: 13 },
    10: { h: 130, w: 29, head: null, muac: null, bp: "105/70", hb: 13 },
    11: { h: 138, w: 33, head: null, muac: null, bp: "108/70", hb: 13 },
    12: { h: 143, w: 37, head: null, muac: null, bp: "110/72", hb: 13 },
    13: { h: 149, w: 41, head: null, muac: null, bp: "110/72", hb: 13 },
    14: { h: 155, w: 45, head: null, muac: null, bp: "112/74", hb: 13 },
    15: { h: 160, w: 49, head: null, muac: null, bp: "114/74", hb: 14 },
    16: { h: 164, w: 53, head: null, muac: null, bp: "115/75", hb: 14 },
    17: { h: 166, w: 55, head: null, muac: null, bp: "115/75", hb: 14 },
    18: { h: 168, w: 57, head: null, muac: null, bp: "116/76", hb: 14 },
    19: { h: 169, w: 58, head: null, muac: null, bp: "118/76", hb: 14 },
    20: { h: 170, w: 59, head: null, muac: null, bp: "118/76", hb: 14 }
};

const RBSK_GROWTH_GIRLS = {
    // Months (2m to 11m) - WHO Standards (Whole Numbers, No BP, No Hb)
    "2m": { h: 53, w: 5, head: 38, muac: 13, bp: "", hb: "" },
    "3m": { h: 55, w: 6, head: 40, muac: 13, bp: "", hb: "" },
    "4m": { h: 57, w: 6, head: 41, muac: 13, bp: "", hb: "" },
    "5m": { h: 60, w: 7, head: 42, muac: 13, bp: "", hb: "" },
    "6m": { h: 62, w: 7, head: 42, muac: 13, bp: "", hb: "" },
    "7m": { h: 63, w: 8, head: 43, muac: 13, bp: "", hb: "" },
    "8m": { h: 66, w: 8, head: 43, muac: 13, bp: "", hb: "" },
    "9m": { h: 67, w: 8, head: 44, muac: 13, bp: "", hb: "" },
    "10m": { h: 70, w: 9, head: 44, muac: 13, bp: "", hb: "" },
    "11m": { h: 71, w: 9, head: 45, muac: 13, bp: "", hb: "" },

    // Years 1 to 5 - RBSK / Indian Reference Standards (AWC - Whole Numbers)
    1: { h: 74, w: 9, head: 45, muac: 13, bp: "", hb: "" },
    2: { h: 86, w: 12, head: 47, muac: 13, bp: "", hb: "" },
    3: { h: 95, w: 14, head: 48, muac: 13, bp: "", hb: "" },
    4: { h: 103, w: 16, head: 49, muac: 13, bp: "", hb: "" },
    5: { h: 109, w: 18, head: 50, muac: 13, bp: "", hb: "" },

    // Years 6 to 20 - School Students (Realistic Moderate Heights & Weights, Normal BP <120/<80, Hb 12-14)
    6: { h: 109, w: 17, head: null, muac: null, bp: "95/65", hb: 12 },
    7: { h: 114, w: 19, head: null, muac: null, bp: "98/65", hb: 12 },
    8: { h: 119, w: 22, head: null, muac: null, bp: "100/68", hb: 12 },
    9: { h: 124, w: 25, head: null, muac: null, bp: "102/68", hb: 13 },
    10: { h: 129, w: 28, head: null, muac: null, bp: "105/70", hb: 13 },
    11: { h: 135, w: 32, head: null, muac: null, bp: "108/70", hb: 13 },
    12: { h: 141, w: 36, head: null, muac: null, bp: "110/72", hb: 13 },
    13: { h: 146, w: 40, head: null, muac: null, bp: "110/72", hb: 13 },
    14: { h: 150, w: 43, head: null, muac: null, bp: "112/74", hb: 13 },
    15: { h: 152, w: 46, head: null, muac: null, bp: "114/74", hb: 13 },
    16: { h: 153, w: 48, head: null, muac: null, bp: "115/75", hb: 13 },
    17: { h: 154, w: 49, head: null, muac: null, bp: "115/75", hb: 13 },
    18: { h: 155, w: 50, head: null, muac: null, bp: "116/76", hb: 13 },
    19: { h: 155, w: 50, head: null, muac: null, bp: "118/76", hb: 14 },
    20: { h: 156, w: 51, head: null, muac: null, bp: "118/76", hb: 14 }
};

/**
 * Standardize gender string
 */
function normalizeGenderValue(g) {
    const s = String(g || "").toLowerCase().trim();
    if (s.startsWith("f") || s.includes("girl") || s.includes("fem") || s.includes("stri") || s.includes("female") || s.includes("mahila")) {
        return "female";
    }
    if (s.startsWith("t") || s.includes("trans")) {
        return "transgender";
    }
    return "male";
}

/**
 * Parse date from various Excel and string formats
 */
function parseDateFlexible(val) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
        return val;
    }

    // Excel serial number (e.g. 42350)
    if (typeof val === "number" || (/^\d{4,6}$/.test(String(val).trim()) && !String(val).includes("-") && !String(val).includes("/"))) {
        const serial = parseInt(val, 10);
        if (serial > 1000 && serial < 100000) {
            // Excel leap year bug adjustment (days - 25569)
            return new Date((serial - 25569) * 86400 * 1000);
        }
    }

    const s = String(val).trim();

    // YYYY-MM-DD or YYYY/MM/DD
    if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(s)) {
        const p = s.split(/[-/.]/);
        return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    }

    // DD-MM-YYYY or DD/MM/YYYY
    if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}$/.test(s)) {
        const p = s.split(/[-/.]/);
        return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    }

    // DDMMYYYY
    if (/^\d{8}$/.test(s)) {
        const d = parseInt(s.slice(0, 2), 10);
        const m = parseInt(s.slice(2, 4), 10);
        const y = parseInt(s.slice(4), 10);
        return new Date(y, m - 1, d);
    }

    const parsed = Date.parse(s);
    if (!isNaN(parsed)) return new Date(parsed);

    return null;
}

/**
 * Calculate age in years & months from DOB, age text or class
 */
function computeStudentAge(dobVal, ageText, classText) {
    // 1. Try DOB first
    const dob = parseDateFlexible(dobVal);
    if (dob && !isNaN(dob.getTime())) {
        const now = new Date();
        let years = now.getFullYear() - dob.getFullYear();
        let months = now.getMonth() - dob.getMonth();
        if (now.getDate() < dob.getDate()) {
            months--;
        }
        if (months < 0) {
            years--;
            months += 12;
        }
        if (years >= 0 && years <= 25) {
            return { years, months, source: "dob" };
        }
    }

    // 2. Try Age text
    const str = String(ageText || "").toLowerCase().trim();
    if (str) {
        const yMatch = str.match(/(\d+)\s*(?:year|yr|y|वर्ष|साल)/);
        const mMatch = str.match(/(\d+)\s*(?:month|mo|m|माह|महीने)/);
        let years = yMatch ? parseInt(yMatch[1], 10) : 0;
        let months = mMatch ? parseInt(mMatch[1], 10) : 0;

        if (!yMatch && !mMatch && /^\d+(\.\d+)?$/.test(str)) {
            const num = parseFloat(str);
            years = Math.floor(num);
            months = Math.round((num - years) * 12);
        }

        if (years > 0 || months > 0) {
            return { years, months, source: "ageText" };
        }
    }

    // 3. Fallback: Estimate from Class name if School
    const cls = String(classText || "").toLowerCase().trim();
    if (cls) {
        const classMatch = cls.match(/\b(1[0-2]|[1-9])\b/);
        if (classMatch) {
            const cNum = parseInt(classMatch[1], 10);
            return { years: cNum + 5, months: 6, source: "class" };
        }
        if (cls.includes("kg") || cls.includes("nursery") || cls.includes("lkg") || cls.includes("ukg")) {
            return { years: 4, months: 6, source: "class" };
        }
    }

    // Default safe fallback (6 years old)
    return { years: 6, months: 0, source: "default" };
}

/**
 * Compute realistic Growth and Clinical Parameters
 */
function computeRbskParameters(student, options = {}) {
    const {
        allowDecimal = false,
        naturalVariation = true,
        fillHb = true,
        fillGrowth = true,
        fillBp = true,
        fillAwcExtra = true,
        hbRange = "normal", // 'normal' (11.5 - 13.8), 'mild' (10.0 - 11.4), 'custom'
        customHbMin = 11.5,
        customHbMax = 13.5,
        fillBirthWeight = true,
        birthWeightFormat = "grams", // 'grams' (2600 to 3100) or 'kg' (2.6 to 3.1)
        awcHbMode = "skip" // 'skip' or 'fill'
    } = options;

    const rawGender = normalizeGenderValue(student.gender);
    const tableGender = rawGender === "female" ? "female" : "male";
    const table = tableGender === "female" ? RBSK_GROWTH_GIRLS : RBSK_GROWTH_BOYS;
    const { years, months } = computeStudentAge(student.dob, student.age, student.class);

    let isUnderSix = false;
    let data = null;
    let ageKey = "";

    if (years === 0 || (years < 1 && months > 0)) {
        const m = Math.max(2, Math.min(11, months || 2));
        ageKey = `${m}m`;
        data = table[ageKey];
        isUnderSix = true;
    } else {
        const y = Math.max(1, Math.min(20, years || 1));
        ageKey = String(y);
        data = table[y];
        isUnderSix = y <= 5;
    }

    if (!data) {
        data = table[1] || RBSK_GROWTH_BOYS[1];
        isUnderSix = true;
    }

    // Variation helper: [-1, 0, 1]
    const getOffset = (range = 1) => {
        if (!naturalVariation) return 0;
        return (Math.floor(Math.random() * (range * 2 + 1)) - range);
    };

    const formatNumber = (val) => {
        if (allowDecimal) {
            const dec = Math.floor(Math.random() * 9) + 1; // .1 to .9
            return `${Math.round(val)}.${dec}`;
        }
        return String(Math.round(val));
    };

    // 1. Height, Weight & BMI
    let calculatedHeight = "";
    let calculatedWeight = "";
    let calculatedBmi = "";

    if (fillGrowth) {
        const baseH = Number(data.h) || 100;
        const baseW = Number(data.w) || 16;
        let finalH = Math.max(45, baseH + getOffset(1));
        let finalW = Math.max(4, baseW + getOffset(1));

        // Natural bounds check: avoid extreme overweight
        if (!isUnderSix && finalH > 0 && finalW > 0) {
            const hM = finalH / 100;
            const currentBmi = finalW / (hM * hM);
            if (currentBmi >= 24.5) {
                finalW = Math.min(finalW, Math.floor(23.5 * hM * hM));
            }
        }

        calculatedHeight = formatNumber(finalH);
        calculatedWeight = formatNumber(finalW);

        const hM = parseFloat(calculatedHeight) / 100;
        const wKg = parseFloat(calculatedWeight);
        if (hM > 0 && wKg > 0) {
            calculatedBmi = (wKg / (hM * hM)).toFixed(2);
        }
    }

    // 2. Blood Pressure (BP): Only for school students (6-20 years)
    let calculatedBp = "";
    if (fillBp && !isUnderSix && data.bp) {
        const bpStr = String(data.bp).trim();
        if (bpStr.includes("/")) {
            const parts = bpStr.split("/");
            let s = parseInt(parts[0], 10);
            let d = parseInt(parts[1], 10);
            if (isNaN(s) || isNaN(d)) {
                calculatedBp = "110/72";
            } else {
                s = Math.min(s, 118);
                d = Math.min(d, 76);
                calculatedBp = `${s}/${d}`;
            }
        } else {
            calculatedBp = "110/72";
        }
    }

    // 3. Hemoglobin (Hb)
    let calculatedHb = "";
    if (fillHb) {
        if (isUnderSix) {
            // AWC child Hb
            if (awcHbMode === "fill") {
                const awcVal = allowDecimal
                    ? (10.8 + Math.random() * 1.4).toFixed(1)
                    : String(Math.round(11 + getOffset(1)));
                calculatedHb = String(awcVal);
            } else {
                calculatedHb = ""; // protocol: skip for AWC
            }
        } else {
            // School Student (6-20 years)
            if (hbRange === "mild") {
                const val = (10.0 + Math.random() * 1.4);
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else if (hbRange === "custom") {
                const min = Math.min(customHbMin, customHbMax);
                const max = Math.max(customHbMin, customHbMax);
                const val = (min + Math.random() * (max - min));
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else {
                // Standard Normal (11.5 - 13.8)
                const baseHb = Number(data.hb) || (years >= 15 ? 13 : 12);
                if (allowDecimal) {
                    const decHb = (baseHb - 0.4 + Math.random() * 1.2).toFixed(1);
                    calculatedHb = String(decHb);
                } else {
                    const intHb = Math.max(11, Math.min(15, baseHb + getOffset(1)));
                    calculatedHb = String(intHb);
                }
            }
        }
    }

    // 4. AWC Specific (Head, MUAC, Birth Weight)
    let calculatedHead = "";
    let calculatedMuac = "";
    let calculatedBirthWeight = "";

    if (isUnderSix && fillAwcExtra) {
        if (data.head) {
            calculatedHead = String(Math.max(35, Math.round(Number(data.head) || 45) + getOffset(1)));
        }
        if (data.muac) {
            calculatedMuac = String(Math.round(Number(data.muac) || 13));
        }
        if (fillBirthWeight) {
            if (birthWeightFormat === "grams") {
                // Realistic birth weight in India: 2600 to 3100 grams
                const gramsList = ["2600", "2700", "2800", "2900", "3000", "3100"];
                calculatedBirthWeight = gramsList[Math.floor(Math.random() * gramsList.length)];
            } else {
                // In kg: 2.6 to 3.1 kg
                const bWeights = ["2.6", "2.7", "2.8", "2.9", "3.0", "3.1"];
                calculatedBirthWeight = bWeights[Math.floor(Math.random() * bWeights.length)];
            }
        }
    }

    return {
        ageYears: years,
        ageMonths: months,
        isUnderSix: isUnderSix,
        gender: rawGender,
        height: calculatedHeight,
        weight: calculatedWeight,
        bmi: calculatedBmi,
        bmiClass: "Normal",
        bmiResult: "Normal",
        bp: calculatedBp,
        bpClass: isUnderSix ? "" : "Normal",
        hb: calculatedHb,
        head: calculatedHead,
        muac: calculatedMuac,
        birthWeight: calculatedBirthWeight
    };
}

// Exports for Browser or Node
if (typeof window !== "undefined") {
    window.RBSK_GROWTH_BOYS = RBSK_GROWTH_BOYS;
    window.RBSK_GROWTH_GIRLS = RBSK_GROWTH_GIRLS;
    window.normalizeGenderValue = normalizeGenderValue;
    window.parseDateFlexible = parseDateFlexible;
    window.computeStudentAge = computeStudentAge;
    window.computeRbskParameters = computeRbskParameters;
}

if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        RBSK_GROWTH_BOYS,
        RBSK_GROWTH_GIRLS,
        normalizeGenderValue,
        parseDateFlexible,
        computeStudentAge,
        computeRbskParameters
    };
}
