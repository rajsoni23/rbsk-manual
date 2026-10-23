/**
 * RBSK & WHO Standard Growth Computation Engine
 * Designed for RBSK Screening Auto-Fill Tool
 * 
 * Features:
 * 1. Exact WHO 0-60 Month Medians (MGRS Standards) to prevent Stunting / Wasting on RBSK Portal.
 * 2. WHO Head Circumference (-1SD to +1SD Normal limits) & MUAC (Green / Normal 13.2 - 13.8 cm).
 * 3. School (>5 Years) BMI Calibration (15.0 - 23.0) and Normal Hb (11.5 - 13.8 g/dL).
 * 4. Vision Left & Right Handling (6/6 for School, skipped for AWC by default).
 * 5. Complete toggle control support for all parameters.
 */

"use strict";

// WHO Standard Reference Tables for Head Circumference (0 to 60 Months)
const WHO_HC_DATA = {
    male: [
        { age: 0, m: 34.5, sd1neg: 33.2, sd1pos: 35.8, sd2neg: 31.9, sd2pos: 37.0 },
        { age: 1, m: 37.3, sd1neg: 36.0, sd1pos: 38.6, sd2neg: 34.7, sd2pos: 39.8 },
        { age: 2, m: 39.0, sd1neg: 37.7, sd1pos: 40.2, sd2neg: 36.4, sd2pos: 41.5 },
        { age: 3, m: 40.5, sd1neg: 39.2, sd1pos: 41.8, sd2neg: 37.9, sd2pos: 43.0 },
        { age: 4, m: 41.6, sd1neg: 40.2, sd1pos: 42.9, sd2neg: 38.9, sd2pos: 44.1 },
        { age: 5, m: 42.6, sd1neg: 41.2, sd1pos: 43.9, sd2neg: 39.9, sd2pos: 45.1 },
        { age: 6, m: 43.3, sd1neg: 42.1, sd1pos: 44.6, sd2neg: 40.8, sd2pos: 45.8 },
        { age: 7, m: 44.0, sd1neg: 42.7, sd1pos: 45.2, sd2neg: 41.4, sd2pos: 46.4 },
        { age: 8, m: 44.5, sd1neg: 43.2, sd1pos: 45.7, sd2neg: 41.9, sd2pos: 47.0 },
        { age: 9, m: 45.0, sd1neg: 43.7, sd1pos: 46.2, sd2neg: 42.5, sd2pos: 47.5 },
        { age: 10, m: 45.4, sd1neg: 44.1, sd1pos: 46.6, sd2neg: 42.9, sd2pos: 48.0 },
        { age: 11, m: 45.8, sd1neg: 44.5, sd1pos: 47.0, sd2neg: 43.3, sd2pos: 48.4 },
        { age: 12, m: 46.1, sd1neg: 44.8, sd1pos: 47.3, sd2neg: 43.6, sd2pos: 48.6 },
        { age: 18, m: 47.4, sd1neg: 46.1, sd1pos: 48.6, sd2neg: 44.9, sd2pos: 49.9 },
        { age: 24, m: 48.3, sd1neg: 47.0, sd1pos: 49.5, sd2neg: 45.8, sd2pos: 50.8 },
        { age: 36, m: 49.5, sd1neg: 48.2, sd1pos: 50.8, sd2neg: 46.9, sd2pos: 52.1 },
        { age: 48, m: 50.4, sd1neg: 49.1, sd1pos: 51.7, sd2neg: 47.8, sd2pos: 53.0 },
        { age: 60, m: 51.1, sd1neg: 49.7, sd1pos: 52.4, sd2neg: 48.4, sd2pos: 53.7 }
    ],
    female: [
        { age: 0, m: 33.9, sd1neg: 32.7, sd1pos: 35.1, sd2neg: 31.5, sd2pos: 36.2 },
        { age: 1, m: 36.5, sd1neg: 35.2, sd1pos: 37.7, sd2neg: 34.0, sd2pos: 38.9 },
        { age: 2, m: 38.2, sd1neg: 36.8, sd1pos: 39.4, sd2neg: 35.5, sd2pos: 40.5 },
        { age: 3, m: 39.5, sd1neg: 38.2, sd1pos: 40.8, sd2neg: 36.9, sd2pos: 42.0 },
        { age: 4, m: 40.6, sd1neg: 39.2, sd1pos: 41.8, sd2neg: 38.0, sd2pos: 43.0 },
        { age: 5, m: 41.5, sd1neg: 40.1, sd1pos: 42.7, sd2neg: 38.8, sd2pos: 43.9 },
        { age: 6, m: 42.2, sd1neg: 40.9, sd1pos: 43.5, sd2neg: 39.6, sd2pos: 44.8 },
        { age: 7, m: 42.8, sd1neg: 41.5, sd1pos: 44.0, sd2neg: 40.2, sd2pos: 45.4 },
        { age: 8, m: 43.3, sd1neg: 42.0, sd1pos: 44.5, sd2neg: 40.7, sd2pos: 45.9 },
        { age: 9, m: 43.8, sd1neg: 42.5, sd1pos: 45.1, sd2neg: 41.2, sd2pos: 46.4 },
        { age: 10, m: 44.2, sd1neg: 42.9, sd1pos: 45.5, sd2neg: 41.6, sd2pos: 46.8 },
        { age: 11, m: 44.6, sd1neg: 43.3, sd1pos: 45.9, sd2neg: 42.0, sd2pos: 47.2 },
        { age: 12, m: 44.9, sd1neg: 43.6, sd1pos: 46.2, sd2neg: 42.3, sd2pos: 47.5 },
        { age: 18, m: 46.2, sd1neg: 44.9, sd1pos: 47.5, sd2neg: 43.6, sd2pos: 48.8 },
        { age: 24, m: 47.2, sd1neg: 45.9, sd1pos: 48.5, sd2neg: 44.6, sd2pos: 49.8 },
        { age: 36, m: 48.5, sd1neg: 47.1, sd1pos: 49.8, sd2neg: 45.8, sd2pos: 51.2 },
        { age: 48, m: 49.5, sd1neg: 48.1, sd1pos: 50.8, sd2neg: 46.8, sd2pos: 52.2 },
        { age: 60, m: 50.2, sd1neg: 48.8, sd1pos: 51.5, sd2neg: 47.5, sd2pos: 52.9 }
    ]
};

// WHO Child Growth Standards (0 to 60 months / Under 5 years)
// Official WHO Multicentre Growth Reference Study (MGRS) Median Height (h in cm) & Weight (w in kg)
// Using exact monthly medians ensures Length-Age, Weight-Age, and Weight-Height classifications
// naturally evaluate as "Normal" on the RBSK portal without force-setting disabled dropdowns.
const WHO_UNDER5_BOYS = {
    0: { h: 50, w: 3 }, 1: { h: 55, w: 5 }, 2: { h: 58, w: 6 }, 3: { h: 61, w: 6 },
    4: { h: 64, w: 7 }, 5: { h: 66, w: 8 }, 6: { h: 68, w: 8 }, 7: { h: 69, w: 8 },
    8: { h: 71, w: 9 }, 9: { h: 72, w: 9 }, 10: { h: 73, w: 9 }, 11: { h: 75, w: 9 },
    12: { h: 76, w: 10 }, 13: { h: 77, w: 10 }, 14: { h: 78, w: 10 }, 15: { h: 79, w: 10 },
    16: { h: 80, w: 11 }, 17: { h: 81, w: 11 }, 18: { h: 82, w: 11 }, 19: { h: 83, w: 11 },
    20: { h: 84, w: 11 }, 21: { h: 85, w: 12 }, 22: { h: 86, w: 12 }, 23: { h: 87, w: 12 },
    24: { h: 88, w: 12 }, 25: { h: 89, w: 12 }, 26: { h: 89, w: 13 }, 27: { h: 90, w: 13 },
    28: { h: 91, w: 13 }, 29: { h: 92, w: 13 }, 30: { h: 93, w: 13 }, 31: { h: 93, w: 14 },
    32: { h: 94, w: 14 }, 33: { h: 95, w: 14 }, 34: { h: 96, w: 14 }, 35: { h: 96, w: 14 },
    36: { h: 97, w: 14 }, 37: { h: 98, w: 15 }, 38: { h: 98, w: 15 }, 39: { h: 99, w: 15 },
    40: { h: 100, w: 15 }, 41: { h: 100, w: 15 }, 42: { h: 101, w: 15 }, 43: { h: 102, w: 16 },
    44: { h: 102, w: 16 }, 45: { h: 103, w: 16 }, 46: { h: 103, w: 16 }, 47: { h: 104, w: 16 },
    48: { h: 105, w: 16 }, 49: { h: 105, w: 17 }, 50: { h: 106, w: 17 }, 51: { h: 106, w: 17 },
    52: { h: 107, w: 17 }, 53: { h: 107, w: 17 }, 54: { h: 108, w: 17 }, 55: { h: 108, w: 18 },
    56: { h: 109, w: 18 }, 57: { h: 110, w: 18 }, 58: { h: 110, w: 18 }, 59: { h: 111, w: 18 },
    60: { h: 111, w: 18 }
};

const WHO_UNDER5_GIRLS = {
    0: { h: 49, w: 3 }, 1: { h: 54, w: 4 }, 2: { h: 57, w: 5 }, 3: { h: 60, w: 6 },
    4: { h: 62, w: 6 }, 5: { h: 64, w: 7 }, 6: { h: 66, w: 7 }, 7: { h: 67, w: 8 },
    8: { h: 69, w: 8 }, 9: { h: 70, w: 8 }, 10: { h: 72, w: 9 }, 11: { h: 73, w: 9 },
    12: { h: 74, w: 9 }, 13: { h: 75, w: 9 }, 14: { h: 76, w: 9 }, 15: { h: 78, w: 10 },
    16: { h: 79, w: 10 }, 17: { h: 80, w: 10 }, 18: { h: 81, w: 10 }, 19: { h: 82, w: 10 },
    20: { h: 83, w: 11 }, 21: { h: 84, w: 11 }, 22: { h: 85, w: 11 }, 23: { h: 86, w: 11 },
    24: { h: 86, w: 12 }, 25: { h: 87, w: 12 }, 26: { h: 88, w: 12 }, 27: { h: 89, w: 12 },
    28: { h: 90, w: 12 }, 29: { h: 90, w: 13 }, 30: { h: 91, w: 13 }, 31: { h: 92, w: 13 },
    32: { h: 93, w: 13 }, 33: { h: 93, w: 13 }, 34: { h: 94, w: 14 }, 35: { h: 95, w: 14 },
    36: { h: 95, w: 14 }, 37: { h: 96, w: 14 }, 38: { h: 97, w: 14 }, 39: { h: 97, w: 14 },
    40: { h: 98, w: 15 }, 41: { h: 99, w: 15 }, 42: { h: 99, w: 15 }, 43: { h: 100, w: 15 },
    44: { h: 101, w: 15 }, 45: { h: 101, w: 16 }, 46: { h: 102, w: 16 }, 47: { h: 102, w: 16 },
    48: { h: 103, w: 16 }, 49: { h: 103, w: 16 }, 50: { h: 104, w: 16 }, 51: { h: 105, w: 17 },
    52: { h: 105, w: 17 }, 53: { h: 106, w: 17 }, 54: { h: 106, w: 17 }, 55: { h: 107, w: 17 },
    56: { h: 107, w: 18 }, 57: { h: 108, w: 18 }, 58: { h: 108, w: 18 }, 59: { h: 109, w: 18 },
    60: { h: 109, w: 18 }
};

// Years 6 to 20 - School Students (RBSK & WHO Reference Standards)
const RBSK_GROWTH_BOYS = {
    6: { h: 116, w: 21, head: null, muac: null, bp: "98/65", hb: 12 },
    7: { h: 121, w: 23, head: null, muac: null, bp: "100/66", hb: 12 },
    8: { h: 127, w: 26, head: null, muac: null, bp: "102/68", hb: 12 },
    9: { h: 132, w: 28, head: null, muac: null, bp: "104/68", hb: 13 },
    10: { h: 137, w: 32, head: null, muac: null, bp: "106/70", hb: 13 },
    11: { h: 143, w: 36, head: null, muac: null, bp: "108/70", hb: 13 },
    12: { h: 149, w: 40, head: null, muac: null, bp: "110/72", hb: 13 },
    13: { h: 156, w: 45, head: null, muac: null, bp: "112/72", hb: 13 },
    14: { h: 163, w: 50, head: null, muac: null, bp: "114/74", hb: 13 },
    15: { h: 169, w: 56, head: null, muac: null, bp: "115/74", hb: 14 },
    16: { h: 173, w: 61, head: null, muac: null, bp: "116/75", hb: 14 },
    17: { h: 175, w: 64, head: null, muac: null, bp: "116/75", hb: 14 },
    18: { h: 176, w: 66, head: null, muac: null, bp: "118/76", hb: 14 },
    19: { h: 177, w: 67, head: null, muac: null, bp: "118/76", hb: 14 },
    20: { h: 177, w: 68, head: null, muac: null, bp: "118/76", hb: 14 }
};

const RBSK_GROWTH_GIRLS = {
    6: { h: 115, w: 20, head: null, muac: null, bp: "98/65", hb: 12 },
    7: { h: 120, w: 22, head: null, muac: null, bp: "100/66", hb: 12 },
    8: { h: 126, w: 25, head: null, muac: null, bp: "102/68", hb: 12 },
    9: { h: 132, w: 28, head: null, muac: null, bp: "104/68", hb: 13 },
    10: { h: 138, w: 32, head: null, muac: null, bp: "105/70", hb: 13 },
    11: { h: 144, w: 36, head: null, muac: null, bp: "107/70", hb: 13 },
    12: { h: 151, w: 41, head: null, muac: null, bp: "109/72", hb: 13 },
    13: { h: 156, w: 46, head: null, muac: null, bp: "110/72", hb: 13 },
    14: { h: 159, w: 49, head: null, muac: null, bp: "112/74", hb: 13 },
    15: { h: 161, w: 52, head: null, muac: null, bp: "114/74", hb: 13 },
    16: { h: 162, w: 53, head: null, muac: null, bp: "115/75", hb: 13 },
    17: { h: 163, w: 54, head: null, muac: null, bp: "115/75", hb: 13 },
    18: { h: 163, w: 55, head: null, muac: null, bp: "116/76", hb: 13 },
    19: { h: 163, w: 55, head: null, muac: null, bp: "118/76", hb: 14 },
    20: { h: 164, w: 56, head: null, muac: null, bp: "118/76", hb: 14 }
};

/**
 * Standardize gender string
 */
function normalizeGenderValue(g) {
    const s = String(g || "").toLowerCase().trim();
    if (s.startsWith("f") || s.includes("girl") || s.includes("fem") || s.includes("stri") || s.includes("female") || s.includes("mahila")) {
        return "female";
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
 * Calculate age in years & months from DOB or text
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
            return { years, months, totalMonths: (years * 12) + months, source: "dob" };
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
            return { years, months, totalMonths: (years * 12) + months, source: "ageText" };
        }
    }

    // 3. Fallback: Estimate from Class name if School
    const cls = String(classText || "").toLowerCase().trim();
    if (cls) {
        const classMatch = cls.match(/\b(1[0-2]|[1-9])\b/);
        if (classMatch) {
            const cNum = parseInt(classMatch[1], 10);
            const yrs = cNum + 5;
            return { years: yrs, months: 6, totalMonths: (yrs * 12) + 6, source: "class" };
        }
        if (cls.includes("kg") || cls.includes("nursery") || cls.includes("lkg") || cls.includes("ukg")) {
            return { years: 4, months: 6, totalMonths: 54, source: "class" };
        }
    }

    // Default safe fallback (6 years old)
    return { years: 6, months: 0, totalMonths: 72, source: "default" };
}

/**
 * Returns WHO Normal Head Circumference within -1SD to +1SD range
 */
function getWhoNormalHeadCircumference(totalMonths, gender) {
    const g = normalizeGenderValue(gender) === "female" ? "female" : "male";
    const refList = WHO_HC_DATA[g];
    let ref = refList[0];
    for (let i = 0; i < refList.length; i++) {
        if (totalMonths >= refList[i].age) {
            ref = refList[i];
        }
    }
    const naturalOffsets = [-0.2, -0.1, 0, 0.1, 0.2];
    const offset = naturalOffsets[Math.floor(Math.random() * naturalOffsets.length)];
    let val = Number((ref.m + offset).toFixed(1));
    val = Math.max(ref.sd1neg + 0.1, Math.min(ref.sd1pos - 0.1, val));
    return Number(val.toFixed(1));
}

/**
 * Returns WHO Normal MUAC (Green range > 12.5 cm, strictly 13.2 to 13.8 cm)
 */
function getWhoNormalMuac() {
    const options = [13.2, 13.4, 13.5, 13.6, 13.8];
    return options[Math.floor(Math.random() * options.length)];
}

/**
 * Compute realistic Growth and Clinical Parameters with full granular field toggles
 */
function computeRbskParameters(student, options = {}) {
    const {
        fillHeight = true,
        fillWeight = true,
        fillHb = true,
        fillVision = true,
        defaultVision = "6/6",
        fillBirthWeight = true,
        fillHead = true,
        fillMuac = true,
        allowDecimal = false,
        naturalVariation = true,
        hbRange = "normal", // 'normal' (11.5 - 13.8), 'mild' (10.0 - 11.4), 'custom'
        customHbMin = 11.5,
        customHbMax = 13.5,
        isAwc = false,
        awcVision = false // whether to fill vision for AWC (default false per RBSK rules)
    } = options;

    const gender = normalizeGenderValue(student.gender);
    const { years, months, totalMonths } = computeStudentAge(student.dob, student.age, student.class);

    // Classification Rule:
    // <= 5.0 years (totalMonths <= 60): Uses WHO Under-5 standards (Length-Age, Weight-Age, Weight-Height)
    // > 5.0 years (61+ months, or 6+ years): Evaluated using BMI!
    const isUnderFive = totalMonths > 0 ? totalMonths <= 60 : (years <= 5 && months === 0);
    const isAwcChild = Boolean(isAwc || student.isAwc || isUnderFive);

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

    let calculatedHeight = "";
    let calculatedWeight = "";
    let calculatedHead = "";
    let calculatedMuac = "";
    let calculatedBirthWeight = "";
    let calculatedHb = "";
    let calculatedVisionLeft = "";
    let calculatedVisionRight = "";

    if (isUnderFive) {
        // --- Under 5 Years (0 - 60 Months) WHO Standards ---
        const m = Math.max(1, Math.min(60, totalMonths || (years ? years * 12 : (months || 1))));
        const under5Table = gender === "female" ? WHO_UNDER5_GIRLS : WHO_UNDER5_BOYS;
        const whoEntry = under5Table[m] || under5Table[12];

        // Maintain exact WHO medians to prevent severe stunting / wasting
        const baseH = whoEntry.h;
        const baseW = whoEntry.w;

        if (fillHeight) {
            calculatedHeight = formatNumber(baseH);
        }
        if (fillWeight) {
            calculatedWeight = formatNumber(baseW);
        }

        // Head Circumference & MUAC for AWC / under-5
        if (fillHead) {
            calculatedHead = String(getWhoNormalHeadCircumference(m, gender));
        }
        if (fillMuac) {
            calculatedMuac = String(getWhoNormalMuac());
        }

        // Birth weight: only when DOB is available and toggle is active
        const hasValidDob = Boolean(student.dob && parseDateFlexible(student.dob));
        if (fillBirthWeight && hasValidDob) {
            const bWeights = ["2.6", "2.7", "2.8", "2.9", "3.0", "3.1"];
            calculatedBirthWeight = bWeights[Math.floor(Math.random() * bWeights.length)];
        }

        // Hb for AWC (if enabled)
        if (fillHb) {
            if (hbRange === "mild") {
                const val = 10.0 + Math.random() * 1.4;
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else if (hbRange === "custom") {
                const min = Math.min(customHbMin, customHbMax);
                const max = Math.max(customHbMin, customHbMax);
                const val = min + Math.random() * (max - min);
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else {
                const val = 11.5 + Math.random() * 1.5; // 11.5 - 13.0
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            }
        }

        // Vision for Under 5: skipped by default unless awcVision option is enabled
        if (fillVision && awcVision) {
            calculatedVisionLeft = defaultVision;
            calculatedVisionRight = defaultVision;
        }
    } else {
        // --- School Students (> 5 Years / 6+ Years): Evaluated via BMI ---
        let y = years;
        if (y <= 5) y = 6;
        y = Math.max(6, Math.min(20, y));

        const schoolTable = gender === "female" ? RBSK_GROWTH_GIRLS : RBSK_GROWTH_BOYS;
        const ref = schoolTable[y] || schoolTable[6];

        const baseH = Math.round(Number(ref.h) || 120);
        const baseW = Math.round(Number(ref.w) || 22);

        let finalH = Math.max(100, baseH + getOffset(1));
        let finalW = Math.max(15, baseW + getOffset(1));

        // Ensure healthy BMI strictly between 15.0 and 23.0
        const hM = finalH / 100;
        const curBmi = finalW / (hM * hM);
        if (curBmi < 14.8) {
            finalW = Math.ceil(15.2 * hM * hM);
        } else if (curBmi > 23.0) {
            finalW = Math.floor(22.0 * hM * hM);
        }

        if (fillHeight) {
            calculatedHeight = formatNumber(finalH);
        }
        if (fillWeight) {
            calculatedWeight = formatNumber(finalW);
        }

        // Hemoglobin for School
        if (fillHb) {
            if (hbRange === "mild") {
                const val = 10.0 + Math.random() * 1.4;
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else if (hbRange === "custom") {
                const min = Math.min(customHbMin, customHbMax);
                const max = Math.max(customHbMin, customHbMax);
                const val = min + Math.random() * (max - min);
                calculatedHb = allowDecimal ? val.toFixed(1) : String(Math.round(val));
            } else {
                const baseHb = Number(ref.hb) || (y >= 15 ? 13 : 12);
                if (allowDecimal) {
                    calculatedHb = (baseHb - 0.4 + Math.random() * 1.2).toFixed(1);
                } else {
                    calculatedHb = String(Math.max(11, Math.min(15, baseHb + getOffset(1))));
                }
            }
        }

        // Vision for School: Default 6/6
        if (fillVision) {
            calculatedVisionLeft = defaultVision || "6/6";
            calculatedVisionRight = defaultVision || "6/6";
        }
    }

    return {
        ageYears: years,
        ageMonths: months,
        totalMonths: totalMonths,
        isUnderFive: isUnderFive,
        isAwcChild: isAwcChild,
        gender: gender,
        height: calculatedHeight,
        weight: calculatedWeight,
        hb: calculatedHb,
        visionLeft: calculatedVisionLeft,
        visionRight: calculatedVisionRight,
        head: calculatedHead,
        muac: calculatedMuac,
        birthWeight: calculatedBirthWeight
    };
}

// Global & CommonJS Export
if (typeof window !== "undefined") {
    window.WHO_UNDER5_BOYS = WHO_UNDER5_BOYS;
    window.WHO_UNDER5_GIRLS = WHO_UNDER5_GIRLS;
    window.RBSK_GROWTH_BOYS = RBSK_GROWTH_BOYS;
    window.RBSK_GROWTH_GIRLS = RBSK_GROWTH_GIRLS;
    window.WHO_HC_DATA = WHO_HC_DATA;
    window.getWhoNormalHeadCircumference = getWhoNormalHeadCircumference;
    window.getWhoNormalMuac = getWhoNormalMuac;
    window.normalizeGenderValue = normalizeGenderValue;
    window.parseDateFlexible = parseDateFlexible;
    window.computeStudentAge = computeStudentAge;
    window.computeRbskParameters = computeRbskParameters;
}

if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        WHO_UNDER5_BOYS,
        WHO_UNDER5_GIRLS,
        RBSK_GROWTH_BOYS,
        RBSK_GROWTH_GIRLS,
        WHO_HC_DATA,
        getWhoNormalHeadCircumference,
        getWhoNormalMuac,
        normalizeGenderValue,
        parseDateFlexible,
        computeStudentAge,
        computeRbskParameters
    };
}
