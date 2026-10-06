const http = require("http");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

const VIDEO_PATH = "C:\\Users\\ravi\\Downloads\\Untitled Project.mp4";
const OUT_DIR = "C:\\Users\\ravi\\Downloads\\Shree_RBSK_Promotion_Kit\\screenshots";
const PORT = 8989;

if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

const TARGET_FRAMES = [
    { time: 2, name: "01_webstore_install.png", label: "Chrome Web Store Installation" },
    { time: 15, name: "02_google_login_activation.png", label: "Google Login & QR Activation" },
    { time: 20, name: "03_main_dashboard_overview.png", label: "Central Analytics Dashboard" },
    { time: 27, name: "04_screening_speed_settings.png", label: "Automation Speed & Auto Growth Settings" },
    { time: 46, name: "05_child_screening_table_controls.png", label: "Child Screening Table with Start/Stop Controls" },
    { time: 58, name: "06_auto_growth_form_filling.png", label: "Automatic Growth, BMI, BP & Hb Calculation" },
    { time: 61, name: "07_screening_preview_submit.png", label: "Automated Screening Preview & 1-Click Submit" },
    { time: 96, name: "08_student_master_table_controls.png", label: "Student Master Control Toolbar & Filters" },
    { time: 112, name: "09_student_master_excel_template.png", label: "Exported Excel for Bulk Registration & ABHA" },
    { time: 167, name: "10_referral_management_portal.png", label: "Referral Management Automation Table" },
    { time: 203, name: "11_smart_screening_excel_dropdowns.png", label: "Smart Excel Template with Native Defect Dropdown" },
    { time: 218, name: "12_realtime_analytics_summary.png", label: "Live Impact Analytics & Modular Switches" }
];

let savedCount = 0;

const server = http.createServer((req, res) => {
    if (req.url === "/video.mp4") {
        const stat = fs.statSync(VIDEO_PATH);
        const range = req.headers.range;
        if (range) {
            const parts = range.replace(/bytes=/, "").split("-");
            const start = parseInt(parts[0], 10);
            const end = parts[1] ? parseInt(parts[1], 10) : stat.size - 1;
            const chunksize = (end - start) + 1;
            const file = fs.createReadStream(VIDEO_PATH, { start, end });
            res.writeHead(206, {
                "Content-Range": `bytes ${start}-${end}/${stat.size}`,
                "Accept-Ranges": "bytes",
                "Content-Length": chunksize,
                "Content-Type": "video/mp4"
            });
            file.pipe(res);
        } else {
            res.writeHead(200, {
                "Content-Length": stat.size,
                "Content-Type": "video/mp4"
            });
            fs.createReadStream(VIDEO_PATH).pipe(res);
        }
        return;
    }

    if (req.url === "/save-frame" && req.method === "POST") {
        let body = "";
        req.on("data", chunk => { body += chunk; });
        req.on("end", () => {
            try {
                const { name, data } = JSON.parse(body);
                const base64Data = data.replace(/^data:image\/png;base64,/, "");
                const filePath = path.join(OUT_DIR, name);
                fs.writeFileSync(filePath, base64Data, "base64");
                savedCount++;
                console.log(`[${savedCount}/${TARGET_FRAMES.length}] Saved ${name}`);
                res.writeHead(200, { "Content-Type": "application/json" });
                res.end(JSON.stringify({ ok: true, saved: savedCount }));

                if (savedCount >= TARGET_FRAMES.length) {
                    console.log("All target frames saved successfully! Closing server...");
                    setTimeout(() => {
                        server.close();
                        process.exit(0);
                    }, 1000);
                }
            } catch (err) {
                console.error("Save error:", err);
                res.writeHead(500);
                res.end(JSON.stringify({ error: err.message }));
            }
        });
        return;
    }

    // Serve extractor HTML
    res.writeHead(200, { "Content-Type": "text/html" });
    res.end(`<!DOCTYPE html>
<html>
<head><title>Frame Extractor</title></head>
<body>
<h2>Extracting Frames...</h2>
<video id="v" src="/video.mp4" crossorigin="anonymous" muted playsinline style="display:none;"></video>
<canvas id="c" style="display:none;"></canvas>
<pre id="log"></pre>
<script>
const frames = ${JSON.stringify(TARGET_FRAMES)};
const v = document.getElementById("v");
const c = document.getElementById("c");
const ctx = c.getContext("2d");
const logEl = document.getElementById("log");

function log(msg) {
    logEl.textContent += msg + "\\n";
    console.log(msg);
}

v.onloadedmetadata = async () => {
    log("Video metadata loaded: " + v.videoWidth + "x" + v.videoHeight + ", duration: " + v.duration + "s");
    c.width = v.videoWidth || 1920;
    c.height = v.videoHeight || 1080;

    for (let i = 0; i < frames.length; i++) {
        const f = frames[i];
        log("Seeking to " + f.time + "s for " + f.name + "...");
        await new Promise(resolve => {
            const onSeeked = () => {
                v.removeEventListener("seeked", onSeeked);
                setTimeout(resolve, 200);
            };
            v.addEventListener("seeked", onSeeked);
            v.currentTime = f.time;
        });

        ctx.drawImage(v, 0, 0, c.width, c.height);
        const dataUrl = c.toDataURL("image/png");

        await fetch("/save-frame", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ name: f.name, data: dataUrl })
        });
        log("Uploaded " + f.name);
    }
    log("Done all frames!");
};
v.load();
</script>
</body>
</html>`);
});

server.listen(PORT, () => {
    console.log(`Frame extraction server running on http://localhost:${PORT}`);
    const chromePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
    const cmd = `"${chromePath}" --headless --disable-gpu --no-sandbox "http://localhost:${PORT}"`;
    console.log("Launching Chrome headless to capture frames...");
    exec(cmd, (err) => {
        if (err && !server.listening) return; // server already closed
        if (err) console.error("Chrome exec error:", err);
    });
});
