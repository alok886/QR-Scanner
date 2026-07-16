// Phoenix Business Advisory QR Scanner V2

let registrations = [];
let attendance = [];
let scannedBookings = new Set();

let verifiedCount = 0;
let rejectedCount = 0;
let duplicateCount = 0;

let html5QrCode = null;
let excelLoaded = false;
let scanLocked = false;
let lastScannedCode = "";
let lastScanTime = 0;

const excelInput = document.getElementById("excelFile");
const loadBtn = document.getElementById("loadExcel");
const scannerBtn = document.getElementById("startScanner");
const statusBox = document.getElementById("statusBox");
const resultCard = document.getElementById("resultCard");

const verifiedCounter = document.getElementById("registeredCount");
const rejectedCounter = document.getElementById("failedCount");
const duplicateCounter = document.getElementById("duplicateCount");
const attendanceBtn = document.getElementById("downloadAttendance");

loadBtn.addEventListener("click", loadExcelFile);
scannerBtn.addEventListener("click", startScanner);
attendanceBtn.addEventListener("click", downloadAttendance);

function loadExcelFile() {
    if (!excelInput.files.length) {
        alert("Please select the Zoho Excel file.");
        return;
    }

    const file = excelInput.files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

            registrations = XLSX.utils.sheet_to_json(firstSheet, {
                defval: ""
            });

            excelLoaded = true;

            showStatus(
                "success",
                `
                    <div class="scan-badge">FILE LOADED</div>
                    <h2>REGISTRATIONS READY</h2>
                    <p><strong>${registrations.length}</strong> registrations loaded successfully.</p>
                    <p>Start the QR scanner to begin check-in.</p>
                `,
                false
            );
        } catch (error) {
            console.error(error);
            alert("Unable to read the Excel file. Please confirm that it is a valid Excel or CSV file.");
        }
    };

    reader.readAsArrayBuffer(file);
}

function startScanner() {
    if (!excelLoaded) {
        alert("Please load the Excel file first.");
        return;
    }

    scannerBtn.disabled = true;
    scannerBtn.textContent = "Starting Camera...";

    html5QrCode = new Html5Qrcode("reader");

    const config = {
        fps: 10,
        qrbox: {
            width: 250,
            height: 250
        },
        aspectRatio: 1
    };

    Html5Qrcode.getCameras()
        .then(cameras => {
            if (!cameras || cameras.length === 0) {
                throw new Error("No camera found.");
            }

            let selectedCamera = cameras.find(camera => {
                const label = String(camera.label || "").toLowerCase();

                return (
                    label.includes("back") ||
                    label.includes("rear") ||
                    label.includes("environment")
                );
            });

            if (!selectedCamera) {
                selectedCamera = cameras[cameras.length - 1];
            }

            return html5QrCode.start(
                selectedCamera.id,
                config,
                onScanSuccess,
                onScanFailure
            );
        })
        .then(() => {
            scannerBtn.textContent = "Scanner Running";
        })
        .catch(error => {
            console.error(error);
            alert("Unable to start or access the camera.");
            scannerBtn.disabled = false;
            scannerBtn.textContent = "Start QR Scanner";
        });
}

function onScanSuccess(decodedText) {
    const bookingID = normalize(decodedText);
    const now = Date.now();

    if (!bookingID || scanLocked) {
        return;
    }

    if (bookingID === lastScannedCode && now - lastScanTime < 4000) {
        return;
    }

    lastScannedCode = bookingID;
    lastScanTime = now;
    scanLocked = true;

    verifyBooking(bookingID);

    setTimeout(() => {
        scanLocked = false;
    }, 1800);
}

function onScanFailure() {
    // Normal camera frame errors are ignored.
}

function verifyBooking(bookingID) {
    const person = registrations.find(row =>
        normalize(row["Booking ID"]) === bookingID
    );

    if (!person) {
        rejectedCount++;
        rejectedCounter.textContent = rejectedCount;

        showStatus(
            "failed",
            `
                <div class="scan-badge">SCAN REJECTED</div>
                <h2>NOT REGISTERED</h2>
                <p>Booking ID</p>
                <p class="person-name">${escapeHtml(bookingID)}</p>
                <p>Record not found in the uploaded Excel file.</p>
            `
        );
        return;
    }

    if (scannedBookings.has(bookingID)) {
        duplicateCount++;
        duplicateCounter.textContent = duplicateCount;

        showStatus(
            "duplicate",
            `
                <div class="scan-badge">ALREADY SCANNED</div>
                <h2>ALREADY CHECKED IN</h2>
                <p class="person-name">${escapeHtml(person["Customer"])}</p>
                <p>Booking ID: <strong>${escapeHtml(bookingID)}</strong></p>
            `
        );
        return;
    }

    const paymentStatus = normalize(person["Payment Status"]).toLowerCase();

    if (paymentStatus !== "paid") {
        rejectedCount++;
        rejectedCounter.textContent = rejectedCount;

        showStatus(
            "failed",
            `
                <div class="scan-badge">PAYMENT NOT VERIFIED</div>
                <h2>NOT CHECKED IN</h2>
                <p class="person-name">${escapeHtml(person["Customer"])}</p>
                <p>Booking ID: <strong>${escapeHtml(bookingID)}</strong></p>
                <p>Payment Status: <strong>${escapeHtml(person["Payment Status"] || "Not Available")}</strong></p>
            `
        );
        return;
    }

    scannedBookings.add(bookingID);

    attendance.push({
        bookingID,
        customer: normalize(person["Customer"]),
        email: normalize(person["Customer Email"]),
        phone: normalize(person["Customer Contact Number"]),
        payment: normalize(person["Payment Status"]),
        time: new Date().toLocaleString()
    });

    verifiedCount++;
    verifiedCounter.textContent = verifiedCount;

    showStatus(
        "success",
        `
            <div class="scan-badge">SCANNED SUCCESSFULLY</div>
            <h2>CHECKED IN</h2>
            <p class="person-name">${escapeHtml(person["Customer"])}</p>

            <div class="result-details">
                <p><span>Booking ID</span><strong>${escapeHtml(bookingID)}</strong></p>
                <p><span>Email</span><strong>${escapeHtml(person["Customer Email"])}</strong></p>
                <p><span>Phone</span><strong>${escapeHtml(person["Customer Contact Number"])}</strong></p>
                <p><span>Payment</span><strong>${escapeHtml(person["Payment Status"])}</strong></p>
            </div>
        `
    );
}

function showStatus(className, html, scrollToResult = true) {
    statusBox.className = className;
    statusBox.innerHTML = html;

    if (scrollToResult) {
        setTimeout(() => {
            resultCard.scrollIntoView({
                behavior: "smooth",
                block: "center"
            });

            statusBox.classList.remove("result-flash");
            void statusBox.offsetWidth;
            statusBox.classList.add("result-flash");
        }, 100);
    }
}

function normalize(value) {
    return String(value ?? "").trim();
}

function escapeHtml(value) {
    return normalize(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function csvCell(value) {
    return `"${normalize(value).replace(/"/g, '""')}"`;
}

function downloadAttendance() {
    if (!attendance.length) {
        alert("No checked-in attendance is available.");
        return;
    }

    const rows = [
        [
            "Booking ID",
            "Customer",
            "Email",
            "Phone",
            "Payment Status",
            "Scan Time"
        ]
    ];

    attendance.forEach(person => {
        rows.push([
            person.bookingID,
            person.customer,
            person.email,
            person.phone,
            person.payment,
            person.time
        ]);
    });

    const csv = rows
        .map(row => row.map(csvCell).join(","))
        .join("\n");

    const blob = new Blob(["\uFEFF" + csv], {
        type: "text/csv;charset=utf-8;"
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "Phoenix_Kadi_Seminar_Attendance.csv";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
}
