// ===========================================
// PHOENIX BUSINESS ADVISORY
// Webinar QR Verification
// Part 1
// ===========================================

// Store all registrations
let registrations = [];

// Store verified attendees
let attendance = [];

// Prevent duplicate scan
let scannedBookings = new Set();

// Counters
let verifiedCount = 0;
let rejectedCount = 0;
let duplicateCount = 0;

// Scanner Object
let html5QrCode;

// Excel Loaded?
let excelLoaded = false;

// HTML Elements
const excelInput = document.getElementById("excelFile");
const loadBtn = document.getElementById("loadExcel");
const scannerBtn = document.getElementById("startScanner");

const statusBox = document.getElementById("statusBox");

const verifiedCounter =
document.getElementById("registeredCount");

const rejectedCounter =
document.getElementById("failedCount");

const duplicateCounter =
document.getElementById("duplicateCount");

const attendanceBtn =
document.getElementById("downloadAttendance");


// ===========================================
// LOAD EXCEL
// ===========================================

loadBtn.addEventListener("click", () => {

    if(excelInput.files.length===0){

        alert("Please Select Zoho Excel File");

        return;

    }

    const file = excelInput.files[0];

    const reader = new FileReader();

    reader.onload = function(e){

        const data = new Uint8Array(e.target.result);

        const workbook =
        XLSX.read(data,{type:'array'});

        const sheet =
        workbook.Sheets[workbook.SheetNames[0]];

        registrations =
        XLSX.utils.sheet_to_json(sheet);

        excelLoaded=true;

        statusBox.className="success";

        statusBox.innerHTML=

        `
        ✅ ${registrations.length}
        registrations loaded successfully.
        <br><br>
        Click Start QR Scanner.
        `;

    }

    reader.readAsArrayBuffer(file);

});


// ===========================================
// PART 2
// QR Scanner
// ===========================================

// Start Scanner

scannerBtn.addEventListener("click", startScanner);

function startScanner(){

    if(!excelLoaded){

        alert("Please load the Excel file first.");

        return;

    }

    scannerBtn.disabled = true;

    scannerBtn.innerHTML = "Starting Camera...";

    html5QrCode = new Html5Qrcode("reader");

    const config = {

        fps:10,

        qrbox:{
            width:250,
            height:250
        },

        aspectRatio:1.0

    };

    Html5Qrcode.getCameras()

    .then(cameras=>{

        if(cameras && cameras.length){

            html5QrCode.start(

                cameras[0].id,

                config,

                onScanSuccess,

                onScanFailure

            );

            scannerBtn.innerHTML="Scanner Running";

        }

        else{

            alert("No Camera Found");

        }

    })

    .catch(err=>{

        console.log(err);

        alert("Unable to Start Camera");

    });

}


// ===========================================
// Called Every Time QR is Read
// ===========================================

function onScanSuccess(decodedText){

    verifyBooking(decodedText.trim());

}


// Ignore Camera Errors

function onScanFailure(error){

    // intentionally empty

}



// ===========================================
// Search Booking ID
// ===========================================

function verifyBooking(bookingID){

    let person = registrations.find(row=>{

        return String(row["Booking ID"]).trim() === bookingID;

    });


    if(!person){

        rejectedCount++;

        rejectedCounter.innerHTML=rejectedCount;

        statusBox.className="failed";

        statusBox.innerHTML=`

            <h2>❌ NOT REGISTERED</h2>

            <br>

            Booking ID

            <br><br>

            <b>${bookingID}</b>

            <br><br>

            Record Not Found.

        `;

        return;

    }


    // Duplicate Check

    if(scannedBookings.has(bookingID)){

        duplicateCount++;

        duplicateCounter.innerHTML=duplicateCount;

        statusBox.className="duplicate";

        statusBox.innerHTML=`

        <h2>⚠ Already Checked In</h2>

        <br>

        ${person["Customer"]}

        <br><br>

        Booking ID

        <br>

        ${bookingID}

        `;

        return;

    }


    // ===========================================
    // PAYMENT VERIFICATION
    // ===========================================

    let paymentStatus = "";

    if(person["Payment Status"] !== undefined &&
       person["Payment Status"] !== null){

        paymentStatus =
        String(person["Payment Status"]).trim().toLowerCase();

    }

    // Payment NOT Paid

    if(paymentStatus !== "paid"){

        rejectedCount++;

        rejectedCounter.innerHTML = rejectedCount;

        statusBox.className = "failed";

        statusBox.innerHTML = `

        <h2>🔴 NOT REGISTERED</h2>

        <br>

        <b>${person["Customer"]}</b>

        <br><br>

        Booking ID :
        ${bookingID}

        <br><br>

        Payment Status :
        <b>${person["Payment Status"]}</b>

        `;

        return;

    }


    // ===========================================
    // REGISTERED
    // ===========================================

    scannedBookings.add(bookingID);

    attendance.push({

        bookingID: bookingID,

        customer: person["Customer"],

        email: person["Customer Email"],

        phone: person["Customer Contact Number"],

        payment: person["Payment Status"],

        time: new Date().toLocaleString()

    });

    verifiedCount++;

    verifiedCounter.innerHTML = verifiedCount;

    statusBox.className = "success";

    statusBox.innerHTML = `

        <h2>✅ REGISTERED</h2>

        <br>

        <b>${person["Customer"]}</b>

        <br><br>

        Booking ID :
        ${bookingID}

        <br>

        Email :
        ${person["Customer Email"]}

        <br>

        Phone :
        ${person["Customer Contact Number"]}

        <br>

        Payment :
        <b>${person["Payment Status"]}</b>

    `;




// ===========================================
// DOWNLOAD ATTENDANCE
// ===========================================

attendanceBtn.addEventListener("click", downloadAttendance);

function downloadAttendance(){

    if(attendance.length===0){

        alert("No Attendance Available");

        return;

    }

    let csv =

"Booking ID,Customer,Email,Phone,Payment Status,Scan Time\n";

    attendance.forEach(person=>{

        csv +=

`${person.bookingID},
${person.customer},
${person.email},
${person.phone},
${person.payment},
${person.time}\n`;

    });

    const blob = new Blob(

        [csv],

        {

            type:"text/csv"

        }

    );

    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");

    a.href = url;

    a.download = "Attendance.csv";

    document.body.appendChild(a);

    a.click();

    document.body.removeChild(a);

    URL.revokeObjectURL(url);

}
}