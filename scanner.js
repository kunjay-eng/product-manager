// =============================
// scanner.js 3.1
// =============================

const video = document.getElementById("video");
const canvas = document.getElementById("canvas");
const ctx = canvas.getContext("2d", { willReadFrequently: true });

const loading = document.getElementById("loading");
const message = document.getElementById("message");

const flashBtn = document.getElementById("flashBtn");
const switchBtn = document.getElementById("switchBtn");
const closeBtn = document.getElementById("closeBtn");

let stream = null;
let detector = null;
let useNative = false;
let facingMode = "environment";
let scanning = false;
let stopped = false;
let rafId = null;
let torchEnabled = false;

// ----------------------------
// BarcodeDetector
// ----------------------------
async function initDetector() {

    if (!("BarcodeDetector" in window)) {
        console.log("BarcodeDetector : Not Supported");
        return;
    }

    try {

        const formats =
            await BarcodeDetector.getSupportedFormats();

        if (formats.includes("qr_code")) {

            detector = new BarcodeDetector({
                formats: ["qr_code"]
            });

            useNative = true;

            console.log("BarcodeDetector Ready");

        }

    }
    catch (e) {

        console.log(e);

    }

}

// ----------------------------
// เปิดกล้อง
// ----------------------------
async function startCamera() {

    stopCamera();

    loading.style.display = "flex";

    message.innerHTML = "กำลังเปิดกล้อง...";

    try {

        stream =
            await navigator.mediaDevices.getUserMedia({

                video: {

                    facingMode: {
                        ideal: facingMode
                    }

                },

                audio: false

            });

    }
    catch (e) {

        loading.style.display = "none";

        alert(
            "เปิดกล้องไม่สำเร็จ\n\n" +
            e.name +
            "\n" +
            e.message
        );

        return;

    }

    video.srcObject = stream;

    await video.play();

    loading.style.display = "none";

    scanning = true;

    requestAnimationFrame(scanLoop);

}

// ----------------------------
// ปิดกล้อง
// ----------------------------
function stopCamera() {

    scanning = false;

    if (rafId)
        cancelAnimationFrame(rafId);

    if (stream) {

        stream
            .getTracks()
            .forEach(track => track.stop());

        stream = null;

    }

}

// ----------------------------
// Scan Loop
// ----------------------------
async function scanLoop() {

    if (!scanning || stopped)
        return;

    if (video.readyState === video.HAVE_ENOUGH_DATA) {

        try {

            // ---------- BarcodeDetector ----------
            if (useNative && detector) {

                const codes =
                    await detector.detect(video);

                if (codes.length > 0 &&
                    codes[0].rawValue) {

                    finish(codes[0].rawValue);
                    return;

                }

            }

            // ---------- jsQR Fallback ----------
            else {

                canvas.width = video.videoWidth;
                canvas.height = video.videoHeight;

                ctx.drawImage(
                    video,
                    0,
                    0,
                    canvas.width,
                    canvas.height
                );

                const img =
                    ctx.getImageData(
                        0,
                        0,
                        canvas.width,
                        canvas.height
                    );

                const qr =
                    jsQR(
                        img.data,
                        img.width,
                        img.height
                    );

                if (qr && qr.data) {

                    finish(qr.data);
                    return;

                }

            }

        }
        catch (err) {

            console.log(err);

        }

    }

    rafId =
        requestAnimationFrame(scanLoop);

}

// ----------------------------
// Scan Success
// ----------------------------
function finish(text) {

    if (stopped)
        return;

    stopped = true;
    scanning = false;

    console.log("QR =", text);

    // สั่น 80ms
    if (navigator.vibrate)
        navigator.vibrate(80);

    stopCamera();

    sendResult(text);

}

// ----------------------------
// ปิดหน้าจอ
// ----------------------------
closeBtn.onclick = function () {

    stopped = true;

    stopCamera();

    window.close();

};

// ----------------------------
// ส่งผลกลับ Apps Script
// ----------------------------
function sendResult(text) {

    debug("sendResult() ถูกเรียก");
    debug("TEXT = " + text);
    debug("SESSION = " + SESSION);
    debug("window.opener = " + !!window.opener);

    if (window.opener && !window.opener.closed) {

        debug("กำลัง postMessage");

        window.opener.postMessage({
            type: "QR_RESULT",
            session: SESSION,
            text: text
        }, "*");

        debug("postMessage ส่งแล้ว");

    } else {

        debug("window.opener ไม่มี");

    }

    // ยังไม่ต้องปิดหน้าต่าง
}



// ----------------------------
// สลับกล้อง
// ----------------------------
switchBtn.onclick = async function () {

    facingMode =
        facingMode === "environment"
            ? "user"
            : "environment";

    stopped = false;

    await startCamera();

};

// ----------------------------
// Torch
// ----------------------------
flashBtn.onclick = async function () {

    if (!stream)
        return;

    const track = stream.getVideoTracks()[0];

    if (!track)
        return;

    const cap = track.getCapabilities();

    if (!cap.torch) {

        alert("เครื่องนี้ไม่รองรับไฟฉาย");

        return;

    }

    torchEnabled = !torchEnabled;

    try {

        await track.applyConstraints({

            advanced: [
                {
                    torch: torchEnabled
                }
            ]

        });

        flashBtn.style.opacity =
            torchEnabled ? "1" : ".5";

    }
    catch (e) {

        console.log(e);

    }

};

// ----------------------------
// โหลดเสร็จ
// ----------------------------
window.onload = async function () {

    await initDetector();

    await startCamera();

};

// ----------------------------
// ออกจากหน้า
// ----------------------------
window.onbeforeunload = function () {

    stopCamera();

};

// ----------------------------
// Visibility
// ----------------------------
document.addEventListener(
    "visibilitychange",
    function () {

        if (document.hidden) {

            stopCamera();

        }

    }


    function debug(msg) {
    const el = document.getElementById("debug");
    if (el) {
        el.innerHTML += "<br>" + msg;
    }
}

);
