const express = require('express');
const cors = require('cors');
const { spawn } = require('child_process');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const axios = require('axios');

// ספריות לטיפול באבטחה של Cloudflare
const { S3Client, GetObjectCommand } = require('@aws-sdk/client-s3');
const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');

// ==========================================================
//              איזור העריכה - הזן את הפרטים כאן
// ==========================================================
const R2_ACCOUNT_ID = "944539d199bcd56d08fd20e2920753c9";
const R2_ACCESS_KEY_ID = "869cd104efd961706ce96b5d051388b3";
const R2_SECRET_ACCESS_KEY = "5ff7e1df459b90aba30e39fd91e04a01b0573014dd224e79036f197fbdf21fcd";
// ==========================================================

// הגדרת החיבור לכספת (R2)
const s3Client = new S3Client({
    region: "auto",
    endpoint: `https://${R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
    credentials: {
        accessKeyId: R2_ACCESS_KEY_ID,
        secretAccessKey: R2_SECRET_ACCESS_KEY,
    },
});

const app = express();
const PORT = 3000;

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

const TEMP_DIR = path.join(__dirname, 'temp_downloads');
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR);

// פונקציה שפותחת את הקישור הנעול
async function getR2SignedUrl(originalUrl) {
    try {
        const urlObj = new URL(originalUrl);
        const pathParts = urlObj.pathname.split('/');
        const bucketName = pathParts[1]; 
        const fileKey = decodeURIComponent(pathParts.slice(2).join('/'));

        const command = new GetObjectCommand({ Bucket: bucketName, Key: fileKey });
        // יצירת מפתח זמני לשעה
        return await getSignedUrl(s3Client, command, { expiresIn: 3600 });
    } catch (e) {
        console.log("⚠️ לא הצלחתי לחתום על הקישור, מנסה רגיל...");
        return originalUrl;
    }
}

app.get('/', (req, res) => res.send('המערכת מחוברת 🚀'));

// === כפתור ורוד: הכנת קבצי הדפסה ===
app.post('/prepare-print', async (req, res) => {
    let { orderId, fileUrl, thickness } = req.body;
    console.log(`\n🌸 בקשה להכנת דפוס: הזמנה ${orderId}`);

    try {
        // אם זה קישור של R2 Cloudflare - נפתח אותו עם המפתח
        if (fileUrl.includes('r2.cloudflarestorage.com')) {
            fileUrl = await getR2SignedUrl(fileUrl);
        } else {
             fileUrl = decodeURIComponent(fileUrl);
        }

        const fileName = `temp_${orderId}_${Date.now()}.pdf`;
        const localFilePath = path.join(TEMP_DIR, fileName);

        console.log(`מוריד קובץ...`);

        const response = await axios({
            method: 'GET',
            url: fileUrl,
            responseType: 'stream',
            decompress: false
        });

        const writer = fs.createWriteStream(localFilePath);
        response.data.pipe(writer);

        await new Promise((resolve, reject) => {
            writer.on('finish', resolve);
            writer.on('error', reject);
        });

        console.log(`הקובץ ירד בהצלחה. מפעיל עיבוד...`);
        
        const pythonScriptPath = path.join(__dirname, 'prepare_print.py');
        const pythonProcess = spawn('python', [pythonScriptPath, localFilePath, orderId, thickness]);

        pythonProcess.stdout.on('data', (data) => console.log(`[Python]: ${data}`));
        pythonProcess.stderr.on('data', (data) => console.error(`[Error]: ${data}`));

        pythonProcess.on('close', (code) => {
            try { fs.unlinkSync(localFilePath); } catch(e) {}
            if (code === 0) res.json({ success: true, message: "הקבצים מוכנים!" });
            else res.status(500).json({ success: false, message: "הסקריפט נכשל" });
        });

    } catch (error) {
        console.error("❌ שגיאה:", error.message);
        if (error.response && error.response.data) {
             // הצגת פרטים אם השרת ענה בשגיאה
             const data = error.response.data;
             console.error("פרטי שגיאה:", Buffer.isBuffer(data) ? data.toString() : data);
        }
        res.status(500).json({ success: false, message: "תקלה בהורדת הקובץ" });
    }
});

// === כפתור סגול: הדמיה ===
app.post('/run-simulation', (req, res) => {
    const orderData = req.body;
    console.log(`\n🟣 בקשה להדמיה: ${orderData.order_id}`);
    const pythonScriptPath = path.join(__dirname, 'main.py');
    const pythonProcess = spawn('python', [pythonScriptPath, JSON.stringify(orderData)]);
    
    pythonProcess.stdout.on('data', (data) => console.log(`[Sim]: ${data}`));
    pythonProcess.on('close', (code) => {
        if (code === 0) res.json({ success: true });
        else res.status(500).json({ success: false });
    });
});

app.listen(PORT, () => {
    console.log(`\n✅ השרת רץ על פורט ${PORT}`);
});