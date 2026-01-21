const express = require('express');

const cors = require('cors');

const { spawn } = require('child_process');

const bodyParser = require('body-parser');

const path = require('path');

const fs = require('fs');

const axios = require('axios');

const { S3Client, GetObjectCommand } = require('@aws-sdk/client-s3');

const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');



// ==========================================================

//              פרטי R2

// ==========================================================

const R2_ACCOUNT_ID = "944539d199bcd56d08fd20e2920753c9";

const R2_ACCESS_KEY_ID = "869cd104efd961706ce96b5d051388b3";

const R2_SECRET_ACCESS_KEY = "5ff7e1df459b90aba30e39fd91e04a01b0573014dd224e79036f197fbdf21fcd";



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

app.use(bodyParser.json({ limit: '200mb' }));

app.use(bodyParser.urlencoded({ limit: '200mb', extended: true }));



const TEMP_DIR = path.join(__dirname, 'temp_downloads');

if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR);



// === הגדרת הפייתון (כדי למנוע שגיאות של ספריות חסרות) ===

const venvPythonPath = path.join(__dirname, 'venv', 'Scripts', 'python.exe');

const PYTHON_EXE = fs.existsSync(venvPythonPath) ? venvPythonPath : 'python';



// --- פונקציות עזר (R2) ---

async function getR2SignedUrl(originalUrl) {

    try {

        const urlObj = new URL(originalUrl);

        const pathParts = urlObj.pathname.split('/');

        const bucketName = pathParts[1]; 

        const fileKey = decodeURIComponent(pathParts.slice(2).join('/'));

        const command = new GetObjectCommand({ Bucket: bucketName, Key: fileKey });

        return await getSignedUrl(s3Client, command, { expiresIn: 3600 });

    } catch (e) {

        return originalUrl;

    }

}



// === כפתור ורוד ===

app.post('/prepare-print', async (req, res) => {

    let { orderId, fileUrl, thickness } = req.body;

    console.log(`\n🌸 בקשה להכנת דפוס: הזמנה ${orderId}`);

    try {

        if (fileUrl.includes('r2.cloudflarestorage.com')) {

            fileUrl = await getR2SignedUrl(fileUrl);

        } else {

            fileUrl = decodeURIComponent(fileUrl);

        }

        const fileName = `temp_${orderId}_${Date.now()}.pdf`;

        const localFilePath = path.join(TEMP_DIR, fileName);

        

        const response = await axios({

            method: 'GET', url: fileUrl, responseType: 'stream', decompress: false

        });

        const writer = fs.createWriteStream(localFilePath);

        response.data.pipe(writer);

        await new Promise((resolve, reject) => {

            writer.on('finish', resolve);

            writer.on('error', reject);

        });



        const pythonScriptPath = path.join(__dirname, 'prepare_print.py');

        const pythonProcess = spawn(PYTHON_EXE, [pythonScriptPath, localFilePath, orderId, thickness], { shell: true });

        

        pythonProcess.stdout.on('data', (data) => console.log(`[Python]: ${data}`));

        pythonProcess.stderr.on('data', (data) => console.error(`[Error]: ${data}`));

        pythonProcess.on('close', (code) => {

            try { if (fs.existsSync(localFilePath)) fs.unlinkSync(localFilePath); } catch(e) {}

            if (code === 0) res.json({ success: true, message: "הקבצים מוכנים!" });

            else res.status(500).json({ success: false, message: "עיבוד הפייתון נכשל" });

        });

    } catch (error) {

        console.error("❌ שגיאה בשרת:", error.message);

        res.status(500).json({ success: false, message: "תקלה" });

    }

});



function saveBase64Image(base64Str, prefix) {

    if (!base64Str || !base64Str.startsWith('data:')) return base64Str;

    try {

        const matches = base64Str.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);

        if (!matches || matches.length !== 3) return base64Str;

        const type = matches[1];

        const data = matches[2];

        const buffer = Buffer.from(data, 'base64');

        let ext = '.png';

        if (type.includes('jpeg') || type.includes('jpg')) ext = '.jpg';

        if (type.includes('pdf')) ext = '.pdf';

        if (type.includes('svg')) ext = '.svg';

        const fileName = `${prefix}_${Date.now()}${ext}`;

        const filePath = path.join(TEMP_DIR, fileName);

        fs.writeFileSync(filePath, buffer);

        console.log(`   > שמרתי קובץ זמני: ${fileName}`);

        return filePath; 

    } catch (e) {

        console.error("Error saving base64:", e);

        return base64Str;

    }

}



// === פונקציית ההרצה (הישנה והטובה) ===

function runSingleSimulation(payloadForPython) {

    return new Promise((resolve, reject) => {

        const pythonScriptPath = path.join(__dirname, 'main.py');

        console.log(`   >> מריץ פייתון...`);

        

        // שימוש בפייתון הנכון (VENV)

        const pythonProcess = spawn(PYTHON_EXE, [pythonScriptPath, JSON.stringify(payloadForPython)]);



        pythonProcess.stdout.on('data', (data) => console.log(`[Sim Python]: ${data}`));

        pythonProcess.stderr.on('data', (data) => console.error(`[Sim Error]: ${data}`));



        pythonProcess.on('close', (code) => {

            if (code === 0) {

                console.log("   V הסתיים בהצלחה");

                resolve();

            } else {

                console.log("   X נכשל");

                reject(new Error(`Python process exited with code ${code}`));

            }

        });

    });

}



// === כפתור סגול: הדמיה ===

// === כפתור סגול: הדמיה (התיקון לאיחוד הקבצים) ===

app.post('/run-simulation', async (req, res) => {

    const { order_id, products, simulation_ad } = req.body;

    console.log(`\n🟣 בקשה להדמיה: הזמנה ${order_id} (${products ? products.length : 0} מוצרים)`);



    if (!products || products.length === 0) {

        return res.status(400).json({ success: false, message: "אין מוצרים" });

    }



    try {

        // 1. יצירת רשימה לאיסוף כל המוצרים המעובדים

        const processedProducts = [];



        for (let i = 0; i < products.length; i++) {

            const prod = products[i];

            console.log(`\n--- מעבד מוצר ${i + 1} ---`);



            // שמירת תמונות כבדות (ללא שינוי)

            if (prod.locations.front?.file_url) prod.locations.front.file_url = saveBase64Image(prod.locations.front.file_url, `front_${i}`);

            if (prod.locations.back?.file_url) prod.locations.back.file_url = saveBase64Image(prod.locations.back.file_url, `back_${i}`);

            if (prod.locations.right_sleeve?.file_url) prod.locations.right_sleeve.file_url = saveBase64Image(prod.locations.right_sleeve.file_url, `right_${i}`);

            if (prod.locations.left_sleeve?.file_url) prod.locations.left_sleeve.file_url = saveBase64Image(prod.locations.left_sleeve.file_url, `left_${i}`);



            // הוספת המוצר לרשימה (במקום לשלוח לפייתון מיד!)

            processedProducts.push({

                item_index: prod.item_index, 

                product_type: prod.product_type,

                product_color_hebrew: prod.product_color_hebrew,

                extra_colors_hebrew: prod.extra_colors_hebrew || [],

                front: prod.locations.front || { exists: false },

                back: prod.locations.back || { exists: false },

                right_sleeve: prod.locations.right_sleeve || { exists: false },

                left_sleeve: prod.locations.left_sleeve || { exists: false }

            });

        }



        // 2. הכנת המידע המלא לפייתון (כל המוצרים יחד)

        const fullBatchData = {

            order_id: order_id,

            products: processedProducts, // כאן נכנסים כל המוצרים

            simulation_ad: simulation_ad || { enabled: false, mode: 'random', selected_products: [] }

        };



        // 3. שליחה לפייתון פעם אחת בלבד!

        await runSingleSimulation(fullBatchData);



        console.log("\n✅ הכל הסתיים בהצלחה!");

        res.json({ success: true, message: "ההדמיות הסתיימו ואוחדו בהצלחה" });



    } catch (error) {

        console.error("❌ תקלה:", error);

        res.status(500).json({ success: false, message: "שגיאה בעיבוד" });

    }

});



app.listen(PORT, () => {

    console.log(`\n✅ השרת רץ על פורט ${PORT}`);

});