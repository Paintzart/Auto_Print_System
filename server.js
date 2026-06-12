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
// Body: orderId, fileUrl, thickness, products (מהמסך), white_print_locations (מערך מ־getWhitePrintLocations).
// white_print_locations: רק מיקומים עם הדפס לבן, למשל [ { product: "1", location: "front" }, { product: "1", location: "left_sleeve" } ].
// thickness_mode: "uniform" (ברירת מחדל – אותו עובי לכולם) | "per_location" (עובי לפי מוצר/צד).
// logo_thicknesses: כש־per_location – מערך { product, location, thickness, variant_index? }.
// thickness נשאר ברירת מחדל לכל מיקום שלא מופיע ב-logo_thicknesses.
// אם ריק [] – לא כותבים קובץ ולא מעבירים ארגומנט; הפייתון יריץ זיהוי אוטומטי.
app.post('/prepare-print', async (req, res) => {
    let { orderId, fileUrl, thickness, white_print_locations, whitePrintLocations, front_print_2pocket, frontPrint2Pocket, item_quantities, itemQuantities, thickness_mode, thicknessMode, logo_thicknesses, logoThicknesses } = req.body;
    // תמיכה גם ב־camelCase מהקליינט
    const listFromBody = white_print_locations ?? whitePrintLocations;
    const front2Pocket = front_print_2pocket ?? frontPrint2Pocket ?? false;
    const thicknessModeVal = String(thickness_mode ?? thicknessMode ?? 'uniform').toLowerCase();
    const logoThicknessList = logo_thicknesses ?? logoThicknesses ?? [];
    const hasPerLocationThickness = thicknessModeVal === 'per_location';
    console.log(`\n🌸 בקשה להכנת דפוס: הזמנה ${orderId}${front2Pocket ? ' (הדפס קידמי 2Pocket – A4 לרוחב)' : ''}${hasPerLocationThickness ? ' (עובי לוגו לפי מוצר/צד)' : ''}`);
    // דיבוג: מה התקבל ב־white_print_locations
    const hasList = listFromBody != null && Array.isArray(listFromBody) && listFromBody.length > 0;
    console.log(`   [דיבוג] white_print_locations: ${listFromBody == null ? 'לא נשלח' : Array.isArray(listFromBody) ? `מערך באורך ${listFromBody.length}` : typeof listFromBody}`);
    if (hasList) console.log(`   [דיבוג] תוכן:`, JSON.stringify(listFromBody));
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
        // קובץ payload להעברת white_print_locations, front_print_2pocket, item_quantities לפייתון (ארגומנט 5)
        let orderPayloadPath = null;
        const itemQtyList = item_quantities ?? itemQuantities ?? [];
        const hasItemQty = Array.isArray(itemQtyList) && itemQtyList.length > 0;
        const payload = {
            white_print_locations: hasList ? listFromBody : [],
            front_print_2pocket: front2Pocket,
            item_quantities: hasItemQty ? itemQtyList : [],
            thickness_mode: thicknessModeVal,
            logo_thicknesses: hasPerLocationThickness && Array.isArray(logoThicknessList) ? logoThicknessList : [],
        };
        if (hasList || front2Pocket || hasItemQty || hasPerLocationThickness) {
            orderPayloadPath = path.join(TEMP_DIR, `order_${orderId}_${Date.now()}_payload.json`);
            fs.writeFileSync(orderPayloadPath, JSON.stringify(payload), 'utf8');
            if (hasList) console.log(`   > רשימת הדפס לבן: ${listFromBody.length} מיקומים, קובץ: ${orderPayloadPath}`);
            if (front2Pocket) console.log(`   > הדפס קידמי 2Pocket (A4 לרוחב) – כפתור סגול`);
            if (hasItemQty) console.log(`   > כמויות פריטים: ${itemQtyList.length} מיקומים`);
            if (hasPerLocationThickness) console.log(`   > עובי לוגו לפי מיקום: ${Array.isArray(logoThicknessList) ? logoThicknessList.length : 0} הגדרות, ברירת מחדל ${thickness || '2px'}`);
        } else {
            console.log(`   > אין רשימת הדפס לבן – יורץ זיהוי אוטומטי`);
        }
        const pythonArgs = [path.join(__dirname, 'prepare_print.py'), localFilePath, orderId, thickness];
        if (orderPayloadPath) pythonArgs.push(orderPayloadPath);
        console.log(`   [דיבוג] מריץ פייתון עם ${pythonArgs.length} ארגומנטים`);
        const pythonProcess = spawn(PYTHON_EXE, pythonArgs, { shell: true, cwd: __dirname });
        pythonProcess.stdout.on('data', (data) => console.log(`[Python]: ${data}`));
        pythonProcess.stderr.on('data', (data) => console.error(`[Error]: ${data}`));
        pythonProcess.on('close', (code) => {
            try { if (fs.existsSync(localFilePath)) fs.unlinkSync(localFilePath); } catch(e) {}
            try { if (orderPayloadPath && fs.existsSync(orderPayloadPath)) fs.unlinkSync(orderPayloadPath); } catch(e) {}
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
        if (type.includes('illustrator') || type.includes('postscript')) ext = '.ai';
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
function getReusePrintFrom(loc) {
    if (!loc) return null;
    return loc.reuse_print_from || loc.reusePrintFrom || null;
}

function normalizeReuseRef(reuse) {
    if (!reuse || typeof reuse !== 'object') return null;
    const product = reuse.product ?? reuse.prod;
    const location = reuse.location ?? reuse.side ?? reuse.loc;
    if (product == null || !location) return null;
    return { product, location };
}

function mapSimulationLocation(loc) {
    if (!loc) return { exists: false };
    const mapped = { ...loc };
    const reuse = normalizeReuseRef(getReusePrintFrom(loc));
    if (reuse) {
        mapped.reuse_print_from = reuse;
        mapped.exists = true;
        delete mapped.reusePrintFrom;
    } else if (mapped.exists === undefined && mapped.file_url) {
        mapped.exists = true;
    }
    return mapped;
}

function maybeSaveLocationFile(loc, prefix) {
    if (!loc?.file_url || getReusePrintFrom(loc)) return;
    loc.file_url = saveBase64Image(loc.file_url, prefix);
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
function isBlobUrl(value) {
    return typeof value === 'string' && value.startsWith('blob:');
}

function collectBlobUrls(value, path = '') {
    const found = [];
    if (value == null) return found;
    if (typeof value === 'string') {
        if (isBlobUrl(value)) found.push(path || '(root)');
        return found;
    }
    if (Array.isArray(value)) {
        value.forEach((item, i) => {
            found.push(...collectBlobUrls(item, `${path}[${i}]`));
        });
        return found;
    }
    if (typeof value === 'object') {
        for (const [key, nested] of Object.entries(value)) {
            const nextPath = path ? `${path}.${key}` : key;
            found.push(...collectBlobUrls(nested, nextPath));
        }
    }
    return found;
}

function getTemplateRaw(loc) {
    if (!loc) return null;
    let raw =
        loc.template_base64 || loc.templateBase64 ||
        loc.template_url || loc.templateUrl ||
        loc.template_file || loc.templateFile;
    if (!raw && Array.isArray(loc.variants) && loc.variants.length > 0) {
        const first = loc.variants[0] || {};
        raw =
            first.template_base64 || first.templateBase64 ||
            first.template_url || first.templateUrl;
    }
    if (!raw || typeof raw !== 'string') return null;
    // base64 גolmi בלי data: prefix
    if (
        !raw.startsWith('data:') &&
        !raw.startsWith('http') &&
        !raw.startsWith('blob:') &&
        !raw.includes('://') &&
        raw.length > 80
    ) {
        return `data:application/postscript;base64,${raw}`;
    }
    return raw;
}

async function resolveDownloadUrl(raw) {
    if (!raw) return null;
    if (isBlobUrl(raw)) return null;
    if (raw.startsWith('data:')) return raw;
    if (raw.includes('r2.cloudflarestorage.com')) {
        return await getR2SignedUrl(raw);
    }
    try {
        return decodeURIComponent(raw);
    } catch (e) {
        return raw;
    }
}

async function resolveStoredFileUrl(raw, savePrefix) {
    if (!raw) return null;
    if (isBlobUrl(raw)) return null;
    if (raw.startsWith('data:')) return saveBase64Image(raw, savePrefix);
    return await resolveDownloadUrl(raw);
}

async function mapVariableLocation(loc, prefix) {
    if (!loc) return { exists: false };
    const isVariable = loc.variable_print || loc.variablePrint || (Array.isArray(loc.variants) && loc.variants.length > 0);
    if (isVariable) {
        const variants = [];
        for (let i = 0; i < (loc.variants || []).length; i++) {
            const v = loc.variants[i];
            const idx = v.index != null ? v.index : i + 1;
            const out = {
                index: idx,
                label: v.label || null,
                req_color_hebrew: v.req_color_hebrew || v.reqColorHebrew || loc.req_color_hebrew || loc.reqColorHebrew,
                no_vectorization: v.no_vectorization ?? v.noVectorization ?? loc.no_vectorization ?? loc.noVectorization ?? false,
                text_overrides: v.text_overrides || v.textOverrides || {},
                image_overrides: {},
            };
            if (v.file_url || v.fileUrl) {
                const raw = v.file_url || v.fileUrl;
                out.file_url = await resolveStoredFileUrl(raw, `${prefix}_var_${idx}`);
            }
            const imgOverrides = v.image_overrides || v.imageOverrides || {};
            for (const [imgName, imgUrl] of Object.entries(imgOverrides)) {
                if (!imgUrl) continue;
                out.image_overrides[imgName] = await resolveStoredFileUrl(
                    imgUrl,
                    `${prefix}_var_${idx}_${imgName}`
                );
            }
            variants.push(out);
        }
        let templateRaw = getTemplateRaw(loc);
        const hasTemplate = !!templateRaw && !isBlobUrl(templateRaw);
        let templateUrl = null;
        if (hasTemplate) {
            templateUrl = await resolveStoredFileUrl(templateRaw, `${prefix}_template`);
        }
        if (templateRaw && isBlobUrl(templateRaw)) {
            console.warn(`   ⚠ template ${prefix}: blob URL is not supported by the server`);
        } else if (hasTemplate && templateUrl) {
            console.log(`   > template ${prefix}: ${String(templateUrl).slice(0, 120)}`);
        } else if (hasTemplate && !templateUrl) {
            console.warn(`   ⚠ template ${prefix}: could not store template file`);
        } else if (loc.template_mode || loc.templateMode) {
            console.warn(`   ⚠ template_mode=true on ${prefix} but no template_url – falling back to file mode`);
        }
        return {
            exists: true,
            variable_print: true,
            req_color_hebrew: loc.req_color_hebrew || loc.reqColorHebrew,
            category: loc.category || 'A4',
            no_vectorization: loc.no_vectorization ?? loc.noVectorization ?? false,
            template_url: templateUrl,
            template_mode: hasTemplate && !!templateUrl ? (loc.template_mode ?? loc.templateMode ?? true) : false,
            outline_text: loc.outline_text ?? loc.outlineText ?? true,
            variants,
        };
    }
    return mapSimulationLocation(loc);
}

function mapProductForSimulation(prod, index) {
    maybeSaveLocationFile(prod.locations?.front, `front_${index}`);
    maybeSaveLocationFile(prod.locations?.back, `back_${index}`);
    maybeSaveLocationFile(prod.locations?.right_sleeve, `right_${index}`);
    maybeSaveLocationFile(prod.locations?.left_sleeve, `left_${index}`);
    return {
        item_index: prod.item_index,
        product_type: prod.product_type,
        product_color_hebrew: prod.product_color_hebrew,
        extra_colors_hebrew: prod.extra_colors_hebrew || [],
        front: mapSimulationLocation(prod.locations?.front),
        back: mapSimulationLocation(prod.locations?.back),
        right_sleeve: mapSimulationLocation(prod.locations?.right_sleeve),
        left_sleeve: mapSimulationLocation(prod.locations?.left_sleeve),
    };
}

async function mapProductForVariable(prod) {
    const loc = prod.locations || {};
    const side = (name) => loc[name] || prod[name];
    return {
        item_index: prod.item_index || '1',
        product_type: prod.product_type,
        product_color_hebrew: prod.product_color_hebrew,
        extra_colors_hebrew: prod.extra_colors_hebrew || [],
        front: await mapVariableLocation(side('front'), 'front'),
        back: await mapVariableLocation(side('back'), 'back'),
        right_sleeve: await mapVariableLocation(side('right_sleeve'), 'right'),
        left_sleeve: await mapVariableLocation(side('left_sleeve'), 'left'),
    };
}

// === כפתור סגול: הדמיה ===
// === כפתור סגול: הדמיה (התיקון לאיחוד הקבצים) ===
app.post('/run-simulation', async (req, res) => {
    const { order_id, products, simulation_ad, is_wholesale } = req.body;
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
            processedProducts.push(mapProductForSimulation(prod, i));
        }
        // 2. הכנת המידע המלא לפייתון (כל המוצרים יחד)
        const fullBatchData = {
            order_id: order_id,
            products: processedProducts, // כאן נכנסים כל המוצרים
            simulation_ad: simulation_ad || { enabled: false, mode: 'random', selected_products: [] },
            is_wholesale: is_wholesale || false
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

// === כפתור כתום: הדמיה + הדפס משתנה (מוצר 1) ===
app.post('/run-variable-simulation', async (req, res) => {
    const { order_id, products, simulation_ad, is_wholesale } = req.body;
    console.log(`\n🟠 בקשה להדמיה משתנה: הזמנה ${order_id}`);
    if (!products || products.length !== 1) {
        return res.status(400).json({ success: false, message: "נדרש בדיוק מוצר 1" });
    }
    const blobPaths = collectBlobUrls(products[0]);
    if (blobPaths.length > 0) {
        console.warn(`   ⚠ blob URLs detected: ${blobPaths.join(', ')}`);
        return res.status(400).json({
            success: false,
            message: 'blob: URL לא נתמך בשרת. המר את קובץ התבנית ל-base64 לפני השליחה.',
            blob_fields: blobPaths,
            fix: {
                template_base64: 'data:application/postscript;base64,...',
                note: 'במקום template_url: blob:... שלח locations.front.template_base64',
            },
        });
    }
    try {
        const prod = await mapProductForVariable(products[0]);
        const hasVariable = ['front', 'back', 'right_sleeve', 'left_sleeve'].some(
            (s) => prod[s]?.variable_print
        );
        if (!hasVariable) {
            return res.status(400).json({ success: false, message: "לא נמצא variable_print באף צד" });
        }
        const fullBatchData = {
            mode: 'variable',
            order_id,
            products: [prod],
            simulation_ad: simulation_ad || { enabled: false, mode: 'random', selected_products: [] },
            is_wholesale: is_wholesale || false,
        };
        await runSingleSimulation(fullBatchData);
        console.log("\n✅ הדמיה משתנה הסתיימה בהצלחה!");
        res.json({ success: true, message: "ההדמיה וההדפסים המשתנים מוכנים" });
    } catch (error) {
        console.error("❌ תקלה (כתום):", error);
        res.status(500).json({ success: false, message: "שגיאה בעיבוד הדפס משתנה" });
    }
});

// === טעינת תיקיית ההדפסה מ-config (לשימוש ב-/download) ===
function getPrintFolderPath() {
    try {
        const configPath = path.join(__dirname, 'config.json');
        if (fs.existsSync(configPath)) {
            const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
            return config.print_folder_path || config.save_folder_path || __dirname;
        }
    } catch (e) {}
    return __dirname;
}

// === הורדת PDF – הממשק קורא ל־localhost:5001/download ===
// Body: orderId או order_id (למשל S026000939). מחזיר את קובץ ה־PDF המסכם (last4_print.pdf).
app.post('/download', (req, res) => {
    const orderId = req.body?.orderId ?? req.body?.order_id;
    if (!orderId) {
        return res.status(400).json({ error: 'חסר orderId או order_id' });
    }
    const last4 = String(orderId).replace(/\D/g, '').slice(-4);
    const printFolder = getPrintFolderPath();
    const pdfName = `${last4}_print.pdf`;
    const pdfPath = path.join(printFolder, pdfName);
    if (!fs.existsSync(pdfPath)) {
        return res.status(404).json({ error: 'קובץ PDF לא נמצא', path: pdfPath });
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${pdfName}"`);
    res.sendFile(pdfPath, (err) => {
        if (err) res.status(500).json({ error: 'שגיאה בשליחת הקובץ' });
    });
});

const PORT_5001 = 5001;
app.listen(PORT, () => {
    console.log(`\n✅ השרת רץ על פורט ${PORT}`);
});
app.listen(PORT_5001, () => {
    console.log(`✅ השרת מאזין גם על פורט ${PORT_5001} (לממשק / הורדת PDF)`);
});
