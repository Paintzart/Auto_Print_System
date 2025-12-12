const express = require('express');
const cors = require('cors');
const { spawn } = require('child_process');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
const PORT = 3000; // השרת ירוץ על פורט 3000

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

// בדיקה שהשרת חי
app.get('/', (req, res) => {
    res.send('המערכת האוטומטית מחוברת ומחכה להוראות! 🚀');
});

// קבלת בקשה מהאתר
app.post('/run-simulation', (req, res) => {
    const orderData = req.body;
    console.log(`\n📦 התקבלה בקשה להדמיה: ${orderData.order_id}`);

    try {
        const jsonData = JSON.stringify(orderData);
        // מניח ש-main.py נמצא באותה תיקייה
        const pythonScriptPath = path.join(__dirname, 'main.py');

        console.log("מפעיל את האוטומציה...");
        
        // הרצת הפייתון
        const pythonProcess = spawn('python', [pythonScriptPath, jsonData]);

        // הצגת לוגים מהפייתון בחלון השחור
        pythonProcess.stdout.on('data', (data) => {
            console.log(`[Python]: ${data.toString('utf8')}`);
        });

        pythonProcess.stderr.on('data', (data) => {
            console.error(`[Error]: ${data}`);
        });

        pythonProcess.on('close', (code) => {
            console.log(`סיום תהליך (קוד ${code})`);
            if (code === 0) {
                res.json({ success: true, message: "בוצע בהצלחה" });
            } else {
                res.status(500).json({ success: false, message: "שגיאה בעיבוד" });
            }
        });

    } catch (error) {
        console.error("Server Error:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// הפעלת ההאזנה
app.listen(PORT, () => {
    console.log(`\n✅ השרת מוכן לעבודה!`);
    console.log(`כתובת: http://localhost:${PORT}`);
    console.log(`לא לסגור את החלון הזה.`);
});