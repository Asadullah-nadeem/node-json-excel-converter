const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

app.use((req, res, next) => {
    const logMessage = `${new Date().toISOString()} - ${req.method} ${req.url}`;
    console.log(logMessage);
    fs.appendFile('logs/server.log', logMessage + '\n', (err) => {
        if (err) console.error('Error writing to log file:', err);
    });
    next();
});

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/convert', upload.single('jsonFile'), async (req, res) => {
    const filePath = req.file.path;
    const jsonType = req.body.jsonType;

    let jsonData;
    try {
        jsonData = JSON.parse(fs.readFileSync(filePath));
    } catch (error) {
        console.error('Error reading JSON file:', error);
        return res.status(400).send("Error parsing JSON file.");
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('JSON Data');

    switch (jsonType) {
        case 'basicArray':
            worksheet.addRow(jsonData);
            break;

        case 'eachObject':
            worksheet.columns = Object.keys(jsonData[0]).map(key => ({ header: key, key }));
            jsonData.forEach((item) => worksheet.addRow(item));
            break;

        case 'describeObject':
            worksheet.columns = Object.keys(jsonData).map(key => ({ header: key, key }));
            worksheet.addRow(jsonData);
            break;

        case 'nested':
            const flattenObject = (obj, prefix = '') =>
                Object.keys(obj).reduce((acc, k) => {
                    const pre = prefix.length ? prefix + '.' : '';
                    if (typeof obj[k] === 'object' && obj[k] !== null) {
                        Object.assign(acc, flattenObject(obj[k], pre + k));
                    } else {
                        acc[pre + k] = obj[k];
                    }
                    return acc;
                }, {});

            const flatData = jsonData.map(flattenObject);
            worksheet.columns = Object.keys(flatData[0]).map(key => ({ header: key, key }));
            flatData.forEach((item) => worksheet.addRow(item));
            break;

        default:
            return res.status(400).send("Invalid JSON type selected.");
    }

    const excelFilePath = path.join('uploads', 'output.xlsx');
    try {
        await workbook.xlsx.writeFile(excelFilePath);
        fs.unlinkSync(filePath); 
    } catch (error) {
        console.error('Error writing Excel file:', error);
        return res.status(500).send("Error converting JSON to Excel.");
    }

    res.download(excelFilePath, 'converted_data.xlsx', (err) => {
        if (err) {
            console.error('Error downloading the file:', err);
        }
        fs.unlinkSync(excelFilePath); 
    });
});

app.listen(3000, () => console.log('Server started on http://localhost:3000'));
