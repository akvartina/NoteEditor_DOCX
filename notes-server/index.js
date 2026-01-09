const express = require('express');
const cors = require('cors');
const multer = require('multer');
const htmlToDocx = require('html-to-docx');
const fs = require('fs');
const { execFile } = require('child_process');
const path = require('path');

const upload = multer({ dest: 'uploads/' });
const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Post-processing HTML to handle highlights, indentation
function postProcessHtml(html) {
    return html
        .replace(/<p class="Quote">/g, '<blockquote>')
}

// ================================
// Export DOCX
// ================================

app.post('/export-docx', async (req, res) => {
    try {
        const buffer = await htmlToDocx(req.body.html);

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        );

        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).send('DOCX export failed');
    }
});

// ================================
// Import DOCX: merge Mammoth + docx-preview
// ================================

/* //first basic draft
app.post('/import-docx', upload.single('file'), async (req, res) => {
    const result = await mammoth.convertToHtml({ path: req.file.path });
    fs.unlinkSync(req.file.path);
    res.json({ html: result.value });
}); */

app.post('/import-docx', upload.single('file'), async (req, res) => {
    const inputPath = req.file.path;
    const scriptPath = path.join(__dirname, 'ruby', 'parse_docx.rb');

    let stdout = '';
    let stderr = '';

    try {
        const parsed = await new Promise((resolve, reject) => {
            const child = execFile(
                'ruby',
                [scriptPath, inputPath],
                { timeout: 15000, maxBuffer: 20 * 1024 * 1024 },
                (err, out, errOut) => {
                    stdout = out || '';
                    stderr = errOut || '';

                    if (err) return reject(err);

                    try {
                        resolve(JSON.parse(stdout));
                    } catch (e) {
                        reject(new Error(`Ruby returned non-JSON.\nSTDOUT=${stdout}\nSTDERR=${stderr}`));
                    }
                }
            );

            child.on('error', reject); // e.g. ruby not found
        });

        if (!parsed.ok || typeof parsed.html !== 'string') {
            throw new Error(`Ruby parse failed: ${JSON.stringify(parsed)}`);
        }

        const html = postProcessHtml(parsed.html);

        res.json({ html });
    } catch (err) {
        console.error(err);
        res.status(500).json({
            error: 'DOCX import failed (Ruby parse_docx.rb)',
            details: err.message,
            stderr,
        });
    } finally {
        try { fs.unlinkSync(inputPath); } catch {}
    }
});

app.listen(3000, () => {
    console.log('DOCX server running on http://localhost:3000');
});