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
    const outputPath = inputPath + '.html';

    try {
        await new Promise((resolve, reject) => {
            execFile(
                'pandoc',
                [
                    inputPath,
                    '-f', 'docx',
                    '-t', 'html',
                    '--standalone',
                    '--wrap=none',
                    '--lua-filter=' + __dirname + '/pandoc-filters/highlight.lua',
                    '-o', outputPath
                ],
                (error) => {
                    if (error) reject(error);
                    else resolve();
                }
            );
        });

        const rawHtml = fs.readFileSync(outputPath, 'utf8');
        const html = postProcessHtml(rawHtml);

        fs.unlinkSync(inputPath);
        fs.unlinkSync(outputPath);

        res.json({ html });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'DOCX import failed (Pandoc)' });
    }
});

app.listen(3000, () => {
    console.log('DOCX server running on http://localhost:3000');
});