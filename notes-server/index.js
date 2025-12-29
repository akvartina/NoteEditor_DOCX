//MAMMOTH STYLE-MAP
const STYLE_MAP = [
    "p[style-name='Heading 1'] => h1:fresh",
    "p[style-name='Heading 2'] => h2:fresh",
    "p[style-name='Heading 3'] => h3:fresh",
    "p[style-name='Quote'] => blockquote:fresh",
    "p[style-name='Bibliography'] => p.bibliography",
    "r[style-name='Emphasis'] => em",
    "r[style-name='Strong'] => strong",
    "r[style-name='Citation'] => span.citation"
];

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const htmlToDocx = require('html-to-docx');
const mammoth = require('mammoth');
const fs = require('fs');

const upload = multer({ dest: 'uploads/' });
const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Post-processing HTML to handle highlights, indentation
function postProcessHtml(html) {
    return html
        // Convert Word highlights to <mark>
        .replace(
            /<span style="background-color:\s*(#[0-9a-fA-F]{3,6}|[^;]+);?">/g,
            '<mark style="background-color:$1">'
        )

        // Convert indentation (pt → px)
        .replace(
            /margin-left:\s*([0-9.]+)pt/g,
            (_, pt) => `margin-left:${pt * 1.333}px`
        )
        // Preserve empty paragraphs
        .replace(/<p>\s*<\/p>/g, '<p>&nbsp;</p>');
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
    try {
        const result = await mammoth.convertToHtml(
            { path: req.file.path },
            {
                styleMap: STYLE_MAP,
                includeDefaultStyleMap: true,
                convertImage: mammoth.images.inline(function (image) {
                    return image.read("base64").then(function (imageBuffer) {
                        return {
                            src: "data:" + image.contentType + ";base64," + imageBuffer
                        };
                    });
                })
            }
        );

        fs.unlinkSync(req.file.path);

        const html = postProcessHtml(result.value);

        res.json({
            html,
            messages: result.messages
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'DOCX import failed' });
    }
});

app.listen(3000, () => {
    console.log('DOCX server running on http://localhost:3000');
});