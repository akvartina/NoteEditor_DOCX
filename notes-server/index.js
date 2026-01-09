const express = require('express');
const cors = require('cors');
const multer = require('multer');
const htmlToDocx = require('html-to-docx');
const XLSX = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const { execFile } = require('child_process');
const path = require('path');
const { v4: uuidv4 } = require('uuid');

const upload = multer({ dest: 'uploads/' });
const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use('/pptx-assets', express.static(path.join(__dirname, 'storage', 'pptx')));

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
// Import DOCX
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

// ================================
// Export XLSX
// ================================
app.post('/export-xlsx', async (req, res) => {
    try {
        const { sheets } = req.body || {};
        if (!Array.isArray(sheets) || sheets.length === 0) {
            return res.status(400).json({ ok: false, error: 'Missing sheets[]' });
        }

        const wb = XLSX.utils.book_new();

        for (const s of sheets) {
            const name = (s && s.name) ? String(s.name).slice(0, 31) : 'Sheet1';
            const cells = Array.isArray(s.cells) ? s.cells : [];
            const ws = XLSX.utils.aoa_to_sheet(cells);
            XLSX.utils.book_append_sheet(wb, ws, name);
        }

        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');
        res.send(buf);
    } catch (err) {
        console.error(err);
        res.status(500).json({ ok: false, error: 'XLSX export failed', details: err.message });
    }
});

// ================================
// Import XLSX
// ================================

app.post('/import-xlsx', upload.single('file'), async (req, res) => {
    const inputPath = req.file.path;
    const scriptPath = path.join(__dirname, 'ruby', 'parse_xlsx.rb');

    let stderr = '';

    try {
        const parsed = await new Promise((resolve, reject) => {
            execFile(
                'ruby',
                [scriptPath, inputPath],
                { timeout: 20000, maxBuffer: 50 * 1024 * 1024 },
                (err, stdout, errOut) => {
                    stderr = errOut || '';
                    if (err) return reject(err);

                    try {
                        resolve(JSON.parse(stdout));
                    } catch (e) {
                        reject(new Error(`Ruby returned non-JSON.\nSTDERR=${stderr}`));
                    }
                }
            );
        });

        if (!parsed.ok) throw new Error(parsed.error || 'Unknown Ruby parse error');

        res.json(parsed); // { ok: true, sheets: [...] }
    } catch (err) {
        console.error(err);
        res.status(500).json({
            ok: false,
            error: 'XLSX import failed (Ruby parse_xlsx.rb)',
            details: err.message,
            stderr,
        });
    } finally {
        try { fs.unlinkSync(inputPath); } catch {}
    }
});

// ================================
// Export PPTX
// ================================
app.post('/export-pptx', async (req, res) => {
    try {
        const { slides } = req.body || {};
        if (!Array.isArray(slides)) return res.status(400).json({ ok:false, error:'slides[] missing' });

        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_WIDE';

        for (const s of slides) {
            const slide = pptx.addSlide();

            const objects = Array.isArray(s.objects) ? s.objects.slice().sort((a,b)=> (a.z||0)-(b.z||0)) : [];
            for (const o of objects) {
                if (o.type === 'text') {
                    slide.addText(o.text || '', {
                        x: o.x || 0, y: o.y || 0, w: o.w || 1, h: o.h || 1,
                        fontSize: 18,
                        valign: 'top'
                    });
                } else if (o.type === 'image') {
                    // src is a URL like /pptx-assets/<deckId>/media/image.png
                    // Convert to local path:
                    const localPath = path.join(__dirname, o.src.replace('/pptx-assets/', 'storage/pptx/'));
                    slide.addImage({
                        path: localPath,
                        x: o.x || 0, y: o.y || 0, w: o.w || 1, h: o.h || 1
                    });
                }
            }
        }

        const buf = await pptx.write('nodebuffer');

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.setHeader('Content-Disposition', 'attachment; filename="export.pptx"');
        res.send(buf);
    } catch (e) {
        console.error(e);
        res.status(500).json({ ok:false, error:'PPTX export failed', details: e.message });
    }
});


// ================================
// Import PPTX
// ================================
app.post('/import-pptx', upload.single('file'), async (req, res) => {
    const inputPath = req.file.path;
    const deckId = uuidv4();
    const deckDir = path.join(__dirname, 'storage', 'pptx', deckId);
    fs.mkdirSync(deckDir, { recursive: true });

    const scriptPath = path.join(__dirname, 'ruby', 'parse_pptx_objects.rb');

    try {
        const parsed = await new Promise((resolve, reject) => {
            execFile(
                'ruby',
                [scriptPath, inputPath, deckDir],
                { timeout: 30000, maxBuffer: 100 * 1024 * 1024 },
                (err, stdout, stderr) => {
                    if (err) return reject(new Error(stderr || err.message));
                    try { resolve(JSON.parse(stdout)); }
                    catch { reject(new Error(`Ruby returned non-JSON. STDERR=${stderr}`)); }
                }
            );
        });

        if (!parsed.ok) throw new Error(parsed.error || 'Parse failed');

        // Persist model
        fs.writeFileSync(path.join(deckDir, 'model.json'), JSON.stringify(parsed, null, 2), 'utf8');

        res.json({ ok: true, deckId, slides: parsed.slides });
    } catch (e) {
        console.error(e);
        res.status(500).json({ ok:false, error:'PPTX import failed', details: e.message });
    } finally {
        try { fs.unlinkSync(inputPath); } catch {}
    }
});

app.listen(3000, () => {
    console.log('Notes server running on http://localhost:3000');
});