import os
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pdfplumber
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter
from io import BytesIO
import fitz
import logging
import json
import re
import telebot

logging.getLogger("pdfminer").setLevel(logging.ERROR)

app = Flask(__name__)
CORS(app)

PATTERNS = [
    # comment
    ("Name",       "comment",   r"%\s*name\s*-\s*(.+)",                       99, "name"),
    ("PRN",        "comment",   r"%\s*roll\s*n\.?o\.?\s*-\s*(\S+)",           1,  "roll n.o"),
    # dashcolon
    ("Name",       "dashcolon", r"(?:^|\s)NAME\s*:-\s*(.+)",                  99, "NAME"),
    ("Name",       "dashcolon", r"(?:^|\s)Name\s*of\s*Student\s*:-\s*(.+)",   3,  "Name of Student"),
    ("PRN",        "dashcolon", r"(?:^|\s)PRN\s*No\.?\s*:-\s*(\S+)",          1,  "PRN No."),
    ("PRN",        "dashcolon", r"(?:^|\s)PRN\s*:-\s*(\S+)",                  1,  "PRN"),
    ("Batch",      "dashcolon", r"(?:^|\s)Batch\s*:-\s*(\S+)",                1,  "Batch"),
    ("Class",      "dashcolon", r"(?:^|\s)Class\s*:-\s*(\S+)",                1,  "Class"),
    ("Div",        "dashcolon", r"(?:^|\s)Div(?:ision)?\s*:-\s*(\S+)",        1,  "Div"),
    # colon
    ("Name",       "colon",     r"(?:^|\s)Name\s*of\s*Student\s*:?\s*(.+)",   3,  "Name of Student"),
    ("Name",       "colon",     r"(?:^|\s)Name\s*:(.+)",                      3,  "Name"),
    ("PRN",        "colon",     r"(?:^|\s)PRN\s*NO\.?\s*:\s*(\S+)",           1,  "PRN NO."),
    ("PRN",        "colon",     r"(?:^|\s)PRN\s*No\.?\s*:\s*(\S+?)\.?(?:\s|$)", 1, "PRN No."),
    ("PRN",        "colon",     r"(?:^|\s)PRN\s*:\s*(\S+)",                   1,  "PRN"),
    ("Batch",      "colon",     r"(?:^|\s)Batch\s*:(\S+?)\.?(?:\s|$)",        1,  "Batch"),
    ("Class",      "colon",     r"(?:^|\s)Class\s*:(\S+?)\.?(?:\s|$)",        1,  "Class"),
    ("Div",        "colon",     r"(?:^|\s)Div(?:ision)?\s*:(\S+?)\.?(?:\s|$)",1, "Division"),
    ("Drawn By",   "colon",     r"(?:^|\s)DRAWN\s*BY\s*:\s*([A-Za-z].+?)(?:\s+PRN|\s*$)", 2, "DRAWN BY"),
    ("Date",       "colon",     r"(?:^|\s)DATE\s*:\s*(\d\S*)",                    1,  "DATE"),
    ("Checked By", "colon",     r"(?:^|\s)CHECKED\s*BY\s*:\s*(?!SCALE|REVISION|MARKS|DRAWN|PAGE|FIRST)([A-Za-z0-9]\S*)", 1, "CHECKED BY"),
    ("Marks",      "colon",     r"(?:^|\s)MARKS\s*:\s*(?!REVISION|SCALE|CHECKED|DRAWN|PAGE|FIRST)([A-Za-z0-9]\S*)",      1, "MARKS"),
    ("Scale",      "colon",     r"(?:^|\s)SCALE\s*:\s*(\S+)",                     1,  "SCALE"),
    ("Sheet No",   "colon",     r"(?:^|\s)SHEET\s*NO\.?\s*:\s*(\S+)",             1,  "SHEET NO."),
    ("Revision",   "colon",     r"(?:^|\s)REVISION\s*NO\.?\s*:\s*(?!SCALE|MARKS|CHECKED|DRAWN|PAGE|FIRST)([A-Za-z0-9]\S*)", 1, "REVISION NO."),
]


def detect_fields(pdf_bytes):
    fields, seen = [], set()
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            lines = {}
            for w in words:
                key = round(w["top"] / 5) * 5
                lines.setdefault(key, []).append(w)
            for _, lw in sorted(lines.items()):
                line = " ".join(w["text"] for w in lw)
                for field, fmt, pat, wc, label_key in PATTERNS:
                    if field in seen:
                        continue
                    m = re.search(pat, line, re.IGNORECASE)
                    if not m:
                        continue
                    raw   = m.group(1).strip().lstrip("-").strip().rstrip(".")
                    parts = raw.split()
                    value = " ".join(parts) if wc == 99 else " ".join(parts[:wc])
                    value = value.lstrip(":").strip()
                    if value:
                        fields.append({"field": field, "value": value, "format": fmt,
                                       "label": label_key, "word_count": len(value.split())})
                        seen.add(field)
    return {"fields": fields}


def sample_colors(pdf_bytes, page_num, bbox):
    try:
        doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[page_num]
        pw, ph = page.rect.width, page.rect.height
        x0, top, x1, bottom = bbox[:4]
        mat = fitz.Matrix(4, 4)

        def avg(clip):
            if clip.is_empty or clip.width < 1 or clip.height < 1:
                return None
            pix = page.get_pixmap(matrix=mat, clip=clip)
            s, n = pix.samples, len(pix.samples) // 3
            if n == 0: return None
            return (sum(s[i*3] for i in range(n))/n/255,
                    sum(s[i*3+1] for i in range(n))/n/255,
                    sum(s[i*3+2] for i in range(n))/n/255)

        bg = None
        for clip in [
            fitz.Rect(x1+5, top, min(pw,x1+120), bottom),
            fitz.Rect(max(0,x0-120), top, x0-2, bottom),
            fitz.Rect(x0, max(0,top-8), x1, top),
            fitz.Rect(x0, bottom, x1, min(ph,bottom+8)),
        ]:
            bg = avg(clip)
            if bg: break
        bg = bg or (1.0, 1.0, 1.0)

        tc = (0.0, 0.0, 0.0)
        inner = fitz.Rect(x0, top, x1, bottom)
        if not inner.is_empty and inner.width > 2:
            pix = page.get_pixmap(matrix=mat, clip=inner)
            s, n = pix.samples, len(pix.samples) // 3
            if n > 0:
                px = sorted([(s[i*3]/255, s[i*3+1]/255, s[i*3+2]/255) for i in range(n)],
                             key=lambda p: p[0]+p[1]+p[2])
                dk = px[:max(1, n//7)]
                tc = (sum(p[0] for p in dk)/len(dk),
                      sum(p[1] for p in dk)/len(dk),
                      sum(p[2] for p in dk)/len(dk))
        doc.close()
        return bg, tc
    except Exception as e:
        print(f"Color error: {e}")
        return (1.0,1.0,1.0), (0.0,0.0,0.0)


def font_info(pl_page, word):
    x0, x1, top, bot = word["x0"], word["x1"], word["top"], word["bottom"]
    pad = 2
    sizes, is_bold = [], False
    for ch in pl_page.chars:
        if (x0-pad)<=ch["x0"]<=(x1+pad) and (top-pad)<=ch["top"]<=(bot+pad):
            if "size" in ch: sizes.append(float(ch["size"]))
            if any(k in ch.get("fontname","").lower() for k in ("bold","heavy","black")):
                is_bold = True
    return (sum(sizes)/len(sizes) if sizes else 10.0), is_bold


def locate_value(pl_page, field_info):
    fmt, label, word_count = field_info["format"], field_info["label"], field_info["word_count"]
    words = pl_page.extract_words()

    def make_bbox(vw):
        f = vw[0]
        if f["text"].startswith(":"):
            cw = (f["x1"]-f["x0"]) / max(len(f["text"]),1)
            ex, dx = f["x0"], f["x0"]+cw
        else:
            ex = dx = f["x0"]
        return (ex, min(v["top"] for v in vw), max(v["x1"] for v in vw),
                max(v["bottom"] for v in vw), dx)

    # ── Comment ──────────────────────────────────────────────────────────────
    if fmt == "comment":
        lp = label.lower().split()
        for w in words:
            if not w["text"].lstrip("%").lower().startswith(lp[0].rstrip("-")):
                continue
            lt = w["top"]
            lw = [ww for ww in words if abs(ww["top"]-lt)<5]
            lt_text = " ".join(ww["text"] for ww in lw)
            if "-" not in lt_text: continue
            fs, _ = font_info(pl_page, w)
            dp = lt_text.index("-")
            vs = lt_text[dp+1:].strip()
            if not vs: continue
            vp = vs.split()
            wc = len(vp) if word_count==99 else min(word_count,len(vp))
            vp = vp[:wc]
            run, dwi = 0, None
            for li2, lw2 in enumerate(lw):
                end = run + len(lw2["text"])
                if run <= dp <= end: dwi=li2; break
                run += len(lw2["text"])+1
            if dwi is None: continue
            si = dwi+1 if lw[dwi]["text"].endswith("-") else dwi
            val_words = []
            for vp2 in vp:
                for cw2 in lw[si:]:
                    cv = cw2["text"].split("-")[-1] if "-" in cw2["text"] else cw2["text"]
                    if cv.lower()==vp2.lower() and cw2 not in val_words:
                        val_words.append(cw2); break
            if not val_words: continue
            x0s = []
            for vi,vw2 in enumerate(val_words):
                x = vw2["x0"]
                if "-" in vw2["text"] and vi==0:
                    di2 = vw2["text"].index("-")
                    cw3 = (vw2["x1"]-vw2["x0"])/max(len(vw2["text"]),1)
                    x = vw2["x0"]+cw3*(di2+1)
                x0s.append(x)
            return (min(x0s), min(v["top"] for v in val_words),
                    max(v["x1"] for v in val_words),
                    max(v["bottom"] for v in val_words), min(x0s)), fs, False

    # ── Dashcolon ─────────────────────────────────────────────────────────────
    elif fmt == "dashcolon":
        lp = label.lower().split()
        for i, w in enumerate(words):
            if w["text"].lower().rstrip(".:") != lp[0].rstrip(".:"):
                continue
            lt = w["top"]
            j, lpi = i+1, 1
            while lpi < len(lp) and j < len(words):
                if abs(words[j]["top"]-lt) > 8: break
                if words[j]["text"].lower().rstrip(".:") == lp[lpi].rstrip(".:"):
                    lpi += 1
                j += 1
            if lpi < len(lp): continue
            while j < len(words) and words[j]["text"] in (":-","-",":"):
                j += 1
            vw = [v for v in words[j:j+word_count] if abs(v["top"]-lt)<8]
            if not vw: continue
            fs, ib = font_info(pl_page, w)
            vx = vw[0]["x0"]
            return (vx, min(v["top"] for v in vw), max(v["x1"] for v in vw),
                    max(v["bottom"] for v in vw), vx), fs, ib

    # ── Colon ─────────────────────────────────────────────────────────────────
    else:
        lw2 = label.lower().split()
        for i, w in enumerate(words):
            if w["text"].split(":")[0].rstrip(",.").lower() != lw2[0].rstrip(":."):
                continue
            lt = w["top"]
            j, li2 = i+1, 1
            llw = w
            while li2 < len(lw2) and j < len(words):
                wj = words[j]
                if abs(wj["top"]-lt) > 8: break
                wb = wj["text"].split(":")[0].rstrip(",.").lower()
                if wb == lw2[li2].rstrip(":."):
                    llw = wj; li2 += 1
                j += 1
            if li2 < len(lw2): continue
            fs, ib = font_info(pl_page, w)
            lt2 = llw["text"]
            if ":" in lt2:
                ac = lt2[lt2.index(":")+1:].strip().rstrip(".")
                if ac:
                    cp = lt2.index(":")
                    cw4 = (llw["x1"]-llw["x0"]) / max(len(lt2),1)
                    vx = llw["x0"] + cw4*(cp+1)
                    ex = [v for v in words[j:j+max(0,word_count-1)] if abs(v["top"]-lt)<8]
                    ax = max([llw["x1"]]+[v["x1"] for v in ex])
                    ab = max([llw["bottom"]]+[v["bottom"] for v in ex])
                    return (vx, llw["top"], ax, ab, vx), fs, ib
            while j < len(words) and words[j]["text"] in (":",":-","-","."):
                j += 1
            vw = [v for v in words[j:j+word_count] if abs(v["top"]-lt)<8]
            if not vw: continue
            return make_bbox(vw), fs, ib

    return None, None, False


def replace_fields_in_pdf(pdf_bytes, replacements_list):
    reader = PdfReader(BytesIO(pdf_bytes))
    writer = PdfWriter()
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(reader.pages):
            pl_page    = pdf.pages[page_num]
            page_w     = float(page.mediabox.width)
            page_h     = float(page.mediabox.height)
            reps = []
            for item in replacements_list:
                bbox, fs, ib = locate_value(pl_page, item)
                if bbox is None:
                    print(f"  Page {page_num+1}: '{item['label']}' not found"); continue
                fmt = item["format"]
                if fmt == "comment":
                    bg, tc = sample_colors(pdf_bytes, page_num, bbox)
                else:
                    bg = (1.0,1.0,1.0)
                    _, tc = sample_colors(pdf_bytes, page_num, bbox)
                print(f"  '{item['label']}' -> '{item['new_value']}' fs={fs:.1f}")
                reps.append((bbox, item["new_value"], fs, bg, tc, fmt, ib))
            if not reps:
                writer.add_page(page); continue
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=(page_w, page_h))
            for bbox, new_text, fs, bg, tc, fmt, ib in reps:
                ex0, top, x1, bot, dx0 = bbox
                ow = x1 - ex0
                h  = bot - top
                ry = page_h - bot
                font = ("Courier-Bold" if ib else "Courier") if fmt=="comment" \
                       else ("Helvetica-Bold" if ib else "Helvetica")

                # ── Per-format padding ───────────────────────────────────────
                if fmt == "comment":
                    pl, pr, pt, pb, tv = 0, 0, 2, 0, 0.25
                elif fmt == "dashcolon":
                    pl, pr, pt, pb, tv = 0, 0, 2, 0, 0.25
                else:
                    pl, pr, pt, pb, tv = 0, 0, 2, 0, 0.25
                # ─────────────────────────────────────────────────────────────

                ntw = c.stringWidth(new_text, font, fs)
                if ntw > ow:
                    fs = fs * ow / ntw; ntw = ow

                # Use font ascent/descent to get exact text coverage
                # This avoids erasing border lines while fully covering old text
                from reportlab.pdfbase.pdfmetrics import getFont
                try:
                    face       = getFont(font).face
                    ascent     = face.ascent  * fs / 1000.0
                    descent    = abs(face.descent) * fs / 1000.0
                except:
                    ascent  = fs * 0.8
                    descent = fs * 0.2

                # rect covers exactly the font's ink area
                # ry is baseline-ish (page_h - bottom of bbox)
                # text baseline = ry + h*tv
                baseline  = ry + h * tv
                rect_bot  = baseline - descent - 0.1    # extra 2pt down to cover letter bottoms
                rect_top  = baseline + ascent
                rect_h    = rect_top - rect_bot

                rx  = ex0 - pl
                rw  = min(ow + pl + pr, page_w - rx)

                c.setFillColorRGB(*bg)
                c.rect(rx, rect_bot, rw, rect_h, fill=True, stroke=False)
                c.setFont(font, fs)
                c.setFillColorRGB(*tc)
                c.drawString(dx0, baseline, new_text)
            c.save()
            packet.seek(0)
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
    out = BytesIO()
    writer.write(out)
    return out.getvalue()


@app.route("/")
def index():
    with open(os.path.join(os.path.dirname(__file__), "index.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.route("/api/detect", methods=["POST"])
def detect():
    if "pdf" not in request.files:
        return jsonify({"error": "No PDF uploaded"}), 400
    try:
        return jsonify(detect_fields(request.files["pdf"].read()))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/replace", methods=["POST"])
def replace():
    if "pdf" not in request.files:
        return jsonify({"error": "No PDF uploaded"}), 400
    pdf_bytes    = request.files["pdf"].read()
    replacements = json.loads(request.form.get("replacements", "[]"))
    try:
        result = replace_fields_in_pdf(pdf_bytes, replacements)
        return send_file(BytesIO(result), mimetype="application/pdf",
                         as_attachment=True, download_name="modified.pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ─── TELEGRAM BOT ────────────────────────────────────────────────────────────
BOT_TOKEN = "8725089715:AAE7pY5hd7ao4FHYB3nXETv2DhLIj7LF_Co"
bot = telebot.TeleBot(BOT_TOKEN)

# Store user state: waiting for field selections after PDF upload
user_sessions = {}  # chat_id -> {pdf_bytes, fields}

@bot.message_handler(commands=["start", "help"])
def start(message):
    bot.reply_to(message, 
        "👋 Welcome to PDF Name Editor Bot!\n\n"
        "📄 Send me a PDF and I'll detect fields like Name, PRN, Batch etc.\n"
        "Then you can edit them and I'll send back the modified PDF!"
    )

@bot.message_handler(content_types=["document"])
def handle_pdf(message):
    chat_id = message.chat.id
    doc = message.document

    if doc.mime_type != "application/pdf":
        bot.reply_to(message, "❌ Please send a PDF file!")
        return

    bot.reply_to(message, "⏳ Analyzing your PDF...")

    # Download the PDF
    file_info = bot.get_file(doc.file_id)
    downloaded = bot.download_file(file_info.file_path)

    # Detect fields using your existing function
    result = detect_fields(downloaded)
    fields = result.get("fields", [])

    if not fields:
        bot.reply_to(message, "⚠️ No editable fields found in this PDF.")
        return

    # Save session
    user_sessions[chat_id] = {
        "pdf_bytes": downloaded,
        "fields": fields,
        "replacements": [],
        "current_index": 0
    }

    # Ask for first field
    ask_next_field(chat_id, message.message_id)

def ask_next_field(chat_id, reply_to=None):
    session = user_sessions.get(chat_id)
    if not session:
        return

    idx = session["current_index"]
    fields = session["fields"]

    if idx >= len(fields):
        # All fields done — apply replacements
        finish_editing(chat_id)
        return

    field = fields[idx]
    text = (
        f"📝 Field {idx+1}/{len(fields)}: *{field['field']}*\n"
        f"Current value: `{field['value']}`\n\n"
        f"Reply with new value, or type `skip` to keep it."
    )
    bot.send_message(chat_id, text, parse_mode="Markdown")

@bot.message_handler(func=lambda m: m.chat.id in user_sessions and 
                     user_sessions[m.chat.id]["current_index"] < len(user_sessions[m.chat.id]["fields"]))
def handle_field_reply(message):
    chat_id = message.chat.id
    session = user_sessions[chat_id]
    idx = session["current_index"]
    field = session["fields"][idx]
    user_input = message.text.strip()

    if user_input.lower() != "skip":
        session["replacements"].append({
            "field":      field["field"],
            "format":     field["format"],
            "label":      field["label"],
            "word_count": field["word_count"],
            "new_value":  user_input
        })

    session["current_index"] += 1
    ask_next_field(chat_id)

def finish_editing(chat_id):
    session = user_sessions.pop(chat_id, None)
    if not session or not session["replacements"]:
        bot.send_message(chat_id, "⚠️ No changes made.")
        return

    bot.send_message(chat_id, "⚙️ Applying changes to your PDF...")

    try:
        result_bytes = replace_fields_in_pdf(session["pdf_bytes"], session["replacements"])
        bot.send_document(
            chat_id,
            ("modified.pdf", BytesIO(result_bytes)),
            caption="✅ Here's your modified PDF!"
        )
    except Exception as e:
        bot.send_message(chat_id, f"❌ Error: {str(e)}")

# Webhook endpoint for Telegram
@app.route("/telegram", methods=["POST"])
def telegram_webhook():
    if request.headers.get("content-type") == "application/json":
        update = telebot.types.Update.de_json(request.get_data(as_text=True))
        bot.process_new_updates([update])
    return "OK", 200
# ─────────────────────────────────────────────────────────────────────────────


if __name__ == "__main__":
    print("\n✅ Server running! Open http://localhost:5000 in your browser\n")
    app.run(debug=True, port=5000)
