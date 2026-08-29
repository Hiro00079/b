"""
Helical Anchor Inventory Trail Lookup
Run: pip install flask openpyxl
Then: python app.py
Open: http://localhost:5000
"""

from flask import Flask, request, render_template_string
import openpyxl
import re
import os
import json
import base64
import tempfile

app = Flask(__name__)
app.secret_key = "ha_inventory_key"

# NOTE: Vercel's Python functions are serverless — there is no guarantee that
# two requests (e.g. /upload then /search) hit the same instance, so we can't
# rely on a global dict or a saved temp file to persist inventory data between
# requests. Instead, the parsed inventory is sent back to the browser as a
# hidden form field (base64-encoded JSON) and posted back on every /search
# request. This keeps the app fully stateless.

def encode_inv(inv):
    return base64.b64encode(json.dumps(inv).encode()).decode()

def decode_inv(blob):
    return json.loads(base64.b64decode(blob.encode()).decode())

# ─────────────────────────────────────────────
# STEP 1: Load Excel into a flat dict
# key = item_code (str), value = {desc, qty}
# ─────────────────────────────────────────────
def load_inventory(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    inv = {}

    rows = ws.iter_rows(values_only=True)
    header = next(rows)  # first row = column names

    def find_col(*names):
        """Find a column index whose header matches any of the given names
        (case-insensitive, trimmed)."""
        for i, h in enumerate(header):
            if h and str(h).strip().lower() in [n.lower() for n in names]:
                return i
        return None

    code_col = find_col("Item No.", "Item No", "Item Code")
    desc_col = find_col("Item Description", "Description")
    qty_col  = find_col("In Stock", "Stock", "Quantity")

    if code_col is None or desc_col is None or qty_col is None:
        raise ValueError(
            f"Could not find required columns in header: {header}"
        )

    for row in rows:
        code = row[code_col] if code_col < len(row) else None
        desc = row[desc_col] if desc_col < len(row) else None
        qty  = row[qty_col]  if qty_col  < len(row) else None
        if code and str(code).strip():
            inv[str(code).strip()] = {
                "desc": str(desc).strip() if desc else "",
                "qty":  qty if qty is not None else 0
            }
    return inv

# ─────────────────────────────────────────────
# STEP 2: Parse HAFG description into parts
#
# LEAD format:    238-254-10L-10-12-14-G
# EXT format:     278-217-7E-G  or  278-217-3E-19-G
#
# Returns dict:
#   od, wall, height, ptype (L/E), flights[], galv (bool)
# ─────────────────────────────────────────────
def parse_desc(desc):
    # Remove trailing -G or G
    raw = desc.strip()
    galv = raw.endswith("-G") or raw.endswith("G")
    raw = re.sub(r"-?G$", "", raw)

    parts = raw.split("-")
    # parts[0]=OD, parts[1]=wall, parts[2]=height+type, rest=flights
    od   = parts[0]   # e.g. 238
    wall = parts[1]   # e.g. 254
    ht_raw = parts[2] # e.g. 10L or 7E or 3E

    # split height and type
    m = re.match(r"([\d./']+)([LE])", ht_raw, re.IGNORECASE)
    if not m:
        return None
    height = m.group(1)
    ptype  = m.group(2).upper()   # L or E

    flights = parts[3:] if len(parts) > 3 else []

    return {
        "od":      od,
        "wall":    wall,
        "height":  height,
        "ptype":   ptype,
        "flights": flights,
        "galv":    galv,
        "raw":     desc
    }

# ─────────────────────────────────────────────
# STEP 3: Search helpers
# Each returns (item_code, desc, qty) or None
# ─────────────────────────────────────────────

def find_exact(inv, prefix, target_desc):
    """Find item where code starts with prefix and desc matches target_desc exactly."""
    for code, v in inv.items():
        if code.startswith(prefix) and v["desc"].strip().lower() == target_desc.strip().lower():
            return (code, v["desc"], v["qty"])
    return None

def find_contains(inv, prefix, *substrings):
    """Find item where code starts with prefix and desc contains ALL substrings (case-insensitive)."""
    for code, v in inv.items():
        if code.startswith(prefix):
            d = v["desc"].lower()
            if all(s.lower() in d for s in substrings):
                return (code, v["desc"], v["qty"])
    return None

def find_all_contains(inv, prefix, *substrings):
    """Like find_contains but returns ALL matches."""
    results = []
    for code, v in inv.items():
        if code.startswith(prefix):
            d = v["desc"].lower()
            if all(s.lower() in d for s in substrings):
                results.append((code, v["desc"], v["qty"]))
    return results

# ─────────────────────────────────────────────
# STEP 4: Build the full trail for one HAFG item
# Returns a list of stage dicts for rendering
# ─────────────────────────────────────────────

def build_trail(hafg_code, inv):
    if hafg_code not in inv:
        return None, f"{hafg_code} not found in inventory"

    desc = inv[hafg_code]["desc"]
    qty  = inv[hafg_code]["qty"]
    p    = parse_desc(desc)
    if not p:
        return None, f"Could not parse description: {desc}"

    od      = p["od"]        # e.g. 238
    wall    = p["wall"]      # e.g. 254
    height  = p["height"]    # e.g. 10
    ptype   = p["ptype"]     # L or E
    flights = p["flights"]   # e.g. ['10','12','14']
    galv    = p["galv"]

    g_suffix = "-G" if galv else ""
    is_lead = (ptype == "L")
    stages  = []

    # ── helper to add a stage row ──────────────
    def row(stage, label, result, note=""):
        if result:
            code, rdesc, rqty = result
            stages.append({
                "stage": stage, "label": label,
                "code": code, "desc": rdesc,
                "qty": rqty, "note": note,
                "found": True
            })
        else:
            stages.append({
                "stage": stage, "label": label,
                "code": "—", "desc": "Not found",
                "qty": 0, "note": note,
                "found": False
            })

    # ── STAGE 0: The FG item itself ────────────
    row("FG", "Finished good (HAFG)",
        (hafg_code, desc, qty))

    # ── STAGE 1: HAFGP — packed ────────────────
    # desc without -G + -G  e.g. 238-254-10L-10-12-14-G
    packed_desc = desc  # same description in HAFGP
    row("HAFGP", "Packed product (HAFGP)",
        find_contains(inv, "HAFGP", od, wall, f"{height}{ptype}",
                      *flights))

    # ── STAGE 2: HAFWFG welding ────────────────
    if is_lead:
        # LEAD: only box + pipe (-B)
        b_desc = f"{od}-{wall}-{height}L-B"
        row("HAFWFG", f"Box welded to pipe (HAFWFG) — {b_desc}",
            find_contains(inv, "HAFWFG", od, wall, f"{height}L", "-B"))
    else:
        # EXTENSION: box only (-B), box+pin (-BP), pin only (-P)
        # Note: search -B but exclude -BP hits by checking desc ends with -B or has -B-
        b_match = find_contains(inv, "HAFWFG", od, wall, f"{height}E", "-B")
        if b_match and "-BP" in b_match[1]:
            b_match = None   # don't count a BP match as a B-only match
        row("HAFWFG", f"Box + pipe — {od}-{wall}-{height}E-B", b_match)
        row("HAFWFG", f"Box + pin — {od}-{wall}-{height}E-BP",
            find_contains(inv, "HAFWFG", od, wall, f"{height}E", "-BP"))
        row("HAFWFG", f"Pin only — {od}-{wall}-{height}E-P",
            find_contains(inv, "HAFWFG", od, wall, f"{height}E", "-P"))

    # ── STAGE 3: HAMWFG — flights welded ───────
    if flights:
        fl_str = "-".join(flights)
        # Match desc must equal EXACTLY od-wall-htptype-flight1-flight2...
        # This prevents "12" matching "10-12" or "12-14"
        def find_hamwfg_exact(inv, od, wall, ht_ptype, flights):
            # Build the full expected description (without galv suffix)
            expected = f"{od}-{wall}-{ht_ptype}-" + "-".join(flights)
            for c, v in inv.items():
                if not c.startswith("HAMWFG"):
                    continue
                d = v["desc"].strip()
                # Strip trailing spaces, galv suffixes, or extra notes
                # Compare core desc: split on space and take first token
                core = d.split(" ")[0].rstrip("-")
                if core.lower() == expected.lower():
                    return (c, v["desc"], v["qty"])
            return None

        row("HAMWFG", f"Flights welded (HAMWFG) — {od}-{wall}-{height}{ptype}-{fl_str}",
            find_hamwfg_exact(inv, od, wall, f"{height}{ptype}", flights))
    else:
        stages.append({
            "stage": "HAMWFG", "label": "Flights welded (HAMWFG)",
            "code": "—", "desc": "No flights on this item",
            "qty": "—", "note": "Extension with no flights", "found": None
        })

    # ── STAGE 4: HACNCFG ─────────────────────
    # box with wall
    row("HACNCFG", f"CNC box — Box{od}-{wall}",
        find_contains(inv, "HACNCFG", f"Box{od}-{wall}"))
    # pin (extension only)
    if not is_lead:
        row("HACNCFG", f"CNC pin — Pin{od}-{wall}",
            find_contains(inv, "HACNCFG", f"Pin{od}-{wall}"))

    # ── STAGE 5: HABSFG ──────────────────────
    # cut box
    row("HABSFG", f"Cut box — Box-{od}-C",
        find_contains(inv, "HABSFG", f"Box-{od}-C"))
    # cut pin (extension only)
    if not is_lead:
        row("HABSFG", f"Cut pin — Pin-{od}-C",
            find_contains(inv, "HABSFG", f"Pin-{od}-C"))
    # cut pipe
    row("HABSFG", f"Cut pipe — {od}-{wall}-{height}{ptype}-C",
        find_contains(inv, "HABSFG", od, wall, f"{height}{ptype}", "-C"))

    # ── STAGE 6: HARM raw material ────────────
    # flights
    for fl in flights:
        row("HARM", f"Flight raw — Flight{od}-{fl}\"",
            find_contains(inv, "HARM", f"Flight{od}", f"-{fl}"))
    # pipe
    row("HARM", f"Pipe raw — Pipe {od}",
        find_contains(inv, "HARM", f"Pipe", od))

    return stages, None

# ─────────────────────────────────────────────
# STEP 5: HTML template
# ─────────────────────────────────────────────

TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Helical Anchor Inventory Trail</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: system-ui, sans-serif; background: #f5f5f0; color: #1a1a1a; }
  .container { max-width: 960px; margin: 0 auto; padding: 2rem 1rem; }
  h1 { font-size: 1.25rem; font-weight: 600; margin-bottom: 1.5rem; }

  /* Upload form */
  .form-card { background: white; border: 1px solid #e0e0d8; border-radius: 10px;
               padding: 1.5rem; margin-bottom: 2rem; }
  .form-row { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; }
  label { font-size: 13px; color: #555; display: block; margin-bottom: 4px; }
  input[type=file] { font-size: 13px; border: 1px solid #ccc; border-radius: 6px;
                     padding: 6px 10px; background: #fafaf8; }
  input[type=text] { font-size: 13px; border: 1px solid #ccc; border-radius: 6px;
                     padding: 7px 10px; width: 260px; }
  button { background: #1a1a1a; color: white; border: none; border-radius: 6px;
           padding: 8px 20px; font-size: 13px; cursor: pointer; }
  button:hover { background: #333; }

  /* Item result block */
  .item-block { background: white; border: 1px solid #e0e0d8; border-radius: 10px;
                padding: 1.5rem; margin-bottom: 1.5rem; }
  .item-header { display: flex; align-items: baseline; gap: 10px; padding-bottom: 10px;
                 border-bottom: 1px solid #eee; margin-bottom: 1rem; flex-wrap: wrap; }
  .item-code { font-size: 15px; font-weight: 600; }
  .item-desc-main { font-size: 13px; color: #666; }
  .badge { font-size: 11px; font-weight: 600; padding: 2px 9px; border-radius: 20px; }
  .badge-lead { background: #dbeafe; color: #1e40af; }
  .badge-ext  { background: #fef3c7; color: #92400e; }

  /* Stage sections */
  .section-title { font-size: 11px; font-weight: 600; color: #888; letter-spacing: .06em;
                   text-transform: uppercase; margin: 1rem 0 6px; }
  .stage-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 8px; }
  .stage-card { border: 1px solid #e8e8e0; border-radius: 8px; padding: 10px 12px; }
  .stage-card.ok   { border-left: 3px solid #16a34a; }
  .stage-card.zero { border-left: 3px solid #dc2626; }
  .stage-card.na   { border-left: 3px solid #9ca3af; }
  .s-label { font-size: 11px; color: #888; margin-bottom: 2px; }
  .s-code  { font-size: 10px; color: #aaa; font-family: monospace; margin-bottom: 3px; }
  .s-desc  { font-size: 12px; color: #222; margin-bottom: 5px; line-height: 1.4; }
  .s-qty   { font-size: 15px; font-weight: 600; }
  .s-qty.ok   { color: #16a34a; }
  .s-qty.zero { color: #dc2626; }
  .s-qty.na   { color: #9ca3af; }
  .s-note  { font-size: 11px; color: #999; margin-top: 3px; font-style: italic; }

  .form-section-title { font-size: 13px; font-weight: 600; color: #333;
                         margin-bottom: 12px; display: flex; align-items: center; gap: 10px; }
  .file-loaded { font-size: 12px; font-weight: 500; color: #16a34a;
                 background: #dcfce7; padding: 2px 10px; border-radius: 20px; }
  .file-none   { font-size: 12px; color: #999; }
  .btn-upload  { background: #1d4ed8; color: white; border: none; border-radius: 6px;
                 padding: 8px 20px; font-size: 13px; cursor: pointer; }
  .btn-upload:hover { background: #1e40af; }
  .error-msg { background: #fef2f2; border: 1px solid #fca5a5; border-radius: 8px;
               padding: 1rem; color: #991b1b; font-size: 13px; margin-bottom: 1rem; }
  .info { font-size: 13px; color: #555; margin-bottom: 1rem; }
</style>
</head>
<body>
<div class="container">
  <h1>Helical Anchor — Inventory Trail Lookup</h1>

  <!-- SECTION 1: Upload file (do this once) -->
  <div class="form-card">
    <div class="form-section-title">
      Step 1 — Upload Excel File
      {% if uploaded_name %}
        <span class="file-loaded">✓ Loaded: {{ uploaded_name }}</span>
      {% else %}
        <span class="file-none">No file loaded</span>
      {% endif %}
    </div>
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <div class="form-row">
        <div>
          <label>Select Excel file (.xlsx)</label>
          <input type="file" name="excel" accept=".xlsx" required>
        </div>
        <div style="padding-top:18px">
          <button type="submit" class="btn-upload">Upload File</button>
        </div>
      </div>
    </form>
  </div>

  <!-- SECTION 2: Search (reuses uploaded file) -->
  <div class="form-card">
    <div class="form-section-title">Step 2 — Search Item Codes</div>
    <form method="POST" action="/search">
      <input type="hidden" name="inv_data" value="{{ inv_data or '' }}">
      <div class="form-row">
        <div style="flex:1">
          <label>Item codes (comma separated)</label>
          <input type="text" name="codes"
                 placeholder="HAFG00016, HAFG00020, HAFG00022"
                 value="{{ codes_input or '' }}"
                 style="width:100%">
        </div>
        <div style="padding-top:18px">
          <button type="submit">Search</button>
        </div>
      </div>
    </form>
  </div>

  {% if error %}
    <div class="error-msg">{{ error }}</div>
  {% endif %}

  {% for item in results %}
    <div class="item-block">
      <div class="item-header">
        <span class="item-code">{{ item.code }}</span>
        <span class="item-desc-main">{{ item.desc }}</span>
        <span class="badge {{ 'badge-lead' if item.is_lead else 'badge-ext' }}">
          {{ 'LEAD' if item.is_lead else 'EXTENSION' }}
        </span>
        {% if item.flights %}
          <span style="font-size:12px;color:#888;">Flights: {{ item.flights | join(', ') }}"</span>
        {% else %}
          <span style="font-size:12px;color:#aaa;">No flights</span>
        {% endif %}
      </div>

      {% if item.error %}
        <div class="error-msg">{{ item.error }}</div>
      {% else %}
        {% set ns = namespace(cur_stage='') %}
        {% for s in item.stages %}
          {% if s.stage != ns.cur_stage %}
            {% set ns.cur_stage = s.stage %}
            <div class="section-title">{{ s.stage }}</div>
            <div class="stage-grid">
          {% endif %}

          {% if s.found == true %}
            {% set card_class = 'ok' if s.qty and s.qty != '—' and s.qty|float > 0 else 'zero' %}
          {% elif s.found == false %}
            {% set card_class = 'zero' %}
          {% else %}
            {% set card_class = 'na' %}
          {% endif %}

          <div class="stage-card {{ card_class }}">
            <div class="s-label">{{ s.label }}</div>
            <div class="s-code">{{ s.code }}</div>
            <div class="s-desc">{{ s.desc }}</div>
            <div class="s-qty {{ card_class }}">
              {% if s.qty == '—' %}—
              {% elif s.found %}{{ s.qty }} nos
              {% else %}0 nos{% endif %}
            </div>
            {% if s.note %}<div class="s-note">{{ s.note }}</div>{% endif %}
          </div>

          {# Close grid when stage changes or last item #}
          {% if loop.last or item.stages[loop.index].stage != ns.cur_stage %}
            </div>
          {% endif %}
        {% endfor %}
      {% endif %}
    </div>
  {% endfor %}

</div>
</body>
</html>
"""

# ─────────────────────────────────────────────
# STEP 6: Flask routes
# ─────────────────────────────────────────────

@app.route("/upload", methods=["POST"])
def upload():
    """Parse the uploaded Excel file and hand the inventory back to the
    browser as a hidden field, since serverless functions can't rely on
    disk or memory persisting between requests."""
    f = request.files.get("excel")
    if not f or f.filename == "":
        return render_template_string(TEMPLATE,
            results=[], error="Please select an Excel file.",
            codes_input="", uploaded_name=None, inv_data="")
    try:
        # Save to a per-request temp path just long enough to parse it
        tmp_dir = tempfile.gettempdir()
        save_path = os.path.join(tmp_dir, f"ha_inv_{os.getpid()}.xlsx")
        f.save(save_path)
        inv = load_inventory(save_path)
        os.remove(save_path)
    except Exception as e:
        return render_template_string(TEMPLATE,
            results=[], error=f"Could not read file: {e}",
            codes_input="", uploaded_name=None, inv_data="")

    return render_template_string(TEMPLATE,
        results=[], error=None, codes_input="",
        uploaded_name=f.filename, inv_data=encode_inv(inv))


@app.route("/", methods=["GET"])
@app.route("/search", methods=["POST"])
def index():
    """Search — uses the inventory data posted back from the upload step."""
    results = []
    error   = None
    codes_input = ""
    inv_data = request.form.get("inv_data", "")
    uploaded_name = "Uploaded file" if inv_data else None

    if request.method == "POST":
        codes_input = request.form.get("codes", "").strip()
        if not inv_data:
            error = "No file uploaded yet. Please upload your Excel file first."
        elif not codes_input:
            error = "Please enter at least one item code."
        else:
            try:
                inv = decode_inv(inv_data)
                codes = [c.strip().upper() for c in codes_input.split(",") if c.strip()]
                for code in codes:
                    if code not in inv:
                        results.append({
                            "code": code, "desc": "—",
                            "is_lead": False, "flights": [],
                            "stages": [], "error": f"{code} not found in the uploaded file."
                        })
                        continue
                    desc = inv[code]["desc"]
                    p    = parse_desc(desc)
                    stages, err = build_trail(code, inv)
                    results.append({
                        "code":    code,
                        "desc":    desc,
                        "is_lead": p["ptype"] == "L" if p else False,
                        "flights": p["flights"] if p else [],
                        "stages":  stages or [],
                        "error":   err
                    })
            except Exception as e:
                error = f"Error processing file: {str(e)}"

    return render_template_string(TEMPLATE,
        results=results, error=error,
        codes_input=codes_input, uploaded_name=uploaded_name,
        inv_data=inv_data)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
