#!/usr/bin/env python3
"""Generate Kalkulator_Premi_MiUHC.html - a masterpiece premium calculator."""
import json
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from generate_excel import extract_syariah_data, extract_konvensional_data

OUTPUT_HTML = os.path.join(SCRIPT_DIR, 'Kalkulator_Premi_MiUHC.html')

# Column order (index 0-12):
# Diamond, Ruby, Emerald, Topaz, Topaz ID, Jade, Jade ID, Sapphire,
# Diamond Smart, Ruby Smart, Emerald Smart, Topaz Smart, Jade Smart
PLAN_NAMES = [
    "Diamond", "Ruby", "Emerald", "Topaz", "Topaz ID",
    "Jade", "Jade ID", "Sapphire", "Diamond Smart",
    "Ruby Smart", "Emerald Smart", "Topaz Smart", "Jade Smart"
]

PLAN_DETAILS = {
    "Diamond": {"wilayah": "Worldwide", "kamar": "1 Bed", "limit": "Unlimited", "deductible": None},
    "Ruby": {"wilayah": "Worldwide (kec. USA)", "kamar": "1 Bed", "limit": "90 M", "deductible": None},
    "Emerald": {"wilayah": "Seluruh Asia", "kamar": "1 Bed", "limit": "35 M", "deductible": None},
    "Topaz": {"wilayah": "Seluruh Asia (kec. JPN/HKG/SGP)", "kamar": "1 Bed", "limit": "18 M", "deductible": None},
    "Topaz ID": {"wilayah": "Indonesia", "kamar": "1 Bed", "limit": "18 M", "deductible": None},
    "Jade": {"wilayah": "Seluruh Asia (kec. JPN/HKG/SGP)", "kamar": "2 Bed", "limit": "10 M", "deductible": None},
    "Jade ID": {"wilayah": "Indonesia", "kamar": "2 Bed", "limit": "10 M", "deductible": None},
    "Sapphire": {"wilayah": "Indonesia", "kamar": "2 Bed (No RI Charges)", "limit": "5 M", "deductible": None},
    "Diamond Smart": {"wilayah": "Worldwide", "kamar": "1 Bed", "limit": "Unlimited", "deductible": "Rp16.000.000"},
    "Ruby Smart": {"wilayah": "Worldwide (kec. USA)", "kamar": "1 Bed", "limit": "90 M", "deductible": "Rp16.000.000"},
    "Emerald Smart": {"wilayah": "Seluruh Asia", "kamar": "1 Bed", "limit": "35 M", "deductible": "Rp8.000.000"},
    "Topaz Smart": {"wilayah": "Seluruh Asia (kec. JPN/HKG/SGP)", "kamar": "1 Bed", "limit": "18 M", "deductible": "Rp8.000.000"},
    "Jade Smart": {"wilayah": "Seluruh Asia (kec. JPN/HKG/SGP)", "kamar": "2 Bed", "limit": "10 M", "deductible": "Rp5.000.000"},
}


def build_html(konv_data, syariah_data):
    """Build the complete HTML string."""
    # Convert data to JSON-serializable format (keys must be strings)
    def convert_data(data):
        result = {}
        for gender in ['PRIA', 'WANITA']:
            result[gender] = {}
            for age in range(86):
                result[gender][str(age)] = data[gender][age]
        return result

    konv_json = json.dumps(convert_data(konv_data), separators=(',', ':'))
    syariah_json = json.dumps(convert_data(syariah_data), separators=(',', ':'))
    plan_details_json = json.dumps(PLAN_DETAILS, separators=(',', ':'))
    plan_names_json = json.dumps(PLAN_NAMES, separators=(',', ':'))

    # Build age options
    age_options = ""
    for age in range(86):
        selected = ' selected' if age == 30 else ''
        age_options += f'<option value="{age}"{selected}>{age} tahun</option>\n'

    # Use template with placeholders to avoid f-string/brace conflicts
    html = HTML_TEMPLATE
    html = html.replace('__DATA_KONV__', konv_json)
    html = html.replace('__DATA_SYARIAH__', syariah_json)
    html = html.replace('__PLAN_NAMES__', plan_names_json)
    html = html.replace('__PLAN_DETAILS__', plan_details_json)
    html = html.replace('__AGE_OPTIONS__', age_options)

    return html


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Kalkulator Premi MiUHC - Manulife Indonesia</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
--primary:#1E7145;
--secondary:#27AE60;
--accent:#2ECC71;
--bg:#f0f7f3;
--card-bg:#ffffff;
--text:#2c3e50;
--text-light:#7f8c8d;
--shadow:0 4px 15px rgba(30,113,69,0.12);
--shadow-hover:0 8px 25px rgba(30,113,69,0.2);
--radius:12px;
--transition:all 0.3s cubic-bezier(0.4,0,0.2,1);
}
body{
font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
background:var(--bg);
color:var(--text);
line-height:1.6;
min-height:100vh;
}
.container{
max-width:1200px;
margin:0 auto;
padding:20px;
}
header{
background:linear-gradient(135deg,var(--primary) 0%,var(--secondary) 100%);
color:#fff;
padding:30px 20px;
text-align:center;
border-radius:0 0 var(--radius) var(--radius);
margin-bottom:30px;
box-shadow:var(--shadow);
}
header h1{
font-size:clamp(1.4rem,4vw,2.2rem);
font-weight:700;
margin-bottom:5px;
letter-spacing:-0.5px;
}
header h2{
font-size:clamp(0.9rem,2.5vw,1.2rem);
font-weight:400;
opacity:0.9;
}
header .subtitle{
font-size:0.85rem;
opacity:0.75;
margin-top:5px;
}
.controls{
background:var(--card-bg);
border-radius:var(--radius);
padding:20px 25px;
box-shadow:var(--shadow);
margin-bottom:25px;
display:flex;
flex-wrap:wrap;
gap:15px;
align-items:center;
}
.control-group{
display:flex;
flex-direction:column;
gap:5px;
}
.control-group label{
font-size:0.8rem;
font-weight:600;
color:var(--primary);
text-transform:uppercase;
letter-spacing:0.5px;
}
.control-group select{
padding:10px 35px 10px 12px;
border:2px solid #e0e0e0;
border-radius:8px;
font-size:0.95rem;
font-family:inherit;
background:#fff;
cursor:pointer;
transition:var(--transition);
appearance:none;
-webkit-appearance:none;
background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%231E7145' d='M6 8L1 3h10z'/%3E%3C/svg%3E");
background-repeat:no-repeat;
background-position:right 12px center;
}
.control-group select:focus{
outline:none;
border-color:var(--primary);
box-shadow:0 0 0 3px rgba(30,113,69,0.15);
}
.tabs{
display:flex;
gap:0;
margin-bottom:25px;
background:var(--card-bg);
border-radius:var(--radius);
padding:5px;
box-shadow:var(--shadow);
}
.tab-btn{
flex:1;
padding:12px 20px;
border:none;
background:transparent;
font-size:0.95rem;
font-weight:600;
font-family:inherit;
color:var(--text-light);
cursor:pointer;
border-radius:8px;
transition:var(--transition);
}
.tab-btn.active{
background:var(--primary);
color:#fff;
box-shadow:0 2px 8px rgba(30,113,69,0.3);
}
.tab-btn:hover:not(.active){
background:rgba(30,113,69,0.08);
color:var(--primary);
}
.section-title{
font-size:clamp(1rem,3vw,1.3rem);
font-weight:700;
color:var(--primary);
margin:25px 0 15px;
padding-left:12px;
border-left:4px solid var(--accent);
}
.section-title .badge{
display:inline-block;
background:var(--accent);
color:#fff;
font-size:0.7rem;
padding:2px 8px;
border-radius:10px;
margin-left:8px;
vertical-align:middle;
font-weight:600;
}
.cards-grid{
display:grid;
grid-template-columns:repeat(auto-fill,minmax(260px,1fr));
gap:18px;
margin-bottom:30px;
}
.plan-card{
background:var(--card-bg);
border-radius:var(--radius);
overflow:hidden;
box-shadow:var(--shadow);
transition:var(--transition);
animation:fadeInUp 0.4s ease forwards;
opacity:0;
}
.plan-card:hover{
transform:translateY(-3px);
box-shadow:var(--shadow-hover);
}
.plan-card .card-header{
background:linear-gradient(135deg,var(--primary),var(--secondary));
color:#fff;
padding:14px 16px;
font-weight:700;
font-size:1rem;
}
.plan-card.smart .card-header{
background:linear-gradient(135deg,#F39C12,#E67E22);
}
.plan-card .card-body{
padding:16px;
}
.plan-card .premi-row{
display:flex;
justify-content:space-between;
align-items:center;
padding:8px 0;
border-bottom:1px solid #f0f0f0;
}
.plan-card .premi-row:last-child{
border-bottom:none;
}
.plan-card .premi-label{
font-size:0.8rem;
color:var(--text-light);
font-weight:500;
}
.plan-card .premi-value{
font-size:0.95rem;
font-weight:700;
color:var(--primary);
}
.plan-card .premi-value.monthly{
color:var(--secondary);
font-size:0.88rem;
}
.plan-card .detail-row{
font-size:0.78rem;
color:var(--text-light);
padding:3px 0;
display:flex;
justify-content:space-between;
}
.plan-card .detail-row .detail-val{
font-weight:600;
color:var(--text);
}
.plan-card .deductible-badge{
display:inline-block;
background:#FFF3CD;
color:#856404;
font-size:0.72rem;
padding:3px 8px;
border-radius:6px;
font-weight:600;
margin-top:6px;
}
.info-box{
background:linear-gradient(135deg,#E8F5EE,#D5F5E3);
border-radius:var(--radius);
padding:18px 22px;
margin:20px 0;
border-left:4px solid var(--accent);
}
.info-box h4{
color:var(--primary);
font-size:0.9rem;
margin-bottom:6px;
}
.info-box p{
font-size:0.82rem;
color:var(--text);
margin:4px 0;
}
.expand-btn{
display:block;
width:100%;
padding:14px;
border:2px dashed var(--secondary);
border-radius:var(--radius);
background:transparent;
color:var(--primary);
font-size:0.95rem;
font-weight:600;
font-family:inherit;
cursor:pointer;
transition:var(--transition);
margin:20px 0;
}
.expand-btn:hover{
background:rgba(30,113,69,0.05);
border-color:var(--primary);
}
.all-plans{
display:none;
}
.all-plans.show{
display:block;
animation:fadeInUp 0.4s ease;
}
footer{
background:var(--card-bg);
border-radius:var(--radius);
padding:20px 25px;
margin-top:30px;
box-shadow:var(--shadow);
font-size:0.78rem;
color:var(--text-light);
line-height:1.8;
}
footer strong{
color:var(--text);
}
.whatsapp-fab{
position:fixed;
bottom:25px;
right:25px;
width:56px;
height:56px;
background:#25D366;
border-radius:50%;
display:flex;
align-items:center;
justify-content:center;
box-shadow:0 4px 15px rgba(37,211,102,0.4);
cursor:pointer;
transition:var(--transition);
z-index:1000;
text-decoration:none;
}
.whatsapp-fab:hover{
transform:scale(1.1);
box-shadow:0 6px 20px rgba(37,211,102,0.6);
}
.whatsapp-fab svg{
width:28px;
height:28px;
fill:#fff;
}
.whatsapp-tooltip{
position:fixed;
bottom:90px;
right:25px;
background:#fff;
color:var(--text);
padding:8px 14px;
border-radius:8px;
font-size:0.75rem;
font-weight:600;
box-shadow:0 2px 10px rgba(0,0,0,0.15);
z-index:1000;
white-space:nowrap;
opacity:0;
transform:translateY(5px);
transition:var(--transition);
pointer-events:none;
}
.whatsapp-fab:hover+.whatsapp-tooltip{
opacity:1;
transform:translateY(0);
}
@keyframes fadeInUp{
from{opacity:0;transform:translateY(15px)}
to{opacity:1;transform:translateY(0)}
}
@media(max-width:600px){
.container{padding:10px}
header{padding:20px 15px;margin-bottom:20px}
.controls{padding:15px;gap:10px}
.cards-grid{grid-template-columns:1fr;gap:12px}
.tabs{flex-direction:column;gap:5px}
.tab-btn{padding:10px}
}
@media print{
.whatsapp-fab,.whatsapp-tooltip{display:none!important}
body{background:#fff}
.plan-card{box-shadow:none;border:1px solid #ddd;break-inside:avoid}
header{box-shadow:none;border-radius:0}
.controls,.tabs{box-shadow:none;border:1px solid #eee}
.expand-btn{display:none}
.all-plans{display:block!important}
footer{box-shadow:none;border-top:1px solid #eee;border-radius:0}
}
</style>
</head>
<body>
<div class="container">
<header>
<h1>Mi Ultimate Healthcare (MiUHC)</h1>
<h2>Kalkulator Premi</h2>
<p class="subtitle">PT Asuransi Jiwa Manulife Indonesia</p>
</header>

<div class="tabs">
<button class="tab-btn active" onclick="switchTab('konvensional')">Konvensional</button>
<button class="tab-btn" onclick="switchTab('syariah')">Syariah</button>
</div>

<div class="controls">
<div class="control-group">
<label>Jenis Kelamin</label>
<select id="gender" onchange="calculate()">
<option value="PRIA">Pria</option>
<option value="WANITA">Wanita</option>
</select>
</div>
<div class="control-group">
<label>Usia</label>
<select id="age" onchange="calculate()">
__AGE_OPTIONS__</select>
</div>
</div>

<div id="results"></div>

<div class="info-box">
<h4>No-Claim BONUS</h4>
<p>Setahun tanpa klaim? LIMIT tahunan kartu Anda naik 10% per tahun hingga 50%</p>
<h4 style="margin-top:10px">No-Claim DISCOUNT</h4>
<p>Setahun tanpa klaim? Dapatkan DISCOUNT 10% di tahun berikutnya, hingga 15%</p>
</div>

<footer>
<strong>Catatan:</strong><br>
*Deductible = biaya yang ditanggung nasabah sendiri sebelum manfaat asuransi berlaku<br>
Premi Tahunan = RS Premi x 1.12 (dibulatkan ke ratusan terdekat)<br>
Premi Bulanan = Premi Tahunan x 1.14 / 12 (dibulatkan ke ratusan terdekat)<br>
Asuransi Jiwa: 30 Juta | Rawat Inap: Sesuai Tagihan<br>
<br>
<strong>PETRUS JAKUB MANULIFE 087781896087</strong>
</footer>
</div>

<a href="https://wa.me/6287781896087" target="_blank" class="whatsapp-fab" aria-label="WhatsApp">
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
</a>
<div class="whatsapp-tooltip">PETRUS JAKUB MANULIFE 087781896087</div>

<script>
(function(){
const DATA_KONV = __DATA_KONV__;
const DATA_SYARIAH = __DATA_SYARIAH__;
const PLAN_NAMES = __PLAN_NAMES__;
const PLAN_DETAILS = __PLAN_DETAILS__;

let currentTab = 'konvensional';

window.switchTab = function(tab) {
    currentTab = tab;
    document.querySelectorAll('.tab-btn').forEach(function(btn, i) {
        btn.classList.toggle('active', (i === 0 && tab === 'konvensional') || (i === 1 && tab === 'syariah'));
    });
    calculate();
};

function formatRp(n) {
    if (n === 0) return 'Rp0';
    return 'Rp' + n.toLocaleString('id-ID');
}

function calcPremi(rs) {
    var tahunan = Math.round(rs * 1.12 / 100) * 100;
    var bulanan = Math.round(tahunan * 1.14 / 12 / 100) * 100;
    return { tahunan: tahunan, bulanan: bulanan };
}

window.toggleAllPlans = function() {
    var section = document.getElementById('all-plans-section');
    var btn = document.getElementById('expand-btn');
    if (section.classList.contains('show')) {
        section.classList.remove('show');
        btn.textContent = 'Lihat Semua Plan (13 Plan)';
    } else {
        section.classList.add('show');
        btn.textContent = 'Sembunyikan';
    }
};

function buildCard(planName, rs, isSmart) {
    var premi = calcPremi(rs);
    var details = PLAN_DETAILS[planName];
    var deductibleHtml = '';
    if (details && details.deductible) {
        deductibleHtml = '<span class="deductible-badge">Deductible: ' + details.deductible + '</span>';
    }
    var wilayah = details ? details.wilayah : '-';
    var kamar = details ? details.kamar : '-';
    var limit = details ? details.limit : '-';
    return '<div class="plan-card' + (isSmart ? ' smart' : '') + '">' +
        '<div class="card-header">' + planName + '</div>' +
        '<div class="card-body">' +
            '<div class="premi-row"><span class="premi-label">Premi Tahunan</span><span class="premi-value">' + formatRp(premi.tahunan) + '</span></div>' +
            '<div class="premi-row"><span class="premi-label">Premi Bulanan</span><span class="premi-value monthly">' + formatRp(premi.bulanan) + '</span></div>' +
            '<div class="detail-row"><span>Wilayah</span><span class="detail-val">' + wilayah + '</span></div>' +
            '<div class="detail-row"><span>Kamar</span><span class="detail-val">' + kamar + '</span></div>' +
            '<div class="detail-row"><span>Limit Kartu</span><span class="detail-val">' + limit + '</span></div>' +
            '<div class="detail-row"><span>Asuransi Jiwa</span><span class="detail-val">30 Juta</span></div>' +
            '<div class="detail-row"><span>Rawat Inap</span><span class="detail-val">Sesuai Tagihan</span></div>' +
            deductibleHtml +
        '</div>' +
    '</div>';
}

window.calculate = function() {
    var data = currentTab === 'konvensional' ? DATA_KONV : DATA_SYARIAH;
    var gender = document.getElementById('gender').value;
    var age = document.getElementById('age').value;
    var row = data[gender][age];
    if (!row) return;

    // Most Wanted: Jade ID (index 6), Jade (5), Topaz ID (4), Topaz (3)
    var mostWanted = [
        { name: 'Jade ID', idx: 6 },
        { name: 'Jade', idx: 5 },
        { name: 'Topaz ID', idx: 4 },
        { name: 'Topaz', idx: 3 }
    ];

    // HEMAT 40%: Jade Smart (12), Topaz Smart (11)
    var hemat = [
        { name: 'Jade Smart', idx: 12 },
        { name: 'Topaz Smart', idx: 11 }
    ];

    var html = '';

    // Most Wanted section
    html += '<div class="section-title">Most Wanted <span class="badge">No Deductible</span></div>';
    html += '<div class="cards-grid">';
    for (var i = 0; i < mostWanted.length; i++) {
        html += buildCard(mostWanted[i].name, row[mostWanted[i].idx], false);
    }
    html += '</div>';

    // HEMAT 40% section
    html += '<div class="section-title">HEMAT 40% <span class="badge" style="background:#F39C12">Smart Plan</span></div>';
    html += '<div class="cards-grid">';
    for (var j = 0; j < hemat.length; j++) {
        html += buildCard(hemat[j].name, row[hemat[j].idx], true);
    }
    html += '</div>';

    // Expand button
    html += '<button class="expand-btn" id="expand-btn" onclick="toggleAllPlans()">Lihat Semua Plan (13 Plan)</button>';

    // All plans section
    html += '<div class="all-plans" id="all-plans-section">';
    html += '<div class="section-title">Semua Plan</div>';
    html += '<div class="cards-grid">';
    for (var k = 0; k < 13; k++) {
        var name = PLAN_NAMES[k];
        var isSmart = name.toLowerCase().indexOf('smart') >= 0;
        html += buildCard(name, row[k], isSmart);
    }
    html += '</div></div>';

    document.getElementById('results').innerHTML = html;

    // Re-trigger animations
    var cards = document.querySelectorAll('.plan-card');
    for (var c = 0; c < cards.length; c++) {
        cards[c].style.animationDelay = (c * 0.05) + 's';
    }
};

// Initial calculation
calculate();
})();
</script>
</body>
</html>"""


def main():
    print("Extracting premium data from PDFs...")
    konv_data = extract_konvensional_data()
    syariah_data = extract_syariah_data()

    print(f"  Konvensional PRIA age 0: {konv_data['PRIA'][0][:3]}...")
    print(f"  Syariah PRIA age 0: {syariah_data['PRIA'][0][:3]}...")
    print(f"  Konvensional PRIA age 30 Jade ID (idx 6): {konv_data['PRIA'][30][6]}")

    print("Generating HTML...")
    html_content = build_html(konv_data, syariah_data)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html_content)

    file_size = os.path.getsize(OUTPUT_HTML)
    print(f"Output: {OUTPUT_HTML}")
    print(f"  File size: {file_size:,} bytes")
    print("Done!")


if __name__ == "__main__":
    main()
