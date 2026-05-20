# ============================================================
# COMPLETE DEALER DASHBOARD GENERATOR
# ============================================================

import pandas as pd
import json
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = "Common Tracker.xlsx"
OUTPUT_HTML = "dealer_dashboard_final.html"

# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    df.columns = [str(col).strip() for col in df.columns]

    for col in df.columns:

        try:
            df[col] = df[col].apply(
                lambda x:
                    int(x)
                    if isinstance(x, (float, int)) and pd.notna(x)
                    else x
            )
        except:
            pass

    df = df.fillna("")

    return df

# ============================================================
# SUMMARY TABLES
# ============================================================

summary_table1 = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Summary",
    header=0,
    usecols="A:J",
    skiprows=0,
    nrows=8
)

summary_table2 = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Summary",
    header=10,
    usecols="A:J",
    nrows=8
)

summary_table3 = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Summary",
    header=20,
    usecols="A:I",
    nrows=8
)

summary_table1 = clean_dataframe(summary_table1)
summary_table2 = clean_dataframe(summary_table2)
summary_table3 = clean_dataframe(summary_table3)

# ============================================================
# OTHER SHEETS
# ============================================================

pwg_df = pd.read_excel(
    EXCEL_FILE,
    sheet_name="PWG Dealers"
)

xper_df = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Xper Tracker"
)

fwk_df = pd.read_excel(
    EXCEL_FILE,
    sheet_name="FWK463 Tracker"
)

pwg_df = clean_dataframe(pwg_df)
xper_df = clean_dataframe(xper_df)
fwk_df = clean_dataframe(fwk_df)

# ============================================================
# JSON
# ============================================================

summary1_json = json.dumps(
    summary_table1.to_dict(orient="records"),
    default=str
)

summary2_json = json.dumps(
    summary_table2.to_dict(orient="records"),
    default=str
)

summary3_json = json.dumps(
    summary_table3.to_dict(orient="records"),
    default=str
)

pwg_json = json.dumps(
    pwg_df.to_dict(orient="records"),
    default=str
)

xper_json = json.dumps(
    xper_df.to_dict(orient="records"),
    default=str
)

fwk_json = json.dumps(
    fwk_df.to_dict(orient="records"),
    default=str
)

# ============================================================
# HTML
# ============================================================

html = """

<!DOCTYPE html>
<html lang="en">

<head>

<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Dealer Dashboard</title>

<script src="https://cdn.tailwindcss.com"></script>

<link rel="preconnect" href="https://fonts.googleapis.com">

<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap"
rel="stylesheet">

<style>

body{
    background:#edf2f7;
    font-family:'Inter',sans-serif;
}

.hidden{
    display:none;
}

.dashboard-card{
    background:white;
    border-radius:28px;
    box-shadow:0 10px 30px rgba(0,0,0,0.08);
}

table{
    border-collapse:collapse;
    width:100%;
}

th,td{
    border:1px solid #d8dee8;
}

thead th{
    position:sticky;
    top:0;
    z-index:2;
}

tbody tr:hover{
    background:#f5f9ff;
}

.filter{
    transition:0.2s;
}

.filter:focus{
    outline:none;
    border-color:#2563eb;
    box-shadow:0 0 0 4px rgba(37,99,235,0.15);
}

.summary-title{
    font-size:28px;
    font-weight:800;
    padding:18px 24px;
    color:white;
}

</style>

</head>

<body>

<div class="max-w-[1900px] mx-auto p-6">

<!-- ===================================================== -->
<!-- HEADER -->
<!-- ===================================================== -->

<div class="dashboard-card p-8 mb-8">

<div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">

<div>

<h1 class="text-5xl font-black text-gray-800">
Dealer Dashboard
</h1>

<p class="text-gray-500 mt-2 text-lg">
Auto Generated From Excel
</p>

</div>

<div class="flex flex-wrap gap-3">

<button
class="tab-btn bg-blue-600 text-white px-6 py-3 rounded-2xl font-bold"
onclick="showTab('summary',event)">
Summary
</button>

<button
class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold"
onclick="showTab('pwg',event)">
PWG Dealers
</button>

<button
class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold"
onclick="showTab('xper',event)">
Xper Dealers
</button>

<button
class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold"
onclick="showTab('fwk',event)">
FWK463 Dealers
</button>

</div>

</div>

</div>

<!-- ===================================================== -->
<!-- SUMMARY -->
<!-- ===================================================== -->

<div id="summary" class="tab-content">

<div class="dashboard-card overflow-auto mb-8">

<div class="summary-title bg-blue-600">
PWG Sales + PSR Sales
</div>

<table class="text-sm">

<thead id="summary-head-1"></thead>
<tbody id="summary-body-1"></tbody>

</table>

</div>

<div class="dashboard-card overflow-auto mb-8">

<div class="summary-title bg-green-600">
Xper Placement Drive Update - P2
</div>

<table class="text-sm">

<thead id="summary-head-2"></thead>
<tbody id="summary-body-2"></tbody>

</table>

</div>

<div class="dashboard-card overflow-auto mb-8">

<div class="summary-title bg-purple-600">
Fevikwik 463 Drive Update
</div>

<table class="text-sm">

<thead id="summary-head-3"></thead>
<tbody id="summary-body-3"></tbody>

</table>

</div>

</div>

<!-- ===================================================== -->
<!-- PWG -->
<!-- ===================================================== -->

<div id="pwg" class="tab-content hidden"></div>

<!-- ===================================================== -->
<!-- XPER -->
<!-- ===================================================== -->

<div id="xper" class="tab-content hidden"></div>

<!-- ===================================================== -->
<!-- FWK -->
<!-- ===================================================== -->

<div id="fwk" class="tab-content hidden"></div>

</div>

<script>

const summaryData1 = __SUMMARY1_DATA__;
const summaryData2 = __SUMMARY2_DATA__;
const summaryData3 = __SUMMARY3_DATA__;

const pwgData = __PWG_DATA__;
const xperData = __XPER_DATA__;
const fwkData = __FWK_DATA__;

function showTab(tabId,event){

document.querySelectorAll('.tab-content').forEach(tab=>{
tab.classList.add('hidden');
});

document.getElementById(tabId).classList.remove('hidden');

document.querySelectorAll('.tab-btn').forEach(btn=>{
btn.classList.remove('bg-blue-600','text-white');
btn.classList.add('bg-gray-200');
});

event.target.classList.remove('bg-gray-200');
event.target.classList.add('bg-blue-600','text-white');

}

// =======================================================
// SUMMARY TABLES
// =======================================================

function renderSummary(data, headId, bodyId){

const head=document.getElementById(headId);
const body=document.getElementById(bodyId);

head.innerHTML='';
body.innerHTML='';

if(!data || data.length===0)
return;

const columns = Object.keys(data[0]);

const cleanedHeaders = [];

columns.forEach(col=>{

if(col.includes('Unnamed')){

cleanedHeaders.push('');

}else{

cleanedHeaders.push(col);

}

});

head.innerHTML = `

<tr class="bg-gray-100">

${cleanedHeaders.map(col=>`

<th class="p-3 font-bold text-gray-800">

${col}

</th>

`).join('')}

</tr>

`;

data.forEach((row,index)=>{

body.innerHTML += `

<tr class="${index % 2 === 0 ? '' : 'bg-gray-50'}">

${columns.map(col=>`

<td class="p-2">

${row[col] ?? ''}

</td>

`).join('')}

</tr>

`;

});

}

// =======================================================
// FILTER COLUMNS
// =======================================================

function getFilterColumns(columns){

return columns.filter(c=>

c.toLowerCase().includes('tty') ||
c.toLowerCase().includes('wss') ||
c.toLowerCase().includes('ssdm') ||
c.toLowerCase().includes('day') ||
c.toLowerCase().includes('cluster') ||
c.toLowerCase().includes('as name')

);

}

// =======================================================
// CALCULATE SUBTOTALS
// =======================================================

function calculateSubtotals(sectionId){

const section=document.getElementById(sectionId);

const totalBox=section.querySelector('.subtotal-total');

const recordBox=section.querySelector('.subtotal-records');

const rows=section.querySelectorAll('tbody tr');

let total=0;
let records=0;

rows.forEach(row=>{

const cells=row.querySelectorAll('td');

if(cells.length>0){

records++;

cells.forEach(cell=>{

const value=parseInt(
String(cell.innerText).replace(/,/g,'')
);

if(!isNaN(value))
total += value;

});

}

});

if(totalBox)
totalBox.innerText=total.toLocaleString();

if(recordBox)
recordBox.innerText=records;

}

// =======================================================
// CREATE TABLE
// =======================================================

function createDealerTable(containerId,data,color='blue'){

const container=document.getElementById(containerId);

if(data.length===0){

container.innerHTML=`

<div class="dashboard-card p-10 text-center">
No Data Found
</div>

`;

return;
}

const columns=Object.keys(data[0]);

const filterColumns=getFilterColumns(columns);

container.innerHTML = `

<div class="dashboard-card p-6 mb-8">

<div class="flex items-center justify-between mb-6">

<h2 class="text-3xl font-black text-gray-800">
${containerId.toUpperCase()} TRACKER
</h2>

<div class="bg-blue-100 text-blue-700 px-4 py-2 rounded-2xl font-bold">
${data.length} Dealers
</div>

</div>

<!-- FILTERS -->

<div class="grid grid-cols-1 md:grid-cols-5 gap-4 mb-6">

${filterColumns.map(col=>`

<select
class="filter border border-gray-300 rounded-2xl p-3"
data-column="${col}"
data-table="${containerId}">

<option value="">
Filter ${col}
</option>

${[...new Set(data.map(r=>r[col]))]
.filter(v=>v!=='')
.map(v=>`
<option value="${v}">
${v}
</option>
`).join('')
}

</select>

`).join('')}

</div>

<!-- SUBTOTALS -->

<div class="grid grid-cols-1 md:grid-cols-4 gap-5 mb-8">

<div class="bg-blue-50 rounded-2xl p-5 border border-blue-200">
<div class="text-sm text-gray-500">Records</div>
<div class="text-3xl font-black text-blue-700 subtotal-records">
${data.length}
</div>
</div>

<div class="bg-green-50 rounded-2xl p-5 border border-green-200">
<div class="text-sm text-gray-500">Numeric Total</div>
<div class="text-3xl font-black text-green-700 subtotal-total">
0
</div>
</div>

</div>

<!-- TABLE -->

<div class="overflow-auto rounded-3xl border border-gray-200">

<table class="w-full text-sm whitespace-nowrap">

<thead class="bg-${color}-600 text-white">

<tr>

${columns.map(col=>`
<th class="p-3">
${col}
</th>
`).join('')}

</tr>

</thead>

<tbody id="${containerId}-body">

${data.map((row,index)=>`

<tr class="${index % 2 === 0 ? '' : 'bg-gray-50'}">

${columns.map(col=>`

<td class="p-2">

${row[col] ?? ''}

</td>

`).join('')}

</tr>

`).join('')}

</tbody>

</table>

</div>

</div>

`;

setupFilters(containerId,data);

calculateSubtotals(containerId);

}

// =======================================================
// FILTERS
// =======================================================

function setupFilters(sectionId,originalData){

const section=document.getElementById(sectionId);

const filters=section.querySelectorAll('.filter');

filters.forEach(filter=>{

filter.addEventListener('change',()=>{

let filtered=[...originalData];

filters.forEach(f=>{

const column=f.dataset.column;
const value=f.value;

if(value!==''){

filtered=filtered.filter(row=>

String(row[column]).trim()===value

);

}

});

renderFiltered(sectionId,filtered);

});

});

}

// =======================================================
// RENDER FILTERED TABLE
// =======================================================

function renderFiltered(sectionId,data){

const tbody=document.getElementById(`${sectionId}-body`);

if(data.length===0){

tbody.innerHTML=`

<tr>
<td colspan="100%" class="p-6 text-center">
No Records Found
</td>
</tr>

`;

return;
}

const columns=Object.keys(data[0]);

tbody.innerHTML=data.map((row,index)=>`

<tr class="${index % 2 === 0 ? '' : 'bg-gray-50'}">

${columns.map(col=>`

<td class="p-2">

${row[col] ?? ''}

</td>

`).join('')}

</tr>

`).join('');

calculateSubtotals(sectionId);

}

// =======================================================
// LOAD
// =======================================================

renderSummary(summaryData1,'summary-head-1','summary-body-1');
renderSummary(summaryData2,'summary-head-2','summary-body-2');
renderSummary(summaryData3,'summary-head-3','summary-body-3');

createDealerTable('pwg',pwgData,'blue');
createDealerTable('xper',xperData,'green');
createDealerTable('fwk',fwkData,'purple');

</script>

</body>
</html>

"""

# ============================================================
# REPLACE JSON PLACEHOLDERS
# ============================================================

html = html.replace(
    "__SUMMARY1_DATA__",
    summary1_json
)

html = html.replace(
    "__SUMMARY2_DATA__",
    summary2_json
)

html = html.replace(
    "__SUMMARY3_DATA__",
    summary3_json
)

html = html.replace(
    "__PWG_DATA__",
    pwg_json
)

html = html.replace(
    "__XPER_DATA__",
    xper_json
)

html = html.replace(
    "__FWK_DATA__",
    fwk_json
)

# ============================================================
# SAVE HTML
# ============================================================

Path(OUTPUT_HTML).write_text(
    html,
    encoding="utf-8"
)

print(f"Dashboard Created Successfully: {OUTPUT_HTML}")
