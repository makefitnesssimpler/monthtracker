# ============================================================
# COMPLETE DEALER DASHBOARD GENERATOR
# (Original logic preserved + file picker wrapper)
# ============================================================

import json
import warnings
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pandas as pd

warnings.filterwarnings("ignore")

OUTPUT_HTML = "index.html"


def clean_dataframe(df):
    df.columns = [str(col).strip() for col in df.columns]
    for col in df.columns:
        try:
            df[col] = df[col].apply(
                lambda x: int(x) if isinstance(x, (float, int)) and pd.notna(x) else x
            )
        except Exception:
            pass
    return df.fillna("")


def build_dashboard(excel_file):
    summary_table1 = pd.read_excel(excel_file, sheet_name="Summary", header=0, usecols="A:J", skiprows=0, nrows=8)
    summary_table2 = pd.read_excel(excel_file, sheet_name="Summary", header=10, usecols="A:J", nrows=8)
    summary_table3 = pd.read_excel(excel_file, sheet_name="Summary", header=20, usecols="A:I", nrows=8)
    pwg_df = pd.read_excel(excel_file, sheet_name="PWG Dealers")
    xper_df = pd.read_excel(excel_file, sheet_name="Xper Tracker")
    fwk_df = pd.read_excel(excel_file, sheet_name="FWK463 Tracker")

    summary_table1 = clean_dataframe(summary_table1)
    summary_table2 = clean_dataframe(summary_table2)
    summary_table3 = clean_dataframe(summary_table3)
    pwg_df = clean_dataframe(pwg_df)
    xper_df = clean_dataframe(xper_df)
    fwk_df = clean_dataframe(fwk_df)

    summary1_json = json.dumps(summary_table1.to_dict(orient="records"), default=str)
    summary2_json = json.dumps(summary_table2.to_dict(orient="records"), default=str)
    summary3_json = json.dumps(summary_table3.to_dict(orient="records"), default=str)
    pwg_json = json.dumps(pwg_df.to_dict(orient="records"), default=str)
    xper_json = json.dumps(xper_df.to_dict(orient="records"), default=str)
    fwk_json = json.dumps(fwk_df.to_dict(orient="records"), default=str)

    html = """
<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dealer Dashboard</title><script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
body{background:#edf2f7;font-family:'Inter',sans-serif}.hidden{display:none}.dashboard-card{background:white;border-radius:20px;box-shadow:0 8px 24px rgba(15,23,42,0.08);border:1px solid #e5e7eb}
table{border-collapse:collapse;width:100%}th,td{border:1px solid #e5e7eb}thead th{position:sticky;top:0;z-index:2}tbody tr:hover{background:#f8fbff}
.filter{transition:0.2s}.filter:focus{outline:none;border-color:#2563eb;box-shadow:0 0 0 4px rgba(37,99,235,0.15)}.summary-title{font-size:24px;font-weight:800;padding:14px 20px;color:white}.sortable{cursor:pointer;user-select:none}
</style></head><body><div class="max-w-[1900px] mx-auto p-6">
<div class="dashboard-card p-8 mb-8"><div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6"><div><h1 class="text-5xl font-black text-gray-800">Dealer Dashboard</h1><p class="text-gray-500 mt-2 text-lg">Auto Generated From Excel</p></div><div class="flex flex-wrap gap-3">
<button class="tab-btn bg-blue-600 text-white px-6 py-3 rounded-2xl font-bold" onclick="showTab('summary',event)">Summary</button>
<button class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold" onclick="showTab('pwg',event)">PWG Dealers</button>
<button class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold" onclick="showTab('xper',event)">Xper Dealers</button>
<button class="tab-btn bg-gray-200 px-6 py-3 rounded-2xl font-bold" onclick="showTab('fwk',event)">FWK463 Dealers</button></div></div></div>
<div id="summary" class="tab-content"><div class="dashboard-card overflow-auto mb-8"><div class="summary-title bg-blue-600">PWG Sales + PSR Sales</div><table class="text-sm"><thead id="summary-head-1"></thead><tbody id="summary-body-1"></tbody></table></div>
<div class="dashboard-card overflow-auto mb-8"><div class="summary-title bg-green-600">Xper Placement Drive Update - P2</div><table class="text-sm"><thead id="summary-head-2"></thead><tbody id="summary-body-2"></tbody></table></div>
<div class="dashboard-card overflow-auto mb-8"><div class="summary-title bg-purple-600">Fevikwik 463 Drive Update</div><table class="text-sm"><thead id="summary-head-3"></thead><tbody id="summary-body-3"></tbody></table></div></div>
<div id="pwg" class="tab-content hidden"></div><div id="xper" class="tab-content hidden"></div><div id="fwk" class="tab-content hidden"></div></div>
<script>
const summaryData1=__SUMMARY1_DATA__,summaryData2=__SUMMARY2_DATA__,summaryData3=__SUMMARY3_DATA__,pwgData=__PWG_DATA__,xperData=__XPER_DATA__,fwkData=__FWK_DATA__; const tableState={};
function showTab(tabId,event){document.querySelectorAll('.tab-content').forEach(t=>t.classList.add('hidden'));document.getElementById(tabId).classList.remove('hidden');document.querySelectorAll('.tab-btn').forEach(b=>{b.classList.remove('bg-blue-600','text-white');b.classList.add('bg-gray-200');});event.target.classList.remove('bg-gray-200');event.target.classList.add('bg-blue-600','text-white');}
function renderSummary(data,headId,bodyId){const h=document.getElementById(headId),b=document.getElementById(bodyId);h.innerHTML='';b.innerHTML='';if(!data?.length)return;const c=Object.keys(data[0]);h.innerHTML=`<tr class="bg-gray-100">${c.map(x=>`<th class="p-3 font-bold text-gray-800">${x.includes('Unnamed')?'':x}</th>`).join('')}</tr>`;b.innerHTML=data.map((r,i)=>`<tr class="${i%2?'bg-gray-50':''}">${c.map(k=>`<td class="p-2">${r[k]??''}</td>`).join('')}</tr>`).join('');}
function getFilterColumns(columns){return columns.filter(c=>c.toLowerCase().includes('tty')||c.toLowerCase().includes('wss')||c.toLowerCase().includes('ssdm')||c.toLowerCase().includes('day')||c.toLowerCase().includes('cluster')||c.toLowerCase().includes('as name'));}
function findMatchingColumn(columns,keywords){return columns.find(col=>keywords.every(k=>String(col).toLowerCase().includes(k)));}
function normalizeFilterValue(column,value){const col=String(column||'').toLowerCase(),raw=String(value??'').trim();if(col.includes('day')||col.includes('drcp')) return raw.toUpperCase();return raw;}
function readCellNumber(v){const p=parseFloat(String(v??'').replace(/,/g,'').trim());return isNaN(p)?0:p;}
function getSortedData(data,sectionId){const st=tableState[sectionId];if(!st?.column)return [...data];const d=st.direction==='asc'?1:-1;return [...data].sort((a,b)=>{const av=a[st.column],bv=b[st.column],an=parseFloat(String(av??'').replace(/,/g,'')),bn=parseFloat(String(bv??'').replace(/,/g,''));if(!isNaN(an)&&!isNaN(bn))return (an-bn)*d;return String(av??'').toLowerCase().localeCompare(String(bv??'').toLowerCase())*d;});}
function calculateSubtotals(sectionId){const sec=document.getElementById(sectionId),rb=sec.querySelector('.subtotal-records'),boxes=sec.querySelectorAll('[data-subtotal-column]'),rows=sec.querySelectorAll('tbody tr');let rc=0;const t={};rows.forEach(r=>{const cells=r.querySelectorAll('td');if(cells.length){rc++;boxes.forEach(b=>{const c=b.dataset.subtotalColumn,idx=Number(b.dataset.columnIndex);if(idx>=0&&cells[idx])t[c]=(t[c]||0)+readCellNumber(cells[idx].innerText);});}});if(rb)rb.innerText=rc;boxes.forEach(b=>{const c=b.dataset.subtotalColumn;b.innerText=(t[c]||0).toLocaleString();});}
function createDealerTable(containerId,data,color='blue'){const container=document.getElementById(containerId);if(!data.length){container.innerHTML=`<div class="dashboard-card p-10 text-center">No Data Found</div>`;return;}
const columns=Object.keys(data[0]);const subtotalColumnsConfig={pwg:[findMatchingColumn(columns,['may','base']),findMatchingColumn(columns,['may','plan']),findMatchingColumn(columns,['may','ach']),findMatchingColumn(columns,['may','balance'])].filter(Boolean),xper:[findMatchingColumn(columns,['no of pouches','april+may']),findMatchingColumn(columns,['balance'])].filter(Boolean),fwk:[]};
if(containerId==='fwk'){const placementCol=findMatchingColumn(columns,['placement'])||findMatchingColumn(columns,['placed']);data=data.map(r=>{const v=placementCol?String(r[placementCol]??'').trim().toLowerCase():'';return {...r,Placement:['yes','y','placed','done','1','true'].includes(v)?'Yes':'No'};});}
const finalColumns=Object.keys(data[0]);const filterColumns=getFilterColumns(finalColumns);const filteredColumns=containerId==='xper'?filterColumns.filter(c=>!String(c).toLowerCase().includes('cluster')):filterColumns;if(containerId==='fwk'&&!filteredColumns.includes('Placement'))filteredColumns.push('Placement');const subtotalColumns=subtotalColumnsConfig[containerId]||[];
container.innerHTML=`<div class="dashboard-card p-6 mb-8"><div class="flex items-center justify-between mb-6"><h2 class="text-3xl font-black text-gray-800">${containerId.toUpperCase()} TRACKER</h2><div class="bg-blue-100 text-blue-700 px-4 py-2 rounded-2xl font-bold">${data.length} Dealers</div></div>
<div class="grid grid-cols-1 md:grid-cols-5 gap-4 mb-6">${filteredColumns.map(col=>`<select class="filter border border-gray-300 rounded-2xl p-3" data-column="${col}"><option value="">Filter ${col}</option>${[...new Set(data.map(r=>normalizeFilterValue(col,r[col])))].filter(v=>v!=='').map(v=>`<option value="${v}">${v}</option>`).join('')}</select>`).join('')}</div>
<div class="grid grid-cols-1 md:grid-cols-5 gap-5 mb-8"><div class="bg-blue-50 rounded-2xl p-5 border border-blue-200"><div class="text-sm text-gray-500">Records</div><div class="text-3xl font-black text-blue-700 subtotal-records">${data.length}</div></div>${subtotalColumns.map(col=>`<div class="bg-green-50 rounded-2xl p-5 border border-green-200"><div class="text-sm text-gray-500">${col} Subtotal</div><div class="text-3xl font-black text-green-700" data-subtotal-column="${col}" data-column-index="${finalColumns.indexOf(col)}">0</div></div>`).join('')}</div>
<div class="overflow-auto rounded-3xl border border-gray-200"><table class="w-full text-sm whitespace-nowrap"><thead class="bg-${color}-600 text-white"><tr>${finalColumns.map(col=>`<th class="p-3 sortable" data-column="${col}" data-section="${containerId}">${col}</th>`).join('')}</tr></thead><tbody id="${containerId}-body">${data.map((r,i)=>`<tr class="${i%2?'bg-gray-50':''}">${finalColumns.map(c=>`<td class="p-2">${r[c]??''}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
setupFilters(containerId,data);setupSorting(containerId,data);calculateSubtotals(containerId);}
function setupFilters(sectionId,originalData){const sec=document.getElementById(sectionId),filters=sec.querySelectorAll('.filter');filters.forEach(f=>f.addEventListener('change',()=>{let filtered=[...originalData];filters.forEach(x=>{const col=x.dataset.column,val=x.value;if(val!=='')filtered=filtered.filter(r=>normalizeFilterValue(col,r[col])===normalizeFilterValue(col,val));});renderFiltered(sectionId,filtered);}));}
function renderFiltered(sectionId,data){const tb=document.getElementById(`${sectionId}-body`);if(!data.length){tb.innerHTML=`<tr><td colspan="100%" class="p-6 text-center">No Records Found</td></tr>`;return;}const sorted=getSortedData(data,sectionId),cols=Object.keys(sorted[0]);tb.innerHTML=sorted.map((r,i)=>`<tr class="${i%2?'bg-gray-50':''}">${cols.map(c=>`<td class="p-2">${r[c]??''}</td>`).join('')}</tr>`).join('');calculateSubtotals(sectionId);}
function setupSorting(sectionId,originalData){const sec=document.getElementById(sectionId),heads=sec.querySelectorAll('th.sortable');heads.forEach(h=>h.addEventListener('click',()=>{const c=h.dataset.column,cur=tableState[sectionId];if(cur&&cur.column===c)tableState[sectionId].direction=cur.direction==='asc'?'desc':'asc';else tableState[sectionId]={column:c,direction:'asc'};let filtered=[...originalData];sec.querySelectorAll('.filter').forEach(f=>{if(f.value!==''){const col=f.dataset.column;filtered=filtered.filter(r=>normalizeFilterValue(col,r[col])===normalizeFilterValue(col,f.value));}});renderFiltered(sectionId,filtered);}));}
renderSummary(summaryData1,'summary-head-1','summary-body-1');renderSummary(summaryData2,'summary-head-2','summary-body-2');renderSummary(summaryData3,'summary-head-3','summary-body-3');createDealerTable('pwg',pwgData,'blue');createDealerTable('xper',xperData,'green');createDealerTable('fwk',fwkData,'purple');
</script></body></html>
"""
    html = html.replace("__SUMMARY1_DATA__", summary1_json).replace("__SUMMARY2_DATA__", summary2_json).replace("__SUMMARY3_DATA__", summary3_json)
    html = html.replace("__PWG_DATA__", pwg_json).replace("__XPER_DATA__", xper_json).replace("__FWK_DATA__", fwk_json)
    Path(OUTPUT_HTML).write_text(html, encoding="utf-8")
    return OUTPUT_HTML


def main():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Common Tracker file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    )
    if not file_path:
        messagebox.showinfo("Cancelled", "No file selected.")
        return
    output = build_dashboard(file_path)
    messagebox.showinfo("Success", f"Dashboard created: {output}")
    print(f"Dashboard Created Successfully: {output}")


if __name__ == "__main__":
    main()
