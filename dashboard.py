import json
import tempfile
import warnings
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pandas as pd
import requests

warnings.filterwarnings("ignore")

OUTPUT_HTML = "index.html"
NETLIFY_TOKEN = "nfp_qdZgVnpR5EvUgX8pkMzwQbnfWzVk3G3wb67f"
NETLIFY_SITE_ID = "fvtracker"


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


def build_dashboard_html(excel_file):
    summary_table1 = clean_dataframe(
        pd.read_excel(excel_file, sheet_name="Summary", header=0, usecols="A:J", nrows=8)
    )
    summary_table2 = clean_dataframe(
        pd.read_excel(excel_file, sheet_name="Summary", header=10, usecols="A:J", nrows=8)
    )
    summary_table3 = clean_dataframe(
        pd.read_excel(excel_file, sheet_name="Summary", header=20, usecols="A:I", nrows=8)
    )
    pwg_df = clean_dataframe(pd.read_excel(excel_file, sheet_name="PWG Dealers"))
    xper_df = clean_dataframe(pd.read_excel(excel_file, sheet_name="Xper Tracker"))
    fwk_df = clean_dataframe(pd.read_excel(excel_file, sheet_name="FWK463 Tracker"))

    html = Path(__file__).with_name("template.html")
    if html.exists():
        template = html.read_text(encoding="utf-8")
    else:
        template = """<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<script src='https://cdn.tailwindcss.com'></script><title>Dealer Dashboard</title></head><body class='bg-gray-100 p-4'>
<h1 class='text-3xl font-bold mb-4'>Dealer Dashboard</h1>
<div id='app'></div>
<script>
const summaryData1=__SUMMARY1_DATA__,summaryData2=__SUMMARY2_DATA__,summaryData3=__SUMMARY3_DATA__;
const pwgData=__PWG_DATA__,xperData=__XPER_DATA__,fwkData=__FWK_DATA__;
const tableState={};
function matchCol(cols, keys){return cols.find(c=>keys.every(k=>String(c).toLowerCase().includes(k)))}
function toNum(v){const n=parseFloat(String(v??'').replace(/,/g,''));return isNaN(n)?0:n}
function norm(col,v){const c=String(col).toLowerCase(),s=String(v??'').trim();if(c.includes('day')||c.includes('drcp')) return s.toUpperCase(); return s}
function filterCols(cols,id){let out=cols.filter(c=>['tty','wss','ssdm','day','cluster','as name'].some(k=>String(c).toLowerCase().includes(k))); if(id==='xper') out=out.filter(c=>!String(c).toLowerCase().includes('cluster')); return out}
function sortRows(rows,id){const st=tableState[id]; if(!st) return [...rows]; const d=st.dir==='asc'?1:-1; return [...rows].sort((a,b)=>{const an=toNum(a[st.col]),bn=toNum(b[st.col]); if(!isNaN(an)&&!isNaN(bn)&&String(a[st.col]).trim()!==''&&String(b[st.col]).trim()!=='') return (an-bn)*d; return String(a[st.col]??'').toLowerCase().localeCompare(String(b[st.col]??'').toLowerCase())*d;});}
function cards(id, cols){if(id==='fwk') return []; if(id==='pwg') return [matchCol(cols,['may','base']),matchCol(cols,['may','plan']),matchCol(cols,['may','ach']),matchCol(cols,['may','balance'])].filter(Boolean); return [matchCol(cols,['no of pouches','april+may']),matchCol(cols,['balance'])].filter(Boolean);}
function renderTable(id, rows, color){
 const app=document.getElementById('app'); const cols=Object.keys(rows[0]||{}); const subtotalCols=cards(id,cols); const filters=filterCols(cols,id);
 app.innerHTML += `<div class='bg-white rounded-xl shadow p-4 mb-5'><h2 class='text-2xl font-bold mb-3'>${id.toUpperCase()} TRACKER</h2><div class='grid grid-cols-1 md:grid-cols-5 gap-3 mb-4'>${filters.map(c=>`<select class='flt border rounded p-2' data-t='${id}' data-c='${c}'><option value=''>Filter ${c}</option>${[...new Set(rows.map(r=>norm(c,r[c])))].filter(v=>v!=='').map(v=>`<option>${v}</option>`).join('')}</select>`).join('')}</div><div class='overflow-auto border rounded'><table class='w-full text-sm'><thead class='bg-${color}-600 text-white'><tr>${cols.map(c=>`<th class='p-2 cursor-pointer sort' data-t='${id}' data-c='${c}'>${c}</th>`).join('')}</tr></thead><tbody id='${id}-body'></tbody></table></div><div class='grid grid-cols-1 md:grid-cols-5 gap-3 mt-4'><div class='bg-blue-50 border rounded p-3'><div class='text-xs text-gray-500'>Records</div><div class='text-2xl font-bold text-blue-700' id='${id}-records'>0</div></div>${subtotalCols.map(c=>`<div class='bg-green-50 border rounded p-3'><div class='text-xs text-gray-500'>${c} Subtotal</div><div class='text-2xl font-bold text-green-700' id='${id}-${c.replace(/\\W+/g,'_')}'>0</div></div>`).join('')}</div></div>`;
 function draw(filtered){ const sorted=sortRows(filtered,id); document.getElementById(`${id}-body`).innerHTML=sorted.map((r,i)=>`<tr class='${i%2?'bg-gray-50':''}'>${cols.map(c=>`<td class='p-2 border'>${r[c]??''}</td>`).join('')}</tr>`).join('')||`<tr><td class='p-4' colspan='100%'>No Records Found</td></tr>`; document.getElementById(`${id}-records`).innerText=filtered.length; subtotalCols.forEach(c=>{const t=filtered.reduce((a,r)=>a+toNum(r[c]),0); const el=document.getElementById(`${id}-${c.replace(/\\W+/g,'_')}`); if(el) el.innerText=t.toLocaleString();});}
 function current(){ let out=[...rows]; document.querySelectorAll(`select.flt[data-t="${id}"]`).forEach(f=>{if(f.value!==''){const c=f.dataset.c; out=out.filter(r=>norm(c,r[c])===norm(c,f.value));}}); return out;}
 draw(rows);
 document.querySelectorAll(`select.flt[data-t="${id}"]`).forEach(f=>f.addEventListener('change',()=>draw(current())));
 document.querySelectorAll(`th.sort[data-t="${id}"]`).forEach(h=>h.addEventListener('click',()=>{const c=h.dataset.c; if(tableState[id]&&tableState[id].col===c){tableState[id].dir=tableState[id].dir==='asc'?'desc':'asc'}else{tableState[id]={col:c,dir:'asc'}}; draw(current());}));
}
document.getElementById('app').innerHTML=`<div class='bg-white rounded-xl shadow p-4 mb-5'><h2 class='font-bold mb-2'>Summary Tables</h2><div class='text-sm text-gray-600'>Summary loaded from workbook.</div></div>`;
if(fwkData.length){const cols=Object.keys(fwkData[0]);const pc=cols.find(c=>String(c).toLowerCase().includes('placement')||String(c).toLowerCase().includes('placed'));for(const r of fwkData){const v=pc?String(r[pc]??'').trim().toLowerCase():''; r['Placement']=['yes','y','placed','done','1','true'].includes(v)?'Yes':'No';}}
if(pwgData.length) renderTable('pwg',pwgData,'blue'); if(xperData.length) renderTable('xper',xperData,'green'); if(fwkData.length) renderTable('fwk',fwkData,'purple');
</script></body></html>"""

    return (
        template.replace("__SUMMARY1_DATA__", json.dumps(summary_table1.to_dict(orient="records"), default=str))
        .replace("__SUMMARY2_DATA__", json.dumps(summary_table2.to_dict(orient="records"), default=str))
        .replace("__SUMMARY3_DATA__", json.dumps(summary_table3.to_dict(orient="records"), default=str))
        .replace("__PWG_DATA__", json.dumps(pwg_df.to_dict(orient="records"), default=str))
        .replace("__XPER_DATA__", json.dumps(xper_df.to_dict(orient="records"), default=str))
        .replace("__FWK_DATA__", json.dumps(fwk_df.to_dict(orient="records"), default=str))
    )


def deploy_to_netlify(site_id, token, html_path):
    with tempfile.TemporaryDirectory() as tmp:
        publish_dir = Path(tmp)
        (publish_dir / "index.html").write_text(Path(html_path).read_text(encoding="utf-8"), encoding="utf-8")
        files = {"file": ("index.html", (publish_dir / "index.html").read_bytes(), "text/html")}
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://api.netlify.com/api/v1/sites/{site_id}/deploys"
        response = requests.post(url, headers=headers, files=files, timeout=120)
        response.raise_for_status()
        return response.json()


def main():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Common Tracker Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    )
    if not file_path:
        messagebox.showinfo("Cancelled", "No file selected.")
        return

    try:
        html_content = build_dashboard_html(file_path)
        Path(OUTPUT_HTML).write_text(html_content, encoding="utf-8")
        deploy = deploy_to_netlify(NETLIFY_SITE_ID, NETLIFY_TOKEN, OUTPUT_HTML)
        deploy_url = deploy.get("deploy_ssl_url") or deploy.get("ssl_url") or "https://fvtracker.netlify.app/"
        messagebox.showinfo("Success", f"Generated {OUTPUT_HTML} and deployed.\nURL: {deploy_url}")
        print(f"Generated {OUTPUT_HTML}")
        print(f"Deployed to: {deploy_url}")
    except Exception as exc:
        messagebox.showerror("Error", str(exc))
        raise


if __name__ == "__main__":
    main()
