(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const DB_NAME = 'collection-local-db', STORE = 'workbooks', KEY = 'main';
  const toAscii = s => String(s || '').replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  function openDB(){ return new Promise((resolve,reject)=>{const r=indexedDB.open(DB_NAME,1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);}); }
  async function getWorkbook(){ const db=await openDB(); return new Promise((resolve,reject)=>{const r=db.transaction(STORE,'readonly').objectStore(STORE).get(KEY);r.onsuccess=()=>{try{resolve(r.result?.data ? XLSX.read(r.result.data,{type:'array',cellDates:true}) : null)}catch(e){reject(e)}};r.onerror=()=>reject(r.error);}); }
  async function saveWorkbook(wb){ const db=await openDB(); const data=XLSX.write(wb,{bookType:'xlsx',type:'array'}); return new Promise((resolve,reject)=>{const tx=db.transaction(STORE,'readwrite');tx.objectStore(STORE).put({data,saved:new Date().toISOString()},KEY);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);}); }
  function addStyles(){
    if($('editAddStyles')) return;
    const s=document.createElement('style');s.id='editAddStyles';s.textContent=`
      #manageModal{position:fixed;inset:0;z-index:99999;background:#10182799;display:flex;align-items:center;justify-content:center;padding:12px}
      #manageModal .manageCard{background:#fff;color:#172338;width:min(760px,100%);max-height:92vh;overflow:auto;border-radius:16px;padding:18px;box-shadow:0 12px 50px #0004}
      #manageModal .manageGrid{display:grid;grid-template-columns:1fr;gap:10px}
      #manageModal label{display:block;font-size:12px;font-weight:700;margin-bottom:5px}
      #manageModal input,#manageModal textarea{width:100%;font:inherit;padding:10px;border:1px solid #d6deea;border-radius:9px;box-sizing:border-box}
      #manageModal textarea{min-height:58px}
      #manageModal .manageActions{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
      #manageModal .manageBtn{border:0;border-radius:9px;padding:10px 14px;background:#e8eef7;color:#14243a;font-weight:700}
      #manageModal .managePrimary{background:#176b87;color:white}
      .manageRowBtn{border:0;border-radius:8px;padding:7px 9px;background:#e8eef7;color:#14243a;font-size:12px;font-weight:700}
      @media(min-width:650px){#manageModal .manageGrid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    `;document.head.appendChild(s);
  }
  function status(text,good=true){const el=$('status');if(el){el.textContent=text;el.style.color=good?'#176b55':'#a33';}}
  function makeModal(title, fields, onSave){
    addStyles();
    $('manageModal')?.remove();
    const modal=document.createElement('div');modal.id='manageModal';
    modal.innerHTML=`<div class="manageCard" role="dialog" aria-modal="true"><h2 style="font-size:18px;margin:0 0 8px">${esc(title)}</h2><p style="font-size:12px;color:#617087;line-height:1.8">تغییرات پس از ذخیره در نسخه محلی همین آیفون ثبت می‌شوند. برای نگهداری نسخه پشتیبان، بعد از تغییرات خروجی Excel بگیرید.</p><form id="manageForm"><div class="manageGrid">${fields.map((f,i)=>`<div><label for="manageField${i}">${esc(f.label)}</label><textarea id="manageField${i}">${esc(f.value)}</textarea></div>`).join('')}</div><div class="manageActions"><button class="manageBtn managePrimary" type="submit">ذخیره</button><button class="manageBtn" type="button" id="manageCancel">انصراف</button></div></form></div>`;
    document.body.appendChild(modal);
    $('manageCancel').onclick=()=>modal.remove();
    $('manageForm').onsubmit=async e=>{e.preventDefault();const values=fields.map((_,i)=>$('manageField'+i).value);const btn=$('manageForm').querySelector('[type=submit]');btn.disabled=true;btn.textContent='در حال ذخیره…';try{await onSave(values);modal.remove();status('تغییرات ذخیره شد.');location.reload();}catch(err){console.error(err);status('ذخیره انجام نشد: '+err.message,false);btn.disabled=false;btn.textContent='ذخیره';}};
  }
  function writeCell(ws,r,c,value){
    const addr=XLSX.utils.encode_cell({r,c});
    if(value===''){delete ws[addr];}
    else {const numeric=String(value).trim()!=='' && Number.isFinite(Number(value));ws[addr]=numeric?{t:'n',v:Number(value)}:{t:'s',v:String(value)};}
    const range=XLSX.utils.decode_range(ws['!ref']||'A1');range.e.r=Math.max(range.e.r,r);range.e.c=Math.max(range.e.c,c);ws['!ref']=XLSX.utils.encode_range(range);
  }
  async function editRecord(sheetName,rowNumber){
    const wb=await getWorkbook();if(!wb)throw new Error('فایل کلکسیون هنوز بارگذاری نشده است.');
    const ws=wb.Sheets[sheetName];if(!ws)throw new Error('شیت پیدا نشد: '+sheetName);
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
    const headers=(rows[0]||[]).map((v,i)=>String(v||`ستون ${i+1}`));while(headers.length&&headers[headers.length-1]==='')headers.pop();
    const row=rows[rowNumber-1]||[];
    makeModal(`ویرایش رکورد — ${sheetName} (ردیف ${rowNumber})`,headers.map((h,i)=>({label:h,value:row[i]??''})),async values=>{headers.forEach((_,i)=>writeCell(ws,rowNumber-1,i,values[i]??''));await saveWorkbook(wb);});
  }
  async function addRecord(){
    const select=$('sheetSelect'),sheetName=select?.value;
    if(!sheetName){status('برای افزودن رکورد، ابتدا یک شیت مشخص از فهرست انتخاب کنید.',false);return;}
    const wb=await getWorkbook();if(!wb)throw new Error('ابتدا فایل Excel را وارد کنید.');
    const ws=wb.Sheets[sheetName];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
    const headers=(rows[0]||[]).map((v,i)=>String(v||`ستون ${i+1}`));while(headers.length&&headers[headers.length-1]==='')headers.pop();
    makeModal(`افزودن رکورد — ${sheetName}`,headers.map(h=>({label:h,value:''})),async values=>{const range=XLSX.utils.decode_range(ws['!ref']||'A1');const target=range.e.r+1;headers.forEach((_,i)=>writeCell(ws,target,i,values[i]??''));await saveWorkbook(wb);});
  }
  function install(){
    if(!$('sheetSelect')||!$('tbody')||!window.XLSX)return;
    if(!$('addRecordExtra')){const b=document.createElement('button');b.id='addRecordExtra';b.className='btn primary';b.textContent='افزودن رکورد';b.type='button';b.onclick=()=>addRecord().catch(e=>status(e.message,false));$('importBtn')?.parentElement?.appendChild(b);}
    const decorate=()=>{
      const head=$('thead')?.querySelector('tr');if(head&&!head.querySelector('.manageHead')){const th=document.createElement('th');th.className='manageHead';th.textContent='ویرایش';head.appendChild(th);}
      const selected=$('sheetSelect').value;
      [...$('tbody').querySelectorAll('tr')].forEach(tr=>{
        if(tr.dataset.manageReady==='1')return;
        const cells=[...tr.querySelectorAll('td')];if(!cells.length||tr.querySelector('.empty'))return;
        const sheetName=selected||cells[0]?.textContent?.trim();
        const rowCell=selected?cells[0]:cells[1];
        const rowNumber=Number(toAscii(rowCell?.textContent?.trim()));
        if(!sheetName||!Number.isFinite(rowNumber)||rowNumber<2)return;
        const td=document.createElement('td');const b=document.createElement('button');b.className='manageRowBtn';b.type='button';b.textContent='ویرایش';
        b.onclick=()=>editRecord(sheetName,rowNumber).catch(e=>status(e.message,false));td.appendChild(b);tr.appendChild(td);tr.dataset.manageReady='1';
      });
    };
    new MutationObserver(decorate).observe($('tbody'),{childList:true,subtree:true});
    new MutationObserver(decorate).observe($('thead'),{childList:true,subtree:true});
    decorate();
  }
  window.addEventListener('load',()=>setTimeout(install,700));
})();
