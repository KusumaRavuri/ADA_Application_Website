/* ADA Application Portal — Multi-step Form JS (9 steps, enhanced) */

let currentStep = 1;
const TOTAL_STEPS = 9;
const STORAGE_KEY = "ada_portal_form_v4";

const STEP_META = [
  { title:"Personal Details", fields:["name","gender","dob","aadhaar","mobile","email","whatsapp","landline","imei","permanent_address","present_address","previous_address","qualification","marital_status","designation","pan","facebook","linkedin","instagram","other_id","bank_name","bank_account","ifsc","epf","esic"],
    labels:{name:"Full Name",gender:"Gender",dob:"Date of Birth",aadhaar:"Aadhaar",mobile:"Mobile",email:"Email",whatsapp:"WhatsApp",landline:"Landline",imei:"IMEI",permanent_address:"Permanent Address",present_address:"Present Address",previous_address:"Previous Address",qualification:"Qualification",marital_status:"Marital Status",designation:"Designation",pan:"PAN",facebook:"Facebook",linkedin:"LinkedIn",instagram:"Instagram",other_id:"Other ID",bank_name:"Bank Name",bank_account:"Account No",ifsc:"IFSC",epf:"EPF",esic:"ESIC"} },
  { title:"Education Details", fields:["tenth_school","tenth_board","tenth_year","tenth_percent","inter_institution","inter_board","inter_year","inter_percent","btech_college","btech_branch","btech_cgpa","btech_year"],
    labels:{tenth_school:"10th School",tenth_board:"10th Board",tenth_year:"10th Year",tenth_percent:"10th %",inter_institution:"Inter/Diploma Institution",inter_board:"Inter Board",inter_year:"Inter Year",inter_percent:"Inter %",btech_college:"BTech College",btech_branch:"Branch",btech_cgpa:"CGPA",btech_year:"Year of Study"} },
  { title:"Family Details", fields:["father_name","father_occupation","father_mobile","mother_name","mother_occupation","mother_mobile","num_siblings","sibling_info"],
    labels:{father_name:"Father Name",father_occupation:"Father Occupation",father_mobile:"Father Mobile",mother_name:"Mother Name",mother_occupation:"Mother Occupation",mother_mobile:"Mother Mobile",num_siblings:"No. of Siblings",sibling_info:"Sibling Info"} },
  { title:"Faculty Coordinator", fields:["faculty_name","faculty_designation","faculty_department","affiliation_id","faculty_email","faculty_contact","faculty_fax"],
    labels:{faculty_name:"Faculty Name",faculty_designation:"Designation",faculty_department:"Department",affiliation_id:"Affiliation ID",faculty_email:"Faculty Email",faculty_contact:"Faculty Contact",faculty_fax:"Faculty Fax"} },
  { title:"College Details", fields:["college_name","principal_name","university_affiliation_name","university_affiliation_no","aicte_code","dte_code","college_email","college_contact","college_fax"],
    labels:{college_name:"College Name",principal_name:"Principal",university_affiliation_name:"University Affiliation",university_affiliation_no:"Affiliation No",aicte_code:"AICTE Code",dte_code:"DTE Code",college_email:"College Email",college_contact:"College Contact",college_fax:"College Fax"} },
  { title:"Project Details", fields:["project_title","guide","area_of_work","duration_from","duration_to","temp_pass","temp_pass_validity","university_reg","letter_no","letter_date"],
    labels:{project_title:"Project Title",guide:"Guide/Directorate",area_of_work:"Area of Work",duration_from:"Duration From",duration_to:"Duration To",temp_pass:"Temp Pass",temp_pass_validity:"Pass Validity",university_reg:"Univ. Reg No.",letter_no:"Letter No.",letter_date:"Letter Date"} },
];

const REQUIRED_MAP = {
  1:["name","gender","dob","aadhaar","mobile","email","permanent_address","present_address","qualification"],
  2:["btech_college","btech_branch"],
  3:[],4:["faculty_name"],5:["college_name"],
  6:["duration_from","duration_to"],7:[],8:[],9:[]
};

function getAllFormFields() {
  const form = document.getElementById("appForm");
  const data = {};
  form.querySelectorAll("input:not([type=file]), select, textarea").forEach(el => {
    if (el.name) data[el.name] = el.value;
  });
  data["__step"] = currentStep;
  return data;
}

function saveFormData() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(getAllFormFields())); } catch(e) {}
}

function restoreFormData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    const form = document.getElementById("appForm");
    Object.keys(data).forEach(name => {
      if (name === "__step") return;
      const el = form.querySelector(`[name="${name}"]`);
      if (el) el.value = data[name];
    });
    const savedStep = parseInt(data["__step"]) || 1;
    if (savedStep > 1 && savedStep <= TOTAL_STEPS) {
      getStep(currentStep).classList.remove("active");
      currentStep = savedStep;
      getStep(currentStep).classList.add("active");
      if (currentStep === TOTAL_STEPS) buildReview();
    }
  } catch(e) {}
}

function clearFormData() {
  try { localStorage.removeItem(STORAGE_KEY); } catch(e) {}
}

function getStep(n) { return document.getElementById("step"+n); }

function setStepIndicator(n) {
  document.querySelectorAll(".prog-step").forEach((el,i) => {
    el.classList.remove("active","completed");
    if (i+1===n) el.classList.add("active");
    else if (i+1<n) el.classList.add("completed");
  });
  const pct = Math.round(((n-1)/(TOTAL_STEPS-1))*100);
  document.getElementById("progressBar").style.width = Math.max(pct,6)+"%";
  document.getElementById("stepCounter").textContent = `Step ${n} of ${TOTAL_STEPS}`;
}

function showNav() {
  document.getElementById("prevBtn").style.visibility = currentStep===1?"hidden":"visible";
  document.getElementById("nextBtn").style.display    = currentStep===TOTAL_STEPS?"none":"inline-flex";
}

function validateStep(n) {
  const required = REQUIRED_MAP[n]||[];
  let ok = true;
  required.forEach(name => {
    const el = document.querySelector(`[name="${name}"]`);
    if (!el) return;
    if (!el.value.trim()) { el.classList.add("is-invalid"); ok=false; }
    el.addEventListener("input",()=>el.classList.remove("is-invalid"),{once:true});
  });
  if (!ok) { shakeCard(n); showToast("Please fill all required fields marked with *","danger"); return false; }

  if (n===7) {
    const pdf   = document.getElementById("rec_letter");
    const photo = document.getElementById("photo");
    let fok = true;
    if (!pdf.files.length)   { document.getElementById("pdfZone").style.borderColor="#e53935"; fok=false; }
    if (!photo.files.length) { document.getElementById("photoZone").style.borderColor="#e53935"; fok=false; }
    if (!fok) { shakeCard(n); showToast("Please upload both required files","danger"); return false; }
  }

  if (n===8) {
    const checks = document.querySelectorAll(".decl-check");
    if (![...checks].every(c=>c.checked)) {
      shakeCard(n); showToast("Please check all declaration checkboxes to proceed","warning"); return false;
    }
  }
  return true;
}

function shakeCard(n) {
  const card = getStep(n)?.querySelector(".glass-card");
  if (!card) return;
  card.style.animation="none"; card.offsetHeight;
  card.style.animation="shake 0.4s ease";
  setTimeout(()=>{ card.style.animation=""; },450);
}
const ss = document.createElement("style");
ss.textContent="@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}";
document.head.appendChild(ss);

function showToast(msg,type="danger") {
  let t = document.getElementById("ada-toast");
  if (!t) {
    t = document.createElement("div"); t.id="ada-toast";
    t.style.cssText="position:fixed;bottom:24px;left:50%;transform:translateX(-50%);z-index:9999;min-width:280px;max-width:460px;padding:0.85rem 1.4rem;border-radius:12px;font-size:0.88rem;font-weight:600;box-shadow:0 8px 32px rgba(0,0,0,0.2);backdrop-filter:blur(12px);transition:opacity .3s;pointer-events:none;text-align:center;";
    document.body.appendChild(t);
  }
  const cols={danger:"rgba(183,28,28,0.92)",warning:"rgba(230,81,0,0.92)",success:"rgba(46,125,50,0.92)"};
  t.style.background=cols[type]||cols.danger; t.style.color="#fff";
  t.style.border="1px solid rgba(255,255,255,0.2)";
  t.textContent=msg; t.style.opacity=1;
  clearTimeout(t._t); t._t=setTimeout(()=>{t.style.opacity=0;},3500);
}

function changeStep(dir) {
  const next = currentStep+dir;
  if (dir===1 && !validateStep(currentStep)) return;
  if (next<1||next>TOTAL_STEPS) return;
  saveFormData();
  getStep(currentStep).classList.remove("active");
  currentStep=next;
  getStep(currentStep).classList.add("active");
  if (currentStep===TOTAL_STEPS) buildReview();
  setStepIndicator(currentStep);
  showNav();
  window.scrollTo({top:0,behavior:"smooth"});
}

function setupFileZone(inputId,zoneId,previewId,allowedExt,isImage) {
  const inp=document.getElementById(inputId);
  const zone=document.getElementById(zoneId);
  const prev=document.getElementById(previewId);
  if (!inp||!zone||!prev) return;
  zone.addEventListener("dragover",e=>{e.preventDefault();zone.classList.add("has-file");});
  zone.addEventListener("dragleave",()=>zone.classList.remove("has-file"));
  zone.addEventListener("drop",e=>{
    e.preventDefault();
    const file=e.dataTransfer.files[0];
    if (file){const dt=new DataTransfer();dt.items.add(file);inp.files=dt.files;handleFile(file,zone,prev,allowedExt,isImage);}
  });
  inp.addEventListener("change",()=>{ if(inp.files[0]) handleFile(inp.files[0],zone,prev,allowedExt,isImage); });
}

function handleFile(file,zone,prev,allowedExt,isImage) {
  const ext=file.name.split(".").pop().toLowerCase();
  if (!allowedExt.includes(ext)){ showToast(`Invalid file. Allowed: ${allowedExt.join(", ")}`,"danger"); return; }
  zone.classList.add("has-file"); zone.style.borderColor="";
  zone.querySelector(".upload-zone-text").textContent="✓ "+file.name;
  zone.querySelector(".upload-zone-hint").textContent=(file.size/1024).toFixed(0)+" KB";
  prev.classList.remove("d-none");
  if (isImage) {
    const r=new FileReader();
    r.onload=e=>{prev.innerHTML=`<img src="${e.target.result}" style="max-height:90px;border-radius:8px;border:1px solid var(--blue-200);"/>`;};
    r.readAsDataURL(file);
  } else {
    prev.innerHTML=`<i class="bi bi-file-earmark-check-fill me-2"></i>PDF ready: ${file.name}`;
  }
}

function buildReview() {
  const form=document.getElementById("appForm");
  const fd=new FormData(form);
  const container=document.getElementById("reviewContent");
  container.innerHTML="";
  STEP_META.forEach(meta=>{
    const sec=document.createElement("div"); sec.className="review-section";
    sec.innerHTML=`<div class="review-section-title">${meta.title}</div><div class="review-fields"></div>`;
    container.appendChild(sec);
    const fieldsDiv=sec.querySelector(".review-fields");
    meta.fields.forEach(f=>{
      const val=fd.get(f)||"";
      const label=meta.labels[f]||f;
      fieldsDiv.innerHTML+=`<div class="review-field"><div class="review-key">${label}</div><div class="review-val ${!val?"review-empty":""}">${val||"—"}</div></div>`;
    });
  });
  const pdfFile=document.getElementById("rec_letter").files[0];
  const photoFile=document.getElementById("photo").files[0];
  const fileSec=document.createElement("div"); fileSec.className="review-section";
  fileSec.innerHTML=`<div class="review-section-title">Documents</div><div class="review-fields">
    <div class="review-field"><div class="review-key">Recommendation Letter</div><div class="review-val ${!pdfFile?"review-empty":""}">${pdfFile?pdfFile.name:"— Not uploaded"}</div></div>
    <div class="review-field"><div class="review-key">Passport Photo</div><div class="review-val ${!photoFile?"review-empty":""}">${photoFile?photoFile.name:"— Not uploaded"}</div></div>
  </div>`;
  container.appendChild(fileSec);
}

async function submitForm() {
  const btn=document.getElementById("submitBtn");
  const errEl=document.getElementById("submitError");
  const pdfFile=document.getElementById("rec_letter").files[0];
  const photoFile=document.getElementById("photo").files[0];
  if (!pdfFile)   { showToast("Recommendation Letter missing. Go back to step 7.","danger"); return; }
  if (!photoFile) { showToast("Passport photo missing. Go back to step 7.","danger"); return; }
  btn.disabled=true;
  btn.innerHTML='<span class="spinner-border spinner-border-sm me-2"></span>Submitting…';
  errEl.classList.add("d-none");
  const fd=new FormData(document.getElementById("appForm"));
  try {
    const resp=await fetch("/submit",{method:"POST",body:fd});
    const data=await resp.json();
    if (data.success) {
      clearFormData();
      // Show modal with App ID and download link
      document.getElementById("modalAppId").textContent = data.app_id;
      document.getElementById("modalDate").textContent  = data.sub_date;
      document.getElementById("modalTime").textContent  = data.sub_time;
      if (data.pdf_available) {
        document.getElementById("modalDownloadBtn").href = `/download_pdf/${data.app_id}`;
        document.getElementById("modalDownloadBtn").style.display = "";
      } else {
        document.getElementById("modalDownloadBtn").style.display = "none";
      }
      const modal = new bootstrap.Modal(document.getElementById("successModal"));
      modal.show();
    } else {
      errEl.textContent=data.error||"Submission failed. Please try again.";
      errEl.classList.remove("d-none");
      btn.disabled=false;
      btn.innerHTML='<i class="bi bi-send-fill me-2"></i>Submit Application';
    }
  } catch(err) {
    errEl.textContent="Network error. Please check your connection and try again.";
    errEl.classList.remove("d-none");
    btn.disabled=false;
    btn.innerHTML='<i class="bi bi-send-fill me-2"></i>Submit Application';
  }
}

document.addEventListener("DOMContentLoaded",()=>{
  restoreFormData();
  setStepIndicator(currentStep); showNav();
  setupFileZone("rec_letter","pdfZone","pdfPreview",["pdf"],false);
  setupFileZone("photo","photoZone","photoPreview",["jpg","jpeg","png"],true);
  document.querySelectorAll(".decl-check").forEach(cb=>{
    cb.addEventListener("change",()=>cb.closest(".check-item").classList.toggle("checked",cb.checked));
  });
  document.querySelectorAll(".glass-input").forEach(el=>{
    el.addEventListener("input",()=>{ el.classList.remove("is-invalid"); saveFormData(); });
    el.addEventListener("change",()=>saveFormData());
  });
});

// ── Dynamic Add-Row: Education ────────────────────────────
let eduRowCount = 0;
function addEduRow() {
  eduRowCount++;
  const idx = eduRowCount;
  const tbody = document.getElementById("extraEduRows");
  const tr = document.createElement("tr");
  tr.id = `eduRow_${idx}`;
  tr.innerHTML = `
    <td>
      <input type="text" class="form-control glass-input-sm" name="edu_level_${idx}" placeholder="e.g. M.Tech, MBA"/>
    </td>
    <td>
      <input type="text" class="form-control glass-input-sm" name="edu_inst_${idx}" placeholder="Institution name"/>
    </td>
    <td>
      <input type="text" class="form-control glass-input-sm" name="edu_board_${idx}" placeholder="Board/University"/>
    </td>
    <td>
      <input type="text" class="form-control glass-input-sm" name="edu_year_${idx}" placeholder="Year"/>
    </td>
    <td style="position:relative;">
      <input type="text" class="form-control glass-input-sm" name="edu_percent_${idx}" placeholder="% / CGPA"/>
      <button type="button" class="btn-remove-tr" onclick="removeRow('eduRow_${idx}')" title="Remove"><i class="bi bi-x-circle-fill"></i></button>
    </td>`;
  tbody.appendChild(tr);
  tr.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ── Dynamic Add-Row: Family ───────────────────────────────
let famRowCount = 0;
function addFamilyRow() {
  famRowCount++;
  const idx = famRowCount;
  const tbody = document.getElementById("extraFamilyRows");
  const tr = document.createElement("tr");
  tr.id = `famRow_${idx}`;
  tr.innerHTML = `
    <td>
      <input type="text" class="form-control glass-input-sm" name="fam_type_${idx}" placeholder="Sibling/Guardian"/>
    </td>
    <td>
      <input type="text" class="form-control glass-input-sm" name="fam_name_${idx}" placeholder="Full name"/>
    </td>
    <td>
      <input type="text" class="form-control glass-input-sm" name="fam_occupation_${idx}" placeholder="Occupation"/>
    </td>
    <td style="position:relative;">
      <input type="tel" class="form-control glass-input-sm" name="fam_mobile_${idx}" placeholder="+91 XXXXXXXXXX"/>
      <button type="button" class="btn-remove-tr" onclick="removeRow('famRow_${idx}')" title="Remove"><i class="bi bi-x-circle-fill"></i></button>
    </td>`;
  tbody.appendChild(tr);
  tr.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function removeRow(id) {
  const el = document.getElementById(id);
  if (el) { el.style.opacity = "0"; el.style.transition = "opacity .2s"; setTimeout(() => el.remove(), 200); }
}
