(function(){
  const TRACKS = {
    xp: { label: 'Extreme Programming', slots: { Y2T3:'BKCHAIN', Y3T1:'IOTDEVT', Y3T2:'RAPIDEV', Y3T3:'DEVOPS' } },
    infosec: { label: 'Information Security', slots: { Y2T3:'INFOSEC', Y3T1:'ITAUDIT', Y3T2:'DIGIFOR', Y3T3:'SECPROG' } },
    datamgmt: { label: 'Data Management', slots: { Y2T3:'BAFBANI', Y3T1:'DATAOPT', Y3T2:'BIGDATA', Y3T3:'DATAVIS' } }
  };

  const BASE_SUBJECTS = [
    // 1st year
    S(1,1,'BICHECO',3),S(1,1,'CSBLIFE',3),S(1,1,'INCOMPU',3),S(1,1,'ISPRGG1',3),S(1,1,'LEADMGT',3),S(1,1,'PEONEFF',2),S(1,1,'PURPCOM',3),S(1,1,'UNDSELF',3),
    S(1,2,'CRITHNK',3),S(1,2,'FUNDSYS',3),S(1,2,'ISFINMA',3),S(1,2,'ISPRGG2',3),S(1,2,'MATWRLD',3),S(1,2,'NSTP-01',3),S(1,2,'PETWODA',2),S(1,2,'SCITECH',3),
    S(1,3,'ASEANST',3),S(1,3,'BUSPROC',3),S(1,3,'DSTALGO',3),S(1,3,'ENTPROG',3),S(1,3,'ITINFNT',3),S(1,3,'NSTP-02',3),S(1,3,'PETRIID',2),S(1,3,'PROFISS',3),
    // 2nd year
    S(2,1,'APPDAET',3),S(2,1,'ARTAPRI',3),S(2,1,'BUSEVAL',3),S(2,1,'PEFORTS',2),S(2,1,'READHIS',3),S(2,1,'SYSANDE',3),S(2,1,'WEBDEVT',3),
    S(2,2,'APPINTR',3),S(2,2,'CONWORLD',3),S(2,2,'IENTARC',3),S(2,2,'INFOMGT',3),S(2,2,'JORIZAL',3),S(2,2,'UIUXDES',3),
    S(2,3,'INFODBM',3), S(2,3,'ISPRMGT',3), S(2,3,'MOBDEVT',3), S(2,3,'REEXECO',3),
    // 3rd year
    S(3,1,'PROCOMM',3), S(3,1,'ISQUANT',3), S(3,1,'ISSRCH',3), S(3,1,'ISSTRAMA',3),
    S(3,2,'GLOBCOM',3), S(3,2,'IETHICS',3), S(3,2,'ISPROJ1',3), S(3,2,'SOFTEST',3),
    S(3,3,'CSBGRAD',1), S(3,3,'MARFRET',3), S(3,3,'ISPROJ2',3),
    // 4th year
    S(4,1,'ISPRACT',6)
  ];

  function S(year, term, code, units){ return { year, term, code, units, isTrackSlot:false }; }
  function Slot(year, term, slotKey, units){ return { year, term, code: slotKey, units, isTrackSlot:true }; }

  const STORAGE_KEY = 'is-grade-calculator';
  function loadState(){
    try{ return JSON.parse(localStorage.getItem(STORAGE_KEY)) || { grades:{}, selectedTrack:'xp' }; }
    catch{ return { grades:{}, selectedTrack:'xp' }; }
  }
  function saveState(state){ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }

  function getActiveSubjects(selectedTrack){
    const slots = TRACKS[selectedTrack].slots;
    return BASE_SUBJECTS.map(sub => {
      if(!sub.isTrackSlot) return sub;
      const code = slots[sub.code];
      return { year: sub.year, term: sub.term, code, units: sub.units, isTrackSlot:false };
    });
  }

  function computeWeightedAverage(items, grades){
    let w=0,u=0;
    for(const it of items){
      const g = grades[it.code];
      if(typeof g === 'number' && g >= 1 && g <= 4){ w += g * it.units; u += it.units; }
    }
    if(u === 0) return null;
    const v = Math.round((w/u)*100)/100;
    return v;
  }

  const state = loadState();

  function render(){
    const app = document.getElementById('app');
    app.innerHTML = '';
    const subjects = getActiveSubjects(state.selectedTrack);

    const byYear = groupBy(subjects, s => s.year);
    Object.keys(byYear).sort((a,b)=>a-b).forEach(y => {
      const year = Number(y);
      const yearDiv = div('year');
      yearDiv.appendChild(div('year-title', `${ordinal(year)} YEAR`));

      const termsWrap = div('terms');
      const byTerm = groupBy(byYear[y], s => s.term);
      [1,2,3].forEach(termNum => {
        const termSubjects = byTerm[termNum] || [];
        if(termSubjects.length === 0) return; // 4th year may have 1 term
        const termCard = div('term');
        termCard.appendChild(div('term-header', `${ordinal(termNum)} TERM`));

        const table = document.createElement('table');
        table.className = 'table';
        table.innerHTML = `<thead><tr><th>SUBJECT</th><th class="units">UNITS</th><th class="grade">GRADE</th></tr></thead>`;
        const tbody = document.createElement('tbody');
        for(const s of termSubjects){
          const tr = document.createElement('tr');
          const codeTd = document.createElement('td');
          codeTd.textContent = s.code;
          const unitsTd = document.createElement('td');
          unitsTd.className='units';
          unitsTd.textContent = String(s.units);
          const gradeTd = document.createElement('td');
          gradeTd.className='grade';
          const input = document.createElement('input');
          input.type = (s.code==='CSBGRAD') ? 'text' : 'number';
          input.className = 'grade';
          if(s.code==='CSBGRAD'){
            input.placeholder = 'P or R';
            const val = state.grades[s.code];
            if(val===4 || val===5) input.value = (val===4?'P':'R');
            input.addEventListener('input', () => {
              const raw = input.value.trim().toUpperCase();
              if(raw==='P'){ state.grades[s.code]=4; input.classList.remove('invalid'); }
              else if(raw==='R'){ state.grades[s.code]=5; input.classList.remove('invalid'); }
              else if(raw===''){ delete state.grades[s.code]; input.classList.remove('invalid'); }
              else { input.classList.add('invalid'); return; }
              updateAverages();
              updateInvalidBadge();
            });
          } else {
            input.step = '0.25';
            input.min = '1';
            input.max = '4';
            if(typeof state.grades[s.code] === 'number') input.value = String(state.grades[s.code]);
            input.addEventListener('input', () => {
              const v = parseFloat(input.value);
              if(Number.isFinite(v) && v >= 1 && v <= 4){ state.grades[s.code] = v; input.classList.remove('invalid'); }
              else if(input.value===''){ delete state.grades[s.code]; input.classList.remove('invalid'); }
              else { input.classList.add('invalid'); }
              updateAverages();
              updateInvalidBadge();
            });
          }
          gradeTd.appendChild(input);
          tr.append(codeTd, unitsTd, gradeTd);
          tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        termCard.appendChild(table);

        const footer = div('term-footer');
        const totalUnits = termSubjects.reduce((sum,s)=>sum+s.units,0);
        const avgSpan = document.createElement('span');
        avgSpan.dataset.year = String(year);
        avgSpan.dataset.term = String(termNum);
        footer.appendChild(span(`Units: ${totalUnits}`));
        footer.appendChild(span('Average: ', avgSpan));
        termCard.appendChild(footer);
        termsWrap.appendChild(termCard);
      });
      yearDiv.appendChild(termsWrap);
      app.appendChild(yearDiv);
    });

    updateAverages();
    updateInvalidBadge();
  }

  function updateAverages(){
    const subjects = getActiveSubjects(state.selectedTrack);
    // per-term
    for(const year of [1,2,3,4]){
      for(const term of [1,2,3]){
        const items = subjects.filter(s=>s.year===year && s.term===term);
        if(items.length===0) continue;
        const avg = computeWeightedAverage(items, state.grades);
        const el = document.querySelector(`.term-footer span[data-year="${year}"][data-term="${term}"]`);
        if(el) el.textContent = avg==null?'—':avg.toFixed(2);
      }
    }
    // overall
    const overallDiv = document.getElementById('overall');
    const overall = computeWeightedAverage(subjects, state.grades);
    const honor = overall==null ? '' : ` — ${honorFromGPA(overall)}`;
    overallDiv.innerHTML = `Overall Average: ${overall==null?'—':overall.toFixed(2)}${honor}` +
      `<span class="note">Honors thresholds (CGPA): Summa 3.800–4.000, Magna 3.600–3.799, Cum Laude 3.400–3.599, Honorable Mention 3.000–3.399.</span>`;
  }

  function groupBy(arr, fn){
    const map = {};
    for(const item of arr){
      const k = fn(item);
      (map[k] ||= []).push(item);
    }
    return map;
  }
  function div(cls, text){ const d=document.createElement('div'); d.className=cls; if(text!=null){ d.textContent=String(text); } return d; }
  function span(text, node){ const s=document.createElement('span'); if(node){ s.textContent=text; s.appendChild(node); return s; } s.textContent=text; return s; }
  function ordinal(n){ return ['','1ST','2ND','3RD','4TH','5TH'][n] || `${n}TH`; }
  function honorFromGPA(g){
    if(g>=3.8) return 'Summa Cum Laude';
    if(g>=3.6) return 'Magna Cum Laude';
    if(g>=3.4) return 'Cum Laude';
    if(g>=3.0) return 'Honorable Mention';
    return 'No Honors';
  }

  function updateInvalidBadge(){
    const inputs = document.querySelectorAll('input.grade');
    let invalidCount = 0;
    inputs.forEach(inp => { if(inp.classList.contains('invalid')) invalidCount++; });
    const badge = document.getElementById('errorBadge');
    if(!badge) return;
    if(invalidCount>0){
      badge.classList.remove('hidden');
      badge.textContent = String(invalidCount);
      badge.title = `${invalidCount} invalid grade${invalidCount>1?'s':''}`;
    } else {
      badge.classList.add('hidden');
    }
  }

  // header controls
  window.addEventListener('DOMContentLoaded', () => {
    const trackSelect = document.getElementById('trackSelect');
    trackSelect.value = state.selectedTrack;
    trackSelect.addEventListener('change', () => {
      state.selectedTrack = trackSelect.value;
      render();
    });

    document.getElementById('saveBtn').addEventListener('click', () => saveState(state));
    document.getElementById('resetBtn').addEventListener('click', () => { localStorage.removeItem(STORAGE_KEY); state.grades={}; state.selectedTrack='xp'; document.getElementById('trackSelect').value='xp'; render(); });
    document.getElementById('excelBtn').addEventListener('click', () => exportExcel());

    render();
  });

  function exportCSV(){
    const subjects = getActiveSubjects(state.selectedTrack);
    
    // Create a well-structured CSV with proper spacing and organization
    const rows = [];
    
    // Header section with metadata
    rows.push(['BENILDE IS GRADE CALCULATOR']);
    rows.push(['']);
    rows.push(['Track:', state.selectedTrack.toUpperCase()]);
    rows.push(['Export Date:', new Date().toLocaleDateString()]);
    rows.push(['']);
    rows.push(['']); // Extra blank row for better separation
    
    // Column headers with proper spacing
    rows.push(['ACADEMIC YEAR', 'TERM', 'SUBJECT CODE', 'UNITS', 'GRADE', 'HONORS STATUS', 'TERM GPA']);
    rows.push(['']); // Extra blank row after headers
    
    // Group by year and term for GPA calculations
    const byYearTerm = {};
    subjects.forEach(s => {
      const key = `${s.year}-${s.term}`;
      if (!byYearTerm[key]) byYearTerm[key] = [];
      byYearTerm[key].push(s);
    });
    
    // Data rows, grouped by term with sub-headers and spacing
    for (const yearTermKey of Object.keys(byYearTerm).sort()) {
      const termSubjects = byYearTerm[yearTermKey];
      const [year, term] = yearTermKey.split('-');
      const termGPA = computeWeightedAverage(termSubjects, state.grades);

      // Add a sub-header for each term
      rows.push([`YEAR ${year} - ${ordinal(term)} TERM`]);
      
      for (const s of termSubjects) {
        const grade = state.grades[s.code];
        let gradeDisplay = '';
        let honorsStatus = '';

        if (s.code === 'CSBGRAD') {
          if (grade === 4) { gradeDisplay = 'P'; honorsStatus = 'PASS'; }
          else if (grade === 5) { gradeDisplay = 'R'; honorsStatus = 'FAIL'; }
        } else {
          if (typeof grade === 'number') {
            gradeDisplay = grade.toFixed(2);
            honorsStatus = honorFromGPA(grade); // Individual subject honors
          }
        }

        rows.push([
          year, // Academic Year
          ordinal(term), // Term
          s.code,
          s.units,
          gradeDisplay,
          honorsStatus,
          termGPA === null ? '' : termGPA.toFixed(2) // Term GPA repeated for each subject in term
        ]);
      }
      rows.push(['']); // Blank row after each term's subjects
    }

    // Summary section
    const overall = computeWeightedAverage(subjects, state.grades);
    rows.push(['SUMMARY']);
    rows.push(['']);
    rows.push(['Total Subjects:', subjects.length]);
    rows.push(['Track:', state.selectedTrack.toUpperCase()]);
    rows.push(['Overall GPA:', overall === null ? '' : overall.toFixed(2)]);
    rows.push(['Honors Level:', overall === null ? '' : honorFromGPA(overall)]);
    
    const csvContent = rows.map(e => e.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const date = new Date().toISOString().split('T')[0];
    link.setAttribute('download', `benilde-grades-${state.selectedTrack}-${date}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function exportExcel(){
    const subjects = getActiveSubjects(state.selectedTrack);
    
    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([]);
    
    // Header section with metadata
    XLSX.utils.sheet_add_aoa(ws, [
      ['BENILDE IS GRADE CALCULATOR'],
      [''],
      ['Track:', state.selectedTrack.toUpperCase()],
      ['Export Date:', new Date().toLocaleDateString()],
      [''],
      [''],
      ['ACADEMIC YEAR', 'TERM', 'SUBJECT CODE', 'UNITS', 'GRADE', 'HONORS STATUS', 'TERM GPA']
    ], { origin: 'A1' });
    
    // Group by year and term
    const byYearTerm = {};
    subjects.forEach(s => {
      const key = `${s.year}-${s.term}`;
      if (!byYearTerm[key]) byYearTerm[key] = [];
      byYearTerm[key].push(s);
    });
    
    let currentRow = 9; // Start after headers
    
    // Data rows with proper grouping
    for (const yearTermKey of Object.keys(byYearTerm).sort()) {
      const termSubjects = byYearTerm[yearTermKey];
      const [year, term] = yearTermKey.split('-');
      const termGPA = computeWeightedAverage(termSubjects, state.grades);

      // Add term header
      XLSX.utils.sheet_add_aoa(ws, [[`YEAR ${year} - ${ordinal(term)} TERM`]], { origin: `A${currentRow}` });
      currentRow++;
      
      // Add subjects for this term
      for (const s of termSubjects) {
        const grade = state.grades[s.code];
        let gradeDisplay = '';
        let honorsStatus = '';

        if (s.code === 'CSBGRAD') {
          if (grade === 4) { gradeDisplay = 'P'; honorsStatus = 'PASS'; }
          else if (grade === 5) { gradeDisplay = 'R'; honorsStatus = 'FAIL'; }
        } else {
          if (typeof grade === 'number') {
            gradeDisplay = grade.toFixed(2);
            honorsStatus = honorFromGPA(grade);
          }
        }

        XLSX.utils.sheet_add_aoa(ws, [[
          year, ordinal(term), s.code, s.units, gradeDisplay, honorsStatus, 
          termGPA === null ? '' : termGPA.toFixed(2)
        ]], { origin: `A${currentRow}` });
        currentRow++;
      }
      currentRow++; // Blank row after term
    }
    
    // Summary section
    const overall = computeWeightedAverage(subjects, state.grades);
    XLSX.utils.sheet_add_aoa(ws, [
      ['SUMMARY'],
      [''],
      ['Total Subjects:', subjects.length],
      ['Track:', state.selectedTrack.toUpperCase()],
      ['Overall GPA:', overall === null ? '' : overall.toFixed(2)],
      ['Honors Level:', overall === null ? '' : honorFromGPA(overall)]
    ], { origin: `A${currentRow}` });
    
    // Apply styling and formatting
    ws['!cols'] = [
      { width: 15 }, // Academic Year
      { width: 12 }, // Term
      { width: 15 }, // Subject Code
      { width: 8 },  // Units
      { width: 10 }, // Grade
      { width: 20 }, // Honors Status
      { width: 12 }  // Term GPA
    ];
    
    // Add the worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Grades');
    
    // Export the file
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `benilde-grades-${state.selectedTrack}-${date}.xlsx`);
  }
})();


