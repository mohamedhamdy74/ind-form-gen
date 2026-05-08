// ========================= STATE =========================
let parsedData = [];
let uploadedFile = null;

// ========================= FILE UPLOAD =========================
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const uploadCard = document.getElementById('upload-card');
const actionsDiv = document.getElementById('actions');

// Drag & Drop
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

dropZone.addEventListener('click', () => {
  fileInput.value = '';
  fileInput.click();
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

function handleFile(file) {
  const validTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'text/csv',
    '.xlsx', '.xls', '.csv'
  ];
  
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('من فضلك ارفع ملف Excel (.xlsx, .xls) أو CSV');
    return;
  }

  uploadedFile = file;
  document.getElementById('file-name').textContent = file.name;
  document.getElementById('file-size').textContent = formatFileSize(file.size);
  fileInfo.style.display = 'flex';

  // Parse Excel
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
      
      if (jsonData.length < 2) {
        alert('الملف فاضي أو مفيهوش بيانات كافية');
        return;
      }

      // ===== SMART COLUMN DETECTION =====
      let headerRowIndex = 0;
      let columnMap = { name: null, idNumber: null, position: null, date: null, company: null };
      let headerRow = [];

      // Scan up to the first 20 rows to find the actual header row
      for (let r = 0; r < Math.min(jsonData.length, 20); r++) {
        const potentialHeader = (jsonData[r] || []).map(h => String(h || ''));
        const tempMap = detectColumns(potentialHeader, false);
        
        // If we detect at least a name or ID column, consider it the header row
        if (tempMap.name !== null || tempMap.idNumber !== null) {
          headerRowIndex = r;
          headerRow = potentialHeader;
          columnMap = tempMap;
          break;
        }
      }

      // If we couldn't find a valid header row, fallback to the first row with positional assumption
      if (columnMap.name === null && columnMap.idNumber === null) {
        headerRowIndex = 0;
        headerRow = (jsonData[0] || []).map(h => String(h || ''));
        columnMap = detectColumns(headerRow, true);
      }
      
      if (columnMap.name === null && columnMap.idNumber === null) {
        alert('مش قادر أتعرف على الأعمدة. تأكد إن فيه عمود للاسم وعمود للرقم القومي.\n\nالأسماء المدعومة:\n- الاسم / Name / الأسم / اسم العامل\n- الرقم القومي / ID / National ID\n- الوظيفة / Position / Job\n- التاريخ / Date\n- الشركة / Company');
        return;
      }

      // Show detected columns info
      console.log('Detected columns:', columnMap, 'at row:', headerRowIndex + 1);
      
      // Skip header row, parse data
      parsedData = [];
      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue;
        
        const name = columnMap.name !== null ? String(row[columnMap.name] || '').trim() : '';
        const idNumber = columnMap.idNumber !== null ? String(row[columnMap.idNumber] || '').trim() : '';
        const position = columnMap.position !== null ? String(row[columnMap.position] || '').trim() : '';
        const company = columnMap.company !== null ? String(row[columnMap.company] || '').trim() : '';
        
        // Skip empty rows (no name and no ID)
        if (!name && !idNumber) continue;
        
        // Handle date
        let dateVal = columnMap.date !== null ? row[columnMap.date] : '';
        if (dateVal instanceof Date) {
          dateVal = formatDate(dateVal);
        } else if (typeof dateVal === 'number') {
          const d = new Date((dateVal - 25569) * 86400 * 1000);
          dateVal = formatDate(d);
        } else {
          dateVal = dateVal ? String(dateVal).trim() : '';
        }

        parsedData.push({ name, idNumber, position, date: dateVal, company });
      }

      document.getElementById('records-count').textContent = parsedData.length;
      actionsDiv.style.display = 'flex';

    } catch (err) {
      alert('حصل خطأ في قراءة الملف: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ========================= SMART COLUMN DETECTION =========================
function normalizeHeader(str) {
  return String(str || '')
    .toLowerCase()
    .trim()
    // Remove diacritics (tashkeel) for Arabic
    .replace(/[\u064B-\u065F\u0670]/g, '')
    // Normalize Arabic characters
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    // Remove extra spaces
    .replace(/\s+/g, ' ')
    // Remove common punctuation
    .replace(/[:\-_\.\/\\,;()#*]/g, '')
    .trim();
}

function detectColumns(headerRow, useFallback = true) {
  // Define all possible aliases for each field
  const fieldAliases = {
    name: [
      // Arabic
      'الاسم', 'الأسم', 'اسم', 'أسم', 'اسم العامل', 'أسم العامل', 'اسم الموظف',
      'الاسم الكامل', 'الاسم بالكامل', 'اسم كامل', 'الاسم الرباعي', 'الاسم الثلاثي',
      'اسم الشخص', 'اسم العميل', 'الأسماء', 'اسماء',
      // English  
      'name', 'full name', 'fullname', 'employee name', 'worker name',
      'person name', 'staff name', 'اname', 'الname',
    ],
    idNumber: [
      // Arabic
      'الرقم القومي', 'الرقم القومى', 'رقم قومي', 'رقم قومى',
      'رقم الهويه', 'رقم الهوية', 'رقم الهويه', 'الرقم الوطني',
      'رقم وطني', 'الهويه', 'الهوية', 'البطاقة', 'رقم البطاقة', 'رقم البطاقه',
      'بطاقه', 'بطاقة', 'الرقم', 'رقم قومى',
      // English
      'id', 'id number', 'idnumber', 'id no', 'id num',
      'national id', 'nationalid', 'national number', 'id card',
      'identification', 'identification number', 'ident',
      'ssn', 'social security', 'civil id',
    ],
    position: [
      // Arabic
      'الوظيفة', 'الوظيفه', 'وظيفة', 'وظيفه', 'المسمى الوظيفي',
      'المسمي الوظيفي', 'المنصب', 'منصب', 'الدور', 'دور',
      'المهنة', 'المهنه', 'مهنة', 'مهنه', 'العمل', 'نوع العمل',
      'الحرفة', 'الحرفه', 'حرفة', 'حرفه', 'التخصص',
      // English
      'position', 'job', 'job title', 'jobtitle', 'title',
      'role', 'occupation', 'profession', 'designation',
      'work', 'trade', 'craft', 'specialty',
    ],
    date: [
      // Arabic
      'التاريخ', 'تاريخ', 'تاريخ المحاضرة', 'تاريخ المحاضره',
      'تاريخ الاندكشن', 'تاريخ التعريف', 'تاريخ الحضور',
      'تاريخ الانضمام', 'تاريخ البدء', 'تاريخ البدايه', 'تاريخ البداية',
      'يوم', 'اليوم', 'التاريخ الميلادي',
      // English
      'date', 'induction date', 'inductiondate', 'joining date',
      'start date', 'startdate', 'training date', 'lecture date',
      'day', 'attendance date',
    ],
    company: [
      // Arabic
      'الشركة', 'الشركه', 'شركة', 'شركه', 'اسم الشركة', 'اسم الشركه',
      'المقاول', 'المقاول الفرعي', 'مقاول', 'جهة العمل', 'جهه العمل',
      'المؤسسة', 'المؤسسه',
      // English
      'company', 'company name', 'companyname', 'firm',
      'contractor', 'subcontractor', 'sub contractor',
      'employer', 'organization', 'organisation', 'org',
    ],
  };

  const result = { name: null, idNumber: null, position: null, date: null, company: null };
  const usedColumns = new Set();

  // Normalize all headers
  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));

  // For each field, try to find a matching column
  // Priority order: name -> idNumber -> position -> date -> company
  for (const [field, aliases] of Object.entries(fieldAliases)) {
    const normalizedAliases = aliases.map(a => normalizeHeader(a));
    
    for (let colIdx = 0; colIdx < normalizedHeaders.length; colIdx++) {
      if (usedColumns.has(colIdx)) continue;
      const header = normalizedHeaders[colIdx];
      if (!header) continue;
      
      // Check exact match first
      if (normalizedAliases.includes(header)) {
        result[field] = colIdx;
        usedColumns.add(colIdx);
        break;
      }
      
      // Check if header contains any alias (for combined headers like "اسم العامل الكامل")
      const matchFound = normalizedAliases.some(alias => 
        alias.length > 2 && (header.includes(alias) || alias.includes(header))
      );
      if (matchFound) {
        result[field] = colIdx;
        usedColumns.add(colIdx);
        break;
      }
    }
  }

  // Fallback: if no columns detected by name, use position-based (first 5 columns)
  if (useFallback && result.name === null && result.idNumber === null) {
    console.warn('No columns detected by header name, falling back to position-based detection');
    const totalCols = headerRow.length;
    if (totalCols >= 1) result.name = 0;
    if (totalCols >= 2) result.idNumber = 1;
    if (totalCols >= 3) result.position = 2;
    if (totalCols >= 4) result.date = 3;
    if (totalCols >= 5) result.company = 4;
  }

  return result;
}

function removeFile() {
  uploadedFile = null;
  parsedData = [];
  fileInput.value = '';
  fileInfo.style.display = 'none';
  actionsDiv.style.display = 'none';
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatDate(d) {
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

// ========================= FORM GENERATION =========================
function generateForms() {
  if (parsedData.length === 0) {
    alert('لا توجد بيانات لتوليد النماذج');
    return;
  }

  // Show loading
  const loading = document.createElement('div');
  loading.className = 'loading-overlay';
  loading.innerHTML = '<div class="loading-spinner"></div><div class="loading-text">جاري توليد النماذج...</div>';
  document.body.appendChild(loading);

  setTimeout(() => {
    const container = document.getElementById('forms-container');
    let formsHtml = '';

    parsedData.forEach((person, index) => {
      // Page 1
      formsHtml += generatePage1(person, index);
      // Page 2
      formsHtml += generatePage2(person, index);
    });

    container.innerHTML = formsHtml;

    document.getElementById('total-forms').textContent = parsedData.length;
    document.getElementById('upload-section').style.display = 'none';
    document.getElementById('forms-section').style.display = 'block';

    loading.remove();
    window.scrollTo(0, 0);
  }, 500);
}

function goBack() {
  document.getElementById('upload-section').style.display = 'flex';
  document.getElementById('forms-section').style.display = 'none';
}

// ========================= CEEC LOGO SVG =========================
function getCEECLogo() {
  return `
    <div class="ceec-logo-container">
      <img src="logo.jpg" alt="CEEC Energy China" style="height:50px; width:auto;" />
    </div>
  `;
}

// ========================= PAGE 1 GENERATOR =========================
function generatePage1(person, index) {
  return `
  <div class="form-page">
    <!-- Header -->
    <div class="form-header">
      ${getCEECLogo()}
      <div class="form-title">AFRE - 1000 MW SOLAR PROJECT</div>
      <div style="width:100px;"></div>
    </div>

    <!-- Main Table -->
    <table class="form-table">
      <!-- Title Row -->
      <tr>
        <th class="col-en" style="font-size:11pt;">HSE Induction</th>
        <th class="col-ar" style="font-size:11pt; font-family:'Cairo','Traditional Arabic',serif;">التعريف ببنود السلامة بالموقع</th>
      </tr>

      <!-- 1. Health & Safety -->
      <tr>
        <td class="col-en">
          <span class="section-num">1.</span> <span class="section-title-en">Health & Safety and Environment targets and policy:</span> The HSE management aimed to prevent accidents to reach average (zero accidents).
        </td>
        <td class="col-ar">
          <span class="section-num">1- أهداف و سياسة ادارة السلامة:</span><br/>
          تهدف إدارة السلامة الى الوصول الى منع الحوادث منها بإنها<br/>
          (Zero Accident)
        </td>
      </tr>

      <!-- 2. PPE -->
      <tr>
        <td class="col-en">
          <span class="section-num">2.</span> <span class="section-title-en">Personal Protective Equipment:</span><br/>
          All the staff (Managers – Engineers – Technicians – labors – Sub Contractors – Visitors) Should to wear the <b>PPE</b> (Safety Shoes – Helmet – Protective Glasses – Reflective Vest – Gloves) in the site all the time.<br/>
          Additional PPE may be required for specific tasks
        </td>
        <td class="col-ar">
          <span class="section-num">2- معدات الوقاية الشخصية:</span>
          <ul class="bullet-list-ar">
            <li>كل العاملين (مديرين – مهندسين – فنيين – عمال – مقاولين من الباطن – الزوار) يجب ارتداء معدات الوقاية الشخصية</li>
            <li>(الحذاء – الخوذة – نظارة السلامة – السترة العاكسة – القفازي) في الموقع كل الوقت.</li>
            <li>قد تحتاج بعض الاعمال المهمات الاضافيه يجب ارتكابها</li>
          </ul>
        </td>
      </tr>

      <!-- 3. Smoking -->
      <tr>
        <td class="col-en">
          <span class="section-num">3.</span> <span class="section-title-en">Smoking:</span><br/>
          Not allowed to smoke in the site except the areas allocated and specific by the HSE department.
        </td>
        <td class="col-ar">
          <span class="section-num">3- التدخين:</span><br/>
          غير مسموح بالتدخين في الموقع فيما عدا المناطق المخصصة لذلك والمحددة بواسطة قسم السلامة و الصحة المهنية.
        </td>
      </tr>

      <!-- 4. Speed Limit -->
      <tr>
        <td class="col-en">
          <span class="section-num">4.</span> <span class="section-title-en">Speed Limit:</span>
          <ul class="bullet-list">
            <li>The set speed is 20 km/h for light and pickup vehicles and 10 km/h for heavy trucks.</li>
            <li>The back alarm must be installed in every cars and equipment onsite.</li>
            <li>Not allowed to ride over trucks under any conditions.</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">4- السرعة المحددة:</span>
          <ul class="bullet-list-ar">
            <li>السرعة المحددة هي 20 كيلو بالقص هذا.</li>
            <li>الانذار الخلفي يجب تركيبه في كل سيارة و معدة في الموقع</li>
            <li>غير مسموح ركوب العاملين فوق السيارات تحت أي ظرف.</li>
          </ul>
        </td>
      </tr>

      <!-- 5. House Keeping -->
      <tr>
        <td class="col-en">
          <span class="section-num">5.</span> <span class="section-title-en">House Keeping:</span><br/>
          The work site should to be cleaned all the time, all wastes have to be removed daily, also nails, remnants of iron and any other sharp material.
        </td>
        <td class="col-ar">
          <span class="section-num">5- النظافة:</span><br/>
          يجب أن يكون موقع العمل نظيفاً في كل اوقات اليوم و يجب ازالة المسامير و المخلفات يوميا ، بالاضافة الى المسامير و بقايا الحديد و اي مواد حادة يومياً
        </td>
      </tr>

      <!-- 6. Electrical Panels -->
      <tr>
        <td class="col-en">
          <span class="section-num">6.</span> <span class="section-title-en">Electrical Panels and Connections:</span><br/>
          Please don't deal with any electrical connection before calling the electrician through the <b>HSE</b> responsible in the area.
        </td>
        <td class="col-ar">
          <span class="section-num">6- اللوحات و التوصيلات الكهربائية:</span><br/>
          يرجو عدم التعامل مع اي وصلة كهربائية قبل الاتصال بمسئول الكهرباء المتخصص و ذلك عن طريق مسئول السلامة في هذه المنطقة.
        </td>
      </tr>

      <!-- 7. Reporting Incidents -->
      <tr>
        <td class="col-en">
          <span class="section-num">7.</span> <span class="section-title-en">Reporting Incidents:</span>
          <ul class="bullet-list">
            <li>The labor's right to stop activities in case of danger for him or for any other one around him.</li>
            <li>All labors should to report for any dangerous situations or unsafe actions to the <b>HSE</b> supervisor to stop the work immediately till take the corrective actions.</li>
            <li>Oil leakage or any harmful materials have to deal with it through professional teamwork.</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">7- الإبلاغ عن الحوادث:</span>
          <ul class="bullet-list-ar">
            <li>من حق العامل إيقاف العمل في حالة وجود خطر عليه او على من حوله.</li>
            <li>على جميع العاملين الابلاغ عن الأوضاع الخطيرة و الأفعال الغير أمنة الى مسئول السلامة لإيقاف العمل فوراً</li>
            <li>تسرب الزيوت او المواد المحظورة يجب ان يتم التعامل معه بواسطة فريق عمل متخصف</li>
          </ul>
        </td>
      </tr>

      <!-- 8. Steps in case of Incident -->
      <tr>
        <td class="col-en">
          <span class="section-num">8.</span> <span class="section-title-en">Steps To Follow in case of Incident In The Site:</span>
          <ul class="bullet-list">
            <li>Have to tell your direct manager immediately, moreover to the <b>HSE</b> responsible or the site first aider.</li>
            <li>Don't move the injured guy.</li>
            <li>Don't give the injured guy any drinks.</li>
            <li>Ensure that the injured guy got enough sufficient air to breath.</li>
            <li>Cover the injured guy and feel him safe.</li>
            <li>In case of incident of electrical shock, have to switch off the electric before touch the shocked guy.</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">8- الخطوات التي يجب اتخاذها في حالة وقوع حادث في الموقع:</span>
          <ul class="bullet-list-ar">
            <li>يجب ان تخبر رئيسك المباشر فوراً ، بالاضافة الى مسئول السلامة أو مسعف الموقع.</li>
            <li>لا تحرك الشخص المصاب.</li>
            <li>لا تعطي الشخص المصاب اي مشروبات</li>
            <li>تأكد من حصول الشخص المصاب على قدر كافي من الهواء للتنفس.</li>
            <li>قم بتغطية المصاب و اطمأنه بحساس بالأمان.</li>
            <li>في حالة وقوع حادث صدمة كهربائية ، يجب فصل التيار الكهربائي قبل لمس المصاب.</li>
          </ul>
        </td>
      </tr>

      <!-- 9. Dangerous Materials -->
      <tr>
        <td class="col-en">
          <span class="section-num">9.</span> <span class="section-title-en">Dealing with Dangerous Materials:</span>
          <ul class="bullet-list">
            <li>Haven't deal with dangerous materials</li>
            <li>Have to get rid of empty containers according the environmental system in the site and put it in the allotted place.</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">9- التعامل مع المواد الخطرة:</span>
          <ul class="bullet-list-ar">
            <li>يجب عدم التعامل المواد الخطرة</li>
            <li>يجب التخلص من الحاويات الفارغة حسب النظام البيئي بالموقع و وضعها في المكان المخصص لها.</li>
          </ul>
        </td>
      </tr>
    </table>

    <div class="page-footer">1 | Page</div>
  </div>
  `;
}

// ========================= PAGE 2 GENERATOR =========================
function generatePage2(person, index) {
  // Build ID number boxes
  const idDigits = person.idNumber.replace(/\D/g, '');
  let idBoxesHtml = '';
  for (let i = 0; i < 14; i++) {
    idBoxesHtml += `<div class="id-box">${idDigits[i] || ''}</div>`;
  }

  return `
  <div class="form-page">
    <!-- Header -->
    <div class="form-header">
      ${getCEECLogo()}
      <div class="form-title">AFRE - 1000 MW SOLAR PROJECT</div>
      <div style="width:100px;"></div>
    </div>

    <!-- Main Table Continued -->
    <table class="form-table">
      <!-- 10. Emergency -->
      <tr>
        <td class="col-en">
          <span class="section-num">• 10.</span> <span class="section-title-en">Emergency:</span>
          <ul class="bullet-list">
            <li>Identify the place of emergency numbers.</li>
            <li>Identify the place of assembly points in case of evacuation.</li>
            <li>Join in trainings of the site for emergency drillings.</li>
            <li>Identify places of fire extinguishers and how to extinguish.</li>
            <li>A personal envelope card will be distributed to each person with instructions</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">10- الطوارئ:</span>
          <ul class="bullet-list-ar">
            <li>تعرف على أماكن أرقام الطوارئ.</li>
            <li>تعرف على أماكن التجمع أثناء الطوارئ.</li>
            <li>شارك في تدريبات الموقع الخاصة بالطوارئ (Emergency Drill)</li>
            <li>تعرف على أماكن طفايات الحريق و كيفية الاطفاء.</li>
            <li>سيتم توزيع كارت معلومات شخصي لكل شخص بالتعليمات</li>
          </ul>
        </td>
      </tr>

      <!-- 11. Work Permits -->
      <tr>
        <td class="col-en">
          <span class="section-num">11.</span> <span class="section-title-en">Work Permits:</span>
          <ul class="bullet-list">
            <li>Before starting any activities ensure that it have work permits<br/>(confined space – hot work – work at height – lifting – excavation )</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">11- تصاريح العمل:</span><br/>
          قبل الشروع في الاعمال التالية تأكد من وجود تصاريح عمل لها:
          <ul class="bullet-list-ar">
            <li>( الاماكن المغلقة ، الاعمال الساخنة ، العمل على ارتفاعات ، عمليات الرفع ، الكهرباء ، الحفر )</li>
          </ul>
        </td>
      </tr>

      <!-- 12. Bad Behaviors -->
      <tr>
        <td class="col-en">
          <span class="section-num">12</span> <span class="section-title-en">Some of Bad Behaviors Have to Prevent It:</span>
          <ul class="bullet-list">
            <li>Strictly prohibited drinking alcohol or any kind of drugs.</li>
            <li>Prohibited to sleep in work site.</li>
            <li>Prohibited use heaters in work sites or in cupboard of cloth keeping.</li>
          </ul>
        </td>
        <td class="col-ar">
          <span class="section-num">12- بعض العادات السيئة التي يجب منعها في الموقع:</span>
          <ul class="bullet-list-ar">
            <li>ممنوع استخدام المشروبات الكحولية أو اقراص مخدرة بجميع انواعها</li>
            <li>ممنوع النوم في اماكن العمل.</li>
            <li>ممنوع استخدام سخانات التدفئة في أماكن العمل أو في الدواليب الخاصة بحفظ الملابس.</li>
          </ul>
        </td>
      </tr>

      <!-- 13. Archeology -->
      <tr>
        <td class="col-en">
          <span class="section-num">13.</span> <span class="section-title-en">Archeology.</span><br/>
          In case of appearing archeology during excavation, immediacy contact your supervisor.
        </td>
        <td class="col-ar">
          <span class="section-num">13- الآثار.</span><br/>
          في حالة ظهور على مواد اثرية اثناء الحفر في الموقع، يتم فوراً التواصل مع المسئول عنك للتبليغ.
        </td>
      </tr>

      <!-- 14. Biodiversity -->
      <tr>
        <td class="col-en">
          <span class="section-num">14.</span> <span class="section-title-en">Biodiversity.</span><br/>
          The animals that appear on the site must be preserved and not exposed to them, and the specialists must be informed of this, and flocks of migratory birds should not be intercepted
        </td>
        <td class="col-ar">
          <span class="section-num">14- التنوع البيولوجي</span><br/>
          يجب الحفاظ على الحيوانات التي تظهر في الموقع و عدم التعرض لها وإبلاغ المختصين بذلك و عدم اعتراض لاسراب الطيور المهاجرة
        </td>
      </tr>
    </table>

    <!-- Acknowledgment -->
    <table class="acknowledgment" style="width:100%; border-collapse:collapse;">
      <tr>
        <td class="ack-en">
          I Approved That I Got The HSE Instructions, and I'll commitment for it and I accept any penalty in case of violations causes any injuring for anyone.
        </td>
        <td class="ack-ar">
          أقر بأنني اطلعت على تعليمات سلامة الموقع، وأنني سوف ألتزم بكل منهود و أقبل أي عقوبات في حال وجود أي انتهاكات وقعت مني أو أنت الي اصابتي أو إصابة من حولي
        </td>
      </tr>
    </table>

    <!-- Form Fields -->
    <table class="form-fields-table">
      <tr>
        <td class="field-label">Name:</td>
        <td class="field-value" colspan="1">${person.name}</td>
        <td class="field-label-ar">الاسم :</td>
      </tr>
      <tr>
        <td class="field-label">ID Number:</td>
        <td colspan="1">
          <div class="id-boxes">${idBoxesHtml}</div>
        </td>
        <td class="field-label-ar">الرقم القومي:</td>
      </tr>
      <tr>
        <td class="field-label">Position:</td>
        <td class="field-value" colspan="1">${person.position}</td>
        <td class="field-label-ar">الوظيفة :</td>
      </tr>
      <tr>
        <td class="field-label">Company:</td>
        <td class="field-value" colspan="1">${person.company || ''}</td>
        <td class="field-label-ar">الشركة :</td>
      </tr>
      <tr>
        <td class="field-label">Date:</td>
        <td class="field-value" colspan="1">${person.date}</td>
        <td class="field-label-ar">التاريخ:</td>
      </tr>
      <tr>
        <td class="field-label">Address:</td>
        <td class="field-value-empty" colspan="1"></td>
        <td class="field-label-ar">العنوان:</td>
      </tr>
      <tr>
        <td class="field-label">Signature:</td>
        <td class="field-value-empty" colspan="1"></td>
        <td class="field-label-ar">التوقيع</td>
      </tr>
    </table>

    <!-- Received Section -->
    <table class="form-fields-table" style="margin-top:0; border-top:2px solid #000;">
      <tr>
        <td class="field-label">Received by</td>
        <td class="field-value-empty" colspan="1"></td>
        <td class="field-label-ar">استلمت من قبل</td>
      </tr>
      <tr>
        <td class="field-label">HSE Department:</td>
        <td class="field-value-empty" colspan="1"></td>
        <td class="field-label-ar">قسم السلامة و الصحة المهنية:</td>
      </tr>
    </table>

    <div class="page-footer">2 | Page</div>
  </div>
  `;
}
