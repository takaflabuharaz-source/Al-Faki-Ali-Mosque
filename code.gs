const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TEST_PHONE_NUMBER = "966502658851"; 

// 🔒 قائمة أرقام المشرفين للإدارة
const ADMIN_PHONES = ["502658851", "0502658851", "555555555"]; 

function doGet() {
  const htmlContent = `
  <!DOCTYPE html>
  <html dir="rtl" lang="ar">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تطبيق مسجد الفكي علي الشيخ حماد</title>
    
    <link rel="manifest" href="manifest.json">
    <meta name="theme-color" content="#1e5631">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">

    <style>
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 0; padding: 12px; direction: rtl; }
      .card { background: #fff; padding: 20px; border-radius: 12px; max-width: 650px; margin: 10px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
      h2, h3, h4 { color: #1e5631; text-align: center; margin-top: 5px; }
      input, select, button { width: 100%; padding: 10px; margin: 6px 0; border-radius: 6px; border: 1px solid #ccc; box-sizing: border-box; font-size: 14px; }
      button { background-color: #1e5631; color: white; border: none; font-weight: bold; cursor: pointer; transition: 0.3s; }
      button:hover { background-color: #143d21; }
      .btn-admin { background-color: #8e44ad !important; }
      .btn-download { background-color: #c0392b !important; }
      .btn-nav { width: 48%; display: inline-block; margin: 1%; }
      .hidden { display: none; }
      table { width: 100%; border-collapse: collapse; margin-top: 12px; }
      th, td { border: 1px solid #ddd; padding: 7px; text-align: center; font-size: 13px; }
      th { background-color: #1e5631; color: white; }
      .badge-paid { color: #27ae60; font-weight: bold; }
      .badge-wait { color: #d35400; font-weight: bold; }
      .note { font-size: 12px; color: #666; margin-bottom: 3px; display: block; }
      .thank-you-card { background: #e8f8f5; border: 1px solid #2ecc71; padding: 15px; border-radius: 8px; text-align: center; color: #117a65; margin: 10px 0; }
      
      /* 🟠 رسالة التنبيه البرتقالية في أسفل الشاشة */
      .orange-notice { 
        background-color: #fff3cd; 
        color: #d35400; 
        border: 1px solid #f39c12; 
        padding: 12px; 
        border-radius: 8px; 
        text-align: center; 
        font-weight: bold; 
        font-size: 13px; 
        line-height: 1.6;
        margin-top: 20px;
      }

      .summary-box { background: #f8f9f9; padding: 10px; border-radius: 6px; font-weight: bold; margin-top: 10px; border: 1px solid #ebedef; }
      .date-filter-box { background: #ffffff; padding: 8px; border: 1px solid #dcdde1; border-radius: 6px; margin: 8px 0; }
    </style>
  </head>
  <body>

    <!-- 1. شاشة الدخول والتسجيل -->
    <div id="authSection" class="card">
      <h2>مسجد الفكي علي الشيخ حماد</h2>

      <div id="loginForm">
        <h3>تسجيل الدخول</h3>
        <input type="tel" id="loginPhone" placeholder="رقم الهاتف (بدون 0 مثلاً: 912345678)">
        <input type="password" id="loginPassword" placeholder="كلمة المرور">
        <button onclick="login()">دخول</button>
        <p style="text-align:center; font-size: 14px;">ليس لديك حساب؟ <a href="#" onclick="toggleAuth('register')">إنشاء حساب جديد</a></p>
      </div>

      <div id="registerForm" class="hidden">
        <h3>إنشاء حساب جديد</h3>
        <input type="text" id="regName" placeholder="الاسم الكامل">
        <input type="tel" id="regPhone" placeholder="رقم الهاتف (بدون 0 مثلاً: 912345678)">
        <input type="password" id="regPassword" placeholder="كلمة المرور">
        <button onclick="register()">تسجيل الحساب</button>
        <p style="text-align:center; font-size: 14px;">لديك حساب بالفعل؟ <a href="#" onclick="toggleAuth('login')">تسجيل الدخول</a></p>
      </div>

      <!-- 🟠 التنبيه أسفل شاشة التسجيل -->
      <div class="orange-notice">
        ⚠️ تنبيه: سيتم إغلاق جميع التبرعات بنهاية يوم 30 من كل شهر، بحيث أنه لن يتم استلام أي تبرع بعد يوم 30 للشهر المنتهي، وسيتم حساب المبلغ للشهر الجديد. وشكراً لتفهمكم.
      </div>
    </div>

    <!-- 2. الشاشة الرئيسية للمستخدم -->
    <div id="appSection" class="card hidden">
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <h3>مرحباً، <span id="userDisplayName"></span></h3>
        <button onclick="logout()" style="width: auto; padding: 5px 12px; background:#7f8c8d;">خروج</button>
      </div>

      <!-- رسالة الشكر والترحيب -->
      <div id="thankYouMessage" class="thank-you-card hidden">
        🎉 **جزاكم الله خيراً وتقبل منا ومنكم صالح الأعمال**<br>
        نشكركم على مساهمتكم الكريمة ودعمكم لمسجد الفكي علي الشيخ حماد.
      </div>

      <!-- تنبيه إغلاق الشهر -->
      <div id="monthClosedNotice" class="hidden" style="background:#fadbd8; color:#78281f; padding:12px; border-radius:6px; text-align:center; font-weight:bold; margin-bottom: 10px;">
        🚫 تم إغلاق استقبال التبرعات لهذا الشهر (اعتباراً من يوم 30). وسيتم فتح باب التبرع للشهر الجديد قريباً.
      </div>

      <!-- نموذج التبرع -->
      <div id="donationFormArea" style="border-bottom: 2px solid #eee; padding-bottom: 15px; margin-top: 10px;">
        <h4>إرسال / تأكيد التبرع</h4>
        
        <label class="note">شهر التبرع:</label>
        <select id="donationMonth">
          <option value="يناير">يناير</option><option value="فبراير">فبراير</option><option value="مارس">مارس</option>
          <option value="أبريل">أبريل</option><option value="مايو">مايو</option><option value="يونيو">يونيو</option>
          <option value="يوليو">يوليو</option><option value="أغسطس" selected>أغسطس</option><option value="سبتمبر">سبتمبر</option>
          <option value="أكتوبر">أكتوبر</option><option value="نوفمبر">نوفمبر</option><option value="ديسمبر">ديسمبر</option>
        </select>

        <label class="note">المبلغ بالجنيه السوداني (SDG):</label>
        <input type="number" id="donationAmount" value="10000">
        
        <label class="note">إرفاق صورة الإشعار (لتغيير الحالة إلى تم الدفع):</label>
        <input type="file" id="receiptFile" accept="image/*">
        
        <button onclick="sendDonation()" style="background-color: #27ae60;">تأكيد الدفع وإرسال التقرير للواتساب 💬</button>
      </div>

      <!-- زر وقسم الإدارة للمصروفات -->
      <div id="adminArea" class="hidden" style="background: #f4ecf7; padding: 12px; border-radius: 8px; margin: 15px 0;">
        <h4 style="color:#8e44ad;">💼 لوحة تسجيل وإدارة المصروفات (للإدارة)</h4>
        
        <!-- تسجيل مصروف جديد -->
        <input type="text" id="expenseTitle" placeholder="بيان المصروف (مثلاً: صيانات، كهرباء...)">
        <input type="number" id="expenseAmount" placeholder="المبلغ بالجنيه السوداني (SDG)">
        <button onclick="addExpense()" class="btn-admin">تسجيل وإرسال تقرير المصروفات للواتساب 📄💬</button>
        
        <hr style="border: 0.5px solid #dcdde1; margin: 15px 0;">
        
        <!-- طباعة المصروفات بالتاريخ -->
        <h5 style="margin: 5px 0; color:#2c3e50; text-align:center;">📅 طباعة المصروفات خلال فترة زمنية</h5>
        <div class="date-filter-box">
          <label class="note">من تاريخ:</label>
          <input type="date" id="startDate">
          <label class="note">إلى تاريخ:</label>
          <input type="date" id="endDate">
          <button onclick="downloadExpensesPDF()" class="btn-download" style="margin-top:8px;">تنزيل PDF المصروفات للفترة المحسوبة 📄</button>
        </div>
      </div>

      <!-- أزرار القوائم -->
      <div style="margin-top: 15px;">
        <button class="btn-nav" onclick="showPaidList()" style="background:#2980b9;">المسددين فقط 🟢</button>
        <button class="btn-nav" onclick="showAllList()" style="background:#34495e;">الكشف الشامل 📋</button>
      </div>

      <!-- 1. عرض قائمة تم الدفع فقط -->
      <div id="paidSection" class="hidden">
        <h4>قائمة المسددين (تم الدفع)</h4>
        <button onclick="downloadPDF('paid')" class="btn-download">تنزيل PDF المسددين 📄</button>
        <table>
          <thead><tr><th>#</th><th>الاسم</th><th>الشهر</th><th>المبلغ (SDG)</th><th>الإشعار</th></tr></thead>
          <tbody id="paidTableBody"></tbody>
        </table>
        <div id="paidSummary" class="summary-box"></div>
      </div>

      <!-- 2. عرض الكشف الشامل -->
      <div id="allSection" class="hidden">
        <h4>الكشف الشامل لجميع المتبرعين</h4>
        <button onclick="downloadPDF('all')" class="btn-download">تنزيل PDF الكشف الشامل 📄</button>
        <table>
          <thead>
            <tr>
              <th>#</th><th>الاسم</th><th>المبلغ (SDG)</th><th>تم الدفع</th><th>في الانتظار</th><th>الإشعار</th>
            </tr>
          </thead>
          <tbody id="allTableBody"></tbody>
        </table>
        <div id="allSummary" class="summary-box"></div>
      </div>

      <!-- 🟠 التنبيه البرتقالي في أسفل الشاشة الرئيسية -->
      <div class="orange-notice">
        ⚠️ تنبيه: سيتم إغلاق جميع التبرعات بنهاية يوم 30 من كل شهر، بحيث أنه لن يتم استلام أي تبرع بعد يوم 30 للشهر المنتهي، وسيتم حساب المبلغ للشهر الجديد. وشكراً لتفهمكم.
      </div>

    </div>

    <script>
      let currentUser = null;

      window.onload = function() {
        const savedUser = localStorage.getItem("msg_user");
        if(savedUser) {
          currentUser = JSON.parse(savedUser);
          showApp();
        }
      };

      function toggleAuth(type) {
        if(type === 'register') {
          document.getElementById('loginForm').classList.add('hidden');
          document.getElementById('registerForm').classList.remove('hidden');
        } else {
          document.getElementById('registerForm').classList.add('hidden');
          document.getElementById('loginForm').classList.remove('hidden');
        }
      }

      function register() {
        const name = document.getElementById('regName').value;
        const phone = document.getElementById('regPhone').value;
        const password = document.getElementById('regPassword').value;

        if(!name || !phone || !password) return alert("يرجى ملء جميع الحقول");

        google.script.run.withSuccessHandler(res => {
          alert(res.message);
          if(res.status === 'success') toggleAuth('login');
        }).processForm({ action: 'register', name, phone, password });
      }

      function login() {
        const phone = document.getElementById('loginPhone').value;
        const password = document.getElementById('loginPassword').value;

        google.script.run.withSuccessHandler(res => {
          if(res.status === 'success') {
            currentUser = res;
            localStorage.setItem("msg_user", JSON.stringify(res));
            showApp();
          } else {
            alert(res.message);
          }
        }).processForm({ action: 'login', phone, password });
      }

      function logout() {
        localStorage.removeItem("msg_user");
        location.reload();
      }

      function showApp() {
        document.getElementById('userDisplayName').innerText = currentUser.userName;
        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appSection').classList.remove('hidden');
        
        if(currentUser.currentAmount) {
          document.getElementById('donationAmount').value = currentUser.currentAmount;
        }

        if(currentUser.isAdmin) {
          document.getElementById('adminArea').classList.remove('hidden');
        }

        if(currentUser.isPaid) {
          document.getElementById('thankYouMessage').classList.remove('hidden');
        }

        const today = new Date().getDate();
        if(today >= 30) {
          document.getElementById('monthClosedNotice').classList.remove('hidden');
          document.getElementById('donationFormArea').classList.add('hidden');
        }

        showPaidList();
      }

      function sendDonation() {
        const amount = document.getElementById('donationAmount').value;
        const month = document.getElementById('donationMonth').value;
        const fileInput = document.getElementById('receiptFile');

        if(!amount) return alert("يرجى إدخال المبلغ");
        if(fileInput.files.length === 0) return alert("يرجى إرفاق صورة الإشعار ليتم تغيير الحالة إلى (تم الدفع ✔)");

        const reader = new FileReader();
        reader.onload = function(e) {
          submitDonationData(amount, month, e.target.result);
        };
        reader.readAsDataURL(fileInput.files[0]);
      }

      function submitDonationData(amount, month, receiptData) {
        alert("جاري التحديث، تنزيل الـ PDF، والفتح في الواتساب...");
        google.script.run.withSuccessHandler(res => {
          if(res.status === 'success') {
            document.getElementById('receiptFile').value = '';
            document.getElementById('thankYouMessage').classList.remove('hidden');
            currentUser.isPaid = true;
            localStorage.setItem("msg_user", JSON.stringify(currentUser));

            const downloadLink = document.createElement('a');
            downloadLink.href = res.pdfDownloadUrl;
            downloadLink.download = "كشف_تبرعات_المسجد.pdf";
            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);

            const message = encodeURIComponent("السلام عليكم ورحمة الله، تم الدفع وإرفاق الإشعار لصالح مسجد الفكي علي الشيخ حماد\\n" +
              "المتبرع: " + currentUser.userName + "\\n" +
              "الشهر: " + month + "\\n" +
              "المبلغ: " + amount + " جنيه سوداني\\n" +
              "مرفق كشف التبرعات المحدث PDF.");

            setTimeout(() => {
              window.open("https://api.whatsapp.com/send?phone=" + res.targetPhone + "&text=" + message, '_blank');
            }, 1000);

            showPaidList();
          } else {
            alert(res.message);
          }
        }).processForm({
          action: 'addDonation',
          userName: currentUser.userName,
          phone: currentUser.phone,
          amount: amount,
          month: month,
          receiptData: receiptData
        });
      }

      function addExpense() {
        const title = document.getElementById('expenseTitle').value;
        const amount = document.getElementById('expenseAmount').value;
        if(!title || !amount) return alert("يرجى إدخال بيان وقيمة المصروف");

        alert("جاري تسجيل المصروف وإعداد التقرير للواتساب...");
        google.script.run.withSuccessHandler(res => {
          if(res.status === 'success') {
            document.getElementById('expenseTitle').value = '';
            document.getElementById('expenseAmount').value = '';

            const downloadLink = document.createElement('a');
            downloadLink.href = res.pdfDownloadUrl;
            downloadLink.download = "كشف_مصروفات_المسجد.pdf";
            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);

            const message = encodeURIComponent("إشعار مصروفات جديد - مسجد الفكي علي الشيخ حماد\\n" +
              "البيان: " + title + "\\n" +
              "المبلغ: " + amount + " جنيه سوداني\\n" +
              "مرفق تقرير المصروفات المحدث PDF.");

            setTimeout(() => {
              window.open("https://api.whatsapp.com/send?phone=" + res.targetPhone + "&text=" + message, '_blank');
            }, 1000);
          } else {
            alert(res.message);
          }
        }).processForm({ action: 'addExpense', title, amount, phone: currentUser.phone });
      }

      function downloadExpensesPDF() {
        const startDate = document.getElementById('startDate').value;
        const endDate = document.getElementById('endDate').value;

        alert("جاري تجهيز وتنزيل كشف المصروفات للفترة المحددة...");
        google.script.run.withSuccessHandler(res => {
          if(res.status === 'error') {
            alert(res.message);
          } else {
            window.open(res.pdfDownloadUrl, '_blank');
          }
        }).processForm({ action: 'generateExpensesPDF', startDate: startDate, endDate: endDate });
      }

      function showPaidList() {
        document.getElementById('allSection').classList.add('hidden');
        document.getElementById('paidSection').classList.remove('hidden');

        google.script.run.withSuccessHandler(data => {
          const tbody = document.getElementById('paidTableBody');
          tbody.innerHTML = '';
          let sum = 0, count = 0;

          data.forEach(item => {
            if(item.status === 'تم الدفع') {
              count++;
              sum += parseFloat(item.amount);
              const receiptBtn = item.receiptUrl ? \`<a href="\${item.receiptUrl}" target="_blank"><button style="padding: 2px 6px; font-size: 11px;">معاينة</button></a>\` : '-';
              tbody.innerHTML += \`<tr><td>\${count}</td><td>\${item.name}</td><td>\${item.month || '-'}</td><td>\${item.amount} SDG</td><td>\${receiptBtn}</td></tr>\`;
            }
          });
          document.getElementById('paidSummary').innerText = \`عدد المسددين: \${count} شخص | إجمالي التبرعات: \${sum} جنيه سوداني\`;
        }).processForm({ action: 'getDonations' });
      }

      function showAllList() {
        document.getElementById('paidSection').classList.add('hidden');
        document.getElementById('allSection').classList.remove('hidden');

        google.script.run.withSuccessHandler(data => {
          const tbody = document.getElementById('allTableBody');
          tbody.innerHTML = '';
          let paidSum = 0, waitSum = 0;
          let paidCount = 0, waitCount = 0;

          data.forEach((item, index) => {
            let paidCol = '-', waitCol = '-';
            if(item.status === 'تم الدفع') {
              paidCol = '<span class="badge-paid">✔ تم الدفع</span>';
              paidSum += parseFloat(item.amount);
              paidCount++;
            } else {
              waitCol = '<span class="badge-wait">⏳ في الانتظار</span>';
              waitSum += parseFloat(item.amount);
              waitCount++;
            }

            const receiptBtn = item.receiptUrl ? \`<a href="\${item.receiptUrl}" target="_blank"><button style="padding: 2px 6px; font-size: 11px;">معاينة</button></a>\` : '-';
            tbody.innerHTML += \`<tr><td>\${index + 1}</td><td>\${item.name}</td><td>\${item.amount} SDG</td><td>\${paidCol}</td><td>\${waitCol}</td><td>\${receiptBtn}</td></tr>\`;
          });

          document.getElementById('allSummary').innerHTML = \`
            المسددين: \${paidCount} (\${paidSum} SDG) | في الانتظار: \${waitCount} (\${waitSum} SDG)<br>
            <strong>المجموع الكلي: \${paidCount + waitCount} شخص | \${paidSum + waitSum} جنيه سوداني</strong>
          \`;
        }).processForm({ action: 'getDonations' });
      }

      function downloadPDF(type) {
        alert("جاري تجهيز وتنزيل كشف الـ PDF...");
        google.script.run.withSuccessHandler(res => {
          window.open(res.pdfDownloadUrl, '_blank');
        }).processForm({ action: 'generatePDF', pdfType: type });
      }
    </script>
  </body>
  </html>
  `;

  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle("تطبيق مسجد الفكي علي الشيخ حماد")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processForm(formObject) {
  const action = formObject.action;
  if (action === "register") return registerUser(formObject.name, formObject.phone, formObject.password);
  if (action === "login") return loginUser(formObject.phone, formObject.password);
  if (action === "addDonation") return addDonation(formObject);
  if (action === "getDonations") return getDonations();
  if (action === "addExpense") return addExpense(formObject);
  if (action === "generateExpensesPDF") return generateExpensesPDF(formObject.startDate, formObject.endDate);
  if (action === "generatePDF") return generatePDF(formObject.pdfType);
}

function registerUser(name, phone, password) {
  const usersSheet = getOrCreateSheet("المستخدمين", ["الاسم", "رقم الهاتف", "كلمة المرور", "تاريخ التسجيل"]);
  const uData = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][1].toString() === phone.toString()) {
      return { status: "error", message: "هذا الرقم مسجل مسبقاً، يرجى تسجيل الدخول بنفس الرقم." };
    }
  }
  
  usersSheet.appendRow([name, phone, password, new Date()]);

  const donationsSheet = getOrCreateSheet("التبرعات", ["الاسم", "رقم الهاتف", "شهر التبرع", "المبلغ", "رابط الإشعار", "الحالة", "التاريخ"]);
  donationsSheet.appendRow([name, phone, "أغسطس", 10000, "", "في الانتظار", new Date()]);

  return { status: "success", message: "تم تسجيل الحساب وإدراج اسمك في الكشف بمبلغ 10000 جنيه سوداني!" };
}

function loginUser(phone, password) {
  const usersSheet = getOrCreateSheet("المستخدمين", ["الاسم", "رقم الهاتف", "كلمة المرور", "تاريخ التسجيل"]);
  const uData = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < uData.length; i++) {
    if (uData[i][1].toString() === phone.toString() && uData[i][2].toString() === password.toString()) {
      let currentAmount = 10000;
      let isPaid = false;
      const donationsSheet = getOrCreateSheet("التبرعات", ["الاسم", "رقم الهاتف", "شهر التبرع", "المبلغ", "رابط الإشعار", "الحالة", "التاريخ"]);
      const dData = donationsSheet.getDataRange().getValues();
      for (let j = 1; j < dData.length; j++) {
        if (dData[j][1].toString() === phone.toString()) {
          currentAmount = dData[j][3];
          if(dData[j][5] === "تم الدفع") isPaid = true;
          break;
        }
      }

      const isAdmin = ADMIN_PHONES.includes(phone.toString());

      return { 
        status: "success", 
        userName: uData[i][0], 
        phone: uData[i][1], 
        currentAmount: currentAmount,
        isPaid: isPaid,
        isAdmin: isAdmin
      };
    }
  }
  return { status: "error", message: "رقم الهاتف أو كلمة المرور غير صحيحة." };
}

function addDonation(data) {
  try {
    const sheet = getOrCreateSheet("التبرعات", ["الاسم", "رقم الهاتف", "شهر التبرع", "المبلغ", "رابط الإشعار", "الحالة", "التاريخ"]);
    const dData = sheet.getDataRange().getValues();
    let fileUrl = "";
    
    if (data.receiptData) {
      const folder = DriveApp.getRootFolder();
      const contentType = data.receiptData.substring(5, data.receiptData.indexOf(';'));
      const bytes = Utilities.base64Decode(data.receiptData.substr(data.receiptData.indexOf(',') + 1));
      const blob = Utilities.newBlob(bytes, contentType, "إشعار_" + data.userName + "_" + Date.now());
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    for (let i = 1; i < dData.length; i++) {
      if (dData[i][1].toString() === data.phone.toString()) {
        sheet.getRange(i + 1, 3).setValue(data.month);
        sheet.getRange(i + 1, 4).setValue(parseFloat(data.amount));
        if (fileUrl) sheet.getRange(i + 1, 5).setValue(fileUrl);
        sheet.getRange(i + 1, 6).setValue("تم الدفع");
        sheet.getRange(i + 1, 7).setValue(new Date());
        break;
      }
    }

    const pdfRes = generatePDF('all');

    return { 
      status: "success", 
      message: "تم تحديث التبرع بنجاح!",
      pdfDownloadUrl: pdfRes.pdfDownloadUrl,
      targetPhone: TEST_PHONE_NUMBER
    };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function addExpense(data) {
  if(!ADMIN_PHONES.includes(data.phone.toString())) {
    return { status: "error", message: "عفواً، لا تملك صلاحية مدير لتسجيل المصروفات." };
  }
  const sheet = getOrCreateSheet("المصروفات", ["البيان", "المبلغ", "التاريخ"]);
  sheet.appendRow([data.title, parseFloat(data.amount), new Date()]);

  const pdfRes = generateExpensesPDF("", "");

  return { 
    status: "success", 
    message: "تم تسجيل المصروف بنجاح!",
    pdfDownloadUrl: pdfRes.pdfDownloadUrl,
    targetPhone: TEST_PHONE_NUMBER
  };
}

function getDonations() {
  const sheet = getOrCreateSheet("التبرعات", ["الاسم", "رقم الهاتف", "شهر التبرع", "المبلغ", "رابط الإشعار", "الحالة", "التاريخ"]);
  const data = sheet.getDataRange().getValues();
  let donations = [];
  
  for (let i = 1; i < data.length; i++) {
    donations.push({
      id: i,
      name: data[i][0],
      month: data[i][2],
      amount: data[i][3],
      receiptUrl: data[i][4],
      status: data[i][5]
    });
  }
  return donations;
}

// إنشاء تقرير مصروفات PDF مع التصفية بالتاريخ
function generateExpensesPDF(startDateStr, endDateStr) {
  const sheet = getOrCreateSheet("المصروفات", ["البيان", "المبلغ", "التاريخ"]);
  const data = sheet.getDataRange().getValues();
  
  let totalExpenses = 0;
  let tableRows = "";
  let rowCounter = 0;

  const startFilter = startDateStr ? new Date(startDateStr) : null;
  const endFilter = endDateStr ? new Date(endDateStr) : null;
  if(endFilter) endFilter.setHours(23, 59, 59, 999);

  for (let i = 1; i < data.length; i++) {
    const expDate = data[i][2] ? new Date(data[i][2]) : null;
    
    // فلترة المصروفات حسب التواريخ المحددة
    if (startFilter && expDate && expDate < startFilter) continue;
    if (endFilter && expDate && expDate > endFilter) continue;

    const amount = parseFloat(data[i][1]) || 0;
    totalExpenses += amount;
    rowCounter++;

    const formattedDate = expDate ? expDate.toLocaleDateString('ar-EG') : '-';

    tableRows += `
      <tr>
        <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${rowCounter}</td>
        <td style="border: 1px solid #ccc; padding: 6px; text-align: right;">${data[i][0]}</td>
        <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${amount} SDG</td>
        <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${formattedDate}</td>
      </tr>`;
  }

  if (rowCounter === 0) {
    return { status: "error", message: "لا توجد مصروفات مسجلة في الفترة الزمنية المحددة." };
  }

  let periodTitle = "كشف جميع المصروفات المسجلة";
  if (startDateStr && endDateStr) {
    periodTitle = `كشف المصروفات للفترة من (${startDateStr}) إلى (${endDateStr})`;
  } else if (startDateStr) {
    periodTitle = `كشف المصروفات اعتباراً من تاريخ (${startDateStr})`;
  } else if (endDateStr) {
    periodTitle = `كشف المصروفات حتى تاريخ (${endDateStr})`;
  }

  const htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <style>
        body { font-family: 'Calibri', sans-serif; font-size: 12pt; margin: 20px; }
        h2 { text-align: center; font-size: 14pt; margin-bottom: 5px; }
        p { text-align: center; margin-top: 0; font-size: 11pt; color: #444; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 12pt; }
        th { border: 1px solid #000; padding: 8px; background-color: #f2f2f2; text-align: center; }
        .footer-row { font-weight: bold; background-color: #e6e6e6; text-align: center; }
      </style>
    </head>
    <body>
      <h2>مسجد الفكي علي الشيخ حماد</h2>
      <p><strong>${periodTitle}</strong></p>
      <p>تاريخ التقرير: ${new Date().toLocaleDateString('ar-EG')}</p>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>بيان المصروف</th>
            <th>المبلغ (جنيه سوداني)</th>
            <th>التاريخ</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
          <tr class="footer-row">
            <td colspan="2">إجمالي المصروفات الكلية</td>
            <td colspan="2">${totalExpenses} SDG</td>
          </tr>
        </tbody>
      </table>
    </body>
    </html>`;

  const blob = Utilities.newBlob(htmlContent, 'text/html', 'كشف_مصروفات_المسجد.html');
  const pdfFile = DriveApp.createFile(blob.getAs('application/pdf')).setName("كشف_مصروفات_المسجد.pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { 
    status: "success",
    pdfUrl: pdfFile.getUrl(),
    pdfDownloadUrl: "https://drive.google.com/uc?export=download&id=" + pdfFile.getId()
  };
}

// إنشاء تقرير الـ PDF مع إضافة التنبيه البرتقالي بأسفل التقرير المطبوع
function generatePDF(pdfType) {
  const sheet = getOrCreateSheet("التبرعات", ["الاسم", "رقم الهاتف", "شهر التبرع", "المبلغ", "رابط الإشعار", "الحالة", "التاريخ"]);
  const data = sheet.getDataRange().getValues();
  
  let paidSum = 0, waitSum = 0;
  let paidCount = 0, waitCount = 0;
  let tableRows = "";

  for (let i = 1; i < data.length; i++) {
    const isPaid = data[i][5] === "تم الدفع";
    const amount = parseFloat(data[i][3]) || 0;

    if (pdfType === 'paid' && !isPaid) continue;

    if (isPaid) { paidCount++; paidSum += amount; } 
    else { waitCount++; waitSum += amount; }

    if (pdfType === 'paid') {
      tableRows += `
        <tr>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${paidCount}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: right;">${data[i][0]}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${data[i][2] || '-'}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${amount} SDG</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">تم الدفع ✔</td>
        </tr>`;
    } else {
      const paidCol = isPaid ? "تم الدفع ✔" : "-";
      const waitCol = !isPaid ? "في الانتظار ⏳" : "-";
      tableRows += `
        <tr>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${i}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: right;">${data[i][0]}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${amount} SDG</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${paidCol}</td>
          <td style="border: 1px solid #ccc; padding: 6px; text-align: center;">${waitCol}</td>
        </tr>`;
    }
  }

  let tableHeader = pdfType === 'paid' ? 
    `<tr><th>#</th><th>اسم المتبرع</th><th>شهر التبرع</th><th>المبلغ</th><th>الحالة</th></tr>` :
    `<tr><th>#</th><th>اسم المتبرع</th><th>المبلغ</th><th>تم الدفع</th><th>في الانتظار</th></tr>`;

  let footerRow = pdfType === 'paid' ?
    `<tr class="footer-row"><td colspan="3">إجمالي المسددين (${paidCount} شخص)</td><td colspan="2">${paidSum} SDG</td></tr>` :
    `<tr class="footer-row"><td colspan="2">المجموع الإجمالي (${paidCount + waitCount} شخص)</td><td>${paidSum + waitSum} SDG</td><td>مدفوع: ${paidSum}</td><td>معلق: ${waitSum}</td></tr>`;

  const htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <style>
        body { font-family: 'Calibri', sans-serif; font-size: 12pt; margin: 20px; }
        h2 { text-align: center; font-size: 14pt; margin-bottom: 5px; }
        p { text-align: center; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 12pt; }
        th { border: 1px solid #000; padding: 8px; background-color: #f2f2f2; text-align: center; }
        .footer-row { font-weight: bold; background-color: #e6e6e6; text-align: center; }
        
        /* 🟠 طباعة التنبيه في أسفل صفحة الـ PDF */
        .pdf-notice {
          margin-top: 25px;
          padding: 10px;
          border: 1px solid #f39c12;
          background-color: #fff3cd;
          color: #d35400;
          text-align: center;
          font-weight: bold;
          font-size: 10.5pt;
          border-radius: 5px;
        }
      </style>
    </head>
    <body>
      <h2>كشف تبرعات مسجد الفكي علي الشيخ حماد (${pdfType === 'paid' ? 'المسددين فقط' : 'الكشف الشامل'})</h2>
      <p>تاريخ التقرير: ${new Date().toLocaleDateString('ar-EG')}</p>
      <table>
        <thead>${tableHeader}</thead>
        <tbody>
          ${tableRows}
          ${footerRow}
        </tbody>
      </table>

      <!-- التنبيه البرتقالي أسفل الـ PDF -->
      <div class="pdf-notice">
        ⚠️ تنبيه: سيتم إغلاق جميع التبرعات بنهاية يوم 30 من كل شهر، بحيث أنه لن يتم استلام أي تبرع بعد يوم 30 للشهر المنتهي، وسيتم حساب المبلغ للشهر الجديد. وشكراً لتفهمكم.
      </div>
    </body>
    </html>`;

  const blob = Utilities.newBlob(htmlContent, 'text/html', 'كشف_تبرعات_المسجد.html');
  const pdfFile = DriveApp.createFile(blob.getAs('application/pdf')).setName("كشف_تبرعات_المسجد.pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { 
    pdfUrl: pdfFile.getUrl(),
    pdfDownloadUrl: "https://drive.google.com/uc?export=download&id=" + pdfFile.getId()
  };
}

function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  return sheet;
}
