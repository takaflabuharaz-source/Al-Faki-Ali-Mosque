function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let data = {};
  
  try {
    if (e && e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    }
  } catch (err) {
    data = (e && e.parameter) ? e.parameter : {};
  }

  if (!data || !data.action) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "API أسرة مجمع الفكي علي الشيخ حماد يعمل بنجاح!"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const result = processForm(data);
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// دالة لتحديد الشهر الحالي والتعامل مع الإغلاق المباشر بنهاية يوم 30 وبداية يوم 1
function getCurrentDonationMonth() {
  const now = new Date();
  const day = now.getDate();
  let monthIndex = now.getMonth(); // 0-based index
  
  // الأشهُر العربية
  const months = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];

  // إذا كان اليوم بعد يوم 30 (أو يوم 31)، يتحول التسجيل تلقائياً إلى الشهر الجديد
  if (day > 30) {
    monthIndex = (monthIndex + 1) % 12;
  }

  return months[monthIndex];
}

function processForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("المستخدمين") || ss.insertSheet("المستخدمين");
  const donationsSheet = ss.getSheetByName("التبرعات") || ss.insertSheet("التبرعات");
  const expensesSheet = ss.getSheetByName("المصروفات") || ss.insertSheet("المصروفات");

  const action = data.action;
  const currentMonth = getCurrentDonationMonth();

  // 1. تسجيل حساب جديد والدخول المباشر
  if (action === "register") {
    const usersData = usersSheet.getDataRange().getValues();
    
    // التحقق من عدم تكرار رقم الهاتف أو الاسم
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString().trim() === data.phone.toString().trim()) {
        return { status: "error", message: "رقم الهاتف مسجل بالفعل!" };
      }
      if (usersData[i][0].toString().trim() === data.name.toString().trim()) {
        return { status: "error", message: "الاسم مسجل من قبل، يرجى اختيار اسم آخر!" };
      }
    }
    
    // إضافة المستخدم في شيت المستخدمين
    usersSheet.appendRow([data.name, data.phone, data.password, "مستخدم"]);

    // إضافة المستخدم تلقائياً للكشف العام بمبلغ 10000 وحالة قيد الانتظار
    donationsSheet.appendRow([
      data.name,
      data.phone,
      currentMonth, 
      10000, 
      "", 
      "في الانتظار", 
      new Date()
    ]);

    return { 
      status: "success", 
      message: "تم إنشاء الحساب بنجاح وإضافتك للكشف العام!",
      userData: {
        userName: data.name,
        phone: data.phone,
        isAdmin: false,
        isPaid: false,
        currentAmount: 10000,
        currentMonth: currentMonth
      }
    };
  }

  // 2. تسجيل الدخول
  if (action === "login") {
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString().trim() === data.phone.toString().trim() && usersData[i][2].toString().trim() === data.password.toString().trim()) {
        const userName = usersData[i][0];
        const isAdmin = usersData[i][3] === "آدمن" || usersData[i][3] === "ادمن" || usersData[i][3] === "Admin";
        
        let isPaid = false;
        let currentAmount = 10000;
        const donationsData = donationsSheet.getDataRange().getValues();
        for (let j = 1; j < donationsData.length; j++) {
          if (donationsData[j][1].toString().trim() === data.phone.toString().trim()) {
            currentAmount = donationsData[j][3] || 10000;
            if (donationsData[j][5] === "تم الدفع") {
              isPaid = true;
            }
          }
        }

        return {
          status: "success",
          userName: userName,
          phone: data.phone,
          isAdmin: isAdmin,
          isPaid: isPaid,
          currentAmount: currentAmount,
          currentMonth: currentMonth
        };
      }
    }
    return { status: "error", message: "رقم الهاتف أو كلمة المرور غير صحيحة!" };
  }

  // 3. جلب قائمة التبرعات
  if (action === "getDonations") {
    const donationsData = donationsSheet.getDataRange().getValues();
    let list = [];
    for (let i = 1; i < donationsData.length; i++) {
      list.push({
        rowIndex: i + 1,
        name: donationsData[i][0],
        phone: donationsData[i][1],
        month: donationsData[i][2] || currentMonth,
        amount: donationsData[i][3] || 10000,
        receiptUrl: donationsData[i][4] || "",
        status: donationsData[i][5] || "في الانتظار"
      });
    }
    return list;
  }

  // 4. جلب قائمة المصروفات
  if (action === "getExpenses") {
    const expensesData = expensesSheet.getDataRange().getValues();
    let list = [];
    for (let i = 1; i < expensesData.length; i++) {
      list.push({
        rowIndex: i + 1,
        date: expensesData[i][0] ? Utilities.formatDate(new Date(expensesData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        title: expensesData[i][1] || "",
        amount: expensesData[i][2] || 0,
        phone: expensesData[i][3] || ""
      });
    }
    return list;
  }

  // 5. إضافة وتحديث تبرع (رفع إشعار التبرع)
  if (action === "addDonation" || action === "addDonations") {
    // التأكد التام من إرفاق الإشعار
    if (!data.receiptData) {
      return { status: "error", message: "إرفاق إشعار التبرع إجباري لإكمال العملية!" };
    }

    let receiptUrl = "";
    if (data.receiptData) {
      const folder = DriveApp.getRootFolder();
      const blob = Utilities.newBlob(Utilities.base64Decode(data.receiptData.split(',')[1]), "image/png", "receipt_" + data.phone + ".png");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      receiptUrl = file.getUrl();
    }

    const donationsData = donationsSheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < donationsData.length; i++) {
      if (donationsData[i][1].toString().trim() === data.phone.toString().trim()) {
        donationsSheet.getRange(i + 1, 3).setValue(data.month || currentMonth);
        donationsSheet.getRange(i + 1, 4).setValue(data.amount || 10000);
        if (receiptUrl) donationsSheet.getRange(i + 1, 5).setValue(receiptUrl);
        donationsSheet.getRange(i + 1, 6).setValue("تم الدفع");
        donationsSheet.getRange(i + 1, 7).setValue(new Date());
        updated = true;
        break;
      }
    }

    if (!updated) {
      donationsSheet.appendRow([
        data.userName, 
        data.phone, 
        data.month || currentMonth, 
        data.amount || 10000, 
        receiptUrl, 
        "تم الدفع", 
        new Date()
      ]);
    }

    const pdfUrl = generatePDFUrl(ss, "paid");
    return {
      status: "success",
      message: "تم تسجيل التبرع ورفع الإشعار بنجاح",
      pdfDownloadUrl: pdfUrl,
      targetPhone: data.phone
    };
  }

  // 6. تعديل سجل بواسطة الإدارة
  if (action === "editDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      if (data.name) donationsSheet.getRange(rowIndex, 1).setValue(data.name);
      if (data.amount !== undefined) donationsSheet.getRange(rowIndex, 4).setValue(data.amount);
      if (data.status) donationsSheet.getRange(rowIndex, 6).setValue(data.status);
      return { status: "success", message: "تم تعديل السجل بنجاح" };
    }
    return { status: "error", message: "رقم الصف غير صحيح" };
  }

  // 7. حذف سجل بواسطة الإدارة
  if (action === "deleteDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      donationsSheet.deleteRow(rowIndex);
      return { status: "success", message: "تم حذف السجل بنجاح" };
    }
    return { status: "error", message: "رقم الصف غير صحيح" };
  }

  // 8. إضافة مصروف (للإدارة فقط)
  if (action === "addExpense") {
    expensesSheet.appendRow([new Date(), data.title, data.amount, data.phone]);
    const pdfUrl = generatePDFUrl(ss, "expenses");
    return {
      status: "success",
      message: "تم تسجيل المصروف",
      pdfDownloadUrl: pdfUrl,
      targetPhone: data.phone
    };
  }

  // 9. توليد كشف PDF
  if (action === "generatePDF" || action === "generateExpensesPDF") {
    const type = action === "generateExpensesPDF" ? "expenses" : (data.pdfType || "all");
    const pdfUrl = generatePDFUrl(ss, type);
    return { status: "success", pdfDownloadUrl: pdfUrl };
  }

  return { status: "error", message: "طلب غير معروف" };
}

function generatePDFUrl(ss, type) {
  try {
    ss.setSharing(SpreadsheetApp.Access.ANYONE_WITH_LINK, SpreadsheetApp.Permission.VIEW);
  } catch(e) {}

  let sheetId = "";
  if (type === "expenses") {
    const sheet = ss.getSheetByName("المصروفات");
    if (sheet) sheetId = "&gid=" + sheet.getSheetId();
  } else {
    const sheet = ss.getSheetByName("التبرعات");
    if (sheet) sheetId = "&gid=" + sheet.getSheetId();
  }
  
  const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&size=letter&portrait=true&fitw=true&gridlines=true&printtitle=false&sheetnames=false&fzr=false' + sheetId;
  return url;
}
