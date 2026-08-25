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
      message: "API أسرة مجمع الفكي يعمل بنجاح!"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const result = processForm(data);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// تحديد الشهر الحالي والتعامل مع التحول التلقائي للشهر الجديد
function getCurrentDonationMonth() {
  const now = new Date();
  const day = now.getDate();
  let monthIndex = now.getMonth();
  
  const months = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];

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

  // فحص وإعادة ضبط حالة الشهر الجديد تلقائياً
  checkAndResetMonthlyStatus(donationsSheet, currentMonth);

  // 1. تسجيل حساب جديد
  if (action === "register") {
    const usersData = usersSheet.getDataRange().getValues();
    const cleanPhone = data.phone.toString().replace(/^0+/, '').trim();
    
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString().trim() === cleanPhone) {
        return { status: "error", message: "رقم الهاتف مسجل بالفعل!" };
      }
      if (usersData[i][0].toString().trim() === data.name.toString().trim()) {
        return { status: "error", message: "الاسم مسجل من قبل، يرجى استخدام اسم مختلف!" };
      }
    }
    
    usersSheet.appendRow([data.name, cleanPhone, data.password, "مستخدم"]);

    // إضافة تلقائية لكشف التبرعات بمبلغ افتراضي 10000
    donationsSheet.appendRow([
      data.name,
      cleanPhone,
      currentMonth, 
      10000, 
      "", 
      "في الانتظار", 
      new Date()
    ]);

    return { 
      status: "success", 
      message: "تم إنشاء الحساب وإضافتك للكشف العام بنجاح!",
      userData: {
        userName: data.name,
        phone: cleanPhone,
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
    const cleanPhone = data.phone.toString().replace(/^0+/, '').trim();

    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString().trim() === cleanPhone && usersData[i][2].toString().trim() === data.password.toString().trim()) {
        const userName = usersData[i][0];
        const isAdmin = usersData[i][3] === "آدمن" || usersData[i][3] === "ادمن" || usersData[i][3] === "Admin";
        
        let isPaid = false;
        let currentAmount = 10000;
        const donationsData = donationsSheet.getDataRange().getValues();
        for (let j = 1; j < donationsData.length; j++) {
          if (donationsData[j][1].toString().trim() === cleanPhone) {
            currentAmount = donationsData[j][3] || 10000;
            if (donationsData[j][5] === "تم الدفع") {
              isPaid = true;
            }
          }
        }

        return {
          status: "success",
          userName: userName,
          phone: cleanPhone,
          isAdmin: isAdmin,
          isPaid: isPaid,
          currentAmount: currentAmount,
          currentMonth: currentMonth
        };
      }
    }
    return { status: "error", message: "رقم الهاتف أو كلمة المرور غير صحيحة!" };
  }

  // 3. جلب التبرعات والإحصائيات
  if (action === "getDonations") {
    const donationsData = donationsSheet.getDataRange().getValues();
    let list = [];
    let paidCount = 0;
    let unpaidCount = 0;
    let totalPaidAmount = 0;

    for (let i = 1; i < donationsData.length; i++) {
      const amount = Number(donationsData[i][3]) || 10000;
      const status = donationsData[i][5] || "في الانتظار";

      if (status === "تم الدفع") {
        paidCount++;
        totalPaidAmount += amount;
      } else {
        unpaidCount++;
      }

      list.push({
        rowIndex: i + 1,
        name: donationsData[i][0],
        phone: donationsData[i][1],
        month: donationsData[i][2] || currentMonth,
        amount: amount,
        receiptUrl: donationsData[i][4] || "",
        status: status
      });
    }

    return {
      donations: list,
      stats: {
        paidCount: paidCount,
        unpaidCount: unpaidCount,
        totalPaidAmount: totalPaidAmount
      }
    };
  }

  // 4. رفع وتحديث الإشعار والمبلغ
  if (action === "addDonation") {
    if (!data.receiptData) {
      return { status: "error", message: "إرفاق إشعار التبرع إجباري!" };
    }

    let receiptUrl = "";
    if (data.receiptData) {
      const folder = DriveApp.getRootFolder();
      const blob = Utilities.newBlob(Utilities.base64Decode(data.receiptData.split(',')[1]), "image/png", "receipt_" + data.phone + ".png");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      receiptUrl = file.getUrl();
    }

    const cleanPhone = data.phone.toString().replace(/^0+/, '').trim();
    const donationsData = donationsSheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < donationsData.length; i++) {
      if (donationsData[i][1].toString().trim() === cleanPhone) {
        donationsSheet.getRange(i + 1, 3).setValue(data.month || currentMonth);
        donationsSheet.getRange(i + 1, 4).setValue(data.amount || 10000); // تحديث بالمبلغ الجديد
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
        cleanPhone, 
        data.month || currentMonth, 
        data.amount || 10000, 
        receiptUrl, 
        "تم الدفع", 
        new Date()
      ]);
    }

    return { status: "success", message: "تم تسجيل وتحديث التبرع بنجاح!" };
  }

  // 5. جلب المصروفات
  if (action === "getExpenses") {
    const expensesData = expensesSheet.getDataRange().getValues();
    let list = [];
    for (let i = 1; i < expensesData.length; i++) {
      list.push({
        rowIndex: i + 1,
        date: expensesData[i][0] ? Utilities.formatDate(new Date(expensesData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        title: expensesData[i][1] || "",
        amount: expensesData[i][2] || 0
      });
    }
    return list;
  }

  // 6. إضافة مصروف
  if (action === "addExpense") {
    expensesSheet.appendRow([new Date(), data.title, data.amount, data.phone]);
    return { status: "success", message: "تم تسجيل المصروف بنجاح!" };
  }

  // 7. تعديل سجل
  if (action === "editDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      if (data.amount !== undefined) donationsSheet.getRange(rowIndex, 4).setValue(data.amount);
      if (data.status) donationsSheet.getRange(rowIndex, 6).setValue(data.status);
      return { status: "success", message: "تم تعديل السجل بنجاح!" };
    }
  }

  // 8. حذف سجل
  if (action === "deleteDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      donationsSheet.deleteRow(rowIndex);
      return { status: "success", message: "تم حذف السجل بنجاح!" };
    }
  }

  // 9. إنشاء رابط PDF
  if (action === "generatePDF") {
    const url = generatePDFUrl(ss, data.pdfType || "donations");
    return { status: "success", pdfDownloadUrl: url };
  }

  return { status: "error", message: "طلب غير معروف!" };
}

function checkAndResetMonthlyStatus(donationsSheet, currentMonth) {
  const data = donationsSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const firstRecordMonth = data[1][2];
  if (firstRecordMonth && firstRecordMonth !== currentMonth) {
    for (let i = 1; i < data.length; i++) {
      donationsSheet.getRange(i + 1, 3).setValue(currentMonth);
      donationsSheet.getRange(i + 1, 5).setValue("");
      donationsSheet.getRange(i + 1, 6).setValue("في الانتظار");
    }
  }
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
  
  return ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&size=letter&portrait=true&fitw=true&gridlines=true&printtitle=false&sheetnames=false&fzr=false' + sheetId;
}
