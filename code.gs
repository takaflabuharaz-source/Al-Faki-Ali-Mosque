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

  // إذا تم فتح الرابط مباشرة في المتصفح بدون برامترات
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

function processForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("المستخدمين") || ss.insertSheet("المستخدمين");
  const donationsSheet = ss.getSheetByName("التبرعات") || ss.insertSheet("التبرعات");
  const expensesSheet = ss.getSheetByName("المصروفات") || ss.insertSheet("المصروفات");

  const action = data.action;

  // 1. تسجيل حساب جديد
  if (action === "register") {
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString() === data.phone.toString()) {
        return { status: "error", message: "رقم الهاتف مسجل بالفعل!" };
      }
    }
    usersSheet.appendRow([data.name, data.phone, data.password, "مستخدم"]);

    donationsSheet.appendRow([
      data.name,
      data.phone,
      "", 
      0, 
      "", 
      "في الانتظار", 
      new Date()
    ]);

    return { status: "success", message: "تم إنشاء الحساب بنجاح! يمكنك الآن تسجيل الدخول." };
  }

  // 2. تسجيل الدخول
  if (action === "login") {
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][1].toString() === data.phone.toString() && usersData[i][2].toString() === data.password.toString()) {
        const userName = usersData[i][0];
        const isAdmin = usersData[i][3] === "آدمن" || usersData[i][3] === "ادمن" || usersData[i][3] === "Admin";
        
        let isPaid = false;
        let currentAmount = 10000;
        const donationsData = donationsSheet.getDataRange().getValues();
        for (let j = 1; j < donationsData.length; j++) {
          if (donationsData[j][1].toString() === data.phone.toString()) {
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
          currentAmount: currentAmount
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
        month: donationsData[i][2] || "",
        amount: donationsData[i][3] || 0,
        receiptUrl: donationsData[i][4] || "",
        status: donationsData[i][5] || "في الانتظار"
      });
    }
    return list;
  }

  // 4. إضافة وتحديث تبرع
  if (action === "addDonation" || action === "addDonations") {
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
      if (donationsData[i][1].toString() === data.phone.toString()) {
        donationsSheet.getRange(i + 1, 3).setValue(data.month || "أغسطس");
        donationsSheet.getRange(i + 1, 4).setValue(data.amount);
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
        data.month || "أغسطس", 
        data.amount, 
        receiptUrl, 
        "تم الدفع", 
        new Date()
      ]);
    }

    const pdfUrl = generatePDFUrl(ss, "paid");
    return {
      status: "success",
      message: "تم تسجيل التبرع بنجاح",
      pdfDownloadUrl: pdfUrl,
      targetPhone: data.phone
    };
  }

  // 5. تعديل سجل بواسطة الإدارة
  if (action === "editDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      if (data.name) donationsSheet.getRange(rowIndex, 1).setValue(data.name);
      if (data.amount) donationsSheet.getRange(rowIndex, 4).setValue(data.amount);
      if (data.status) donationsSheet.getRange(rowIndex, 6).setValue(data.status);
      return { status: "success", message: "تم تعديل السجل بنجاح" };
    }
    return { status: "error", message: "رقم الصف غير صحيح" };
  }

  // 6. حذف سجل بواسطة الإدارة
  if (action === "deleteDonation") {
    const rowIndex = parseInt(data.rowIndex);
    if (rowIndex > 1) {
      donationsSheet.deleteRow(rowIndex);
      return { status: "success", message: "تم حذف السجل بنجاح" };
    }
    return { status: "error", message: "رقم الصف غير صحيح" };
  }

  // 7. إضافة مصروف
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

  // 8. توليد كشف PDF
  if (action === "generatePDF" || action === "generateExpensesPDF") {
    const pdfUrl = generatePDFUrl(ss, data.pdfType || "all");
    return { status: "success", pdfDownloadUrl: pdfUrl };
  }

  return { status: "error", message: "طلب غير معروف" };
}

function generatePDFUrl(ss, type) {
  try {
    ss.setSharing(SpreadsheetApp.Access.ANYONE_WITH_LINK, SpreadsheetApp.Permission.VIEW);
  } catch(e) {}
  
  const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&size=letter&portrait=true&fitw=true&gridlines=true&printtitle=false&sheetnames=false&fzr=false';
  return url;
}
