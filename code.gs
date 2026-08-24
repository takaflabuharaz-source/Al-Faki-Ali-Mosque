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
    } else if (e && e.parameter) {
      data = e.parameter;
    }
  } catch (err) {
    data = (e && e.parameter) ? e.parameter : {};
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
            currentAmount = donationsData[j][2];
            if (donationsData[j][3] === "تم الدفع") {
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

// 3. جلب قائمة التبرعات (معدل ليطابق شيت جوجل الخاص بك)
  if (action === "getDonations") {
    const donationsData = donationsSheet.getDataRange().getValues();
    let list = [];
    for (let i = 1; i < donationsData.length; i++) {
      list.push({
        name: donationsData[i][0],         // العمود A: الاسم
        phone: donationsData[i][1],        // العمود B: رقم الهاتف
        month: donationsData[i][2] || "",  // العمود C: الشهر
        amount: donationsData[i][3],       // العمود D: المبلغ
        receiptUrl: donationsData[i][4] || "", // العمود E: رابط الإشعار
        status: donationsData[i][5]        // العمود F: الحالة
      });
    }
    return list;
  }

  // 4. إضافة وتحديث تبرع (معدل ليطابق شيت جوجل الخاص بك)
  if (action === "addDonations" || action === "addDonation") {
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
        donationsSheet.getRange(i + 1, 3).setValue(data.month || "أغسطس"); // العمود C: الشهر
        donationsSheet.getRange(i + 1, 4).setValue(data.amount);            // العمود D: المبلغ
        if (receiptUrl) donationsSheet.getRange(i + 1, 5).setValue(receiptUrl); // العمود E: رابط الإشعار
        donationsSheet.getRange(i + 1, 6).setValue("تم الدفع");             // العمود F: الحالة
        donationsSheet.getRange(i + 1, 7).setValue(new Date());              // العمود G: التاريخ
        updated = true;
        break;
      }
    }

    if (!updated) {
      // إدراج صف جديد بالترتيب الصحيح [A, B, C, D, E, F, G]
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
      targetPhone: "249912345678"
    };
  }

  // 5. إضافة مصروف
  if (action === "addExpense") {
    expensesSheet.appendRow([new Date(), data.title, data.amount, data.phone]);
    const pdfUrl = generatePDFUrl(ss, "expenses");
    return {
      status: "success",
      message: "تم تسجيل المصروف",
      pdfDownloadUrl: pdfUrl,
      targetPhone: "249912345678"
    };
  }

  // 6. توليد كشف PDF
  if (action === "generatePDF") {
    const pdfUrl = generatePDFUrl(ss, data.pdfType);
    return { status: "success", pdfDownloadUrl: pdfUrl };
  }

  return { status: "error", message: "طلب غير معروف" };
}

function generatePDFUrl(ss, type) {
  const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&size=letter&portrait=true&fitw=true&gridlines=true&printtitle=false&sheetnames=false&fzr=false';
  return url;
}
