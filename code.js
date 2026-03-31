// إعداد مركزي
var CONFIG = {
  COMPANY_SHEETS: {
    "CNX القاهرة الجديدة": "CNX New Cairo 23-12",
    "الشركة الثانية": "اسم_الشيت_2",
  },
  COL_DRIVER_ID: 7, // index starts from 0
  COL_PASSWORD: 8, // index starts from 0
  ELIGIBLE_PAIRS_SHEET: "EligiblePairs",
  SWAP_REQUESTS_SHEET: "SwapRequests",
  ADMINS_SHEET: "Admins",
  SHIFT_KEY_HEADER: "ShiftKey",
  STATUS_PENDING: "Pending",
  STATUS_APPROVED: "Approved",
  STATUS_REJECTED: "Rejected",
  PHONE_HEADER_CANDIDATES: [
    "رقم الهاتف",
    "الهاتف",
    "تليفون",
    "موبايل",
    "phone",
  ],
  DRIVER_NAME_HEADER_CANDIDATES: [
    "اسم السائق",
    "اسم الكابتن",
    "الاسم",
    "driver",
  ],
  DATE_HEADER_CANDIDATES: ["التاريخ", "تاريخ", "date"],
  TIME_HEADER_CANDIDATES: ["الوقت", "وقت", "موعد", "النقطة", "time"],
  TRIP_HEADER_CANDIDATES: ["الرحلة", "الرحله", "رحلة", "رحله", "trip"],
};

var ELIGIBLE_PAIRS_HEADERS = [
  "CompanyKey",
  "DriverA_Id",
  "DriverB_Id",
  "Active",
];
var SWAP_REQUESTS_HEADERS = [
  "RequestId",
  "CompanyKey",
  "RequesterDriverId",
  "PartnerDriverId",
  "ShiftKey",
  "ShiftLabel",
  "Status",
  "CreatedAt",
  "ResolvedAt",
  "ResolverEmail",
  "Notes",
];
var ADMINS_HEADERS = ["Email", "CompanyKey"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

function onOpen() {
  // intentionally left blank: magic-link menu is disabled
}

function showGenerateTokensDialog() {
  SpreadsheetApp.getUi().alert("تم إيقاف ميزة الرابط السحري.");
}

function getCompanies() {
  return Object.keys(CONFIG.COMPANY_SHEETS);
}

function checkLogin(driverId, password, companyName, queryDate) {
  var normalizedDriverId = normalizeValue(driverId);
  var normalizedPassword = normalizeValue(password);
  var normalizedQueryDate = normalizeValue(queryDate);

  if (!normalizedDriverId) {
    return JSON.stringify({
      success: false,
      message: "يرجى إدخال رقم السائق.",
    });
  }
  if (!normalizedPassword) {
    return JSON.stringify({
      success: false,
      message: "يرجى إدخال كلمة المرور.",
    });
  }
  if (!normalizedQueryDate) {
    return JSON.stringify({
      success: false,
      message: "يرجى اختيار التاريخ المطلوب.",
    });
  }

  try {
    var sheet = getCompanySheetOrThrow(companyName);
    var loaded = loadCompanySheetData(sheet);
    if (!loaded.ok) {
      return JSON.stringify({ success: false, message: loaded.message });
    }

    var matches = filterDriverRows(
      loaded,
      normalizedDriverId,
      normalizedPassword,
    );
    if (!matches.length) {
      return JSON.stringify({
        success: false,
        message:
          "رقم السائق أو كلمة المرور غير صحيحة للشركة المختارة. يرجى التحقق والمحاولة مرة أخرى.",
      });
    }

    var scriptTimeZone = Session.getScriptTimeZone();
    var selectedDayKey = parseQueryDateToDayKey(normalizedQueryDate);
    if (!selectedDayKey) {
      return JSON.stringify({
        success: false,
        message: "التاريخ المختار غير صالح.",
      });
    }

    var sorted = sortDriverRows(matches, loaded.headers);
    var headers = loaded.headers;
    var formattedRows = [headers];
    var backgrounds = [loaded.headerStyle.backgrounds];
    var fontColors = [loaded.headerStyle.fontColors];
    var fontWeights = [loaded.headerStyle.fontWeights];
    var fontStyles = [loaded.headerStyle.fontStyles];
    var horizontalAlignments = [loaded.headerStyle.horizontalAlignments];
    var shiftEntries = [];
    var selectedDayCount = 0;
    var cumulativeCount = 0;
    var cumulativeFromDayKey = "";
    var cumulativeToDayKey = selectedDayKey;

    for (var i = 0; i < sorted.length; i++) {
      var item = sorted[i];
      var effectiveDayKey = getEffectiveDayKey(
        item.row,
        loaded.columns,
        scriptTimeZone,
      );
      if (!effectiveDayKey) {
        continue;
      }
      if (effectiveDayKey <= selectedDayKey) {
        cumulativeCount++;
        if (!cumulativeFromDayKey || effectiveDayKey < cumulativeFromDayKey) {
          cumulativeFromDayKey = effectiveDayKey;
        }
      }
      if (effectiveDayKey === selectedDayKey) {
        selectedDayCount++;
        var rowCopy = item.row.slice();
        formattedRows.push(formatRowForDisplay(rowCopy, headers));
        backgrounds.push(item.background.slice());
        fontColors.push(item.fontColor.slice());
        fontWeights.push(item.fontWeight.slice());
        fontStyles.push(item.fontStyle.slice());
        horizontalAlignments.push(item.horizontalAlignment.slice());
        shiftEntries.push({
          shiftKey: item.shiftKey,
          shiftLabel: item.shiftLabel,
          rowIndex: item.rowIndex,
        });
      }
    }

    var profile = getDriverProfileFromLoadedData(loaded, normalizedDriverId);
    var cumulativeFromDate = cumulativeFromDayKey
      ? formatDayKeyToDisplay(cumulativeFromDayKey)
      : "";

    return JSON.stringify({
      success: true,
      fileName: SpreadsheetApp.getActiveSpreadsheet().getName(),
      sheetName: sheet.getName(),
      companyName: companyName,
      driverId: normalizedDriverId,
      driverName: profile.name || "",
      driverPhone: profile.phone || "",
      data: formattedRows,
      backgrounds: backgrounds,
      fontColors: fontColors,
      fontWeights: fontWeights,
      fontStyles: fontStyles,
      horizontalAlignments: horizontalAlignments,
      shiftEntries: shiftEntries,
      selectedDayCount: selectedDayCount,
      cumulativeCount: cumulativeCount,
      cumulativeFromDate: cumulativeFromDate,
      cumulativeToDate: formatDayKeyToDisplay(cumulativeToDayKey),
      queryDateDisplay: formatDayKeyToDisplay(selectedDayKey),
    });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.message });
  }
}

function getEligiblePartners(companyName, driverId, password) {
  var auth = authenticateDriver(companyName, driverId, password);
  if (!auth.success) {
    return JSON.stringify(auth);
  }

  ensureMetaSheets();

  var eligiblePairsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.ELIGIBLE_PAIRS_SHEET,
  );
  var partnerIds = getEligiblePartnerIdsFromPairs(
    eligiblePairsSheet,
    companyName,
    driverId,
  );
  var partners = [];

  for (var i = 0; i < partnerIds.length; i++) {
    var profile = getDriverProfileFromLoadedData(
      auth.loadedData,
      partnerIds[i],
    );
    partners.push({
      driverId: partnerIds[i],
      driverName: profile.name || "",
      phone: profile.phone || "",
    });
  }

  return JSON.stringify({ success: true, partners: partners });
}

function requestSwap(
  companyName,
  driverId,
  password,
  partnerDriverId,
  shiftKeys,
  note,
) {
  var auth = authenticateDriver(companyName, driverId, password);
  if (!auth.success) {
    return JSON.stringify(auth);
  }

  var normalizedPartnerId = normalizeValue(partnerDriverId);
  var normalizedShiftKeys = normalizeShiftKeysInput(shiftKeys);
  if (!normalizedPartnerId) {
    return JSON.stringify({
      success: false,
      message: "يرجى اختيار السائق البديل.",
    });
  }
  if (!normalizedShiftKeys.length) {
    return JSON.stringify({
      success: false,
      message: "يرجى اختيار شيفت واحد على الأقل.",
    });
  }
  if (normalizedPartnerId === normalizeValue(driverId)) {
    return JSON.stringify({
      success: false,
      message: "لا يمكن إنشاء طلب تبديل مع نفس السائق.",
    });
  }

  ensureMetaSheets();
  var pairSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.ELIGIBLE_PAIRS_SHEET,
  );
  if (!isPairEligible(pairSheet, companyName, driverId, normalizedPartnerId)) {
    return JSON.stringify({
      success: false,
      message:
        "هذا السائق غير مسموح له بالتبديل معك حسب إعدادات EligiblePairs.",
    });
  }

  var requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SWAP_REQUESTS_SHEET,
  );
  var existingRows = getSheetBodyValues(requestsSheet);
  for (var i = 0; i < existingRows.length; i++) {
    var existing = mapRequestRow(existingRows[i]);
    var existingShiftKeys = deserializeShiftKeys(existing.shiftKey);
    if (
      normalizeValue(existing.companyKey) === normalizeValue(companyName) &&
      normalizeValue(existing.requesterDriverId) === normalizeValue(driverId) &&
      normalizeValue(existing.status) === normalizeValue(CONFIG.STATUS_PENDING)
    ) {
      for (var j = 0; j < normalizedShiftKeys.length; j++) {
        if (existingShiftKeys.indexOf(normalizedShiftKeys[j]) !== -1) {
          return JSON.stringify({
            success: false,
            message: "يوجد طلب تبديل Pending بالفعل لأحد الشيفتات المختارة.",
          });
        }
      }
    }
  }

  var shiftLabels = [];
  for (var k = 0; k < normalizedShiftKeys.length; k++) {
    var shiftInfo = findShiftInfoByKey(
      auth.sheet,
      auth.loadedData.headers,
      normalizedShiftKeys[k],
    );
    if (!shiftInfo.found) {
      return JSON.stringify({
        success: false,
        message: "أحد الشيفتات المختارة غير موجود.",
      });
    }

    var rowDriverId = normalizeValue(
      shiftInfo.row[auth.loadedData.columns.driverIdCol],
    );
    if (rowDriverId !== normalizeValue(driverId)) {
      return JSON.stringify({
        success: false,
        message: "أحد الشيفتات المختارة ليس مملوكًا للسائق الحالي.",
      });
    }
    shiftLabels.push(shiftInfo.shiftLabel);
  }

  var batchSummary = shiftLabels[0] || normalizedShiftKeys[0];
  if (normalizedShiftKeys.length > 1) {
    batchSummary += " + " + (normalizedShiftKeys.length - 1) + " شيفت";
  }

  var now = new Date();
  var requestId = Utilities.getUuid();
  requestsSheet.appendRow([
    requestId,
    companyName,
    normalizeValue(driverId),
    normalizedPartnerId,
    JSON.stringify(normalizedShiftKeys),
    batchSummary,
    CONFIG.STATUS_PENDING,
    now,
    "",
    "",
    note || "",
  ]);

  return JSON.stringify({
    success: true,
    message:
      "تم إرسال طلب تبديل " +
      normalizedShiftKeys.length +
      " شيفت بنجاح وحالته الآن Pending.",
    requestId: requestId,
  });
}

function listDriverSwapRequests(companyName, driverId, password) {
  var auth = authenticateDriver(companyName, driverId, password);
  if (!auth.success) {
    return JSON.stringify(auth);
  }

  ensureMetaSheets();
  var requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SWAP_REQUESTS_SHEET,
  );
  var rows = getSheetBodyValues(requestsSheet);
  var requests = [];
  var currentDriverId = normalizeValue(driverId);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var req = mapRequestRow(row);
    if (normalizeValue(req.companyKey) !== normalizeValue(companyName))
      continue;
    if (
      normalizeValue(req.requesterDriverId) !== currentDriverId &&
      normalizeValue(req.partnerDriverId) !== currentDriverId
    ) {
      continue;
    }
    requests.push(req);
  }

  requests.sort(function (a, b) {
    return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
  });

  return JSON.stringify({ success: true, requests: requests });
}

function listPendingSwapRequests(companyName) {
  ensureMetaSheets();
  var adminCheck = assertSheetAdmin(companyName);
  if (!adminCheck.success) {
    return JSON.stringify(adminCheck);
  }

  var requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SWAP_REQUESTS_SHEET,
  );
  var rows = getSheetBodyValues(requestsSheet);
  var pending = [];
  for (var i = 0; i < rows.length; i++) {
    var req = mapRequestRow(rows[i]);
    if (
      normalizeValue(req.companyKey) === normalizeValue(companyName) &&
      normalizeValue(req.status) === normalizeValue(CONFIG.STATUS_PENDING)
    ) {
      pending.push(req);
    }
  }
  return JSON.stringify({ success: true, pending: pending });
}

function getAdminAccessContext() {
  ensureMetaSheets();
  var context = buildAdminAccessContext();
  if (!context.success) {
    return JSON.stringify(context);
  }
  return JSON.stringify({
    success: true,
    isAdmin: context.isAdmin,
    email: context.email,
    allCompanies: context.allCompanies,
    allowedCompanies: context.allowedCompanies,
    companyOptions: context.companyOptions,
  });
}

function listPendingSwapRequestsForAdmin(companyName) {
  ensureMetaSheets();
  var context = buildAdminAccessContext();
  if (!context.success) {
    return JSON.stringify(context);
  }
  if (!context.isAdmin) {
    return JSON.stringify({
      success: false,
      message: "ليس لديك صلاحية مسؤول.",
    });
  }

  var normalizedFilterCompany = normalizeValue(companyName);
  if (
    normalizedFilterCompany &&
    !context.allCompanies &&
    context.allowedCompanies.indexOf(normalizedFilterCompany) === -1
  ) {
    return JSON.stringify({
      success: false,
      message: "هذه الشركة ليست ضمن صلاحياتك.",
    });
  }

  var requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SWAP_REQUESTS_SHEET,
  );
  var rows = getSheetBodyValues(requestsSheet);
  var pending = [];
  for (var i = 0; i < rows.length; i++) {
    var req = mapRequestRow(rows[i]);
    if (normalizeValue(req.status) !== normalizeValue(CONFIG.STATUS_PENDING)) {
      continue;
    }
    var reqCompany = normalizeValue(req.companyKey);
    if (normalizedFilterCompany && reqCompany !== normalizedFilterCompany) {
      continue;
    }
    if (
      !context.allCompanies &&
      context.allowedCompanies.indexOf(reqCompany) === -1
    ) {
      continue;
    }
    pending.push(req);
  }

  pending.sort(function (a, b) {
    return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
  });

  return JSON.stringify({
    success: true,
    pending: pending,
    context: {
      allCompanies: context.allCompanies,
      allowedCompanies: context.allowedCompanies,
      companyOptions: context.companyOptions,
      email: context.email,
    },
  });
}

function resolveSwapRequest(requestId, companyName, approve, note) {
  ensureMetaSheets();
  var adminCheck = assertSheetAdmin(companyName);
  if (!adminCheck.success) {
    return JSON.stringify(adminCheck);
  }

  var requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SWAP_REQUESTS_SHEET,
  );
  var bodyRows = getSheetBodyValues(requestsSheet);
  var target = null;
  var targetSheetRow = -1;

  for (var i = 0; i < bodyRows.length; i++) {
    var req = mapRequestRow(bodyRows[i]);
    if (
      normalizeValue(req.requestId) === normalizeValue(requestId) &&
      normalizeValue(req.companyKey) === normalizeValue(companyName)
    ) {
      target = req;
      targetSheetRow = i + 2;
      break;
    }
  }

  if (!target) {
    return JSON.stringify({
      success: false,
      message: "طلب التبديل غير موجود.",
    });
  }
  if (normalizeValue(target.status) !== normalizeValue(CONFIG.STATUS_PENDING)) {
    return JSON.stringify({
      success: false,
      message: "تمت معالجة هذا الطلب مسبقًا.",
    });
  }

  var nextStatus = approve ? CONFIG.STATUS_APPROVED : CONFIG.STATUS_REJECTED;

  if (approve) {
    var companySheet = getCompanySheetOrThrow(companyName);
    var loaded = loadCompanySheetData(companySheet);
    var shiftKeys = deserializeShiftKeys(target.shiftKey);
    if (!shiftKeys.length) {
      return JSON.stringify({
        success: false,
        message: "لا توجد مفاتيح شيفت صالحة داخل الطلب.",
      });
    }
    var partnerProfile = getDriverProfileFromLoadedData(
      loaded,
      target.partnerDriverId,
    );
    if (!partnerProfile.found) {
      return JSON.stringify({
        success: false,
        message: "تعذر العثور على بيانات السائق البديل في شيت الشركة.",
      });
    }

    var updates = [];
    for (var j = 0; j < shiftKeys.length; j++) {
      var shiftInfo = findShiftInfoByKey(
        companySheet,
        loaded.headers,
        shiftKeys[j],
      );
      if (!shiftInfo.found) {
        return JSON.stringify({
          success: false,
          message: "تعذر تطبيق التبديل: أحد الشيفتات غير موجود.",
        });
      }
      var ownerId = normalizeValue(shiftInfo.row[loaded.columns.driverIdCol]);
      if (ownerId !== normalizeValue(target.requesterDriverId)) {
        return JSON.stringify({
          success: false,
          message:
            "تعذر تطبيق التبديل: أحد الشيفتات لا يملكه السائق مقدم الطلب حاليًا.",
        });
      }
      updates.push(shiftInfo.rowIndex);
    }

    for (var k = 0; k < updates.length; k++) {
      var rowIndex = updates[k];
      companySheet
        .getRange(rowIndex, loaded.columns.driverIdCol + 1)
        .setValue(partnerProfile.driverId);
      companySheet
        .getRange(rowIndex, loaded.columns.passwordCol + 1)
        .setValue(partnerProfile.password);
      if (loaded.columns.driverNameCol >= 0) {
        companySheet
          .getRange(rowIndex, loaded.columns.driverNameCol + 1)
          .setValue(partnerProfile.name || "");
      }
      if (loaded.columns.phoneCol >= 0) {
        companySheet
          .getRange(rowIndex, loaded.columns.phoneCol + 1)
          .setValue(partnerProfile.phone || "");
      }
    }
  }

  requestsSheet.getRange(targetSheetRow, 7).setValue(nextStatus);
  requestsSheet.getRange(targetSheetRow, 9).setValue(new Date());
  requestsSheet.getRange(targetSheetRow, 10).setValue(adminCheck.email);
  requestsSheet.getRange(targetSheetRow, 11).setValue(note || "");

  return JSON.stringify({
    success: true,
    message:
      nextStatus === CONFIG.STATUS_APPROVED
        ? "تم قبول الطلب وتحديث جميع الشيفتات المختارة بنجاح."
        : "تم رفض طلب التبديل.",
  });
}

function loginByToken(token) {
  return JSON.stringify({
    success: false,
    message: "تم إيقاف تسجيل الدخول بالرابط السحري.",
  });
}

function generateRandomToken() {
  return "";
}

function generateTokensForCompany(companyName) {
  return {
    success: false,
    message: "تم إيقاف ميزة توليد الروابط السحرية.",
    count: 0,
  };
}

function getMagicLinkForCaptain(driverId, password, companyName) {
  return JSON.stringify({
    success: false,
    message: "تم إيقاف ميزة الرابط السحري.",
  });
}

function ensureMetaSheets() {
  ensureSheetWithHeaders(CONFIG.ELIGIBLE_PAIRS_SHEET, ELIGIBLE_PAIRS_HEADERS);
  ensureSheetWithHeaders(CONFIG.SWAP_REQUESTS_SHEET, SWAP_REQUESTS_HEADERS);
  ensureSheetWithHeaders(CONFIG.ADMINS_SHEET, ADMINS_HEADERS);
}

function ensureSheetWithHeaders(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  var existingHeaders = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  var needWrite = false;
  for (var i = 0; i < headers.length; i++) {
    if (normalizeValue(existingHeaders[i]) !== normalizeValue(headers[i])) {
      needWrite = true;
      break;
    }
  }
  if (needWrite || existingHeaders.length < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getCompanySheetOrThrow(companyName) {
  if (!companyName || !CONFIG.COMPANY_SHEETS[companyName]) {
    throw new Error("الشركة غير صحيحة. يرجى اختيار شركة من القائمة.");
  }
  var sheetName = CONFIG.COMPANY_SHEETS[companyName];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("شيت الشركة غير موجود: " + sheetName);
  }
  return sheet;
}

function loadCompanySheetData(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) {
    return { ok: false, message: "الشيت لا يحتوي على بيانات كافية." };
  }

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var fontColors = range.getFontColors();
  var fontWeights = range.getFontWeights();
  var fontStyles = range.getFontStyles();
  var horizontalAlignments = range.getHorizontalAlignments();
  var headers = values[0];

  var columns = {
    driverIdCol: CONFIG.COL_DRIVER_ID,
    passwordCol: CONFIG.COL_PASSWORD,
    driverNameCol: findColumnByHeader(
      headers,
      CONFIG.DRIVER_NAME_HEADER_CANDIDATES,
    ),
    phoneCol: findColumnByHeader(headers, CONFIG.PHONE_HEADER_CANDIDATES),
    shiftKeyCol: findColumnByHeader(headers, [CONFIG.SHIFT_KEY_HEADER]),
    dateCol: findColumnByHeader(headers, CONFIG.DATE_HEADER_CANDIDATES),
    timeCol: findColumnByHeader(headers, CONFIG.TIME_HEADER_CANDIDATES),
    tripCol: findColumnByHeader(headers, CONFIG.TRIP_HEADER_CANDIDATES),
  };

  return {
    ok: true,
    headers: headers,
    rows: values,
    backgrounds: backgrounds,
    fontColors: fontColors,
    fontWeights: fontWeights,
    fontStyles: fontStyles,
    horizontalAlignments: horizontalAlignments,
    columns: columns,
    headerStyle: {
      backgrounds: backgrounds[0],
      fontColors: fontColors[0],
      fontWeights: fontWeights[0],
      fontStyles: fontStyles[0],
      horizontalAlignments: horizontalAlignments[0],
    },
  };
}

function filterDriverRows(loaded, driverId, password) {
  var list = [];
  for (var i = 1; i < loaded.rows.length; i++) {
    var row = loaded.rows[i];
    var rowDriverId = normalizeValue(row[loaded.columns.driverIdCol]);
    var rowPassword = normalizeValue(row[loaded.columns.passwordCol]);
    if (
      rowDriverId === normalizeValue(driverId) &&
      rowPassword === normalizeValue(password)
    ) {
      var rowIndex = i + 2;
      var shiftKey = buildShiftKey(row, loaded.columns.shiftKeyCol, rowIndex);
      list.push({
        row: row,
        rowIndex: rowIndex,
        shiftKey: shiftKey,
        shiftLabel: buildShiftLabel(row, loaded.headers, shiftKey),
        background: loaded.backgrounds[i],
        fontColor: loaded.fontColors[i],
        fontWeight: loaded.fontWeights[i],
        fontStyle: loaded.fontStyles[i],
        horizontalAlignment: loaded.horizontalAlignments[i],
      });
    }
  }
  return list;
}

function sortDriverRows(rowsWithMeta, headers) {
  var dateCol = findColumnByHeader(headers, CONFIG.DATE_HEADER_CANDIDATES);
  var timeCol = findColumnByHeader(headers, CONFIG.TIME_HEADER_CANDIDATES);
  var tripCol = findColumnByHeader(headers, CONFIG.TRIP_HEADER_CANDIDATES);
  if (dateCol < 0) dateCol = 0;

  rowsWithMeta.sort(function (a, b) {
    var keyA = getSortKeyMs(a.row, dateCol, timeCol);
    var keyB = getSortKeyMs(b.row, dateCol, timeCol);
    if (keyA !== keyB) return keyA - keyB;
    var tripA = tripCol >= 0 ? getRihlaSortOrder(a.row[tripCol]) : 2;
    var tripB = tripCol >= 0 ? getRihlaSortOrder(b.row[tripCol]) : 2;
    return tripA - tripB;
  });
  return rowsWithMeta;
}

function formatRowForDisplay(row, headers) {
  for (var j = 0; j < row.length; j++) {
    var cell = row[j];
    var header = headers[j] ? headers[j].toString().trim() : "";
    if (typeof cell === "number" && cell < 1 && cell >= 0) {
      var totalMinutes = Math.round(cell * 24 * 60);
      var hours = Math.floor(totalMinutes / 60);
      var minutes = totalMinutes % 60;
      row[j] =
        (hours < 10 ? "0" + hours : hours) +
        ":" +
        (minutes < 10 ? "0" + minutes : minutes);
      continue;
    }
    if (cell instanceof Date) {
      var isTimeOnly = cell.getFullYear() <= 1900;
      if (
        header.indexOf("الوقت") !== -1 ||
        header.indexOf("النقطة") !== -1 ||
        header.indexOf("موعد") !== -1 ||
        isTimeOnly
      ) {
        row[j] = Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "HH:mm",
        );
      } else if (
        header.indexOf("تاريخ") !== -1 ||
        header.indexOf("التاريخ") !== -1
      ) {
        row[j] = Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy",
        );
      } else {
        row[j] = Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "yyyy/MM/dd HH:mm",
        );
      }
    }
  }
  return row;
}

function buildShiftKey(row, shiftKeyCol, rowIndex) {
  if (shiftKeyCol >= 0 && normalizeValue(row[shiftKeyCol])) {
    return normalizeValue(row[shiftKeyCol]);
  }
  return "ROW:" + rowIndex;
}

function buildShiftLabel(row, headers, fallbackKey) {
  var dateCol = findColumnByHeader(headers, CONFIG.DATE_HEADER_CANDIDATES);
  var timeCol = findColumnByHeader(headers, CONFIG.TIME_HEADER_CANDIDATES);
  var tripCol = findColumnByHeader(headers, CONFIG.TRIP_HEADER_CANDIDATES);
  var parts = [];
  if (dateCol >= 0) parts.push(formatCellLite(row[dateCol]));
  if (timeCol >= 0) parts.push(formatCellLite(row[timeCol]));
  if (tripCol >= 0) parts.push(formatCellLite(row[tripCol]));
  var label = parts.filter(Boolean).join(" | ");
  return label || fallbackKey;
}

function authenticateDriver(companyName, driverId, password) {
  try {
    var sheet = getCompanySheetOrThrow(companyName);
    var loaded = loadCompanySheetData(sheet);
    if (!loaded.ok) {
      return { success: false, message: loaded.message };
    }
    var rows = filterDriverRows(loaded, driverId, password);
    if (!rows.length) {
      return {
        success: false,
        message: "بيانات الدخول غير صحيحة للشركة المختارة.",
      };
    }
    return { success: true, sheet: sheet, loadedData: loaded };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getEligiblePartnerIdsFromPairs(pairSheet, companyName, driverId) {
  var rows = getSheetBodyValues(pairSheet);
  var partnerMap = {};
  var myId = normalizeValue(driverId);

  for (var i = 0; i < rows.length; i++) {
    var company = normalizeValue(rows[i][0]);
    var a = normalizeValue(rows[i][1]);
    var b = normalizeValue(rows[i][2]);
    var activeCell = rows[i][3];
    var isActive = parseTruthy(activeCell);
    if (!isActive) continue;
    if (company !== normalizeValue(companyName)) continue;
    if (a === myId && b) partnerMap[b] = true;
    if (b === myId && a) partnerMap[a] = true;
  }
  return Object.keys(partnerMap);
}

function isPairEligible(pairSheet, companyName, driverA, driverB) {
  var a = normalizeValue(driverA);
  var b = normalizeValue(driverB);
  var rows = getSheetBodyValues(pairSheet);
  for (var i = 0; i < rows.length; i++) {
    var company = normalizeValue(rows[i][0]);
    var p1 = normalizeValue(rows[i][1]);
    var p2 = normalizeValue(rows[i][2]);
    var active = parseTruthy(rows[i][3]);
    if (!active) continue;
    if (company !== normalizeValue(companyName)) continue;
    if ((p1 === a && p2 === b) || (p1 === b && p2 === a)) {
      return true;
    }
  }
  return false;
}

function findShiftInfoByKey(sheet, headers, shiftKey) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return { found: false };
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var shiftKeyCol = findColumnByHeader(headers, [CONFIG.SHIFT_KEY_HEADER]);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;
    var rowKey = buildShiftKey(row, shiftKeyCol, rowIndex);
    if (normalizeValue(rowKey) === normalizeValue(shiftKey)) {
      return {
        found: true,
        rowIndex: rowIndex,
        row: row,
        shiftLabel: buildShiftLabel(row, headers, rowKey),
      };
    }
  }
  return { found: false };
}

function getDriverProfileFromLoadedData(loaded, driverId) {
  var id = normalizeValue(driverId);
  for (var i = 1; i < loaded.rows.length; i++) {
    var row = loaded.rows[i];
    if (normalizeValue(row[loaded.columns.driverIdCol]) === id) {
      return {
        found: true,
        driverId: id,
        password: normalizeValue(row[loaded.columns.passwordCol]),
        name:
          loaded.columns.driverNameCol >= 0
            ? String(row[loaded.columns.driverNameCol] || "")
            : "",
        phone:
          loaded.columns.phoneCol >= 0
            ? String(row[loaded.columns.phoneCol] || "")
            : "",
      };
    }
  }
  return { found: false, driverId: id, password: "", name: "", phone: "" };
}

function assertSheetAdmin(companyName) {
  var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.ADMINS_SHEET,
  );
  var email = normalizeValue(Session.getActiveUser().getEmail());
  if (!email) {
    return {
      success: false,
      message:
        "لا يمكن التحقق من هوية المسؤول. تأكد من تنفيذ السكربت بحساب Google Workspace صحيح.",
    };
  }
  var rows = getSheetBodyValues(adminSheet);
  for (var i = 0; i < rows.length; i++) {
    var rowEmail = normalizeValue(rows[i][0]);
    var rowCompany = normalizeValue(rows[i][1]);
    if (!rowEmail) continue;
    if (rowEmail !== email) continue;
    if (rowCompany === "*" || rowCompany === normalizeValue(companyName)) {
      return { success: true, email: email };
    }
  }
  return {
    success: false,
    message: "ليس لديك صلاحية اعتماد طلبات التبديل لهذه الشركة.",
  };
}

function buildAdminAccessContext() {
  var email = normalizeValue(Session.getActiveUser().getEmail());
  if (!email) {
    return {
      success: false,
      message:
        "لا يمكن التحقق من هوية المسؤول. تأكد من تنفيذ السكربت بحساب Google Workspace صحيح.",
    };
  }

  var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.ADMINS_SHEET,
  );
  var rows = getSheetBodyValues(adminSheet);
  var allCompanies = false;
  var allowedCompaniesMap = {};

  for (var i = 0; i < rows.length; i++) {
    var rowEmail = normalizeValue(rows[i][0]);
    var rowCompany = normalizeValue(rows[i][1]);
    if (!rowEmail || rowEmail !== email) continue;
    if (rowCompany === "*") {
      allCompanies = true;
      break;
    }
    if (rowCompany) {
      allowedCompaniesMap[rowCompany] = true;
    }
  }

  var configuredCompanies = Object.keys(CONFIG.COMPANY_SHEETS);
  var allowedCompanies = allCompanies
    ? configuredCompanies.slice()
    : Object.keys(allowedCompaniesMap).filter(function (companyName) {
        return configuredCompanies.indexOf(companyName) !== -1;
      });

  return {
    success: true,
    isAdmin: allCompanies || allowedCompanies.length > 0,
    email: email,
    allCompanies: allCompanies,
    allowedCompanies: allowedCompanies,
    companyOptions: allCompanies
      ? configuredCompanies.slice()
      : allowedCompanies.slice(),
  };
}

function mapRequestRow(row) {
  var shiftKeys = deserializeShiftKeys(row[4] || "");
  return {
    requestId: row[0] || "",
    companyKey: row[1] || "",
    requesterDriverId: row[2] || "",
    partnerDriverId: row[3] || "",
    shiftKey: row[4] || "",
    shiftKeys: shiftKeys,
    shiftCount: shiftKeys.length,
    shiftLabel: row[5] || "",
    status: row[6] || "",
    createdAt: row[7] || "",
    resolvedAt: row[8] || "",
    resolverEmail: row[9] || "",
    notes: row[10] || "",
  };
}

function getSheetBodyValues(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function normalizeShiftKeysInput(shiftKeys) {
  var raw = [];
  if (Array.isArray(shiftKeys)) {
    raw = shiftKeys;
  } else if (typeof shiftKeys === "string" && shiftKeys.trim()) {
    try {
      var parsed = JSON.parse(shiftKeys);
      if (Array.isArray(parsed)) {
        raw = parsed;
      } else {
        raw = shiftKeys.split(",");
      }
    } catch (e) {
      raw = shiftKeys.split(",");
    }
  }

  var seen = {};
  var normalized = [];
  for (var i = 0; i < raw.length; i++) {
    var key = normalizeValue(raw[i]);
    if (!key || seen[key]) continue;
    seen[key] = true;
    normalized.push(key);
  }
  return normalized;
}

function deserializeShiftKeys(shiftKeyCell) {
  var value = normalizeValue(shiftKeyCell);
  if (!value) return [];
  try {
    var parsed = JSON.parse(value);
    if (Array.isArray(parsed)) {
      return normalizeShiftKeysInput(parsed);
    }
  } catch (e) {
    // Backward compatibility for older single shift values.
  }
  return normalizeShiftKeysInput([value]);
}

function normalizeValue(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function parseTruthy(value) {
  var v = normalizeValue(value).toLowerCase();
  return v === "true" || v === "1" || v === "yes" || v === "y";
}

function findColumnByHeader(headers, candidates) {
  var normalizedCandidates = candidates.map(function (candidate) {
    return normalizeHeaderToken(candidate);
  });
  for (var i = 0; i < headers.length; i++) {
    var header = normalizeHeaderToken(headers[i]);
    for (var j = 0; j < normalizedCandidates.length; j++) {
      if (
        header === normalizedCandidates[j] ||
        header.indexOf(normalizedCandidates[j]) !== -1
      ) {
        return i;
      }
    }
  }
  return -1;
}

function normalizeHeaderToken(value) {
  var v = normalizeValue(value).toLowerCase();
  // توحيد أشهر اختلافات الكتابة العربية لتفادي فشل المطابقة
  v = v.replace(/[أإآ]/g, "ا");
  v = v.replace(/ة/g, "ه");
  v = v.replace(/ى/g, "ي");
  return v;
}

function formatCellLite(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "dd/MM HH:mm",
    );
  }
  if (typeof value === "number" && value < 1 && value >= 0) {
    var totalMinutes = Math.round(value * 24 * 60);
    var hours = Math.floor(totalMinutes / 60);
    var minutes = totalMinutes % 60;
    return (
      (hours < 10 ? "0" + hours : hours) +
      ":" +
      (minutes < 10 ? "0" + minutes : minutes)
    );
  }
  return normalizeValue(value);
}

function parseQueryDateToDayKey(queryDate) {
  var v = normalizeValue(queryDate);
  if (!v) return "";
  var m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  return (
    m[1] +
    "-" +
    (m[2].length === 1 ? "0" + m[2] : m[2]) +
    "-" +
    (m[3].length === 1 ? "0" + m[3] : m[3])
  );
}

function formatDayKeyToDisplay(dayKey) {
  if (!dayKey) return "";
  var parts = dayKey.split("-");
  if (parts.length !== 3) return "";
  return parts[2] + "/" + parts[1] + "/" + parts[0];
}

function getEffectiveDayKey(row, columns, timeZone) {
  if (columns.dateCol < 0) return "";
  var baseDayKey = extractDateDayKey(row[columns.dateCol], timeZone);
  if (!baseDayKey) return "";

  // قاعدة 24-hour: شيفت انصراف يبدأ من 00:00 حتى 10:00 يُنسب لليوم السابق
  var tripText = "";
  if (columns.tripCol >= 0) {
    tripText = normalizeValue(row[columns.tripCol]).toLowerCase();
  }
  var isDropShift =
    tripText.indexOf("انصراف") !== -1 ||
    tripText.indexOf("checkout") !== -1 ||
    tripText.indexOf("drop") !== -1;
  if (!isDropShift) return baseDayKey;

  var startMinutes = extractStartMinutes(row, columns);
  if (startMinutes >= 0 && startMinutes <= 10 * 60) {
    return shiftDayKey(baseDayKey, -1);
  }
  return baseDayKey;
}

function extractStartMinutes(row, columns) {
  if (columns.timeCol >= 0) {
    var parsed = parseTimeToMinutes(row[columns.timeCol]);
    if (parsed >= 0) return parsed;
  }
  var dateCell = row[columns.dateCol];
  if (dateCell instanceof Date) {
    return dateCell.getHours() * 60 + dateCell.getMinutes();
  }
  return -1;
}

// تحويل أي قيمة وقت إلى دقائق من منتصف الليل (0-1440)
function parseTimeToMinutes(cell) {
  if (cell === null || cell === undefined || cell === "") return -1;
  if (cell instanceof Date) return cell.getHours() * 60 + cell.getMinutes();
  if (typeof cell === "number") {
    if (cell < 1 && cell >= 0) return Math.round(cell * 24 * 60);
    return -1;
  }
  if (typeof cell === "string") {
    var trimmed = String(cell).trim();
    if (!trimmed) return -1;

    // يدعم: 00:00, 7:30, 07:30:00, 12:00 AM, 12:00 ص, 1:15 pm
    var match = trimmed.match(
      /^(\d{1,2}):(\d{1,2})(?::\d{1,2})?\s*([AaPp][Mm]|ص|م)?$/,
    );
    if (match) {
      var hours = parseInt(match[1], 10);
      var minutes = parseInt(match[2], 10);
      var meridiem = (match[3] || "").toLowerCase();
      if (minutes < 0 || minutes > 59) return -1;

      if (meridiem === "am" || meridiem === "ص") {
        if (hours === 12) hours = 0;
      } else if (meridiem === "pm" || meridiem === "م") {
        if (hours < 12) hours += 12;
      }

      if (hours < 0 || hours > 24) return -1;
      if (hours === 24 && minutes > 0) return -1;
      if (hours === 24) return 24 * 60;
      return hours * 60 + minutes;
    }
  }
  return -1;
}

function extractDateDayKey(cell, timeZone) {
  if (cell === null || cell === undefined || cell === "") return "";
  if (cell instanceof Date) {
    if (cell.getFullYear() <= 1900) return "";
    return Utilities.formatDate(
      cell,
      timeZone || Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }
  if (typeof cell === "number") {
    if (cell < 1) return "";
    // Excel serial date -> milliseconds
    var parsedDate = new Date((cell - 25569) * 86400000);
    return Utilities.formatDate(
      parsedDate,
      timeZone || Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }
  if (typeof cell === "string") {
    var trimmed = String(cell).trim();
    if (!trimmed) return "";
    var ddmmyyyy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (ddmmyyyy) {
      return (
        ddmmyyyy[3] +
        "-" +
        String(parseInt(ddmmyyyy[2], 10)).padStart(2, "0") +
        "-" +
        String(parseInt(ddmmyyyy[1], 10)).padStart(2, "0")
      );
    }
    var yyyymmdd = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (yyyymmdd) {
      return (
        yyyymmdd[1] +
        "-" +
        String(parseInt(yyyymmdd[2], 10)).padStart(2, "0") +
        "-" +
        String(parseInt(yyyymmdd[3], 10)).padStart(2, "0")
      );
    }
  }
  return "";
}

function shiftDayKey(dayKey, dayDelta) {
  var m = normalizeValue(dayKey).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  var date = new Date(
    parseInt(m[1], 10),
    parseInt(m[2], 10) - 1,
    parseInt(m[3], 10),
  );
  date.setDate(date.getDate() + (dayDelta || 0));
  return (
    date.getFullYear() +
    "-" +
    String(date.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(date.getDate()).padStart(2, "0")
  );
}

// تحويل أي قيمة تاريخ إلى milliseconds لبداية اليوم
function parseDateToMs(cell) {
  if (cell === null || cell === undefined || cell === "")
    return Number.MAX_VALUE;
  if (cell instanceof Date) {
    if (cell.getFullYear() <= 1900) return Number.MAX_VALUE;
    return new Date(
      cell.getFullYear(),
      cell.getMonth(),
      cell.getDate(),
    ).getTime();
  }
  if (typeof cell === "number") {
    if (cell < 1 && cell >= 0) return Number.MAX_VALUE;
    if (cell > 1) {
      var parsedDate = new Date((cell - 25569) * 86400000);
      return new Date(
        parsedDate.getFullYear(),
        parsedDate.getMonth(),
        parsedDate.getDate(),
      ).getTime();
    }
    return Number.MAX_VALUE;
  }
  if (typeof cell === "string") {
    var trimmed = String(cell).trim();
    if (!trimmed) return Number.MAX_VALUE;
    var ddmmyyyy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (ddmmyyyy) {
      return new Date(
        parseInt(ddmmyyyy[3], 10),
        parseInt(ddmmyyyy[2], 10) - 1,
        parseInt(ddmmyyyy[1], 10),
      ).getTime();
    }
    var yyyymmdd = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (yyyymmdd) {
      return new Date(
        parseInt(yyyymmdd[1], 10),
        parseInt(yyyymmdd[2], 10) - 1,
        parseInt(yyyymmdd[3], 10),
      ).getTime();
    }
  }
  return Number.MAX_VALUE;
}

function getSortKeyMs(row, dateCol, timeCol) {
  var dateMs = parseDateToMs(row[dateCol]);
  if (dateMs === Number.MAX_VALUE) return Number.MAX_VALUE;
  var timeMins = timeCol >= 0 ? parseTimeToMinutes(row[timeCol]) : 0;
  if (timeMins < 0) timeMins = 0;
  return dateMs + timeMins * 60000;
}

function getRihlaSortOrder(val) {
  var v = String(val || "").trim();
  if (v.indexOf("انصراف") !== -1) return 0;
  if (v.indexOf("حضور") !== -1) return 1;
  return 2;
}
