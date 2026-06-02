// ============================================
// 経費申請システム - Google Apps Script バックエンド
// ============================================
// 【セットアップ手順】（たった3ステップ！）
// 1. Google Apps Script (https://script.google.com) で新規プロジェクト作成
// 2. このファイルの内容をCode.gsにコピー
// 3. 関数選択で「fullSetup」を選んで ▶ で実行
//    → フォルダ、スプレッドシート、スクリプトプロパティが全て自動作成されます！
// 4. Web Appとしてデプロイ（「ウェブアプリ」→「全員がアクセス可能」）
// 5. デプロイURLをフロントエンドの script.js にセット
//
// オプション設定（スクリプトプロパティに手動追加）:
//    - GEMINI_API_KEY: Gemini APIキー（インボイス検索用）
//    - GOOGLE_CHAT_WEBHOOK: Google Chat Webhook URL
// ============================================


// === 定数 ===
const PROPS = PropertiesService.getScriptProperties();

function getMasterSpreadsheet() {
  let masterId = PROPS.getProperty('MASTER_SPREADSHEET_ID');
  if (!masterId) {
    // 初期設定が済んでいない場合は自動セットアップを実行
    fullSetup();
    masterId = PROPS.getProperty('MASTER_SPREADSHEET_ID');
  }
  return SpreadsheetApp.openById(masterId);
}

// ============================================
// Web App エンドポイント
// ============================================
function doGet(e) {
  const mode = e.parameter.mode || '';
  let result;

  try {
    switch (mode) {
      case 'get_employees':
        result = getEmployees();
        break;
      case 'get_expenses':
        result = getExpenses(e.parameter.employee, e.parameter.month);
        break;
      case 'get_pending':
        result = getPendingExpenses(e.parameter.approver, e.parameter.department);
        break;
      case 'get_final_pending':
        result = getFinalPendingExpenses();
        break;
      case 'get_admin_completed':
        result = getAdminCompletedExpenses();
        break;
      default:
        result = { error: 'Unknown mode: ' + mode };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;

  try {
    const payload = JSON.parse(e.postData.contents);
    const mode = payload.mode || '';

    switch (mode) {
      case 'submit_expense':
        result = submitExpense(payload.data);
        break;
      case 'submit_expenses_batch':
        result = submitExpensesBatch(payload.data);
        break;
      case 'withdraw_expenses_batch':
        result = withdrawExpensesBatch(payload.expenseIds, payload.employeeName);
        break;
      case 'approve_expense':
        result = approveExpense(payload.expenseId, payload.approver);
        break;
      case 'reject_expense':
        result = rejectExpense(payload.expenseId, payload.rejector, payload.reason);
        break;
      case 'batch_approve_reject':
        result = batchApproveReject(payload.approvals, payload.rejections, payload.approver, payload.employeeName, payload.rejectReason);
        break;
      case 'final_confirm':
        result = finalConfirmExpense(payload.expenseId, payload.confirmer);
        break;
      default:
        result = { error: 'Unknown mode: ' + mode };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 従業員マスタ取得
// ============================================
function getEmployees() {
  const ss = getMasterSpreadsheet();
  const sheet = ss.getSheetByName('従業員マスタ');
  if (!sheet) return { employees: [] };

  const data = sheet.getDataRange().getValues();
  // ヘッダー: 名前, メール, 所属課, 権限(employee/manager/admin), 上席名, 従業員番号
  const employees = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    employees.push({
      name: data[i][0],
      email: data[i][1],
      department: data[i][2],
      role: data[i][3],
      supervisor: data[i][4] || '',
      empId: data[i][5] ? String(data[i][5]) : '',
      position: data[i][6] ? data[i][6] : ''
    });
  }
  return { employees };
}

// ============================================
// 経費申請（登録）
// ============================================
function submitExpense(data) {
  // 1. 経費IDを生成
  const expenseId = 'EXP-' + new Date().getTime();

  // 2. レシート画像をDriveに保存
  let receiptUrl = '';
  if (data.receiptImage) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data.receiptImage),
        'image/jpeg',
        `receipt_${expenseId}.jpg`
      );
      const folderId = PROPS.getProperty('RECEIPT_FOLDER_ID');
      const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      receiptUrl = file.getUrl();
    } catch (err) {
      Logger.log('Receipt save error: ' + err.message);
    }
  }

  // 3. 従業員の経費台帳スプレッドシートを取得/作成
  const year = new Date(data.date).getFullYear();
  const month = new Date(data.date).getMonth() + 1;
  const ssName = `${data.employeeName}（経費台帳）（${year}年度）`;

  const expenseSS = getOrCreateExpenseSpreadsheet(ssName, data.employeeName);
  const sheetName = `${month}月経費`;
  let sheet = expenseSS.getSheetByName(sheetName);
  if (!sheet) {
    sheet = expenseSS.insertSheet(sheetName);
    // ヘッダー行追加
    sheet.appendRow([
      'ID', '利用日', '経費科目', '金額', '店名/会社名',
      'インボイス番号', 'ステータス', '申請日', '承認者',
      '承認日', '備考', 'レシート画像URL', '棄却理由', '支払種別'
    ]);
    // ヘッダー書式
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setFrozenRows(1);
  }

  // 4. データ行追加
  sheet.appendRow([
    expenseId,
    data.date,
    data.category,
    data.amount,
    data.store,
    data.invoiceNumber || '',
    '申請中',
    new Date().toISOString(),
    '',    // 承認者
    '',    // 承認日
    data.memo || '',
    receiptUrl,
    '',    // 棄却理由
    data.paymentMethod || '' // 支払種別
  ]);

  // 5. 上席に通知
  if (data.supervisor) {
    const supervisorEmail = getEmployeeEmail(data.supervisor);
    if (supervisorEmail) {
      sendNotification(
        supervisorEmail,
        `【経費申請】${data.employeeName}さんから経費申請がありました`,
        `${data.employeeName}さんが経費を申請しました。\n\n` +
        `■ 利用日: ${data.date}\n` +
        `■ 科目: ${data.category}\n` +
        `■ 金額: ¥${parseInt(data.amount).toLocaleString()}\n` +
        `■ 店名: ${data.store}\n\n` +
        `経費申請システムにログインして承認をお願いします。`
      );
    }
  }

  return { success: true, expenseId: expenseId };
}

// ============================================
// 経費一括申請（登録）
// ============================================
function submitExpensesBatch(dataArray) {
  if (!dataArray || dataArray.length === 0) return { success: true };

  const results = [];
  const employeeMap = {}; // name -> supervisor

  dataArray.forEach(data => {
    // 1. 経費IDを生成 (フロントで作成した一時IDを上書き)
    const expenseId = 'EXP-' + new Date().getTime() + Math.floor(Math.random() * 1000);

    // 2. レシート画像をDriveに保存
    let receiptUrl = '';
    if (data.receiptImage) {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(data.receiptImage),
          'image/jpeg',
          `receipt_${expenseId}.jpg`
        );
        const folderId = PROPS.getProperty('RECEIPT_FOLDER_ID');
        const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        receiptUrl = file.getUrl();
      } catch (err) {
        Logger.log('Receipt save error: ' + err.message);
      }
    }

    // 3. 従業員の経費台帳スプレッドシートを取得/作成
    const year = new Date(data.date).getFullYear();
    const month = new Date(data.date).getMonth() + 1;
    const ssName = `${data.employeeName}（経費台帳）（${year}年度）`;

    const expenseSS = getOrCreateExpenseSpreadsheet(ssName, data.employeeName);
    const sheetName = `${month}月経費`;
    let sheet = expenseSS.getSheetByName(sheetName);
    if (!sheet) {
      sheet = expenseSS.insertSheet(sheetName);
      // ヘッダー行追加
      sheet.appendRow([
        'ID', '利用日', '経費科目', '金額', '店名/会社名',
        'インボイス番号', 'ステータス', '申請日', '承認者',
        '承認日', '備考', 'レシート画像URL', '棄却理由', '支払種別'
      ]);
      // ヘッダー書式
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setFrozenRows(1);
    }

    // 4. データ行追加
    sheet.appendRow([
      expenseId,
      data.date,
      data.category,
      data.amount,
      data.store,
      data.invoiceNumber || '',
      '申請中',
      new Date().toISOString(),
      data.supervisor || '',    // 承認者（※申請中は提出先として使用）
      '',    // 承認日
      data.memo || '',
      receiptUrl,
      '',    // 棄却理由
      data.paymentMethod || '' // 支払種別
    ]);
    
    // まとめて通知するため記録
    if (data.supervisor) {
      employeeMap[data.employeeName] = data.supervisor;
    }
    
    results.push(expenseId);
  });

  // 5. 上席にまとめて通知
  Object.keys(employeeMap).forEach(empName => {
    try {
      const supervisorName = employeeMap[empName];
      const supervisorEmail = getEmployeeEmail(supervisorName);
      if (supervisorEmail) {
        sendNotification(
          supervisorEmail,
          `【経費申請】${empName}さんから申請がありました`,
          `${empName}さんから複数の経費申請がありました。\n\n` +
          `経費申請システムにログインして、承認待ち一覧をご確認ください。`
        );
      }
    } catch (e) {
      Logger.log('Notification error for ' + empName + ': ' + e.message);
    }
  });

  return { success: true, count: results.length };
}

// ============================================
// 一括取り下げ
// ============================================
function withdrawExpensesBatch(expenseIds, employeeName) {
  if (!expenseIds || expenseIds.length === 0) return { success: true };

  const year = new Date().getFullYear();
  const ssName = `${employeeName}（経費台帳）（${year}年度）`;
  
  try {
    const ss = findSpreadsheetByName(ssName);
    if (!ss) return { error: '台帳が見つかりません' };

    let count = 0;
    ss.getSheets().forEach(sheet => {
      if (sheet.getName() === '経費マスタ') return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const currentId = data[i][0];
        if (expenseIds.includes(currentId)) {
          sheet.getRange(i + 1, 7).setValue('取り下げ');
          count++;
        }
      }
    });

    return { success: true, count: count };
  } catch (err) {
    Logger.log('Withdraw batch error: ' + err.message);
    return { error: err.message };
  }
}

// ============================================
// 経費一覧取得
// ============================================
function getExpenses(employeeName, monthFilter) {
  if (!employeeName) return { expenses: [] };

  const year = new Date().getFullYear();
  const ssName = `${employeeName}（経費台帳）（${year}年度）`;

  try {
    const ss = findSpreadsheetByName(ssName);
    if (!ss) return { expenses: [] };

    const expenses = [];
    const sheets = ss.getSheets();

    sheets.forEach(sheet => {
      if (sheet.getName() === '経費マスタ') return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        expenses.push({
          id: data[i][0],
          employeeName: employeeName,
          date: data[i][1],
          category: data[i][2],
          amount: data[i][3],
          store: data[i][4],
          invoiceNumber: data[i][5],
          status: data[i][6],
          submittedAt: data[i][7],
          approvedBy: data[i][8],
          approvedAt: data[i][9],
          memo: data[i][10],
          receiptUrl: data[i][11],
          rejectReason: data[i][12],
          paymentMethod: data[i][13] || ''
        });
      }
    });

    return { expenses };
  } catch (err) {
    Logger.log('Get expenses error: ' + err.message);
    return { expenses: [] };
  }
}

// ============================================
// 承認待ち経費取得（上席用）
// ============================================
function getPendingExpenses(approverName, department) {
  const year = new Date().getFullYear();
  const employees = getEmployees().employees;

  const pending = [];
  const processedSheets = new Set();

  employees.forEach(emp => {
    // adminの申請は承認の概念がないか別枠とする
    if (emp.role === 'admin') return;

    const ssName = `${emp.name}（経費台帳）（${year}年度）`;
    if (processedSheets.has(ssName)) return; // 既に処理した台帳はスキップ（マスタ二重登録対策）
    processedSheets.add(ssName);

    try {
      const ss = findSpreadsheetByName(ssName);
      if (!ss) return;

      ss.getSheets().forEach(sheet => {
        if (sheet.getName() === '経費マスタ') return;
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][6] === '申請中') {
            const requestedApprover = data[i][8];
            
            // 提出先が明記されている場合は、自分宛てのみ表示
            if (requestedApprover) {
              if (requestedApprover !== approverName) continue;
            } else {
              // 過去のデータ（提出先未指定）の場合は、同一部署の一般社員の物のみ表示
              if (emp.department !== department || emp.role === 'manager') continue;
            }
            pending.push({
              id: data[i][0],
              employeeName: emp.name,
              date: data[i][1],
              category: data[i][2],
              amount: data[i][3],
              store: data[i][4],
              invoiceNumber: data[i][5],
              status: data[i][6],
              memo: data[i][10],
              receiptUrl: data[i][11],
              paymentMethod: data[i][13] || ''
            });
          }
        }
      });
    } catch (err) {
      Logger.log('Pending fetch error for ' + emp.name + ': ' + err.message);
    }
  });

  return { expenses: pending };
}

// ============================================
// 最終確認待ち経費取得（総務長用）
// ============================================
function getFinalPendingExpenses() {
  const year = new Date().getFullYear();
  const employees = getEmployees().employees;

  const pending = [];
  const processedSheets = new Set();
  
  employees.forEach(emp => {
    const ssName = `${emp.name}（経費台帳）（${year}年度）`;
    if (processedSheets.has(ssName)) return; // 重複スキップ
    processedSheets.add(ssName);
    
    try {
      const ss = findSpreadsheetByName(ssName);
      if (!ss) return;

      ss.getSheets().forEach(sheet => {
        if (sheet.getName() === '経費マスタ') return;
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][6] === '承認') {
            pending.push({
              id: data[i][0],
              employeeName: emp.name,
              date: data[i][1],
              category: data[i][2],
              amount: data[i][3],
              store: data[i][4],
              invoiceNumber: data[i][5],
              status: data[i][6],
              memo: data[i][10],
              receiptUrl: data[i][11],
              paymentMethod: data[i][13] || ''
            });
          }
        }
      });
    } catch (err) {
      Logger.log('Final pending fetch error: ' + err.message);
    }
  });

  return { expenses: pending };
}

// ============================================
// 経費承認
// ============================================
function approveExpense(expenseId, approverName) {
  const result = findAndUpdateExpense(expenseId, (row, rowIndex, sheet) => {
    sheet.getRange(rowIndex, 7).setValue('承認');       // ステータス
    sheet.getRange(rowIndex, 9).setValue(approverName); // 承認者
    sheet.getRange(rowIndex, 10).setValue(new Date().toISOString()); // 承認日
    return row;
  });

  if (!result.found) return { error: '経費データが見つかりません' };

  // 申請者の全経費が承認されたかチェック → 総務長に通知
  checkAllApproved(result.employeeName);

  // 申請者に承認通知
  const empEmail = getEmployeeEmail(result.employeeName);
  if (empEmail) {
    sendNotification(
      empEmail,
      `【経費承認】経費が承認されました`,
      `あなたの経費申請（${expenseId}）が${approverName}により承認されました。`
    );
  }

  return { success: true };
}

// ============================================
// 経費棄却
// ============================================
function rejectExpense(expenseId, rejectorName, reason) {
  const result = findAndUpdateExpense(expenseId, (row, rowIndex, sheet) => {
    sheet.getRange(rowIndex, 7).setValue('棄却');    // ステータス
    sheet.getRange(rowIndex, 13).setValue(reason);   // 棄却理由
    return row;
  });

  if (!result.found) return { error: '経費データが見つかりません' };

  // 申請者に棄却通知
  const empEmail = getEmployeeEmail(result.employeeName);
  if (empEmail) {
    sendNotification(
      empEmail,
      `【経費棄却】経費が棄却されました`,
      `あなたの経費申請（${expenseId}）が${rejectorName}により棄却されました。\n\n` +
      `■ 棄却理由: ${reason}\n\n` +
      `内容を修正の上、再度申請してください。`
    );
  }

  return { success: true };
}

// ============================================
// 最終確認
// ============================================
function finalConfirmExpense(expenseId, confirmerName) {
  const result = findAndUpdateExpense(expenseId, (row, rowIndex, sheet) => {
    sheet.getRange(rowIndex, 7).setValue('確認済み');
    return row;
  });

  if (!result.found) return { error: '経費データが見つかりません' };

  return { success: true };
}

// ============================================
// 一括承認・棄却
// ============================================
function batchApproveReject(approvals, rejections, approverName, employeeName, rejectReason) {
  let approvedCount = 0;
  let rejectedCount = 0;

  const year = new Date().getFullYear();
  const ssName = `${employeeName}（経費台帳）（${year}年度）`;
  
  try {
    const ss = findSpreadsheetByName(ssName);
    if (!ss) return { error: '台帳が見つかりません' };

    ss.getSheets().forEach(sheet => {
      if (sheet.getName() === '経費マスタ') return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const currentId = data[i][0];
        if (approvals && approvals.includes(currentId)) {
          sheet.getRange(i + 1, 7).setValue('承認');
          sheet.getRange(i + 1, 9).setValue(approverName);
          sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
          approvedCount++;
        } else if (rejections && rejections.includes(currentId)) {
          sheet.getRange(i + 1, 7).setValue('棄却');
          sheet.getRange(i + 1, 13).setValue(rejectReason || '一括棄却による');
          rejectedCount++;
        }
      }
    });

    // 申請者に通知
    const empEmail = getEmployeeEmail(employeeName);
    if (empEmail) {
      let msg = `あなたの経費申請の一部（または全て）が処理されました。\n\n`;
      if (approvedCount > 0) msg += `■ 承認された件数: ${approvedCount}件\n`;
      if (rejectedCount > 0) {
        msg += `■ 棄却された件数: ${rejectedCount}件\n`;
        msg += `■ 棄却理由: ${rejectReason || '一括処理による棄却'}\n\n`;
      }
      msg += `経費申請システムで詳細をご確認ください。`;

      sendNotification(
        empEmail,
        `【経費申請結果通知】経費申請の処理が完了しました`,
        msg
      );
    }

    // 承認されたものがある場合、総務長と社長に通知する
    if (approvedCount > 0) {
      const employees = getEmployees().employees;
      let notifyEmails = [];
      employees.forEach(e => {
        if (e.role === 'admin' || e.position === '社長') {
          if (e.email) notifyEmails.push(e.email);
        }
      });
      
      notifyEmails = [...new Set(notifyEmails)];
      if (notifyEmails.length > 0) {
        sendNotification(
          notifyEmails.join(','),
          `【経費承認完了】${employeeName}さんの経費が上席により承認されました`,
          `${employeeName}さんの経費（${approvedCount}件）が、${approverName}により承認されました。\n\n` +
          `経費申請システムにログインして最終確認をお願いします。`
        );
      }
    }

    return { success: true, approvedCount, rejectedCount };
  } catch (err) {
    Logger.log('Batch approve/reject error: ' + err.message);
    return { error: err.message };
  }
}

// ============================================
// ============================================
// 承認済み経費取得（Admin用）
// ============================================
function getAdminCompletedExpenses() {
  const year = new Date().getFullYear();
  const employees = getEmployees().employees;

  const completed = [];
  const processedSheets = new Set();
  
  employees.forEach(emp => {
    const ssName = `${emp.name}（経費台帳）（${year}年度）`;
    if (processedSheets.has(ssName)) return; // 重複スキップ
    processedSheets.add(ssName);
    
    try {
      const ss = findSpreadsheetByName(ssName);
      if (!ss) return;

      ss.getSheets().forEach(sheet => {
        if (sheet.getName() === '経費マスタ') return;
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][6] === '確認済み') {
            completed.push({
              id: data[i][0],
              employeeName: emp.name,
              date: data[i][1],
              category: data[i][2],
              amount: data[i][3],
              store: data[i][4],
              invoiceNumber: data[i][5],
              status: data[i][6],
              approver: data[i][8] || '', // 承認者名
              approvedAt: data[i][9] || '', // 承認日時
              memo: data[i][10],
              receiptUrl: data[i][11],
              paymentMethod: data[i][13] || ''
            });
          }
        }
      });
    } catch (err) {
      Logger.log('Admin completed fetch error: ' + err.message);
    }
  });

  return { expenses: completed };
}

// ============================================
// ヘルパー関数
// ============================================

// 経費台帳スプレッドシートを取得または作成
function getOrCreateExpenseSpreadsheet(ssName, employeeName) {
  let ss = findSpreadsheetByName(ssName);
  if (ss) return ss;

  // 新規作成
  const folderId = PROPS.getProperty('EXPENSE_FOLDER_ID');
  ss = SpreadsheetApp.create(ssName);

  if (folderId) {
    const file = DriveApp.getFileById(ss.getId());
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }

  // 経費マスタシート作成
  let masterSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  masterSheet.setName('経費マスタ');
  masterSheet.appendRow(['従業員名', employeeName]);
  masterSheet.appendRow(['年度', new Date().getFullYear()]);
  masterSheet.appendRow(['作成日', new Date().toISOString()]);

  // 1〜12月シートを作成
  for (let m = 1; m <= 12; m++) {
    const monthSheet = ss.insertSheet(`${m}月経費`);
    monthSheet.appendRow([
      'ID', '利用日', '経費科目', '金額', '店名/会社名',
      'インボイス番号', 'ステータス', '申請日', '承認者',
      '承認日', '備考', 'レシート画像URL', '棄却理由', '支払種別'
    ]);
    monthSheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f0f0f0');
    monthSheet.setFrozenRows(1);
  }

  return ss;
}

// スプレッドシートを名前で検索
function findSpreadsheetByName(name) {
  const folderId = PROPS.getProperty('EXPENSE_FOLDER_ID');
  let files;
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    files = folder.getFilesByName(name);
  } else {
    files = DriveApp.getFilesByName(name);
  }

  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  return null;
}

// 従業員のメールアドレスを取得
function getEmployeeEmail(name) {
  const employees = getEmployees().employees;
  const emp = employees.find(e => e.name === name);
  return emp ? emp.email : null;
}

// 経費データを検索・更新
function findAndUpdateExpense(expenseId, updateFn) {
  const year = new Date().getFullYear();
  const employees = getEmployees().employees;

  for (const emp of employees) {
    const ssName = `${emp.name}（経費台帳）（${year}年度）`;
    try {
      const ss = findSpreadsheetByName(ssName);
      if (!ss) continue;

      const sheets = ss.getSheets();
      for (const sheet of sheets) {
        if (sheet.getName() === '経費マスタ') continue;
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === expenseId) {
            updateFn(data[i], i + 1, sheet);
            return { found: true, employeeName: emp.name };
          }
        }
      }
    } catch (err) {
      Logger.log('Find expense error: ' + err.message);
    }
  }

  return { found: false };
}

// 全経費が承認されたかチェック → 総務長に通知
function checkAllApproved(employeeName) {
  const year = new Date().getFullYear();
  const month = new Date().getMonth() + 1;
  const ssName = `${employeeName}（経費台帳）（${year}年度）`;

  try {
    const ss = findSpreadsheetByName(ssName);
    if (!ss) return;

    const sheet = ss.getSheetByName(`${month}月経費`);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    let allApproved = true;
    let hasExpenses = false;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      hasExpenses = true;
      if (data[i][6] !== '承認') {
        allApproved = false;
        break;
      }
    }

    if (hasExpenses && allApproved) {
      // 総務長に通知
      const employees = getEmployees().employees;
      const admin = employees.find(e => e.role === 'admin');
      if (admin && admin.email) {
        sendNotification(
          admin.email,
          `【経費全承認】${employeeName}さんの${month}月経費が全て承認されました`,
          `${employeeName}さんの${month}月分の経費が全て上席により承認されました。\n\n` +
          `経費申請システムにログインして最終確認をお願いします。`
        );
      }
    }
  } catch (err) {
    Logger.log('Check all approved error: ' + err.message);
  }
}

// ============================================
// 通知送信
// ============================================
function sendNotification(email, subject, body) {
  // メール通知
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });
  } catch (err) {
    Logger.log('Mail send error: ' + err.message);
  }

  // Google Chat Webhook（設定されている場合）
  const webhookUrl = PROPS.getProperty('GOOGLE_CHAT_WEBHOOK');
  if (webhookUrl) {
    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: `*${subject}*\n\n${body}`
        })
      });
    } catch (err) {
      Logger.log('Chat webhook error: ' + err.message);
    }
  }
}

// ============================================
// ★ 初期セットアップ（これを1回実行するだけで全て構築）
// ============================================
// 【使い方】
// 1. Google Apps Script で新規プロジェクトを作成
// 2. このファイルの内容を貼り付け
// 3. 関数選択で「fullSetup」を選んで ▶ で実行
// 4. Google Drive / Spreadsheet への権限を承認
// 5. ログ（表示 → ログ）に出力されたWeb App URLを
//    フロントエンドの script.js の GAS_WEBAPP_URL に設定
// ============================================

/**
 * ★ ワンクリックフルセットアップ
 * 以下を全て自動作成します：
 *   1. Google Drive に「経費申請システム」フォルダ
 *   2. サブフォルダ「経費台帳」「レシート画像」
 *   3. マスタ用スプレッドシート（従業員マスタ、経費科目マスタ、インボイスキャッシュ）
 *   4. サンプル従業員の経費台帳スプレッドシート
 *   5. スクリプトプロパティを全て自動設定
 */
function fullSetup() {
  Logger.log('=== 経費申請システム フルセットアップ開始 ===');

  // ─── 1. ルートフォルダ作成 ───
  let rootFolder;
  const existingFolders = DriveApp.getFoldersByName('経費申請システム');
  if (existingFolders.hasNext()) {
    rootFolder = existingFolders.next();
    Logger.log('✓ 既存フォルダ「経費申請システム」を使用');
  } else {
    rootFolder = DriveApp.createFolder('経費申請システム');
    Logger.log('✓ フォルダ「経費申請システム」を作成');
  }

  // ─── 2. サブフォルダ作成 ───
  const expenseFolder = getOrCreateSubFolder_(rootFolder, '経費台帳');
  const receiptFolder = getOrCreateSubFolder_(rootFolder, 'レシート画像');
  Logger.log('✓ サブフォルダ「経費台帳」「レシート画像」を作成/確認');

  // ─── 3. マスタ用スプレッドシート作成 ───
  let masterSS = findSpreadsheetInFolder_(rootFolder, '経費申請システム_マスタ');
  if (!masterSS) {
    masterSS = SpreadsheetApp.create('経費申請システム_マスタ');
    moveFileToFolder_(masterSS.getId(), rootFolder);
    Logger.log('✓ マスタスプレッドシートを作成');
  } else {
    Logger.log('✓ 既存マスタスプレッドシートを使用');
  }

  // --- 従業員マスタシート ---
  let empSheet = masterSS.getSheetByName('従業員マスタ');
  if (!empSheet) {
    // 既存のSheet1をリネームして使用
    const defaultSheet = masterSS.getSheetByName('Sheet1') || masterSS.getSheets()[0];
    if (defaultSheet && defaultSheet.getLastRow() === 0) {
      defaultSheet.setName('従業員マスタ');
      empSheet = defaultSheet;
    } else {
      empSheet = masterSS.insertSheet('従業員マスタ');
    }
    empSheet.appendRow(['名前', 'メール', '所属課', '権限', '上席名', '従業員番号']);
    empSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f0f0');
    empSheet.setFrozenRows(1);

    // サンプル従業員（後でメールアドレスを書き換えてください）
    empSheet.appendRow(['塩野谷圭介', 'shionoya@example.com', '工事課', 'employee', '山田太郎', '1001']);
    empSheet.appendRow(['山田太郎', 'yamada@example.com', '工事課', 'manager', '', '1002']);
    empSheet.appendRow(['佐藤花子', 'sato@example.com', '総務課', 'admin', '', '1003']);

    // 列幅を調整
    empSheet.setColumnWidth(1, 150);
    empSheet.setColumnWidth(2, 250);
    empSheet.setColumnWidth(3, 120);
    empSheet.setColumnWidth(4, 100);
    empSheet.setColumnWidth(5, 150);
    empSheet.setColumnWidth(6, 100);

    Logger.log('✓ 従業員マスタシートを作成（サンプル3名登録済み）');
  }

  // --- 経費科目マスタシート ---
  let catSheet = masterSS.getSheetByName('経費科目マスタ');
  if (!catSheet) {
    catSheet = masterSS.insertSheet('経費科目マスタ');
    catSheet.appendRow(['科目コード', '科目名']);
    catSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');
    catSheet.setFrozenRows(1);

    const categories = [
      '交通費', '宿泊費', '会議費', '交際費', '消耗品費',
      '通信費', '図書費', '雑費', '飲食費', '備品費', 'その他'
    ];
    categories.forEach((cat, i) => {
      catSheet.appendRow([`CAT-${String(i + 1).padStart(3, '0')}`, cat]);
    });
    Logger.log('✓ 経費科目マスタシートを作成（11科目登録済み）');
  }

  // --- インボイスキャッシュシート ---
  let invSheet = masterSS.getSheetByName('インボイスキャッシュ');
  if (!invSheet) {
    invSheet = masterSS.insertSheet('インボイスキャッシュ');
    invSheet.appendRow(['会社名/店名', 'インボイス登録番号', '登録日']);
    invSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
    invSheet.setFrozenRows(1);
    Logger.log('✓ インボイスキャッシュシートを作成');
  }

  // 不要なSheet1を削除
  try {
    const sheet1 = masterSS.getSheetByName('Sheet1');
    if (sheet1 && masterSS.getSheets().length > 1) {
      masterSS.deleteSheet(sheet1);
    }
  } catch(e) { /* 無視 */ }

  // ─── 4. サンプル従業員の経費台帳を作成 ───
  const year = new Date().getFullYear();
  const sampleEmployees = ['塩野谷圭介', '山田太郎', '佐藤花子'];

  sampleEmployees.forEach(name => {
    const ssName = `${name}（経費台帳）（${year}年度）`;
    let existingSS = findSpreadsheetInFolder_(expenseFolder, ssName);
    if (!existingSS) {
      const newSS = SpreadsheetApp.create(ssName);
      moveFileToFolder_(newSS.getId(), expenseFolder);

      // 経費マスタシート（Sheet1をリネーム）
      const ms = newSS.getSheets()[0];
      ms.setName('経費マスタ');
      ms.appendRow(['従業員名', name]);
      ms.appendRow(['年度', year]);
      ms.appendRow(['作成日', new Date().toLocaleDateString('ja-JP')]);

      // 1月〜12月のシートを作成
      for (let m = 1; m <= 12; m++) {
        const monthSheet = newSS.insertSheet(`${m}月経費`);
        monthSheet.appendRow([
          'ID', '利用日', '経費科目', '金額', '店名/会社名',
          'インボイス番号', 'ステータス', '申請日', '承認者',
          '承認日', '備考', 'レシート画像URL', '棄却理由', '支払種別'
        ]);
        monthSheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f0f0f0');
        monthSheet.setFrozenRows(1);
        // 列幅調整
        monthSheet.setColumnWidth(1, 160);  // ID
        monthSheet.setColumnWidth(2, 120);  // 利用日
        monthSheet.setColumnWidth(3, 100);  // 経費科目
        monthSheet.setColumnWidth(4, 100);  // 金額
        monthSheet.setColumnWidth(5, 180);  // 店名
        monthSheet.setColumnWidth(6, 160);  // インボイス番号
        monthSheet.setColumnWidth(7, 80);   // ステータス
      }

      Logger.log(`✓ 経費台帳「${ssName}」を作成（12ヶ月分のシート付き）`);
    } else {
      Logger.log(`✓ 経費台帳「${ssName}」は既に存在`);
    }
  });

  // ─── 5. スクリプトプロパティを自動設定 ───
  PROPS.setProperty('MASTER_SPREADSHEET_ID', masterSS.getId());
  PROPS.setProperty('EXPENSE_FOLDER_ID', expenseFolder.getId());
  PROPS.setProperty('RECEIPT_FOLDER_ID', receiptFolder.getId());
  Logger.log('✓ スクリプトプロパティを自動設定完了');

  // ─── 結果サマリー ───
  Logger.log('');
  Logger.log('============================================');
  Logger.log('★ セットアップ完了！');
  Logger.log('============================================');
  Logger.log('');
  Logger.log('📁 Google Drive 構成:');
  Logger.log(`   経費申請システム/`);
  Logger.log(`   ├── 経費申請システム_マスタ (SS)`);
  Logger.log(`   ├── 経費台帳/`);
  sampleEmployees.forEach(name => {
    Logger.log(`   │   └── ${name}（経費台帳）（${year}年度）`);
  });
  Logger.log(`   └── レシート画像/`);
  Logger.log('');
  Logger.log('📝 スクリプトプロパティ:');
  Logger.log(`   MASTER_SPREADSHEET_ID = ${masterSS.getId()}`);
  Logger.log(`   EXPENSE_FOLDER_ID     = ${expenseFolder.getId()}`);
  Logger.log(`   RECEIPT_FOLDER_ID     = ${receiptFolder.getId()}`);
  Logger.log('');
  Logger.log('【次のステップ】');
  Logger.log('1. 従業員マスタのメールアドレスを実際のものに更新してください');
  Logger.log(`   → ${masterSS.getUrl()}`);
  Logger.log('2. このプロジェクトをWeb Appとしてデプロイしてください');
  Logger.log('   → デプロイ → 新しいデプロイ → ウェブアプリ → アクセス:全員');
  Logger.log('3. デプロイURLをフロントエンドの script.js の GAS_WEBAPP_URL に設定');
  Logger.log('============================================');
}

// ─── セットアップ用 ヘルパー関数 ───

function getOrCreateSubFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function findSpreadsheetInFolder_(folder, name) {
  const files = folder.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  return null;
}

function moveFileToFolder_(fileId, folder) {
  const file = DriveApp.getFileById(fileId);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}
