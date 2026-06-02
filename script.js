// ============================================
// 経費申請システム - メインスクリプト
// ============================================

// === 設定 ===
// GAS Web App URLをここに設定（デプロイ後に更新）
const GAS_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbyjD_TYlD0od6110kE7zTRXKjfuSsmQjPE5wPmDeVd56ZhSS9vpJZl159t0v-jsar61/exec';

// Gemini API Key（OCR用）
const GEMINI_API_KEY = 'AIzaSyDcTpPVlwqhfMk8tVliS11cyhFlyoihZmw';
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';

// ============================================
// ユーティリティ
// ============================================
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

function formatCurrency(num) {
    if (!num && num !== 0) return '-';
    return '¥ ' + parseInt(num).toLocaleString();
}

function formatDateJP(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

function getStatusBadge(status) {
    const map = {
        '申請中': 'status-pending',
        '承認': 'status-approved',
        '棄却': 'status-rejected',
        '確認済み': 'status-confirmed'
    };
    const cls = map[status] || 'status-pending';
    return `<span class="status-badge ${cls}">${status}</span>`;
}

// ============================================
// アプリ状態管理
// ============================================
let currentUser = null; // { name, email, department, role, supervisor }
let capturedImageBase64 = null;

function saveLogin(user) {
    currentUser = user;
    localStorage.setItem('expenseUser', JSON.stringify(user));
}

function loadLogin() {
    const saved = localStorage.getItem('expenseUser');
    if (saved) {
        currentUser = JSON.parse(saved);
        return true;
    }
    return false;
}

function logout() {
    currentUser = null;
    localStorage.removeItem('expenseUser');
    document.getElementById('login-view').style.display = 'block';
    document.getElementById('main-view').style.display = 'none';
}

// ============================================
// Googleログインコールバック
// ============================================
window.handleGoogleLogin = function(response) {
    if (!window.globalEmployees || window.globalEmployees.length === 0) {
        alert('従業員データを読み込み中です。少々お待ちください。');
        return;
    }
    
    try {
        if (!response || !response.credential) {
            alert('Googleからの認証情報が受け取れませんでした。');
            return;
        }

        const base64Url = response.credential.split('.')[1];
        let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const padLength = (4 - (base64.length % 4)) % 4;
        base64 += '='.repeat(padLength);
        const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
        
        const payload = JSON.parse(jsonPayload);
        const email = payload.email.toLowerCase();
        
        const user = window.globalEmployees.find(e => e.email.toLowerCase() === email);
        if (!user) {
            alert(`未登録のアカウントです:\n${email}`);
            return;
        }
        
        saveLogin(user);
        if (window.showMainViewFunc) {
            window.showMainViewFunc();
            showToast(`${user.name}さん、ようこそ！`, 'success');
        } else {
            alert('初期化エラー: 画面の切り替え準備ができていません。');
        }
        
    } catch (error) {
        console.error('Login error:', error);
        alert('ログイン処理内部でエラーが発生しました:\n' + error.message);
    }
};

// ============================================
// メインロジック
// ============================================
document.addEventListener('DOMContentLoaded', async () => {

    // --- DOM要素 ---
    const loginView = document.getElementById('login-view');
    const mainView = document.getElementById('main-view');
    const btnLogout = document.getElementById('btn-logout');
    const currentUserName = document.getElementById('current-user-name');

    // タブ
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    const mainTabs = document.getElementById('main-tabs');

    // カメラ
    const cameraArea = document.getElementById('camera-area');
    const cameraPlaceholder = document.getElementById('camera-placeholder');
    const receiptInput = document.getElementById('receipt-input');
    const receiptPreview = document.getElementById('receipt-preview');
    const cameraActions = document.getElementById('camera-actions');
    const btnRetake = document.getElementById('btn-retake');
    const btnOcr = document.getElementById('btn-ocr');
    const ocrLoading = document.getElementById('ocr-loading');

    // フォーム
    const expenseForm = document.getElementById('expense-form');
    const expenseDate = document.getElementById('expense-date');
    const expenseCategory = document.getElementById('expense-category');
    const expensePaymentMethod = document.getElementById('expense-payment-method');
    const expenseAmount = document.getElementById('expense-amount');
    const expenseStore = document.getElementById('expense-store');
    const expenseInvoice = document.getElementById('expense-invoice');
    const expenseMemo = document.getElementById('expense-memo');

    // 一覧
    const listLoading = document.getElementById('list-loading');
    const expenseTableWrapper = document.getElementById('expense-table-wrapper');
    const expenseTbody = document.getElementById('expense-tbody');
    const monthDropdown = document.getElementById('month-dropdown');

    // 承認
    const approveLoading = document.getElementById('approve-loading');
    const approveTableWrapper = document.getElementById('approve-table-wrapper');
    const approveTbody = document.getElementById('approve-tbody');
    const approveTitle = document.getElementById('approve-title');
    const tabApprovalBtn = document.querySelector('.tab-approval');

    // モーダル
    const rejectModal = document.getElementById('reject-modal');
    const rejectReason = document.getElementById('reject-reason');
    const btnRejectCancel = document.getElementById('btn-reject-cancel');
    const btnRejectConfirm = document.getElementById('btn-reject-confirm');
    const detailModal = document.getElementById('detail-modal');
    const detailBody = document.getElementById('detail-body');
    const btnDetailClose = document.getElementById('btn-detail-close');

    // 下書きリスト関連
    let draftExpenses = [];
    const draftListSection = document.getElementById('draft-list-section');
    const draftTbody = document.getElementById('draft-tbody');
    const draftCount = document.getElementById('draft-count');
    const btnSubmitAll = document.getElementById('btn-submit-all');
    const submissionTarget = document.getElementById('submission-target');
    const draftTotalAmountSpan = document.getElementById('draft-total-amount');
    const expenseTotalAmountSpan = document.getElementById('expense-total-amount');
    const approveContainer = document.getElementById('approve-container');

    // --- 発行日初期値 ---
    const today = new Date().toISOString().split('T')[0];
    expenseDate.value = today;

    // ============================================
    // 金額フォーマット（全角→半角、カンマ区切り）
    // ============================================
    let isAmountComposing = false;
    expenseAmount.addEventListener('compositionstart', () => { isAmountComposing = true; });
    expenseAmount.addEventListener('compositionend', (e) => {
        isAmountComposing = false;
        formatAmountInput(e);
    });
    expenseAmount.addEventListener('input', (e) => {
        if (isAmountComposing) return;
        formatAmountInput(e);
    });

    function formatAmountInput(e) {
        let value = e.target.value.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        let rawStr = value.replace(/,/g, '').replace(/[^0-9]/g, '');
        if (rawStr !== '') {
            e.target.value = parseInt(rawStr, 10).toLocaleString();
        } else {
            e.target.value = '';
        }
    }



    // ============================================
    // 従業員マスタ取得 & ログイン（キャッシュ付き高速化）
    // ============================================
    const EMPLOYEE_CACHE_KEY = 'cachedEmployees';
    const EMPLOYEE_CACHE_TS_KEY = 'cachedEmployeesTimestamp';
    const CACHE_MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24時間

    function getCachedEmployees() {
        try {
            const cached = localStorage.getItem(EMPLOYEE_CACHE_KEY);
            const ts = localStorage.getItem(EMPLOYEE_CACHE_TS_KEY);
            if (cached && ts) {
                const age = Date.now() - parseInt(ts);
                if (age < CACHE_MAX_AGE_MS) {
                    return JSON.parse(cached);
                }
            }
        } catch (e) { /* ignore */ }
        return null;
    }

    function setCachedEmployees(employees) {
        try {
            localStorage.setItem(EMPLOYEE_CACHE_KEY, JSON.stringify(employees));
            localStorage.setItem(EMPLOYEE_CACHE_TS_KEY, String(Date.now()));
        } catch (e) { /* ignore */ }
    }

    async function fetchEmployeesFromGAS() {
        try {
            const res = await fetch(`${GAS_WEBAPP_URL}?mode=get_employees`);
            if (!res.ok) throw new Error('Failed to fetch employees');
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            if (data.employees && data.employees.length > 0) {
                return data.employees;
            }
        } catch (e) {
            console.error('Employee fetch error:', e);
            showToast('通信エラー: ' + e.message, 'error');
            const select = document.getElementById('local-login-user');
            if (select) select.innerHTML = `<option value="">設定エラーが発生しています（${e.message}）</option>`;
        }
        return [];
    }

    async function loadEmployees() {
        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const demoEmployees = [
                { name: '塩野谷圭介', email: 'shionoya@example.com', department: '工事課', role: 'employee', supervisor: '山田太郎', empId: '1001' },
                { name: '山田太郎', email: 'yamada@example.com', department: '工事課', role: 'manager', supervisor: '', empId: '1002' },
                { name: '佐藤花子', email: 'sato@example.com', department: '総務課', role: 'admin', supervisor: '', empId: '1003' },
            ];
            return demoEmployees;
        }

        // キャッシュがあれば即座に返し、バックグラウンドで最新を取得
        const cached = getCachedEmployees();
        if (cached && cached.length > 0) {
            // バックグラウンドで最新データを取得・キャッシュ更新
            fetchEmployeesFromGAS().then(fresh => {
                if (fresh && fresh.length > 0) {
                    setCachedEmployees(fresh);
                    window.globalEmployees = fresh;
                }
            });
            return cached;
        }

        // キャッシュがない場合はGASから取得して待つ
        const employees = await fetchEmployeesFromGAS();
        if (employees && employees.length > 0) {
            setCachedEmployees(employees);
        }
        return employees;
    }

    function applyEmployees(employees) {
        window.globalEmployees = employees;
        
        if (submissionTarget) {
            // 既存のoptionをクリア（デフォルトの「選択してください」以外）
            while (submissionTarget.options.length > 1) {
                submissionTarget.remove(1);
            }
            employees.forEach(emp => {
                if (emp.position === '課長' || emp.role === 'manager') {
                    const opt = document.createElement('option');
                    opt.value = emp.name;
                    opt.textContent = `${emp.name}（${emp.department}）`;
                    submissionTarget.appendChild(opt);
                }
            });
        }
        
        window.employeesLoaded = true;
    }

    loadEmployees().then(employees => {
        applyEmployees(employees);
    });

    // 自動ログインチェック
    if (loadLogin()) {
        showMainView();
    }

    // ローカル環境（file://）でのテストログインフォールバック
    if (window.location.protocol === 'file:') {
        const fallbackDiv = document.getElementById('local-login-fallback');
        if (fallbackDiv) {
            fallbackDiv.style.display = 'block';
            
            // 従業員データがロードされるのを待つ
            const checkInterval = setInterval(() => {
                if (window.employeesLoaded) {
                    clearInterval(checkInterval);
                    const select = document.getElementById('local-login-user');
                    if (window.globalEmployees && window.globalEmployees.length > 0) {
                        select.innerHTML = '<option value="">ユーザーを選択してログイン</option>';
                        window.globalEmployees.forEach(emp => {
                            const opt = document.createElement('option');
                            opt.value = emp.email;
                            opt.textContent = `${emp.name}（${emp.department}/${emp.role}）`;
                            select.appendChild(opt);
                        });
                    } else if (select.innerHTML.includes('読み込み中')) {
                        select.innerHTML = '<option value="">データがありません（GAS側のエラーの可能性があります）</option>';
                    }
                }
            }, 500);

            document.getElementById('btn-local-login').addEventListener('click', () => {
                const select = document.getElementById('local-login-user');
                const email = select.value;
                if (!email) {
                    alert('ユーザーを選択してください');
                    return;
                }
                const user = window.globalEmployees.find(e => e.email === email);
                if (user) {
                    saveLogin(user);
                    showMainView();
                    showToast(`${user.name}さんとしてテストログインしました`, 'success');
                } else {
                    alert('選択されたユーザーが見つかりません。');
                }
            });
        }
    }

    btnLogout.addEventListener('click', () => {
        logout();
        showToast('ログアウトしました', 'info');
    });

    // 他から呼べるようにする
    window.showMainViewFunc = showMainView;
    function showMainView() {
        loginView.style.display = 'none';
        mainView.style.display = 'block';
        currentUserName.textContent = `${currentUser.name}（${currentUser.department}）`;

        // 権限に応じて承認系のタブ表示
        const tabCompletedBtn = document.querySelector('.tab-completed');
        if (currentUser.role === 'manager' || currentUser.role === 'admin' || currentUser.position === '社長') {
            tabApprovalBtn.style.display = 'inline-block';
            if (currentUser.role === 'admin' || currentUser.position === '社長') {
                approveTitle.textContent = '最終確認待ち経費';
                if (tabCompletedBtn) tabCompletedBtn.style.display = 'inline-block';
            } else {
                approveTitle.textContent = '承認待ち経費';
                if (tabCompletedBtn) tabCompletedBtn.style.display = 'none';
            }
        } else {
            tabApprovalBtn.style.display = 'none';
            if (tabCompletedBtn) tabCompletedBtn.style.display = 'none';
        }
    }

    // ============================================
    // タブ切り替え
    // ============================================
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.dataset.tab;
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            tabContents.forEach(tc => tc.style.display = 'none');
            document.getElementById(targetTab).style.display = 'block';

            // 幅切替
            if (targetTab === 'tab-new') {
                mainTabs.classList.remove('wide');
            } else {
                mainTabs.classList.add('wide');
            }

            // データロード
            if (targetTab === 'tab-list') {
                loadExpenseList();
            } else if (targetTab === 'tab-approve') {
                loadPendingApprovals();
            } else if (targetTab === 'tab-completed') {
                if (typeof loadAdminCompletedExpenses === 'function') {
                    loadAdminCompletedExpenses();
                }
            }
        });
    });

    // ============================================
    // カメラ撮影 & OCR
    // ============================================
    cameraArea.addEventListener('click', () => {
        receiptInput.click();
    });

    receiptInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (ev) => {
            receiptPreview.src = ev.target.result;
            receiptPreview.style.display = 'block';
            cameraPlaceholder.style.display = 'none';
            cameraActions.style.display = 'flex';

            // Base64データを保存（data:image/...;base64,を除く）
            capturedImageBase64 = ev.target.result.split(',')[1];
        };
        reader.readAsDataURL(file);
    });

    btnRetake.addEventListener('click', () => {
        receiptPreview.style.display = 'none';
        receiptPreview.src = '';
        cameraPlaceholder.style.display = 'block';
        cameraActions.style.display = 'none';
        capturedImageBase64 = null;
        receiptInput.value = '';
    });

    btnOcr.addEventListener('click', async () => {
        if (!capturedImageBase64) {
            showToast('レシートを撮影してください', 'error');
            return;
        }

        if (GEMINI_API_KEY.includes('YOUR_GEMINI')) {
            // デモモード
            showToast('デモモード：ダミーOCRデータを入力します', 'info');
            expenseDate.value = today;
            expenseCategory.value = '飲食費';
            expenseAmount.value = '1,500';
            expenseStore.value = 'スターバックス渋谷店';
            expenseInvoice.value = 'T1234567890123';
            return;
        }

        ocrLoading.style.display = 'block';
        cameraActions.style.display = 'none';
        btnOcr.disabled = true;

        try {
            const response = await fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{
                        parts: [
                            {
                                text: `このレシート／領収書の画像から以下の情報をJSON形式で抽出してください。
必ず以下のキーで返してください：
{
  "date": "YYYY-MM-DD形式の利用日",
  "category": "経費科目（交通費/宿泊費/会議費/交際費/消耗品費/通信費/図書費/雑費/飲食費/備品費/その他 のいずれか）",
  "paymentMethod": "支払種別（現金/クレジットカード（会社）/クレジットカード（個人）のいずれか。不明な場合は空文字を返してください。Visa, Masterなどのブランド名やクレジット売上表記があればクレジットカード（個人）としてください）",
  "amount": 合計金額（税込、数値のみ）,
  "store": "店名または会社名",
  "invoice_number": "インボイス登録番号（T+13桁の数字。記載がなければ空文字）"
}
JSONのみを返してください。説明文は不要です。`
                            },
                            {
                                inlineData: {
                                    mimeType: 'image/jpeg',
                                    data: capturedImageBase64
                                }
                            }
                        ]
                    }]
                })
            });

            if (!response.ok) throw new Error(`Gemini API error: ${response.status}`);
            const data = await response.json();

            // レスポンスからJSONを抽出
            let text = data.candidates[0].content.parts[0].text;
            // コードブロックを除去
            text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
            const ocrResult = JSON.parse(text);

            // フォームに反映
            if (ocrResult.date) expenseDate.value = ocrResult.date;
            if (ocrResult.category) expenseCategory.value = ocrResult.category;
            if (ocrResult.paymentMethod) expensePaymentMethod.value = ocrResult.paymentMethod;
            if (ocrResult.amount) {
                expenseAmount.value = parseInt(ocrResult.amount).toLocaleString();
            }
            if (ocrResult.store) expenseStore.value = ocrResult.store;
            if (ocrResult.invoice_number) {
                expenseInvoice.value = ocrResult.invoice_number;
            }

            showToast('レシートの内容を読み取りました！', 'success');

        } catch (error) {
            console.error('OCR error:', error);
            showToast('レシートの読み取りに失敗しました: ' + error.message, 'error');
        } finally {
            ocrLoading.style.display = 'none';
            cameraActions.style.display = 'flex';
            btnOcr.disabled = false;
        }
    });

    // ============================================



    // ============================================
    // 経費下書きへ追加
    // ============================================
    expenseForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const rawAmount = expenseAmount.value.replace(/,/g, '');
        if (!rawAmount || isNaN(parseInt(rawAmount))) {
            showToast('金額を正しく入力してください', 'error');
            return;
        }

        const inputDate = expenseDate.value;
        const inputCategory = expenseCategory.value;
        const inputStore = expenseStore.value.trim();
        const inputAmt = parseInt(rawAmount);

        // 重複チェック
        const isDraftDuplicate = draftExpenses.some(exp => exp.date === inputDate && exp.category === inputCategory && exp.amount === inputAmt && exp.store === inputStore);
        const isHistoryDuplicate = (window.currentExpenses || []).some(exp => exp.date === inputDate && exp.category === inputCategory && exp.amount === inputAmt && exp.store === inputStore);
        
        if (isDraftDuplicate || isHistoryDuplicate) {
            const confirmMsg = "同じ利用日・科目・金額・店名の経費がすでに存在します。\n\n本当にリストへ追加しますか？\n（※同じ日に複数回利用した場合はOKを押してください）";
            if (!confirm(confirmMsg)) {
                return;
            }
        }

        const expenseData = {
            id: 'EXP-TEMP-' + Date.now() + Math.floor(Math.random() * 1000),
            employeeName: currentUser.name,
            employeeEmail: currentUser.email,
            department: currentUser.department,
            supervisor: currentUser.supervisor || '',
            date: expenseDate.value,
            category: expenseCategory.value,
            paymentMethod: expensePaymentMethod.value,
            amount: parseInt(rawAmount),
            store: expenseStore.value.trim(),
            invoiceNumber: expenseInvoice.value.trim(),
            memo: expenseMemo.value.trim(),
            receiptImage: capturedImageBase64 || '',
            status: '申請中',
            submittedAt: new Date().toISOString()
        };

        draftExpenses.push(expenseData);
        updateDraftUI();
        showToast('リストに追加しました。さらに自動入力で連続追加できます', 'success');
        resetForm();
    });

    function updateDraftUI() {
        if (draftExpenses.length > 0) {
            draftListSection.style.display = 'block';
            draftCount.textContent = draftExpenses.length;
            
            draftTbody.innerHTML = '';
            let total = 0;
            draftExpenses.forEach((exp, index) => {
                if (exp.paymentMethod !== 'クレジットカード（会社）') {
                    total += parseInt(exp.amount) || 0;
                }
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${formatDateJP(exp.date)}</td>
                    <td>${exp.category}</td>
                    <td style="text-align:right;">${formatCurrency(exp.amount)}</td>
                    <td>
                        <button type="button" class="action-btn" onclick="editDraft(${index})" style="margin-right:5px; background:var(--primary);">編集</button>
                        <button type="button" class="action-btn reject" onclick="removeDraft(${index})">削除</button>
                    </td>
                `;
                draftTbody.appendChild(tr);
            });
            if (draftTotalAmountSpan) draftTotalAmountSpan.textContent = formatCurrency(total);
        } else {
            draftListSection.style.display = 'none';
        }
    }

    window.removeDraft = function(index) {
        draftExpenses.splice(index, 1);
        updateDraftUI();
    };

    window.editDraft = function(index) {
        const exp = draftExpenses[index];
        expenseDate.value = exp.date;
        expenseCategory.value = exp.category;
        expensePaymentMethod.value = exp.paymentMethod;
        expenseAmount.value = exp.amount.toLocaleString();
        expenseStore.value = exp.store;
        expenseInvoice.value = exp.invoiceNumber || '';
        expenseMemo.value = exp.memo || '';
        if (exp.receiptImage) {
            capturedImageBase64 = exp.receiptImage;
            receiptPreview.src = 'data:image/jpeg;base64,' + exp.receiptImage;
            receiptPreview.style.display = 'block';
            cameraPlaceholder.style.display = 'none';
            cameraActions.style.display = 'flex';
        } else {
            receiptPreview.style.display = 'none';
            receiptPreview.src = '';
            cameraPlaceholder.style.display = 'block';
            cameraActions.style.display = 'none';
            capturedImageBase64 = null;
        }
        draftExpenses.splice(index, 1);
        updateDraftUI();
        showToast('内容をフォームに戻しました。修正して再度リストに追加してください', 'info');
        window.scrollTo(0, 0); // 上にスクロール
    };

    // ============================================
    // まとめて送信
    // ============================================
    btnSubmitAll.addEventListener('click', async () => {
        if (draftExpenses.length === 0) return;
        if (!submissionTarget.value) {
            showToast('提出先の課長を選択してください。', 'error');
            return;
        }
        if (!confirm(`${draftExpenses.length}件の経費をまとめて申請しますか？`)) return;

        draftExpenses.forEach(exp => exp.supervisor = submissionTarget.value);

        const originalText = btnSubmitAll.textContent;
        btnSubmitAll.textContent = '送信中...';
        btnSubmitAll.disabled = true;

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const saved = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            draftExpenses.forEach(exp => saved.push(exp));
            localStorage.setItem('demoExpenses', JSON.stringify(saved));
            showToast(`${draftExpenses.length}件の経費を申請しました`, 'success');
            draftExpenses = [];
            updateDraftUI();
            btnSubmitAll.textContent = originalText;
            btnSubmitAll.disabled = false;
            return;
        }

        try {
            const response = await fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    mode: 'submit_expenses_batch',
                    data: draftExpenses
                })
            });

            const result = await response.json();
            if (result.error) throw new Error(result.error);

            showToast(`${draftExpenses.length}件の経費を一括申請しました！上席に通知されました。`, 'success');
            draftExpenses = [];
            updateDraftUI();
            
            // ★送信完了後に「申請済み経費一覧」を自動でリロードして最新状態にする
            await loadExpenseList();
        } catch (error) {
            console.error('Batch submit error:', error);
            showToast('申請に失敗しました: ' + error.message, 'error');
        } finally {
            btnSubmitAll.textContent = originalText;
            btnSubmitAll.disabled = false;
        }
    });

    function resetForm() {
        expenseForm.reset();
        expenseDate.value = today;
        receiptPreview.style.display = 'none';
        receiptPreview.src = '';
        cameraPlaceholder.style.display = 'block';
        cameraActions.style.display = 'none';
        capturedImageBase64 = null;
        receiptInput.value = '';
    }

    // ============================================
    // 経費一覧表示（キャッシュ＆重複防止つき）
    // ============================================
    let currentMonthFilter = 'ALL';
    let isLoadingExpenseList = false; // 重複呼び出し防止ガード
    const EXPENSE_CACHE_KEY = 'cachedExpenseList';

    monthDropdown.addEventListener('change', (e) => {
        currentMonthFilter = e.target.value;
        loadExpenseList();
    });

    function renderExpenseRows(expenses) {
        expenseTbody.innerHTML = '';

        // フィルタ適用
        expenses = expenses.filter(exp => exp.status !== '削除');

        if (currentMonthFilter !== 'ALL') {
            expenses = expenses.filter(exp => {
                const d = new Date(exp.date);
                return (d.getMonth() + 1) === parseInt(currentMonthFilter);
            });
        }

        if (expenses.length > 0) {
            expenses.sort((a, b) => new Date(b.date) - new Date(a.date));
            let total = 0;
            expenses.forEach(exp => {
                if (exp.paymentMethod !== 'クレジットカード（会社）') {
                    total += parseInt(exp.amount) || 0;
                }
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="text-align:center;">
                        ${(exp.status === '申請中' || exp.status === '棄却') ? `<input type="checkbox" class="withdraw-chk" value="${exp.id}">` : ''}
                    </td>
                    <td>${formatDateJP(exp.date)}</td>
                    <td>${exp.category || '-'}</td>
                    <td style="text-align:right;font-weight:500;">${formatCurrency(exp.amount)}</td>
                    <td>${exp.store || '-'}</td>
                    <td>${getStatusBadge(exp.status)}</td>
                    <td>
                        <button class="action-btn" onclick="showExpenseDetail('${exp.id || ''}')">詳細</button>
                    </td>
                `;
                expenseTbody.appendChild(tr);
            });
            if (expenseTotalAmountSpan) expenseTotalAmountSpan.textContent = formatCurrency(total);
        } else {
            expenseTbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:40px; color:var(--google-text-sub);">経費データがありません</td></tr>';
            if (expenseTotalAmountSpan) expenseTotalAmountSpan.textContent = '¥ 0';
        }

        listLoading.style.display = 'none';
        expenseTableWrapper.style.display = 'block';
        updateWithdrawActions();
    }

    async function loadExpenseList() {
        // 重複呼び出し防止
        if (isLoadingExpenseList) return;
        isLoadingExpenseList = true;

        listLoading.style.display = 'block';
        expenseTableWrapper.style.display = 'none';
        expenseTbody.innerHTML = '';

        let expenses = [];

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            expenses = JSON.parse(localStorage.getItem('demoExpenses') || '[]')
                .filter(exp => exp.employeeName === currentUser.name);
            window.currentExpenses = expenses;
            renderExpenseRows(expenses);
            isLoadingExpenseList = false;
            return;
        }

        // キャッシュがあれば即座に表示
        try {
            const cached = localStorage.getItem(EXPENSE_CACHE_KEY);
            if (cached) {
                const cachedData = JSON.parse(cached);
                if (cachedData && cachedData.length > 0) {
                    window.currentExpenses = cachedData;
                    renderExpenseRows(cachedData);
                }
            }
        } catch (e) { /* ignore */ }

        // GASから最新データを取得
        try {
            const res = await fetch(`${GAS_WEBAPP_URL}?mode=get_expenses&employee=${encodeURIComponent(currentUser.name)}&month=${currentMonthFilter}&t=${Date.now()}`);
            const data = await res.json();
            expenses = data.expenses || [];
            window.currentExpenses = expenses;

            // キャッシュ更新
            try {
                localStorage.setItem(EXPENSE_CACHE_KEY, JSON.stringify(expenses));
            } catch (e) { /* ignore */ }

            // 最新データで再描画
            renderExpenseRows(expenses);
        } catch (e) {
            console.error('Load expense list error:', e);
            showToast('経費一覧の取得に失敗しました', 'error');
            // キャッシュで既に表示済みならエラーでも表示を維持
            if (!expenseTbody.hasChildNodes()) {
                listLoading.style.display = 'none';
                expenseTableWrapper.style.display = 'block';
            }
        } finally {
            isLoadingExpenseList = false;
        }
    }

    // ============================================
    // 一括取り下げ・再編集フロー
    // ============================================
    window.toggleAllWithdraw = function() {
        const checkAll = document.getElementById('check-all-withdraw');
        const checkboxes = document.querySelectorAll('.withdraw-chk');
        checkboxes.forEach(cb => cb.checked = checkAll.checked);
        updateWithdrawActions();
    };

    expenseTbody.addEventListener('change', (e) => {
        if (e.target.classList.contains('withdraw-chk')) {
            updateWithdrawActions();
        }
    });

    const withdrawActions = document.getElementById('withdraw-actions');
    const btnBulkWithdraw = document.getElementById('btn-bulk-withdraw');

    window.updateWithdrawActions = function() {
        const checked = document.querySelectorAll('.withdraw-chk:checked');
        if (checked.length > 0) {
            withdrawActions.style.display = 'block';
            btnBulkWithdraw.textContent = `選択した ${checked.length} 件の経費を一括取り下げ（修正用リストへ戻す）`;
        } else {
            withdrawActions.style.display = 'none';
        }
    };

    if(btnBulkWithdraw) {
        btnBulkWithdraw.addEventListener('click', async () => {
            const checkboxes = document.querySelectorAll('.withdraw-chk:checked');
            if (checkboxes.length === 0) return;
            
            if (!confirm(`選択した ${checkboxes.length} 件の経費を取り下げ、未送信リスト（修正用）に戻しますか？`)) return;

            const idsToWithdraw = Array.from(checkboxes).map(cb => cb.value);
            btnBulkWithdraw.textContent = '取り下げ処理中...';
            btnBulkWithdraw.disabled = true;

            if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
                const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
                idsToWithdraw.forEach(id => {
                    const idx = all.findIndex(e => e.id === id);
                    if (idx >= 0) {
                        const clone = JSON.parse(JSON.stringify(all[idx]));
                        clone.id = 'EXP-TEMP-' + Date.now() + Math.random();
                        clone.status = '申請中';
                        clone.employeeName = currentUser.name;
                        draftExpenses.push(clone);
                        all[idx].status = '取り下げ';
                    }
                });
                localStorage.setItem('demoExpenses', JSON.stringify(all));
                updateDraftUI();
                await loadExpenseList();
                document.querySelector('.tab-btn[data-tab="tab-new"]').click();
                btnBulkWithdraw.disabled = false;
                showToast(`${checkboxes.length}件を修正リストに戻しました`, 'success');
                return;
            }

            try {
                const response = await fetch(GAS_WEBAPP_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({
                        mode: 'withdraw_expenses_batch',
                        expenseIds: idsToWithdraw,
                        employeeName: currentUser.name
                    })
                });

                const result = await response.json();
                if (result.error) throw new Error(result.error);

                idsToWithdraw.forEach(id => {
                    const targetExp = (window.currentExpenses || []).find(e => e.id === id);
                    if (targetExp) {
                        const expClone = JSON.parse(JSON.stringify(targetExp));
                        expClone.id = 'EXP-TEMP-' + Date.now() + Math.floor(Math.random() * 1000);
                        expClone.status = '申請中';
                        expClone.employeeName = currentUser.name;
                        expClone.receiptImage = targetExp.receiptUrl ? '' : ''; 
                        draftExpenses.push(expClone);
                    }
                });

                updateDraftUI();
                await loadExpenseList(); // 再取得してリストを更新
                const checkAll = document.getElementById('check-all-withdraw');
                if (checkAll) checkAll.checked = false;
                
                document.querySelector('.tab-btn[data-tab="tab-new"]').click();
                window.scrollTo(0, 0); 
                
                showToast(`${checkboxes.length}件の経費を取り下げ、修正画面へ戻しました`, 'success');

            } catch (error) {
                console.error('Withdraw error:', error);
                showToast('取り下げに失敗しました: ' + error.message, 'error');
            } finally {
                btnBulkWithdraw.textContent = '選択した経費を一括取り下げ（修正用リストへ戻す）';
                btnBulkWithdraw.disabled = false;
            }
        });
    }

    // ============================================
    // 承認画面
    // ============================================
    async function loadPendingApprovals() {
        approveLoading.style.display = 'block';
        if (approveContainer) approveContainer.style.display = 'none';
        if (approveContainer) approveContainer.innerHTML = '';

        let pending = [];

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            // デモモード：全申請中データ
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            if (currentUser.role === 'admin') {
                pending = all.filter(exp => exp.status === '承認');
            } else if (currentUser.role === 'manager') {
                pending = all.filter(exp => exp.status === '申請中' && exp.department === currentUser.department);
            }
        } else {
            try {
                const mode = currentUser.role === 'admin' ? 'get_final_pending' : 'get_pending';
                const res = await fetch(`${GAS_WEBAPP_URL}?mode=${mode}&approver=${encodeURIComponent(currentUser.name)}&department=${encodeURIComponent(currentUser.department)}`);
                const data = await res.json();
                pending = data.expenses || [];
            } catch (e) {
                console.error('Load pending approvals error:', e);
                showToast('承認待ちデータの取得に失敗しました', 'error');
            }
        }
        window.currentPending = pending;

        if (pending.length > 0) {
            // 申請者ごとにグループ化
            const grouped = {};
            pending.forEach(exp => {
                const emp = exp.employeeName || '不明なユーザー';
                if (!grouped[emp]) grouped[emp] = [];
                grouped[emp].push(exp);
            });

            for (const [empName, exps] of Object.entries(grouped)) {
                const section = document.createElement('div');
                section.className = 'approval-group card glass';
                section.style.marginBottom = '30px';
                
                let html = `<h3 style="margin-top:0; border-bottom:1px solid rgba(0,0,0,0.1); padding-bottom:10px;">
                                ${empName} さんの申請 (${exps.length}件)
                            </h3>
                            <div class="history-table-wrapper" style="box-shadow:none; padding:0; background:transparent;">
                                <table class="history-table">
                                    <thead>
                                        <tr>
                                            <th style="text-align:center;"><input type="checkbox" onchange="toggleAllApprove(this, '${empName}')"><br><span style="font-size:0.65rem;font-weight:normal;">全選択</span></th>
                                            <th>利用日</th>
                                            <th>経費科目</th>
                                            <th>金額</th>
                                            <th>店名/会社名</th>
                                            <th>インボイス番号</th>
                                            <th>操作</th>
                                        </tr>
                                    </thead>
                                    <tbody>`;
                
                exps.forEach(exp => {
                    html += `
                        <tr>
                            <td style="text-align:center;"><input type="checkbox" class="approve-chk-${empName}" value="${exp.id}"></td>
                            <td>${formatDateJP(exp.date)}</td>
                            <td>${exp.category || '-'}</td>
                            <td style="text-align:right;font-weight:500;">${formatCurrency(exp.amount)}</td>
                            <td>${exp.store || '-'}</td>
                            <td style="font-size:0.8rem;color:#888;">${exp.invoiceNumber || '-'}</td>
                            <td>
                                ${exp.receiptImage ? `<button class="action-btn" onclick="showReceiptImage('${exp.id}')">レシート</button>` : '-'}
                                <button class="action-btn" onclick="showExpenseDetail('${exp.id}')">詳細</button>
                            </td>
                        </tr>
                    `;
                });

                html += `</tbody></table></div>`;
                
                if (currentUser.role === 'admin') {
                    html += `
                    <div style="margin-top: 15px; display:flex; gap:10px; justify-content:flex-end;">
                        <button class="btn submit-btn" onclick="batchFinalConfirm('${empName}')" style="margin:0;">選択した経費を最終確認</button>
                    </div>`;
                } else {
                    html += `
                    <div style="margin-top: 15px; display:flex; gap:10px; justify-content:flex-end;">
                        <button class="btn reject" onclick="batchReject('${empName}')" style="margin:0;">選択を棄却</button>
                        <button class="btn submit-btn" onclick="batchApprove('${empName}')" style="margin:0;">選択を承認（未選択は棄却）</button>
                    </div>`;
                }

                section.innerHTML = html;
                approveContainer.appendChild(section);
            }
        } else {
            approveContainer.innerHTML = '<div style="text-align:center; padding:40px; color:#888;">承認待ちの経費はありません</div>';
        }

        approveLoading.style.display = 'none';
        if (approveContainer) approveContainer.style.display = 'block';
    }

    // 全選択トグル
    window.toggleAllApprove = function(chk, empName) {
        const checkboxes = document.querySelectorAll('.approve-chk-' + empName);
        checkboxes.forEach(cb => cb.checked = chk.checked);
    };

    // 一括承認・未選択は棄却
    window.batchApprove = async function(empName) {
        const checkboxes = Array.from(document.querySelectorAll('.approve-chk-' + empName));
        const approvals = checkboxes.filter(cb => cb.checked).map(cb => cb.value);
        const rejections = checkboxes.filter(cb => !cb.checked).map(cb => cb.value);
        
        const total = approvals.length + rejections.length;
        if (!confirm(`${empName} さんの申請を処理しますか？\n\n承認: ${approvals.length}件\n棄却: ${rejections.length}件`)) return;

        await submitApproveReject(empName, approvals, rejections, '一括処理による自動棄却');
    };

    // 一括棄却
    window.batchReject = async function(empName) {
        const checkboxes = Array.from(document.querySelectorAll('.approve-chk-' + empName));
        const selected = checkboxes.filter(cb => cb.checked).map(cb => cb.value);
        
        if (selected.length === 0) {
            showToast('棄却する経費を選択してください', 'error');
            return;
        }

        const reason = prompt('棄却理由を入力してください:', '内容に不備があるため');
        if (reason === null) return; // キャンセル

        if (!confirm(`選択した ${selected.length}件 を棄却しますか？`)) return;

        await submitApproveReject(empName, [], selected, reason);
    };

    // 一括最終確認（総務長用）
    window.batchFinalConfirm = async function(empName) {
        const checkboxes = Array.from(document.querySelectorAll('.approve-chk-' + empName));
        const selected = checkboxes.filter(cb => cb.checked).map(cb => cb.value);
        
        if (selected.length === 0) {
            showToast('確認済みにする経費を選択してください', 'error');
            return;
        }

        if (!confirm(`選択した ${selected.length}件 を最終確認済みにしますか？`)) return;

        // 総務長の確認フローを回す（今回は個別APIをループ呼び出し、または専用にする。ここでは簡易的に一括処理用APIを利用）
        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            showToast('デモモード：最終確認完了', 'success');
            loadPendingApprovals();
            return;
        }

        try {
            // 今回は便宜上、forループでAPIを呼んでしまうか一括で行うか
            for(const id of selected) {
                await fetch(GAS_WEBAPP_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ mode: 'final_confirm', expenseId: id, confirmer: currentUser.name })
                });
            }
            showToast(`${selected.length}件の最終確認が完了しました`, 'success');
            loadPendingApprovals();
        } catch(e) {
            showToast('通信エラー', 'error');
        }
    };

    async function submitApproveReject(empName, approvals, rejections, rejectReason) {
        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            
            approvals.forEach(id => {
                const idx = all.findIndex(e => e.id === id);
                if (idx >= 0) {
                    all[idx].status = '承認';
                    all[idx].approvedBy = currentUser.name;
                    all[idx].approvedAt = new Date().toISOString();
                }
            });

            rejections.forEach(id => {
                const idx = all.findIndex(e => e.id === id);
                if (idx >= 0) {
                    all[idx].status = '棄却';
                    all[idx].rejectedBy = currentUser.name;
                    all[idx].rejectReason = rejectReason;
                }
            });

            localStorage.setItem('demoExpenses', JSON.stringify(all));
            showToast('処理が完了しました', 'success');
            loadPendingApprovals();
            return;
        }

        try {
            const res = await fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    mode: 'batch_approve_reject',
                    approvals: approvals,
                    rejections: rejections,
                    employeeName: empName,
                    approver: currentUser.name,
                    rejectReason: rejectReason
                })
            });
            const data = await res.json();
            if (data.error) throw new Error(data.error);

            showToast('処理が完了しました', 'success');
            loadPendingApprovals();
        } catch (e) {
            showToast('通信に失敗しました: ' + e.message, 'error');
        }
    }

    // ============================================
    // 承認・棄却・最終確認アクション
    // ============================================
    let pendingRejectId = null;

    window.approveExpense = async function(expenseId) {
        if (!confirm('この経費を承認しますか？')) return;

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            const idx = all.findIndex(e => e.id === expenseId);
            if (idx >= 0) {
                all[idx].status = '承認';
                all[idx].approvedBy = currentUser.name;
                all[idx].approvedAt = new Date().toISOString();
                localStorage.setItem('demoExpenses', JSON.stringify(all));
            }
            showToast('経費を承認しました', 'success');
            loadPendingApprovals();
            return;
        }

        try {
            const res = await fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    mode: 'approve_expense',
                    expenseId: expenseId,
                    approver: currentUser.name
                })
            });
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            showToast('経費を承認しました', 'success');
            loadPendingApprovals();
        } catch (e) {
            showToast('承認処理に失敗しました: ' + e.message, 'error');
        }
    };

    window.openRejectModal = function(expenseId) {
        pendingRejectId = expenseId;
        rejectReason.value = '';
        rejectModal.style.display = 'flex';
    };

    btnRejectCancel.addEventListener('click', () => {
        rejectModal.style.display = 'none';
        pendingRejectId = null;
    });

    btnRejectConfirm.addEventListener('click', async () => {
        const reason = rejectReason.value.trim();
        if (!reason) {
            showToast('棄却理由を入力してください', 'error');
            return;
        }

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            const idx = all.findIndex(e => e.id === pendingRejectId);
            if (idx >= 0) {
                all[idx].status = '棄却';
                all[idx].rejectedBy = currentUser.name;
                all[idx].rejectReason = reason;
                localStorage.setItem('demoExpenses', JSON.stringify(all));
            }
            rejectModal.style.display = 'none';
            showToast('経費を棄却しました', 'info');
            loadPendingApprovals();
            return;
        }

        try {
            const res = await fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    mode: 'reject_expense',
                    expenseId: pendingRejectId,
                    rejector: currentUser.name,
                    reason: reason
                })
            });
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            rejectModal.style.display = 'none';
            showToast('経費を棄却しました', 'info');
            loadPendingApprovals();
        } catch (e) {
            showToast('棄却処理に失敗しました: ' + e.message, 'error');
        }
    });

    window.finalConfirm = async function(expenseId) {
        if (!confirm('この経費を確認済みにしますか？')) return;

        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            const idx = all.findIndex(e => e.id === expenseId);
            if (idx >= 0) {
                all[idx].status = '確認済み';
                all[idx].confirmedBy = currentUser.name;
                all[idx].confirmedAt = new Date().toISOString();
                localStorage.setItem('demoExpenses', JSON.stringify(all));
            }
            showToast('最終確認が完了しました', 'success');
            loadPendingApprovals();
            return;
        }

        try {
            const res = await fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    mode: 'final_confirm',
                    expenseId: expenseId,
                    confirmer: currentUser.name
                })
            });
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            showToast('最終確認が完了しました', 'success');
            loadPendingApprovals();
        } catch (e) {
            showToast('確認処理に失敗しました: ' + e.message, 'error');
        }
    };

    // ============================================
    // 承認済み（管理用）画面
    // ============================================
    window.loadAdminCompletedExpenses = async function() {
        const completedLoading = document.getElementById('completed-loading');
        const completedTableWrapper = document.getElementById('completed-table-wrapper');
        const completedTbody = document.getElementById('completed-tbody');
        if (!completedLoading || !completedTableWrapper || !completedTbody) return;

        completedLoading.style.display = 'block';
        completedTableWrapper.style.display = 'none';
        completedTbody.innerHTML = '';

        try {
            const res = await fetch(`${GAS_WEBAPP_URL}?mode=get_admin_completed&t=${Date.now()}`);
            const data = await res.json();
            
            let expenses = data.expenses || [];
            if (expenses.length > 0) {
                expenses.sort((a, b) => new Date(b.date) - new Date(a.date));
                expenses.forEach(exp => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${exp.employeeName}</td>
                        <td>${formatDateJP(exp.date)}</td>
                        <td>${exp.category || '-'}</td>
                        <td style="text-align:right;">${formatCurrency(exp.amount)}</td>
                        <td>${exp.approver || '-'}</td>
                        <td>${exp.approvedAt ? formatDateJP(exp.approvedAt) : '-'}</td>
                    `;
                    completedTbody.appendChild(tr);
                });
                completedTableWrapper.style.display = 'block';
            } else {
                completedTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:40px; color:#888;">承認済みの経費データがありません</td></tr>';
                completedTableWrapper.style.display = 'block';
            }
        } catch (e) {
            console.error('completed list error:', e);
            showToast('承認済み一覧の取得に失敗しました', 'error');
        } finally {
            completedLoading.style.display = 'none';
        }
    };

    // ============================================
    // 詳細表示
    // ============================================
    window.showExpenseDetail = function(expenseId) {
        let expense = null;
        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            expense = all.find(e => e.id === expenseId);
        } else {
            expense = (window.currentExpenses || []).find(e => e.id === expenseId) || 
                      (window.currentPending || []).find(e => e.id === expenseId);
        }
        if (!expense) {
            showToast('詳細データが見つかりません', 'error');
            return;
        }

        detailBody.innerHTML = `
            <div class="detail-row"><span class="detail-label">利用日</span><span class="detail-value">${formatDateJP(expense.date)}</span></div>
            <div class="detail-row"><span class="detail-label">経費科目</span><span class="detail-value">${expense.category || '-'}</span></div>
            <div class="detail-row"><span class="detail-label">支払種別</span><span class="detail-value">${expense.paymentMethod || '-'}</span></div>
            <div class="detail-row"><span class="detail-label">金額</span><span class="detail-value">${formatCurrency(expense.amount)}</span></div>
            <div class="detail-row"><span class="detail-label">店名/会社名</span><span class="detail-value">${expense.store || '-'}</span></div>
            <div class="detail-row"><span class="detail-label">インボイス番号</span><span class="detail-value">${expense.invoiceNumber || '-'}</span></div>
            <div class="detail-row"><span class="detail-label">ステータス</span><span class="detail-value">${getStatusBadge(expense.status)}</span></div>
            <div class="detail-row"><span class="detail-label">備考</span><span class="detail-value">${expense.memo || '-'}</span></div>
            ${expense.rejectReason ? `<div class="detail-row"><span class="detail-label">棄却理由</span><span class="detail-value" style="color:var(--status-rejected);">${expense.rejectReason}</span></div>` : ''}
            ${expense.receiptImage ? `<img src="data:image/jpeg;base64,${expense.receiptImage}" class="detail-receipt-img" alt="レシート画像">` : ''}
        `;
        detailModal.style.display = 'flex';
    };

    window.showReceiptImage = function(expenseId) {
        let expense = null;
        if (GAS_WEBAPP_URL.includes('YOUR_GAS')) {
            const all = JSON.parse(localStorage.getItem('demoExpenses') || '[]');
            expense = all.find(e => e.id === expenseId);
        } else {
            expense = (window.currentExpenses || []).find(e => e.id === expenseId) || 
                      (window.currentPending || []).find(e => e.id === expenseId);
        }
        if (!expense || !expense.receiptImage) {
            showToast('レシート画像がありません', 'info');
            return;
        }
        detailBody.innerHTML = `<img src="data:image/jpeg;base64,${expense.receiptImage}" class="detail-receipt-img" alt="レシート画像" style="max-height:500px;">`;
        detailModal.style.display = 'flex';
    };

    btnDetailClose.addEventListener('click', () => {
        detailModal.style.display = 'none';
    });

    // モーダル外クリックで閉じる
    rejectModal.addEventListener('click', (e) => {
        if (e.target === rejectModal) rejectModal.style.display = 'none';
    });
    detailModal.addEventListener('click', (e) => {
        if (e.target === detailModal) detailModal.style.display = 'none';
    });

});
