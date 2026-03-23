document.addEventListener('DOMContentLoaded', function() {
    // DOM 元素 - 匹配 HTML 中的實際 ID
    const caseList = document.getElementById('caseList');
    const addCaseBtn = document.getElementById('addCaseBtn');
    const excelBox = document.getElementById('excelBox');
    const chatMessages = document.getElementById('chatMessages');
    const userInput = document.getElementById('userInput');
    const sendBtn = document.querySelector('.chat-input button');
    const fileInput = document.getElementById('fileInput');
    const importBtn = document.getElementById('importBtn');
    const excelContainer = document.getElementById('luckysheet-container');
    const caseTitle = document.getElementById('caseTitle');


    // 全域變數
    let activeCaseId = null;
    let cases = {};
    let caseCounter = 0;
    let isLuckysheetReady = false;
    let luckysheetInitialized = false;
    let activeTypeMenu = null; // 追蹤當前開啟的類型選單

    // 初始化
    init();

    // ── 側邊欄收納 ──────────────────────────────────────────
    const sidebarToggleBtn = document.getElementById('sidebarToggleBtn');
    const sidebar = document.getElementById('sidebar');

    if (sidebarToggleBtn && sidebar) {
        sidebarToggleBtn.addEventListener('click', function () {
            const isCollapsed = sidebar.classList.toggle('collapsed');
            sidebarToggleBtn.title = isCollapsed ? '展開側邊欄' : '收納側邊欄';
            // 收納後觸發 Luckysheet resize 以填滿新空間
            setTimeout(function () {
                if (typeof luckysheet !== 'undefined') {
                    try { luckysheet.resize(); } catch (e) { /* ignore */ }
                }
                window.dispatchEvent(new Event('resize'));
            }, 320);
        });
    }

    async function init() {
        initLuckysheetWhenReady();
        bindEvents();
        bindAuthEvents();
        await checkSession();
    }


    function initLuckysheetWhenReady() {
        console.log('開始檢查 Luckysheet 載入狀態...');
        
        if (typeof luckysheet !== 'undefined') {
            console.log('Luckysheet 已載入，開始初始化');
            initLuckysheet();
        } else {
            console.log('Luckysheet 尚未載入，等待中...');
            let retryCount = 0;
            const maxRetries = 20; // 增加重試次數
            
            const checkInterval = setInterval(() => {
                console.log(`檢查 Luckysheet 載入狀態 - 第 ${retryCount + 1} 次`);
                
                if (typeof luckysheet !== 'undefined') {
                    clearInterval(checkInterval);
                    console.log('Luckysheet 載入成功，開始初始化');
                    initLuckysheet();
                } else if (retryCount >= maxRetries) {
                    clearInterval(checkInterval);
                    console.error('Luckysheet 載入失敗，超過最大重試次數');
                    showLuckysheetError();
                }
                retryCount++;
            }, 250); // 縮短檢查間隔
        }
    }

    function showLuckysheetError() {
        console.log('Luckysheet 載入失敗，自動切換到完整表格模式');
        
        // 直接使用完整表格顯示模式
        createSimpleTable(null);
        
        // 如果當試算表有 Excel 數據，立即顯示
        if (activeCaseId && cases[activeCaseId] && cases[activeCaseId].excelData) {
            createSimpleTable(cases[activeCaseId].excelData);
        }
    }

    // ── Session / Auth ──────────────────────────────────────

    function setLoggedInUI(email) {
        document.getElementById('emailLoginArea').style.display = 'none';
        document.getElementById('emailUserArea').style.display = 'flex';
        document.getElementById('emailDisplay').textContent = email;
    }

    function setLoggedOutUI() {
        document.getElementById('emailLoginArea').style.display = 'flex';
        document.getElementById('emailUserArea').style.display = 'none';
        document.getElementById('emailInput').value = '';
        cases = {};
        activeCaseId = null;
        caseCounter = 0;
        renderCaseList();
        renderChat();
        if (caseTitle) caseTitle.textContent = '請選擇或新增試算表';
    }

    async function checkSession() {
        try {
            const res  = await fetch('/api/auth/me');
            const data = await res.json();
            if (data.status === 'ok') {
                setLoggedInUI(data.email);
                await loadCasesFromDB();
            } else {
                setLoggedOutUI();
            }
        } catch (e) {
            setLoggedOutUI();
        }
    }

    async function loginWithEmail() {
        const email = document.getElementById('emailInput').value.trim();
        if (!email) { alert('請輸入 Email'); return; }

        const btn = document.getElementById('loginBtn');
        btn.textContent = '登入中...';
        btn.disabled = true;

        try {
            const res  = await fetch('/api/auth/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email })
            });
            const data = await res.json();
            if (data.status === 'success') {
                setLoggedInUI(data.email);
                await loadCasesFromDB();
            } else {
                alert(data.error || '登入失敗');
            }
        } catch (e) {
            alert('登入失敗，請稍後再試');
        } finally {
            btn.textContent = '登入';
            btn.disabled = false;
        }
    }

    async function logoutUser() {
        await fetch('/api/auth/logout', { method: 'POST' });
        setLoggedOutUI();
    }

    function bindAuthEvents() {
        document.getElementById('loginBtn').addEventListener('click', loginWithEmail);
        document.getElementById('logoutBtn').addEventListener('click', logoutUser);
        document.getElementById('emailInput').addEventListener('keypress', e => {
            if (e.key === 'Enter') loginWithEmail();
        });
    }

    // ── 案場（DB 版）───────────────────────────────────────

    async function loadCasesFromDB() {
        const res  = await fetch('/api/cases');
        const data = await res.json();
        if (data.status !== 'success') return;

        cases = {};
        for (const c of data.cases) {
            const key = String(c.id);
            cases[key] = {
                id:                   c.id,
                name:                 c.name,
                siteType:             c.site_type,
                messages:             [],
                excelData:            null,
                excelChecked:         false,  // 尚未嘗試從伺服器載入，切換時會自動嘗試
                hasExcel:             !!c.excel_filename,
                excelFileName:        c.excel_filename
                                        ? `${c.name}_${c.excel_filename}` : null,
                excelOriginalFileName: c.excel_filename || null,
            };
        }

        // 讓計數器從當前已有的案場數量開始，避免跨用戶累加
        caseCounter = Object.keys(cases).length;

        renderCaseList();

        if (Object.keys(cases).length === 0) {
            await addNewCase();
        } else {
            switchActiveCase(Object.keys(cases)[0]);
        }
    }

    // saveCases() 保留為空操作（各操作已即時寫 DB）
    function saveCases() {}

    // 供外部 script 讀取當前本地案場 key（非 DB ID）
    window.getActiveCaseId = function() { return activeCaseId; };

    // 供外部 script 呼叫：將訊息存入指定案場並（若仍在該案場）顯示
    // caseId 必須是本地 key（用 window.getActiveCaseId() 取得），省略時退回到 activeCaseId
    window.appendMessageToCase = function(caseId, role, text, type) {
        const targetId = caseId || activeCaseId;
        if (!targetId || !cases[targetId]) return;
        const msg = { role: role || 'bot', text };
        cases[targetId].messages.push(msg);
        // 只有目前仍在該案場才即時顯示
        if (activeCaseId === targetId) {
            appendMessage(msg.role, msg.text);
        }
        // 存 DB
        const dbId = cases[targetId].id;
        if (dbId) {
            fetch(`/api/cases/${dbId}/messages`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ messages: [{ role: msg.role, content: text }] })
            }).catch(() => {});
        }
    };

    // 舊介面保留相容（不指定 caseId → 用 activeCaseId）
    window.saveBotMessageToDB = function(text) {
        window.appendMessageToCase(activeCaseId, 'bot', text);
    };

    // 初始化 Luckysheet - 優化版本
    function initLuckysheet(data = null) {
        console.log('開始初始化 Luckysheet...');

        // 避免重複初始化
        if (luckysheetInitialized) {
            console.log('Luckysheet 已初始化，更新數據');
            updateLuckysheetData(data);
            return;
        }

        // 確保 Luckysheet 已載入
        if (typeof luckysheet === 'undefined') {
            console.warn('Luckysheet 尚未載入，延遲初始化...');
            setTimeout(() => initLuckysheet(data), 500);
            return;
        }

        // 清理容器
        const container = document.getElementById('luckysheet-container');
        if (!container) {
            console.error('找不到 Luckysheet 容器');
            return;
        }

        // 先顯示白色背景避免黑屏
        container.innerHTML = '<div style="width: 100%; height: 100%; background: white; display: flex; align-items: center; justify-content: center;"><div style="color: #666;">載入中...</div></div>';
        console.log('容器已清理');

        // 設定預設數據
        const defaultData = data || [{
            name: "工作表1",
            color: "",
            index: "0",
            status: "1",
            order: "0",
            celldata: [],
            config: {
                merge: {},
                rowlen: {},
                columnlen: {},
                rowhidden: {},
                colhidden: {},
                borderInfo: [],
                authority: {}
            },
            row: 20,
            column: 10,
            pivotTable: null,
            isPivotTable: false,
            luckysheet_select_save: [],
            calcChain: [],
            hyperlink: {},
            dataVerification: {}
        }];

        const options = {
            container: 'luckysheet-container',
            data: defaultData,
            lang: 'zh',
            allowCopy: true,
            allowUpdate: true,
            showsheetbar: true,
            showstatisticBar: false,
            enableAddRow: true,
            enableAddCol: true,
            showtoolbar: true,
            showinfobar: false,
            showsheetbarConfig: {
                add: true,
                menu: true,
                sheet: true
            },
            hook: {
                workbookCreateBefore: function() {
                    console.log('Luckysheet 開始創建...');
                },
                workbookCreateAfter: function() {
                    isLuckysheetReady = true;
                    luckysheetInitialized = true;
                    console.log('Luckysheet 初始化完成');
                    
                    // 隱藏可能的載入提示
                    const loadingElement = document.querySelector('.luckysheet-loading');
                    if (loadingElement) {
                        loadingElement.style.display = 'none';
                    }
                },
                cellRenderAfter: function() {
                    console.log('單元格渲染完成');
                }
            }
        };

        try {
            console.log('開始創建 Luckysheet 實例...');
            luckysheet.create(options);

            // 設定超時檢查
            setTimeout(() => {
                if (!luckysheetInitialized) {
                    console.error('Luckysheet 初始化超時');
                    showLuckysheetError();
                }
            }, 10000); // 10秒超時

        } catch (error) {
            console.error('Luckysheet 初始化失敗:', error);
            luckysheetInitialized = false;
            showLuckysheetError();
        }
    }

    function showNoExcelPlaceholder() {
        // 顯示「請匯入 Excel」提示（僅在 Luckysheet 未初始化時使用）
        const container = document.getElementById('luckysheet-container');
        if (container) {
            // 重置旗標，因為 container 即將被自定義 HTML 替換
            isLuckysheetReady = false;
            luckysheetInitialized = false;
            container.innerHTML = `
                <div style="display: flex; align-items: center; justify-content: center; height: 100%; background: white; flex-direction: column;">
                    <div style="text-align: center; padding: 20px;">
                        <div style="font-size: 48px; margin-bottom: 20px;">📄</div>
                        <h3 style="color: #2c3e50; margin: 10px 0;">空白試算表</h3>
                        <p style="color: #666; margin: 10px 0;">請匯入 Excel 檔案開始使用</p>
                        <button onclick="document.getElementById('fileInput').click()"
                                style="background: #3498db; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px;">
                            📁 選擇 Excel 檔案
                        </button>
                    </div>
                </div>
            `;
        }
    }

    function updateLuckysheetData(data) {
        // 沒有數據：若 Luckysheet 已初始化就重置為空白工作表（不破壞實例）
        // 若未初始化則顯示提示 HTML
        if (!data) {
            if (isLuckysheetReady) {
                // 保留 Luckysheet 實例，只清空資料 → 後續 setAllSheetData 可直接用
                try {
                    luckysheet.setAllSheetData([{
                        name: '工作表1', index: '0', status: '1', order: '0',
                        celldata: [], row: 20, column: 10,
                        config: { merge: {}, rowlen: {}, columnlen: {}, borderInfo: [] },
                        pivotTable: null, isPivotTable: false,
                        luckysheet_select_save: [], calcChain: [],
                        hyperlink: {}, dataVerification: {}
                    }]);
                } catch (e) {
                    console.warn('清空 Luckysheet 失敗，改用提示頁面:', e);
                    showNoExcelPlaceholder();
                }
            } else {
                showNoExcelPlaceholder();
            }
            return;
        }

        if (!isLuckysheetReady) {
            console.log('Luckysheet 未準備好，重新初始化');
            luckysheetInitialized = false;
            initLuckysheet(data);
            return;
        }

        try {
            console.log('更新 Luckysheet 數據...');
            if (luckysheet && luckysheet.setAllSheetData) {
                luckysheet.setAllSheetData(data);
                console.log('數據更新成功');
            } else {
                console.warn('setAllSheetData 方法不可用，重新初始化');
                luckysheetInitialized = false;
                initLuckysheet(data);
            }
        } catch (error) {
            console.warn('更新 Luckysheet 數據失敗，重新初始化:', error);
            luckysheetInitialized = false;
            initLuckysheet(data);
        }
    }

    // 試算表管理
    async function addNewCase(siteType = 'single') {
        caseCounter++;
        const caseName = `表 ${caseCounter}`;

        const res  = await fetch('/api/cases', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: caseName, site_type: siteType })
        });
        const data = await res.json();
        if (data.status !== 'success') { console.error('新增案場失敗', data); return; }

        const key = String(data.case.id);
        cases[key] = {
            id:                    data.case.id,
            name:                  data.case.name,
            siteType:              data.case.site_type,
            messages:              [],
            excelData:             null,
            excelChecked:          true,   // 新案場確認沒有 Excel，不需要自動嘗試載入
            hasExcel:              false,
            excelFileName:         null,
            excelOriginalFileName: null,
        };

        renderCaseList();
        switchActiveCase(key);
    }

    async function deleteCase(caseId) {
        if (Object.keys(cases).length === 1) {
            alert('至少需要保留一試算表');
            return;
        }

        if (!confirm(`確定要刪除試算表「${cases[caseId].name}」嗎？`)) return;

        const dbId = cases[caseId].id;
        const originalFileName = cases[caseId].excelOriginalFileName;

        // 刪除後端 Excel 檔案
        if (originalFileName) {
            fetch('/api/delete_excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: `${cases[caseId].name}_${originalFileName}` })
            }).catch(() => {});
        }

        // 刪除 DB 記錄（含聊天紀錄 cascade）
        await fetch(`/api/cases/${dbId}`, { method: 'DELETE' });

        delete cases[caseId];

        let nextId = activeCaseId;
        if (activeCaseId === caseId) {
            activeCaseId = null;
            nextId = Object.keys(cases)[0];
        }
        renderCaseList();
        switchActiveCase(nextId);
    }

    async function renameCase(caseId) {
        const currentName = cases[caseId].name;
        const newName = prompt('請輸入新的試算表名稱:', currentName);
        if (!newName || !newName.trim() || newName.trim() === currentName) return;

        const trimmed = newName.trim();
        const dbId = cases[caseId].id;
        const originalFileName = cases[caseId].excelOriginalFileName;

        // 更新 DB 名稱
        await fetch(`/api/cases/${dbId}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: trimmed })
        });

        // 重新命名 Excel 檔案
        if (originalFileName) {
            const oldFullName = `${currentName}_${originalFileName}`;
            const newFullName = `${trimmed}_${originalFileName}`;
            fetch('/api/rename_excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ old_filename: oldFullName, new_filename: newFullName })
            }).then(r => r.json()).then(data => {
                if (data.status === 'success') {
                    cases[caseId].excelFileName = newFullName;
                }
            }).catch(() => {});
        }

        cases[caseId].name = trimmed;
        if (window.currentCase && window.currentCase.id === caseId) {
            window.currentCase.name = trimmed;
        }
        renderCaseList();
        if (caseTitle && activeCaseId === caseId) caseTitle.textContent = trimmed;
    }

    function renderCaseList() {
        if (!caseList) return;

        // 使用 DocumentFragment 提高性能
        const fragment = document.createDocumentFragment();

        Object.keys(cases).forEach(caseId => {
            const caseData = cases[caseId];
            const li = document.createElement('li');
            li.className = 'case-item';
            li.dataset.caseId = caseId;

            // 根據站點類型添加對應的 class（預設為 single）
            const siteType = caseData.siteType || 'single';
            li.classList.add(siteType === 'multi' ? 'multi-site' : 'single-site');

            if (caseId === activeCaseId) {
                li.classList.add('active');
            }

            // Excel 狀態指示器
            const excelIndicator = document.createElement('span');
            excelIndicator.className = 'excel-indicator';
            if (caseData.hasExcel) {
                excelIndicator.innerHTML = '📊';
                excelIndicator.title = `包含 Excel 資料: ${caseData.excelFileName || '未知檔案'}`;
            } else {
                excelIndicator.innerHTML = '📄';
                excelIndicator.title = '無 Excel 資料';
            }

            const span = document.createElement('span');
            span.textContent = caseData.name;

            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'delete-case-btn';
            deleteBtn.innerHTML = '×';
            deleteBtn.title = '刪除試算表';

            li.appendChild(excelIndicator);
            li.appendChild(span);
            li.appendChild(deleteBtn);

            // 使用事件委託減少事件監聯器
            li.addEventListener('click', handleCaseClick);
            li.addEventListener('contextmenu', handleCaseRightClick);

            fragment.appendChild(li);
        });

        // 一次性更新 DOM
        caseList.innerHTML = '';
        caseList.appendChild(fragment);
    }

    function handleCaseClick(e) {
        const caseId = e.currentTarget.dataset.caseId;
        if (e.target.classList.contains('delete-case-btn')) {
            e.stopPropagation();
            deleteCase(caseId);
        } else if (e.target.closest('.case-type-menu')) {
            // 點擊的是類型選單內部，不做處理
            return;
        } else {
            // 先切換到該案場
            switchActiveCase(caseId);
            // 顯示類型切換選單
            showCaseTypeMenu(e.currentTarget, caseId);
        }
    }

    // 顯示案場類型切換選單
    function showCaseTypeMenu(caseElement, caseId) {
        // 先關閉已開啟的選單
        closeCaseTypeMenu();

        const caseData = cases[caseId];
        const currentType = caseData.siteType || 'single';

        // 創建選單
        const menu = document.createElement('div');
        menu.className = 'case-type-menu show';
        menu.id = 'caseTypeMenu';

        menu.innerHTML = `
            <div class="case-type-menu-header">切換案場類型</div>
            <button class="case-type-menu-item ${currentType === 'single' ? 'active' : ''}" data-type="single">
                單站
            </button>
            <button class="case-type-menu-item ${currentType === 'multi' ? 'active' : ''}" data-type="multi">
                多站
            </button>
        `;

        // 添加點擊事件
        menu.querySelectorAll('.case-type-menu-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                const newType = item.dataset.type;
                changeCaseType(caseId, newType);
                closeCaseTypeMenu();
            });
        });

        caseElement.appendChild(menu);
        activeTypeMenu = menu;

        // 點擊其他地方關閉選單
        setTimeout(() => {
            document.addEventListener('click', handleOutsideClick);
        }, 0);
    }

    // 關閉類型選單
    function closeCaseTypeMenu() {
        if (activeTypeMenu) {
            activeTypeMenu.remove();
            activeTypeMenu = null;
        }
        document.removeEventListener('click', handleOutsideClick);
    }

    // 處理點擊外部關閉選單
    function handleOutsideClick(e) {
        if (activeTypeMenu && !activeTypeMenu.contains(e.target)) {
            closeCaseTypeMenu();
        }
    }

    // 根據當前案場的站點類型更新「匯入表格」按鈕狀態
    function updateImportSheetsBtnState() {
        const btn = document.getElementById('importSheetsBtn');
        if (!btn) return;
        const siteType = (activeCaseId && cases[activeCaseId])
            ? (cases[activeCaseId].siteType || 'single')
            : 'single';
        btn.disabled = (siteType !== 'multi');
    }

    // 更改案場類型
    function changeCaseType(caseId, newType) {
        const caseData = cases[caseId];
        const oldType = caseData.siteType || 'single';

        if (oldType === newType) return;

        caseData.siteType = newType;
        saveCases();
        renderCaseList();
        if (caseId === activeCaseId) updateImportSheetsBtnState();

        // 顯示提示訊息
        const typeText = newType === 'multi' ? '多站' : '單站';
        const message = {
            role: 'bot',
            text: `案場類型已切換為「${typeText}」`
        };
        caseData.messages.push(message);
        renderChat();
        saveCases();
    }

    function handleCaseRightClick(e) {
        e.preventDefault();
        const caseId = e.currentTarget.dataset.caseId;
        renameCase(caseId);
    }

    async function switchActiveCase(caseId) {
        if (activeCaseId === caseId) return;

        activeCaseId = caseId;
        updateImportSheetsBtnState();

        // 設置全局變量，供價金滾算等功能使用
        window.currentCase = {
            id: cases[caseId].id,           // DB 的數字 ID，用於後端定址
            name: cases[caseId].name,
            original_filename: cases[caseId].excelOriginalFileName || '',
            sheet_name: cases[caseId].sheetName || null
        };

        // 從 DB 載入該案場的完整聊天紀錄
        try {
            const dbId = cases[caseId].id;
            const res  = await fetch(`/api/cases/${dbId}/messages`);
            const data = await res.json();
            if (data.status === 'success') {
                cases[caseId].messages = data.messages.map(m => ({
                    role: m.role, text: m.content
                }));
            }
        } catch (e) {
            console.warn('載入聊天紀錄失敗:', e);
        }

        // 只要記憶體中沒有 excelData 且尚未確認過，就嘗試從伺服器靜默載入
        // 不依賴 hasExcel，因為舊資料的 excel_filename 可能為 null
        // excelChecked 旗標避免每次切換都重打 API（失敗後就不再重試）
        if (!cases[caseId].excelData && !cases[caseId].excelChecked) {
            reloadExcelData(caseId, true);
        }

        // 批量更新 UI
        requestAnimationFrame(() => {
            renderCaseList();
            renderChat();

            // 更新標題
            if (caseTitle) {
                caseTitle.textContent = cases[caseId].name;
            }
            
            // 載入該案場的 Excel 數據（統一使用靜態 HTML 表格）
            if (cases[caseId].excelData) {
                createSimpleTable(cases[caseId].excelData);
            } else if (!cases[caseId].excelChecked) {
                // excelData 為 null 但尚未確認（reloadExcelData 正在進行中）
                // 不觸發任何清空動作，保留 Luckysheet 現狀，等 fetch 完成後 setAllSheetData
                console.log('Excel 載入中，保持 Luckysheet 現狀等待資料...');
            } else {
                // 確認此案場無 Excel 檔案，顯示空白提示
                createSimpleTable(null);
            }
        });
    }

    // 聊天功能
    function renderChat() {
        if (!chatMessages) return;
        
        chatMessages.innerHTML = '';
        if (activeCaseId && cases[activeCaseId]) {
            cases[activeCaseId].messages.forEach(message => {
                appendMessage(message.role, message.text);
            });
        }
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    // 掛到 window 讓 price_rolling.js 等外部 script 也能使用
    window._renderMarkdown = function(text) { return renderMarkdown(text); };
    function isTableLine(line) {
        var s = line.trim();
        if (!s) return false;
        if (/^\s*\|/.test(line)) return true;
        if (s.indexOf('|') !== -1) return true;
        return false;
    }

    function normalizeMarkdown(text) {
        return text
            .replace(/([^\n])(#{1,3} )/g, '$1\n$2')
            .replace(/([^\n])(- \*\*)/g, '$1\n$2')
            .replace(/([^\n])((?<!\|)---(?![-|]))/g, '$1\n$2')
            .replace(/([^\n])(\*\*[^*]+\*\*[^|*\n]{3,})/g, '$1\n$2');
    }

    function renderMarkdown(text) {
        const lines = normalizeMarkdown(text).split('\n');
        const output = [];
        let i = 0;

        while (i < lines.length) {
            const line = lines[i];

            if (isTableLine(line)) {
                const tableLines = [];
                while (i < lines.length && isTableLine(lines[i])) {
                    tableLines.push(lines[i]);
                    i++;
                }
                output.push(buildTable(tableLines));
                continue;
            }

            if (line.startsWith('```')) {
                i++;
                const codeLines = [];
                while (i < lines.length && !lines[i].startsWith('```')) {
                    codeLines.push(escapeHtml(lines[i]));
                    i++;
                }
                i++;
                output.push(`<pre><code>${codeLines.join('\n')}</code></pre>`);
                continue;
            }

            if (/^#{1,3}\s/.test(line)) {
                const level = line.match(/^(#{1,3})/)[1].length;
                const content = inlineMarkdown(line.replace(/^#{1,3}\s/, ''));
                output.push(`<h${level} class="md-h">${content}</h${level}>`);
                i++; continue;
            }

            const trimmed = line.trim();
            if (trimmed === '') {
                output.push('<br>');
            } else {
                output.push(`<p class="md-p">${inlineMarkdown(trimmed)}</p>`);
            }
            i++;
        }

        return output.join('');
    }

    function buildTable(lines) {
        const rows = lines
            .filter(l => !/^[-|:\s]+$/.test(l.trim()))
            .map(l =>
                l.trim()
                 .replace(/^\|/, '').replace(/\|$/, '')
                 .split('|')
                 .map(cell => inlineMarkdown(cell.trim()))
            );

        if (rows.length === 0) return '';

        const [header, ...body] = rows;
        const thCells = header.map(h => `<th>${h}</th>`).join('');
        const tbRows  = body.map(row => {
            const cells = header.map((_, ci) =>
                `<td>${row[ci] !== undefined ? row[ci] : ''}</td>`
            ).join('');
            return `<tr>${cells}</tr>`;
        }).join('');

        return `<div class="md-table-wrap"><table class="md-table"><thead><tr>${thCells}</tr></thead><tbody>${tbRows}</tbody></table></div>`;
    }

    function inlineMarkdown(text) {
        return text
            .replace(/\*\*(.+?)\*\*/g, '\x00strong\x01$1\x00/strong\x01')
            .replace(/\*(.+?)\*/g,     '\x00em\x01$1\x00/em\x01')
            .replace(/`(.+?)`/g,       '\x00code class="md-code"\x01$1\x00/code\x01')
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/\x00(\/??[a-z][^\x01]*)\x01/g, '<$1>');
    }

    function escapeHtml(s) {
        return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }

    function appendMessage(role, text) {
        if (!chatMessages) return;

        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${role}`;

        // bot / assistant 訊息套用 Markdown 渲染；使用者訊息維持純文字
        if (role === 'bot' || role === 'assistant') {
            messageDiv.innerHTML = renderMarkdown(text);
        } else {
            messageDiv.textContent = text;
        }

        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    function sendMessage() {
        if (!userInput || !chatMessages) return;

        const text = userInput.value.trim();
        if (!text || !activeCaseId) return;


        // 顯示使用者訊息（含停止按鈕）
        const userMessage = { role: 'user', text: text };
        cases[activeCaseId].messages.push(userMessage);

        const abortController = new AbortController();
        const userMsgDiv = document.createElement('div');
        userMsgDiv.className = 'message user';

        const stopBtn = document.createElement('button');
        stopBtn.className = 'stop-btn';
        stopBtn.title = '停止回答';
        stopBtn.textContent = '⏹';
        stopBtn.onclick = () => abortController.abort();

        const textSpan = document.createElement('span');
        textSpan.textContent = text;

        userMsgDiv.appendChild(stopBtn);
        userMsgDiv.appendChild(textSpan);
        chatMessages.appendChild(userMsgDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;

        userInput.value = '';
        userInput.placeholder = '輸入訊息...';
        saveCases();

        // 顯示 "思考中..." 提示
        const loadingId = 'loading-' + Date.now();
        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'message bot';
        loadingDiv.id = loadingId;
        loadingDiv.textContent = '正在分析數據';
        chatMessages.appendChild(loadingDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;

        // 使用預設參數（敏感性分析功能已移除）
        const costAdj = 1.0;  // 預設不調整
        const rentAdj = 1.0;  // 預設不調整
        const loanAmount = '';  // 預設空值

        // 在發送請求前先鎖定當前案場 ID，避免非同步回調時 activeCaseId 已切換
        const senderCaseId = activeCaseId;

        // 獲取當前案場的 Excel 檔案資訊
        const caseName = cases[senderCaseId].name || '';
        const originalFilename = cases[senderCaseId].excelOriginalFileName || '';

        // 呼叫後端 Agent API
        fetch('/api/agent_chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            signal: abortController.signal,
            body: JSON.stringify({
                query: text,
                equipment_cost_adj: costAdj,
                rent_adj: rentAdj,
                loan_amount: loanAmount,
                case_id: cases[senderCaseId].id,
                case_name: caseName,
                original_filename: originalFilename,
                sheet_name: cases[senderCaseId].sheetName || null
            })
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(err => {
                    throw new Error(err.error || '未知的後端錯誤');
                });
            }
            return response.json();
        })
        .then(data => {
            stopBtn.remove();
            const loadingElement = document.getElementById(loadingId);
            if (loadingElement) loadingElement.remove();

            const botMessage = { role: 'bot', text: data.response };
            // 回答永遠存到原本那個聊天室
            if (cases[senderCaseId]) {
                cases[senderCaseId].messages.push(botMessage);
            }
            // 只有使用者還停在原聊天室時才即時顯示；否則等切回去時 renderChat 會顯示
            if (activeCaseId === senderCaseId) {
                appendMessage(botMessage.role, botMessage.text);
            }

            // 存入 DB（用原本的案場 ID）
            const dbId = cases[senderCaseId] && cases[senderCaseId].id;
            if (dbId) {
                fetch(`/api/cases/${dbId}/messages`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ messages: [
                        { role: 'user', content: text },
                        { role: 'bot',  content: data.response }
                    ]})
                }).catch(() => {});
            }

            if (data.excel_modified && cases[senderCaseId]) {
                reloadExcelData(senderCaseId);
            }
        })
        .catch(error => {
            stopBtn.remove();
            const loadingElement = document.getElementById(loadingId);
            if (loadingElement) loadingElement.remove();

            // 使用者主動停止 → 安靜清理，不顯示錯誤
            if (error.name === 'AbortError') {
                return;
            }

            console.error('錯誤:', error);
            const errorMessage = { role: 'bot', text: '系統發生問題，請稍後再試。如持續發生請聯繫管理員。' };
            if (cases[senderCaseId]) {
                cases[senderCaseId].messages.push(errorMessage);
            }
            if (activeCaseId === senderCaseId) {
                appendMessage(errorMessage.role, errorMessage.text);
            }
        });
    }

    // 重新載入 Excel 數據（silent=true 時不在聊天室顯示更新訊息）
    // 改為下載原始 xlsx 二進位 → 前端 XLSX.js 解析，與上傳走同一條路，樣式完整保留
    function reloadExcelData(caseId, silent = false) {
        console.log(`正在重新載入案場 ${caseId} 的 Excel 數據...`);

        const caseName = cases[caseId].name;
        const originalFileName = cases[caseId].excelOriginalFileName;

        // 使用 download_excel 端點取得原始二進位檔案
        fetch(`/api/download_excel?case_id=${encodeURIComponent(cases[caseId].id || '')}&case_name=${encodeURIComponent(caseName)}&original_filename=${encodeURIComponent(originalFileName || '')}`)
            .then(response => {
                if (!response.ok) {
                    if (cases[caseId]) cases[caseId].excelChecked = true;
                    throw new Error('找不到 Excel 檔案');
                }
                return response.arrayBuffer();
            })
            .then(arrayBuffer => {
                // 與上傳路徑完全相同：XLSX.read + convertToLuckysheet → 樣式/合併格/欄寬均保留
                const uint8 = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(uint8, { type: 'array', cellStyles: true });
                const luckysheetData = convertToLuckysheet(workbook);

                cases[caseId].excelData = luckysheetData;
                cases[caseId].excelChecked = true;
                cases[caseId].hasExcel = true;

                if (caseId === activeCaseId) {
                    createSimpleTable(luckysheetData);
                    console.log('Excel 顯示已更新（含樣式）');

                    if (!silent) {
                        const updateMessage = { role: 'bot', text: 'Excel 表格已自動更新' };
                        cases[activeCaseId].messages.push(updateMessage);
                        appendMessage(updateMessage.role, updateMessage.text);
                    }
                }
            })
            .catch(error => {
                console.error('重新載入 Excel 失敗:', error);
                if (cases[caseId]) cases[caseId].excelChecked = true;
                if (!silent) {
                    const errorMsg = { role: 'bot', text: `Excel 更新失敗: ${error.message}` };
                    if (cases[activeCaseId]) {
                        cases[activeCaseId].messages.push(errorMsg);
                        appendMessage(errorMsg.role, errorMsg.text);
                    }
                }
            });
    }

    // Excel 處理功能
    function handleExcelFile(file) {
        if (!file) return;

        // 先上傳檔案到後端
        uploadFileToBackend(file);

        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array', cellStyles: true});
                
                // 轉換為 Luckysheet 格式
                const luckysheetData = convertToLuckysheet(workbook);

                // 保存到當前
                cases[activeCaseId].excelData = luckysheetData;
                cases[activeCaseId].hasExcel = true;
                cases[activeCaseId].excelChecked = true;
                cases[activeCaseId].excelFileName = file.name;
                cases[activeCaseId].excelOriginalFileName = file.name;  // 儲存原始檔名
                // 統一使用靜態 HTML 表格顯示
                createSimpleTable(luckysheetData);
                
                // 更新 UI（不顯示前端預覽訊息，等上傳完成後再顯示）
                renderCaseList();
                saveCases();
                
                console.log('Excel 檔案處理完成');
                
            } catch (error) {
                console.error('處理 Excel 檔案時發生錯誤:', error);
                alert('無法處理此 Excel 檔案，請確認檔案格式正確。');
                
                // 添加錯誤訊息到聊天
                const errorMessage = { 
                    role: 'bot', 
                    text: `Excel 檔案載入失敗：${error.message}`
                };
                cases[activeCaseId].messages.push(errorMessage);
                renderChat();
                saveCases();
            }
        };
        
        reader.readAsArrayBuffer(file);
    }
    // 上傳檔案到後端
    async function uploadFileToBackend(file) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('case_id', activeCaseId);
        formData.append('case_name', cases[activeCaseId].name);  // 傳送前端顯示名稱
        formData.append('original_filename', file.name);

        fetch('/api/upload_excel', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(err => {
                    throw new Error(err.error || '上傳失敗');
                });
            }
            return response.json();
        })
        .then(async data => {
            console.log('檔案上傳成功:', data);

            if(data.original_filename){
                cases[activeCaseId].excelOriginalFileName = data.original_filename;  // 儲存原始檔名

                // ── 優先用 Excel B2 的案場名稱，否則 fallback 到去副檔名的檔名 ──
                const rawName = data.original_filename;
                const newCaseName = (data.project_name && data.project_name.trim())
                    ? data.project_name.trim()
                    : rawName.replace(/\.xlsx$/i, '').replace(/\.xls$/i, '').trim();
                const oldCaseName = cases[activeCaseId].name;

                if (newCaseName && newCaseName !== oldCaseName) {
                    const dbId = cases[activeCaseId].id;

                    // 更新前端 state
                    cases[activeCaseId].name = newCaseName;
                    if (window.currentCase && window.currentCase.id === activeCaseId) {
                        window.currentCase.name = newCaseName;
                    }

                    // 更新 DB 案場名稱
                    if (dbId) {
                        fetch(`/api/cases/${dbId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ name: newCaseName })
                        }).catch(() => {});
                    }

                    // 後端上傳時已直接用 project_name 重命名磁碟檔案
                    // 前端只需更新 excelFileName 為後端回傳的 filename
                    if (data.filename) {
                        cases[activeCaseId].excelFileName = data.filename;
                    }

                    // 同步更新 UI（側邊欄名稱、標題）
                    renderCaseList();
                    const caseTitle = document.getElementById('case-title');
                    if (caseTitle && activeCaseId === String(activeCaseId)) {
                        caseTitle.textContent = newCaseName;
                    }
                }

                // 同步更新 DB 案場的 excel_filename
                const dbId = cases[activeCaseId].id;
                if (dbId) {
                    fetch(`/api/cases/${dbId}`, {
                        method: 'PUT',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ excel_filename: data.original_filename })
                    }).catch(() => {});
                    cases[activeCaseId].hasExcel = true;
                }

                // 同步更新 window.currentCase
                if (window.currentCase && window.currentCase.id === activeCaseId) {
                    window.currentCase.original_filename = data.original_filename;
                }
            }

            // 上傳成功訊息（包含功能說明）
            const successText = `Excel 檔案已上傳到伺服器：${data.original_filename}\n\n` +
                      `功能說明：\n` +
                      `- 價金滾算：請點擊上方的「價金滾算」按鈕來執行滾算計算\n` +
                      `- 匯入表格：點選上方「匯入表格」可以插入其他案場的資訊\n` +
                      `- 匯出表格：若想匯出結果，可以點擊上方的「匯出表格」`;
            const successMessage = { role: 'bot', text: successText };
            cases[activeCaseId].messages.push(successMessage);
            renderChat();

            // 將通知訊息存入 DB，切換案場後仍可看到
            const uploadDbId = cases[activeCaseId].id;
            if (uploadDbId) {
                fetch(`/api/cases/${uploadDbId}/messages`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ messages: [{ role: 'bot', content: successText }] })
                }).catch(() => {});
            }
        })
        .catch(error => {
            console.error('檔案上傳錯誤:', error);

            const errorMessage = {
                role: 'bot',
                text: `檔案上傳到伺服器失敗：${error.message}`
            };
            cases[activeCaseId].messages.push(errorMessage);
            renderChat();
            saveCases();
        });
    }

    function convertToLuckysheet(workbook) {
        // 將 xlsx 顏色物件轉為 #RRGGBB 字串
        function toHexColor(colorObj) {
            if (!colorObj) return null;
            const raw = colorObj.argb || colorObj.rgb || '';
            if (raw.length === 8) return '#' + raw.slice(2); // ARGB → RGB
            if (raw.length === 6) return '#' + raw;
            return null;
        }

        const sheets = [];

        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];

            if (!worksheet['!ref']) {
                sheets.push({
                    name: sheetName, index: index.toString(),
                    status: index === 0 ? "1" : "0", order: index.toString(),
                    celldata: [], row: 20, column: 10, config: {},
                    pivotTable: null, isPivotTable: false,
                    luckysheet_select_save: [], calcChain: [],
                    hyperlink: {}, dataVerification: {}
                });
                return;
            }

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const celldata = [];

            for (let row = range.s.r; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                    const cell = worksheet[cellAddress];

                    // 跳過完全空白且沒有樣式的儲存格
                    if (!cell) continue;
                    const hasValue = cell.v !== undefined && cell.v !== null && cell.v !== '';
                    const hasStyle = !!cell.s;
                    if (!hasValue && !hasStyle) continue;

                    const vObj = {
                        v: hasValue ? cell.v : '',
                        ct: { fa: 'General', t: cell.t === 'n' ? 'n' : 's' },
                        m: cell.w || (hasValue ? String(cell.v) : '')
                    };

                    // 數字格式
                    if (cell.t === 'n' && cell.z) vObj.ct.fa = cell.z;

                    // 樣式（需要 XLSX.read 帶 cellStyles: true）
                    if (cell.s) {
                        const s = cell.s;

                        // 背景色
                        const bg = toHexColor(s.fgColor);
                        if (bg) vObj.bg = bg;

                        // 字型
                        if (s.font) {
                            const fc = toHexColor(s.font.color);
                            if (fc) vObj.fc = fc;
                            if (s.font.bold)      vObj.bl = 1;
                            if (s.font.italic)    vObj.it = 1;
                            if (s.font.underline) vObj.un = 1;
                            if (s.font.strike)    vObj.cl = 1;
                            if (s.font.sz)        vObj.fs = s.font.sz;
                            if (s.font.name)      vObj.ff = s.font.name;
                        }

                        // 對齊
                        if (s.alignment) {
                            const htMap = { left: 1, center: 2, right: 3, justify: 4 };
                            const vtMap = { top: 1, center: 2, bottom: 3 };
                            if (s.alignment.horizontal) vObj.ht = htMap[s.alignment.horizontal] || 0;
                            if (s.alignment.vertical)   vObj.vt = vtMap[s.alignment.vertical]   || 0;
                            if (s.alignment.wrapText)   vObj.tb = 2;
                        }
                    }

                    celldata.push({ r: row, c: col, v: vObj });
                }
            }

            // 合併格：XLSX 格式 → Luckysheet 格式
            const mergeConfig = {};
            if (Array.isArray(worksheet['!merges'])) {
                worksheet['!merges'].forEach(m => {
                    mergeConfig[`${m.s.r}_${m.s.c}`] = {
                        r: m.s.r, c: m.s.c,
                        rs: m.e.r - m.s.r + 1,
                        cs: m.e.c - m.s.c + 1
                    };
                });
            }

            // 欄寬
            const columnlen = {};
            if (worksheet['!cols']) {
                worksheet['!cols'].forEach((col, i) => {
                    if (col) {
                        const w = col.wpx || (col.wch ? Math.round(col.wch * 7) : null);
                        if (w) columnlen[i] = w;
                    }
                });
            }

            // 列高
            const rowlen = {};
            if (worksheet['!rows']) {
                worksheet['!rows'].forEach((row, i) => {
                    if (row) {
                        const h = row.hpx || (row.hpt ? Math.round(row.hpt * 1.333) : null);
                        if (h) rowlen[i] = h;
                    }
                });
            }

            sheets.push({
                name: sheetName,
                index: index.toString(),
                status: index === 0 ? "1" : "0",
                order: index.toString(),
                celldata: celldata,
                row: Math.max(range.e.r + 1, 20),
                column: Math.max(range.e.c + 1, 10),
                config: {
                    merge: mergeConfig,
                    rowlen: rowlen,
                    columnlen: columnlen,
                    rowhidden: {}, colhidden: {},
                    borderInfo: [], authority: {}
                },
                pivotTable: null, isPivotTable: false,
                luckysheet_select_save: [], calcChain: [],
                hyperlink: {}, dataVerification: {}
            });
        });

        return sheets;
    }

    // 事件綁定函數 - 統一管理所有事件
    function bindEvents() {
        // 案場新增下拉選單（單站 / 多站）
        const addCaseDropdown = document.getElementById('addCaseDropdown');
        if (addCaseBtn && addCaseDropdown) {
            // 點擊 + 按鈕顯示/隱藏下拉選單
            addCaseBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                addCaseDropdown.classList.toggle('show');
            });

            // 點擊下拉選單項目
            addCaseDropdown.querySelectorAll('.add-case-dropdown-item').forEach(item => {
                item.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const siteType = item.dataset.siteType;
                    addNewCase(siteType);
                    addCaseDropdown.classList.remove('show');
                });
            });

            // 點擊其他地方關閉下拉選單
            document.addEventListener('click', (e) => {
                if (!addCaseBtn.contains(e.target) && !addCaseDropdown.contains(e.target)) {
                    addCaseDropdown.classList.remove('show');
                }
            });
        }
        
        // 聊天相關事件
        if (sendBtn) {
            sendBtn.addEventListener('click', sendMessage);
        }
        
        if (userInput) {
            userInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    sendMessage();
                }
            });
        }

        // Excel 匯入事件
        if (importBtn) {
            importBtn.addEventListener('click', () => {
                if (fileInput) {
                    fileInput.click();
                }
            });
        }

        if (fileInput) {
            fileInput.addEventListener('change', handleFileChange);
        }

        // 拖拽功能
        const excelContainer = document.getElementById('luckysheet-container');
        if (excelContainer) {
            excelContainer.addEventListener('dragover', handleDragOver);
            excelContainer.addEventListener('dragleave', handleDragLeave);
            excelContainer.addEventListener('drop', handleDrop);
        }
    }

    function handleFileChange(event) {
        const file = event.target.files[0];
        if (file && activeCaseId) {
            // 處理檔案（不顯示載入提示）
            handleExcelFile(file);
        }
        // 重置檔案輸入
        fileInput.value = '';
    }

    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.style.backgroundColor = '#f0f8ff';
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.style.backgroundColor = '';
    }

    function handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.style.backgroundColor = '';
        
        const files = e.dataTransfer.files;
        if (files.length > 0 && activeCaseId) {
            const file = files[0];
            if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
                handleExcelFile(file);
            } else {
                alert('請選擇 Excel 檔案 (.xlsx 或 .xls)');
            }
        }
    }

    // Luckysheet 初始化失敗時的備用表格顯示
    function createSimpleTable(data) {
        const container = document.getElementById('luckysheet-container');
        if (!container) return;

        if (!data || data.length === 0) {
            container.innerHTML = `
                <div style="display:flex;align-items:center;justify-content:center;height:100%;flex-direction:column;background:white;">
                    <div style="text-align:center;color:#666;">
                        <div style="font-size:48px;margin-bottom:20px;">📄</div>
                        <h3 style="color:#2c3e50;margin:10px 0;">空白試算表</h3>
                        <p style="color:#666;margin:10px 0;">請匯入 Excel 檔案開始使用</p>
                    </div>
                </div>`;
            return;
        }

        // ── 建立分頁式 UI ────────────────────────────────────────────
        container.innerHTML = `
            <div style="display:flex;flex-direction:column;height:100%;background:white;border:1px solid #ddd;border-radius:4px;overflow:hidden;">
                <div id="sheet-tab-bar" style="display:flex;align-items:flex-end;background:#f5f5f5;border-bottom:2px solid #2c7be5;overflow-x:auto;flex-shrink:0;padding:0 8px;gap:2px;min-height:36px;"></div>
                <div style="display:flex;justify-content:space-between;align-items:center;padding:4px 12px;background:#fff;border-bottom:1px solid #eee;flex-shrink:0;">
                    <span id="excel-info" style="font-size:11px;color:#888;"></span>
                    <span id="sheet-cell-count" style="font-size:11px;color:#888;"></span>
                </div>
                <div id="sheet-content-area" style="flex:1;overflow:auto;"></div>
            </div>`;

        let activeSheetIndex = 0;

        function renderTabs() {
            const bar = document.getElementById('sheet-tab-bar');
            if (!bar) return;
            bar.innerHTML = '';
            data.forEach((sheet, i) => {
                const tab = document.createElement('div');
                tab.style.cssText = [
                    'padding:6px 14px 5px',
                    'cursor:pointer',
                    'font-size:12px',
                    'border:1px solid ' + (i === activeSheetIndex ? '#2c7be5' : '#ccc'),
                    'border-bottom:' + (i === activeSheetIndex ? '2px solid #fff' : '1px solid #ccc'),
                    'border-radius:4px 4px 0 0',
                    'background:' + (i === activeSheetIndex ? '#fff' : '#eee'),
                    'color:' + (i === activeSheetIndex ? '#2c7be5' : '#555'),
                    'font-weight:' + (i === activeSheetIndex ? 'bold' : 'normal'),
                    'white-space:nowrap',
                    'user-select:none',
                    'margin-bottom:-2px',
                    'transition:background .15s'
                ].join(';');
                tab.textContent = sheet.name || ('工作表' + (i + 1));
                tab.title = sheet.name || '';
                tab.addEventListener('click', () => {
                    activeSheetIndex = i;
                    renderTabs();
                    renderSheetContent(i);
                });
                tab.addEventListener('mouseenter', () => {
                    if (i !== activeSheetIndex) tab.style.background = '#e0e0e0';
                });
                tab.addEventListener('mouseleave', () => {
                    if (i !== activeSheetIndex) tab.style.background = '#eee';
                });
                bar.appendChild(tab);
            });
        }

        function renderSheetContent(idx) {
            // 同步更新案場目前觀看的分頁名稱，讓 AI 知道使用者在看哪個分頁
            if (cases[activeCaseId] && data[idx]) {
                cases[activeCaseId].sheetName = data[idx].name || null;
            }
            displaySheetData(data[idx], idx);
        }

        renderTabs();
        renderSheetContent(0);

        const infoEl = document.getElementById('excel-info');
        if (infoEl) infoEl.textContent = `共 ${data.length} 個工作表`;
    }

    // 渲染單一工作表到 #sheet-content-area，完整還原 xlsx 樣式並套用千分位格式
    function displaySheetData(sheet, sheetIndex) {
        const tableContainer = document.getElementById('sheet-content-area');
        if (!tableContainer) return;

        function fmtNumber(v) {
            if (v === null || v === undefined || v === '') return '';
            const n = Number(v);
            if (isNaN(n)) return String(v);
            return Math.round(n).toLocaleString('zh-TW');
        }

        function isMoneyValue(v, fa) {
            if (typeof v !== 'number') return false;
            if (Number.isInteger(v) && v >= 1900 && v <= 2100) return false;
            if (fa && fa.includes('%')) return false;
            if (Math.abs(v) < 1 && !Number.isInteger(v)) return false;
            return true;
        }

        const cells = sheet.celldata || [];
        const merges = (sheet.config && sheet.config.merge) ? sheet.config.merge : {};
        const colWidths = (sheet.config && sheet.config.columnlen) ? sheet.config.columnlen : {};
        const rowHeights = (sheet.config && sheet.config.rowlen) ? sheet.config.rowlen : {};

        const countEl = document.getElementById('sheet-cell-count');
        if (countEl) countEl.textContent = `${cells.length} 個儲存格`;

        if (!cells.length) {
            tableContainer.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#aaa;font-size:13px;">此工作表沒有內容</div>';
            return;
        }

        const maxRow = Math.max(...cells.map(c => c.r)) + 1;
        const maxCol = Math.max(...cells.map(c => c.c)) + 1;

        const cellMap = new Map();
        cells.forEach(c => cellMap.set(`${c.r}-${c.c}`, c));

        const covered = new Set();
        Object.values(merges).forEach(m => {
            for (let dr = 0; dr < m.rs; dr++) {
                for (let dc = 0; dc < m.cs; dc++) {
                    if (dr === 0 && dc === 0) continue;
                    covered.add(`${m.r + dr}_${m.c + dc}`);
                }
            }
        });

        let tbl = '<div style="overflow:auto;height:100%;"><table style="border-collapse:collapse;font-size:12px;table-layout:fixed;">';

        // colgroup
        tbl += '<colgroup><col style="width:36px">';
        for (let c = 0; c < maxCol; c++) {
            tbl += `<col style="width:${colWidths[c] || 80}px">`;
        }
        tbl += '</colgroup>';

        // 欄標題
        tbl += '<thead><tr style="background:#ecf0f1;position:sticky;top:0;z-index:2;"><th style="border:1px solid #bdc3c7;padding:3px 5px;text-align:center;background:#ecf0f1;"></th>';
        for (let c = 0; c < maxCol; c++) {
            let colName = '';
            let tmp = c;
            do { colName = String.fromCharCode(65 + tmp % 26) + colName; tmp = Math.floor(tmp / 26) - 1; } while (tmp >= 0);
            tbl += `<th style="border:1px solid #bdc3c7;padding:3px 5px;text-align:center;background:#ecf0f1;">${colName}</th>`;
        }
        tbl += '</tr></thead><tbody>';

        for (let r = 0; r < maxRow; r++) {
            const rh = rowHeights[r] ? `height:${rowHeights[r]}px;` : '';
            tbl += `<tr style="${rh}"><td style="border:1px solid #bdc3c7;padding:3px 5px;background:#ecf0f1;font-weight:bold;text-align:center;min-width:36px;position:sticky;left:0;z-index:1;">${r + 1}</td>`;

            for (let c = 0; c < maxCol; c++) {
                const covKey = `${r}_${c}`;
                if (covered.has(covKey)) continue;

                const mergeInfo = merges[`${r}_${c}`];
                const rowspan = mergeInfo ? mergeInfo.rs : 1;
                const colspan = mergeInfo ? mergeInfo.cs : 1;

                const cell = cellMap.get(`${r}-${c}`);
                const v = cell ? cell.v : null;

                let tdStyle = 'border:1px solid #bdc3c7;padding:3px 6px;overflow:hidden;vertical-align:middle;';
                tdStyle += `max-width:${colWidths[c] || 80}px;`;

                if (v) {
                    if (v.bg) {
                        tdStyle += `background:${v.bg};`;
                        // 年份列補白字：值是 2000~2100 的整數（年份數字）或「年份」標籤
                        if (!v.fc) {
                            const _raw = v.v;
                            const _isYearNum   = typeof _raw === 'number' && Number.isInteger(_raw) && _raw >= 2000 && _raw <= 2100;
                            const _isYearLabel = _raw === '年份' || v.m === '年份';
                            if (_isYearNum || _isYearLabel) tdStyle += 'color:#FFFFFF;';
                        }
                    }
                    if (v.fc) tdStyle += `color:${v.fc};`;
                    if (v.bl) tdStyle += 'font-weight:bold;';
                    if (v.it) tdStyle += 'font-style:italic;';
                    if (v.un) tdStyle += 'text-decoration:underline;';
                    if (v.fs) tdStyle += `font-size:${v.fs}pt;`;
                    if (v.ff) tdStyle += `font-family:${v.ff};`;
                    const htMap = {1:'left',2:'center',3:'right',4:'justify'};
                    const vtMap = {1:'top',2:'middle',3:'bottom'};
                    if (v.ht) tdStyle += `text-align:${htMap[v.ht]||'left'};`;
                    else if (v.ct && v.ct.t === 'n') tdStyle += 'text-align:right;';
                    if (v.vt) tdStyle += `vertical-align:${vtMap[v.vt]||'middle'};`;
                    tdStyle += (v.tb === 2) ? 'white-space:normal;word-break:break-all;' : 'white-space:nowrap;';
                } else {
                    tdStyle += 'white-space:nowrap;';
                }

                let displayVal = '';
                if (v) {
                    const raw = v.v;
                    const fa = (v.ct && v.ct.fa) ? v.ct.fa : '';
                    if (isMoneyValue(raw, fa)) {
                        displayVal = fmtNumber(raw);
                    } else if (raw !== null && raw !== undefined && raw !== '') {
                        displayVal = (v.m !== undefined && v.m !== '') ? v.m : String(raw);
                    }
                }

                const safe = String(displayVal)
                    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
                    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');

                const rsAttr = rowspan > 1 ? ` rowspan="${rowspan}"` : '';
                const csAttr = colspan > 1 ? ` colspan="${colspan}"` : '';

                tbl += `<td${rsAttr}${csAttr} style="${tdStyle}">${safe || '&nbsp;'}</td>`;
            }
            tbl += '</tr>';
        }

        tbl += '</tbody></table></div>';
        tableContainer.innerHTML = tbl;
    }

    // 舊名保留，供其他地方呼叫相容
    function displayCompleteExcelData(sheets) {
        if (sheets && sheets.length > 0) displaySheetData(sheets[0], 0);
    }

    // ========== 匯入表格功能 ==========

    // 儲存選擇的 sheets
    let selectedSheets = [];

    // 開啟匯入對話框
    window.openImportDialog = function() {
        const overlay = document.getElementById('importDialogOverlay');
        overlay.classList.add('show');
        selectedSheets = [];
        updateSelectedCount();
        loadCaseSheetsForImport();
    };

    // 關閉匯入對話框
    window.closeImportDialog = function() {
        const overlay = document.getElementById('importDialogOverlay');
        overlay.classList.remove('show');
        selectedSheets = [];
    };

    // 載入案場和 sheets 資料
    async function loadCaseSheetsForImport() {
        const listContainer = document.getElementById('importCaseList');
        listContainer.innerHTML = '<div class="import-loading">載入中...</div>';

        try {
            const response = await fetch('/api/list_case_sheets');
            const data = await response.json();

            if (data.status === 'success' && data.cases) {
                renderCaseSheetList(data.cases);
            } else {
                listContainer.innerHTML = `
                    <div class="import-empty-message">
                        <div class="icon">📭</div>
                        <div>沒有可匯入的案場</div>
                    </div>
                `;
            }
        } catch (error) {
            console.error('載入案場列表失敗:', error);
            listContainer.innerHTML = `
                <div class="import-empty-message">
                    <div class="icon">❌</div>
                    <div>載入失敗: ${error.message}</div>
                </div>
            `;
        }
    }

    // 渲染案場和 sheet 列表
    function renderCaseSheetList(casesData) {
        const listContainer = document.getElementById('importCaseList');

        // 過濾掉當前案場
        const filteredCases = casesData.filter(c => {
            // 比對案場名稱，排除當前案場
            if (activeCaseId && cases[activeCaseId]) {
                return c.case_name !== cases[activeCaseId].name;
            }
            return true;
        });

        if (filteredCases.length === 0) {
            listContainer.innerHTML = `
                <div class="import-empty-message">
                    <div class="icon">📭</div>
                    <div>沒有其他案場可匯入</div>
                </div>
            `;
            return;
        }

        let html = '';
        filteredCases.forEach((caseData, index) => {
            const siteType = caseData.site_type || 'single';
            const siteTypeText = siteType === 'multi' ? '多站' : '單站';
            const sheetCount = caseData.sheets ? caseData.sheets.length : 0;

            html += `
                <div class="import-case-item" data-case-name="${caseData.case_name}" data-filename="${caseData.filename}">
                    <div class="import-case-header" onclick="toggleImportCase(this)">
                        <span class="import-case-toggle">▶</span>
                        <span class="import-case-name">${caseData.display_name || caseData.case_name}</span>
                        <span class="import-case-type ${siteType}">${siteTypeText}</span>
                        <span class="import-sheet-count">${sheetCount} 個 Sheet</span>
                    </div>
                    <div class="import-sheet-list">
            `;

            if (caseData.sheets && caseData.sheets.length > 0) {
                caseData.sheets.forEach((sheet, sheetIndex) => {
                    const sheetId = `sheet_${index}_${sheetIndex}`;
                    html += `
                        <div class="import-sheet-item" data-sheet-name="${sheet}">
                            <input type="checkbox" id="${sheetId}"
                                   onchange="toggleSheetSelection('${caseData.case_name}', '${sheet}', '${caseData.filename}', this.checked)">
                            <label for="${sheetId}">${sheet}</label>
                        </div>
                    `;
                });
            } else {
                html += '<div style="color: #7f8c8d; padding: 10px;">此案場沒有 Sheet</div>';
            }

            html += `
                    </div>
                </div>
            `;
        });

        listContainer.innerHTML = html;
    }

    // 展開/收合案場
    window.toggleImportCase = function(headerElement) {
        const caseItem = headerElement.closest('.import-case-item');
        caseItem.classList.toggle('expanded');
    };

    // 切換 sheet 選擇狀態
    window.toggleSheetSelection = function(caseName, sheetName, filename, isSelected) {
        const sheetKey = `${caseName}|${sheetName}|${filename}`;

        if (isSelected) {
            if (!selectedSheets.find(s => s.key === sheetKey)) {
                selectedSheets.push({
                    key: sheetKey,
                    case_name: caseName,
                    sheet_name: sheetName,
                    filename: filename
                });
            }
        } else {
            selectedSheets = selectedSheets.filter(s => s.key !== sheetKey);
        }

        // 更新 UI
        const sheetItem = document.querySelector(`.import-sheet-item[data-sheet-name="${sheetName}"]`);
        if (sheetItem) {
            sheetItem.classList.toggle('selected', isSelected);
        }

        updateSelectedCount();
    };

    // 更新已選擇數量
    function updateSelectedCount() {
        const countElement = document.getElementById('importSelectedCount');
        const submitBtn = document.querySelector('.import-btn-submit');

        countElement.textContent = `已選擇 ${selectedSheets.length} 個 Sheet`;
        submitBtn.disabled = selectedSheets.length === 0;
    }

    // 執行匯入
    window.executeImportSheets = async function() {
        if (selectedSheets.length === 0) {
            alert('請選擇至少一個 Sheet');
            return;
        }

        if (!activeCaseId || !cases[activeCaseId]) {
            alert('請先選擇目標案場');
            return;
        }

        const targetCase = cases[activeCaseId];
        const targetCaseName = targetCase.name;
        const targetFilename = targetCase.excelOriginalFileName || null;

        // 先保存選擇的 sheets 資訊（因為關閉對話框會清空）
        const importedSheets = [...selectedSheets];
        const importedCount = selectedSheets.length;

        // 顯示載入中
        const submitBtn = document.querySelector('.import-btn-submit');
        const originalText = submitBtn.textContent;
        submitBtn.textContent = '匯入中...';
        submitBtn.disabled = true;

        // 收集財務參數
        const loanRatioVal = document.getElementById('importLoanRatio').value;
        const repayPeriodsVal = document.getElementById('importRepayPeriods').value;
        const dividendRatioVal = document.getElementById('importDividendRatio').value;
        const capReductionVal = document.getElementById('importCapReductionPeriods').value;
        const bankRateVal = document.getElementById('importBankRate').value;

        const financeParams = {
            loan_ratio: loanRatioVal !== '' ? parseFloat(loanRatioVal) : null,
            repay_periods: repayPeriodsVal !== '' ? parseInt(repayPeriodsVal) : null,
            dividend_ratio: dividendRatioVal !== '' ? parseFloat(dividendRatioVal) : null,
            cap_reduction_periods: capReductionVal !== '' ? parseInt(capReductionVal) : null,
            bank_rate: bankRateVal !== '' ? parseFloat(bankRateVal) : null
        };

        try {
            const response = await fetch('/api/import_sheets', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    target_case_name: targetCaseName,
                    target_filename: targetFilename,
                    sheets_to_import: importedSheets.map(s => ({
                        case_name: s.case_name,
                        sheet_name: s.sheet_name,
                        filename: s.filename
                    })),
                    finance_params: financeParams
                })
            });

            const data = await response.json();

            if (data.status === 'success') {
                // 關閉對話框
                closeImportDialog();

                // 更新案場的 Excel 資訊
                if (data.new_filename) {
                    cases[activeCaseId].excelOriginalFileName = data.new_filename;
                    cases[activeCaseId].hasExcel = true;
                    cases[activeCaseId].excelFileName = data.new_filename;

                    // 同步更新 window.currentCase
                    if (window.currentCase && window.currentCase.id === activeCaseId) {
                        window.currentCase.original_filename = data.new_filename;
                    }
                }

                // 顯示成功訊息（使用保存的資訊）
                const sheetList = importedSheets.map(s => `- ${s.case_name}_${s.sheet_name}`).join('\n');
                const successMessage = {
                    role: 'bot',
                    text: `成功匯入 ${importedCount} 個 Sheet\n\n匯入的 Sheet：\n${sheetList}`
                };
                cases[activeCaseId].messages.push(successMessage);
                saveCases();
                renderChat();

                // 重新載入 Excel 數據
                reloadExcelData(activeCaseId);

                // 更新案場列表顯示
                renderCaseList();
            } else {
                throw new Error(data.error || '匯入失敗');
            }
        } catch (error) {
            console.error('匯入失敗:', error);
            alert(`匯入失敗: ${error.message}`);
        } finally {
            submitBtn.textContent = originalText;
            submitBtn.disabled = false;
        }
    };

    // 綁定匯入按鈕事件
    const importSheetsBtn = document.getElementById('importSheetsBtn');
    if (importSheetsBtn) {
        importSheetsBtn.addEventListener('click', openImportDialog);
    }

    // 初始化匯入表格按鈕狀態
    updateImportSheetsBtnState();
});