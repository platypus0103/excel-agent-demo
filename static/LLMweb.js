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

    function init() {
        loadCases();
        if (Object.keys(cases).length === 0) {
            addNewCase();
        } else {
            const firstCaseId = Object.keys(cases)[0];
            switchActiveCase(firstCaseId);
        }
        
        // 優化的 Luckysheet 初始化
        initLuckysheetWhenReady();
        bindEvents();
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

    // 本地儲存
    function saveCases() {
        localStorage.setItem('llmWebCases', JSON.stringify(cases));
    }

    function loadCases() {
        const saved = localStorage.getItem('llmWebCases');
        if (saved) {
            try {
                cases = JSON.parse(saved);
                const caseIds = Object.keys(cases);
                if (caseIds.length > 0) {
                    caseCounter = Math.max(...caseIds.map(id => parseInt(id.replace('case', '')))) + 1;
                }
            } catch (e) {
                console.error('載數據時發生錯誤:', e);
                cases = {};
                caseCounter = 0;
            }
        }
    }

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

    function updateLuckysheetData(data) {
        // 如果沒有數據，顯示友好提示而不是重新初始化
        if (!data) {
            console.log('無 Excel 數據，顯示空白提示');
            const container = document.getElementById('luckysheet-container');
            if (container) {
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
    function addNewCase(siteType = 'single') {
        const caseId = `case${caseCounter}`;  // 先取得 ID
        const caseName = `表 ${caseCounter + 1}`;  // 顯示名稱
        caseCounter++;  // 最後遞增

        cases[caseId] = {
            name: caseName,
            siteType: siteType,  // 站點類型：'single'（單站）或 'multi'（多站）
            messages: [],
            excelData: null,
            hasExcel: false,
            excelFileName: null,
            excelOriginalFileName: null  // Excel 原始檔名（例如：公版.xlsx）
        };

        saveCases();
        renderCaseList();
        switchActiveCase(caseId);
    }

    function deleteCase(caseId) {
        if (Object.keys(cases).length === 1) {
            alert('至少需要保留一試算表');
            return;
        }

        if (confirm(`確定要刪除試算表「${cases[caseId].name}」嗎？`)) {
            // 如果有 Excel 檔案，先刪除後端的 Excel
            const caseName = cases[caseId].name;
            const originalFileName = cases[caseId].excelOriginalFileName;

            if (originalFileName) {
                const fullFileName = `${caseName}_${originalFileName}`;

                // 呼叫後端 API 刪除 Excel 檔案
                fetch('/api/delete_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        case_id: caseId,
                        filename: fullFileName
                    })
                })
                .then(response => response.json())
                .then(data => {
                    if (data.status === 'success') {
                        console.log('後端 Excel 檔案已刪除:', fullFileName);
                    } else {
                        console.warn('後端 Excel 刪除失敗:', data.error);
                    }
                })
                .catch(error => {
                    console.error('刪除 Excel 請求失敗:', error);
                });
            }

            // 刪除前端案場資料
            delete cases[caseId];

            // 判斷是否需要切換到其他案場
            let newActiveCaseId = activeCaseId;
            if (activeCaseId === caseId) {
                // 刪除的是當前案場，需要切換到其他案場
                const remainingCases = Object.keys(cases);
                newActiveCaseId = remainingCases[0];
                // 重要：先將 activeCaseId 設為 null，避免 switchActiveCase 提前 return
                activeCaseId = null;
            }

            saveCases();
            renderCaseList();
            switchActiveCase(newActiveCaseId);
        }
    }

    function renameCase(caseId) {
        const currentName = cases[caseId].name;
        const newName = prompt('請輸入新的試算表名稱:', currentName);

        if (newName && newName.trim() && newName !== currentName) {
            const originalFileName = cases[caseId].excelOriginalFileName;  // 原始檔名（例如：公版.xlsx）

            if(originalFileName){
                // 組成舊檔名和新檔名
                const oldFullName = `${currentName}_${originalFileName}`;  // 例如：表 1_公版.xlsx
                const newFullName = `${newName.trim()}_${originalFileName}`;  // 例如：財報_公版.xlsx

                fetch('/api/rename_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        case_id: caseId,
                        old_filename: oldFullName,
                        new_filename: newFullName
                    })
                })
                .then(response => response.json())
                .then(data => {
                    if (data.status === 'success') {
                        console.log('後端檔案已重新命名:', newFullName);
                    } else {
                        console.error('後端重新命名失敗:', data.error);
                        alert(`後端檔案重新命名失敗: ${data.error || '未知錯誤'}`);
                    }
                })
                .catch(error => {
                    console.error('重新命名請求失敗:', error);
                    alert(`重新命名請求失敗: ${error.message}`);
                });

                // 更新儲存的完整檔名
                cases[caseId].excelFileName = newFullName;
            }

            cases[caseId].name = newName.trim();

            // 同步更新 window.currentCase（供 price_rolling.js 等使用）
            if (window.currentCase && window.currentCase.id === caseId) {
                window.currentCase.name = newName.trim();
            }

            saveCases();
            renderCaseList();
        }
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

    // 更改案場類型
    function changeCaseType(caseId, newType) {
        const caseData = cases[caseId];
        const oldType = caseData.siteType || 'single';

        if (oldType === newType) return;

        caseData.siteType = newType;
        saveCases();
        renderCaseList();

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

    function switchActiveCase(caseId) {
        // 避免重複切換同一
        if (activeCaseId === caseId) return;

        activeCaseId = caseId;

        // 設置全局變量，供價金滾算等功能使用
        window.currentCase = {
            id: caseId,
            name: cases[caseId].name,
            original_filename: cases[caseId].excelOriginalFileName || cases[caseId].originalFilename || '',
            sheet_name: cases[caseId].sheetName || null
        };

        // 批量更新 UI - 避免多次重繪
        requestAnimationFrame(() => {
            renderCaseList();
            renderChat();

            // 更新標題
            if (caseTitle) {
                caseTitle.textContent = cases[caseId].name;
            }
            
            // 載入該案場的 Excel 數據
            if (cases[caseId].excelData) {
                if (luckysheetInitialized && isLuckysheetReady) {
                    updateLuckysheetData(cases[caseId].excelData);
                } else {
                    console.log('使用完整表格顯示已保存的 Excel 數據');
                    createSimpleTable(cases[caseId].excelData);
                }
            } else {
                // 案場沒有 Excel 數據，顯示空白表格
                if (luckysheetInitialized && isLuckysheetReady) {
                    // 清空 Luckysheet，顯示空白表格
                    console.log('切換到沒有 Excel 數據的案場，顯示空白表格');
                    updateLuckysheetData(null);
                } else {
                    // Luckysheet 未初始化，先顯示簡化表格避免黑屏
                    console.log('Luckysheet 未初始化，顯示簡化表格');
                    const container = document.getElementById('luckysheet-container');
                    if (container) {
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
                    // 不再嘗試初始化 Luckysheet，避免黑屏
                }
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

    function appendMessage(role, text) {
        if (!chatMessages) return;
        
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${role}`;
        // 處理換行符，保持格式
        messageDiv.innerHTML = text.replace(/\n/g, '<br>');
        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    function sendMessage() {
        if (!userInput || !chatMessages) return;

        const text = userInput.value.trim();
        if (!text || !activeCaseId) return;

        // 前端輸入驗證：輸入過短可能是錯字
        if (text.length < 3) {
            userInput.placeholder = '請輸入至少 3 個字元描述您的需求';
            userInput.value = '';
            return;
        }

        // 顯示使用者訊息
        const userMessage = { role: 'user', text: text };
        cases[activeCaseId].messages.push(userMessage);
        appendMessage(userMessage.role, userMessage.text);
        userInput.value = '';
        userInput.placeholder = '輸入訊息...'; // 重置提示
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

        // 獲取當前案場的 Excel 檔案資訊
        const caseName = cases[activeCaseId].name || '';
        const originalFilename = cases[activeCaseId].excelOriginalFileName || '';

        // 呼叫後端 Agent API
        fetch('/api/agent_chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                query: text,
                equipment_cost_adj: costAdj,
                rent_adj: rentAdj,
                loan_amount: loanAmount, // 將借款金額傳遞給後端
                case_name: caseName, // 傳遞案場名稱
                original_filename: originalFilename // 傳遞原始檔名
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
            // 移除載入提示
            const loadingElement = document.getElementById(loadingId);
            if (loadingElement) loadingElement.remove();

            // 顯示 Agent 回應
            const botMessage = { role: 'bot', text: data.response };
            cases[activeCaseId].messages.push(botMessage);
            appendMessage(botMessage.role, botMessage.text);
            saveCases();

            // 檢查是否需要重新載入 Excel
            if (data.excel_modified && activeCaseId) {
                console.log('檢測到 Excel 已被修改，準備重新載入...');
                reloadExcelData(activeCaseId);
            }
        })
        .catch(error => {
            console.error('錯誤:', error);
            // 移除載入提示
            const loadingElement = document.getElementById(loadingId);
            if (loadingElement) loadingElement.remove();

            const errorMessage = { role: 'bot', text: `發生錯誤: ${error.message}` };
            cases[activeCaseId].messages.push(errorMessage);
            appendMessage(errorMessage.role, errorMessage.text);
            saveCases();
        });
    }

    // 重新載入 Excel 數據
    function reloadExcelData(caseId) {
        console.log(`正在重新載入案場 ${caseId} 的 Excel 數據...`);

        // 傳送前端名稱和原始檔名
        const caseName = cases[caseId].name;
        const originalFileName = cases[caseId].excelOriginalFileName;

        fetch(`/api/read_excel/${caseId}?case_name=${encodeURIComponent(caseName)}&original_filename=${encodeURIComponent(originalFileName || '')}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error('無法讀取 Excel 檔案');
                }
                return response.json();
            })
            .then(data => {
                if (data.status === 'success' && data.data) {
                    console.log('Excel 數據已成功載入');

                    // 更新案場的 Excel 數據
                    cases[caseId].excelData = data.data;
                    saveCases();

                    // 如果是當前案場，更新 Luckysheet 顯示
                    if (caseId === activeCaseId) {
                        if (luckysheetInitialized && isLuckysheetReady) {
                            updateLuckysheetData(data.data);
                            console.log('Luckysheet 已更新顯示');
                        } else {
                            createSimpleTable(data.data);
                            console.log('表格已更新顯示');
                        }

                        // 顯示更新提示訊息
                        const updateMessage = {
                            role: 'bot',
                            text: 'Excel 表格已自動更新'
                        };
                        cases[activeCaseId].messages.push(updateMessage);
                        appendMessage(updateMessage.role, updateMessage.text);
                        saveCases();
                    }
                } else {
                    console.warn('Excel 數據格式不正確或為空');
                }
            })
            .catch(error => {
                console.error('重新載入 Excel 失敗:', error);
                // 如果找不到檔案，可能是因為還沒上傳，不顯示錯誤
                if (!error.message.includes('404')) {
                    const errorMsg = {
                        role: 'bot',
                        text: `Excel 更新失敗: ${error.message}`
                    };
                    cases[activeCaseId].messages.push(errorMsg);
                    appendMessage(errorMsg.role, errorMsg.text);
                    saveCases();
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
                const workbook = XLSX.read(data, {type: 'array'});
                
                // 轉換為 Luckysheet 格式
                const luckysheetData = convertToLuckysheet(workbook);

                // 保存到當前
                cases[activeCaseId].excelData = luckysheetData;
                cases[activeCaseId].hasExcel = true;
                cases[activeCaseId].excelFileName = file.name;
                cases[activeCaseId].excelOriginalFileName = file.name;  // 儲存原始檔名 
                // 嘗試在 Luckysheet 中載入數據，失敗則使用完整表格顯示
                if (luckysheetInitialized && isLuckysheetReady) {
                    updateLuckysheetData(luckysheetData);
                } else {
                    console.log('Luckysheet 不可用，使用完整表格顯示模式');
                    createSimpleTable(luckysheetData);
                }
                
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
    function uploadFileToBackend(file) {
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
        .then(data => {
            console.log('檔案上傳成功:', data);

            if(data.original_filename){
                cases[activeCaseId].excelOriginalFileName = data.original_filename;  // 儲存原始檔名
                saveCases();

                // 同步更新 window.currentCase
                if (window.currentCase && window.currentCase.id === activeCaseId) {
                    window.currentCase.original_filename = data.original_filename;
                }
            }

            // 上傳成功訊息（包含功能說明）
            const successMessage = {
                role: 'bot',
                text: `Excel 檔案已上傳到伺服器：${data.original_filename}\n\n` +
                      `功能說明：\n` +
                      `- 價金滾算：請點擊上方的「價金滾算」按鈕來執行滾算計算\n` +
                      `- 匯入表格：點選上方「匯入表格」可以插入其他案場的資訊\n` +
                      `- 匯出表格：若想匯出結果，可以點擊上方的「匯出表格」`
            };
            cases[activeCaseId].messages.push(successMessage);
            renderChat();
            saveCases();
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
        const sheets = [];
        
        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];
            
            // 處理空工作表
            if (!worksheet['!ref']) {
                sheets.push({
                    name: sheetName,
                    index: index.toString(),
                    status: index === 0 ? "1" : "0",
                    order: index.toString(),
                    celldata: [],
                    row: 20,
                    column: 10,
                    config: {},
                    pivotTable: null,
                    isPivotTable: false,
                    luckysheet_select_save: [],
                    calcChain: [],
                    hyperlink: {},
                    dataVerification: {}
                });
                return;
            }
            
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const celldata = [];
            
            // 單元格轉換 - 處理有值或有公式的單元格
            for (let row = range.s.r; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                    const cell = worksheet[cellAddress];

                    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;

                    const vObj = {
                        v: cell.v,
                        ct: {
                            fa: "General",
                            t: cell.t === 'n' ? 'n' : 's'
                        },
                        m: cell.w || String(cell.v)
                    };

                    // 處理數字格式
                    if (cell.t === 'n' && cell.z) {
                        vObj.ct.fa = cell.z;
                    }

                    celldata.push({ r: row, c: col, v: vObj });
                }
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
                    merge: worksheet['!merges'] || {},
                    rowlen: {},
                    columnlen: {},
                    rowhidden: {},
                    colhidden: {},
                    borderInfo: [],
                    authority: {}
                },
                pivotTable: null,
                isPivotTable: false,
                luckysheet_select_save: [],
                calcChain: [],
                hyperlink: {},
                dataVerification: {}
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

        container.innerHTML = `
            <div style="padding: 10px; border: 1px solid #ddd; background: #f9f9f9; border-radius: 4px; height: 100%;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                    <h4 style="margin: 0; color: #333;">📊 Excel 內容完整顯示</h4>
                    <div style="font-size: 12px; color: #666;">
                        <span id="excel-info">等待載入 Excel 檔案...</span>
                    </div>
                </div>
                <div id="simple-table-container" style="height: calc(100% - 60px); overflow: auto; background: white; border: 1px solid #ccc; padding: 10px;"></div>
                <div style="margin-top: 10px; font-size: 11px; color: #666; text-align: center;">
                    ℹ️ 完整數據模式 - 所有文字和數值都已保留，可與 LLM 討論分析
                </div>
            </div>
        `;

        if (data && data.length > 0) {
            displayCompleteExcelData(data);
        } else {
            // 如果沒有數據，顯示提示訊息
            const simpleTableContainer = document.getElementById('simple-table-container');
            if (simpleTableContainer) {
                simpleTableContainer.innerHTML = `
                    <div style="display: flex; align-items: center; justify-content: center; height: 100%; flex-direction: column;">
                        <div style="text-align: center; color: #666;">
                            <div style="font-size: 48px; margin-bottom: 20px;">📄</div>
                            <h4 style="margin: 10px 0; color: #2c3e50;">尚未匯入 Excel 檔案</h4>
                            <p style="margin: 10px 0;">請點擊上方「匯入 Excel」按鈕選擇檔案</p>
                        </div>
                    </div>
                `;
            }
        }
    }

    // 完整顯示 Excel 數據，確保所有內容都被抓取
    function displayCompleteExcelData(sheets) {
        const tableContainer = document.getElementById('simple-table-container');
        const infoElement = document.getElementById('excel-info');
        
        if (!tableContainer) return;
        
        let html = '';
        let totalCells = 0;
        let totalSheets = sheets.length;
        
        sheets.forEach((sheet, sheetIndex) => {
            const sheetCells = sheet.celldata ? sheet.celldata.length : 0;
            totalCells += sheetCells;
            
            html += `
                <div style="margin-bottom: 25px; border: 1px solid #e0e0e0; border-radius: 6px;">
                    <div style="background: #34495e; color: white; padding: 8px 12px; border-radius: 6px 6px 0 0;">
                        <strong>📋 ${sheet.name}</strong> 
                        <span style="font-size: 11px; opacity: 0.8;">(${sheetCells} 個儲存格)</span>
                    </div>
                    <div style="padding: 15px;">
            `;
            
            if (sheet.celldata && sheet.celldata.length > 0) {
                // 建立完整的儲存格網格
                const maxRow = Math.max(...sheet.celldata.map(cell => cell.r)) + 1;
                const maxCol = Math.max(...sheet.celldata.map(cell => cell.c)) + 1;
                
                // 創建儲存格數據映射
                const cellMap = new Map();
                sheet.celldata.forEach(cell => {
                    const key = `${cell.r}-${cell.c}`;
                    cellMap.set(key, cell);
                });
                
                html += `
                    <div style="margin-bottom: 10px; font-size: 12px; color: #666;">
                        📐 工作表大小: ${maxRow} 行 × ${maxCol} 列
                    </div>
                    <table style="border-collapse: collapse; width: 100%; margin-bottom: 15px; font-size: 12px;">
                `;
                
                // 添加列標題 (A, B, C...)
                html += '<tr style="background: #ecf0f1;"><th style="border: 1px solid #bdc3c7; padding: 4px 6px; min-width: 30px; font-weight: bold;"></th>';
                for (let c = 0; c < maxCol; c++) {
                    const colName = String.fromCharCode(65 + (c % 26));
                    html += `<th style="border: 1px solid #bdc3c7; padding: 4px 6px; font-weight: bold; background: #ecf0f1;">${colName}</th>`;
                }
                html += '</tr>';
                
                // 顯示所有行的數據
                for (let r = 0; r < maxRow; r++) {
                    html += `<tr>`;
                    // 行號
                    html += `<td style="border: 1px solid #bdc3c7; padding: 4px 6px; background: #ecf0f1; font-weight: bold; text-align: center;">${r + 1}</td>`;
                    
                    for (let c = 0; c < maxCol; c++) {
                        const key = `${r}-${c}`;
                        const cell = cellMap.get(key);
                        
                        let cellValue = '';
                        let cellStyle = 'border: 1px solid #bdc3c7; padding: 4px 8px; vertical-align: top; max-width: 200px; word-wrap: break-word;';
                        
                        if (cell && cell.v) {
                            // 優先使用格式化的文字，然後是原始值
                            cellValue = cell.v.m || cell.v.v || '';
                            
                            // 根據數據類型設定樣式
                            if (cell.v.t === 'n') {
                                cellStyle += ' text-align: right; font-family: monospace;';
                            }
                            
                            // 如果是重要數據，加粗顯示
                            if (typeof cellValue === 'string' && (
                                cellValue.includes('%') || 
                                cellValue.includes('$') || 
                                cellValue.includes('總') || 
                                cellValue.includes('合計') ||
                                cellValue.includes('收入') ||
                                cellValue.includes('成本') ||
                                cellValue.includes('費用')
                            )) {
                                cellStyle += ' font-weight: bold; background: #fff3cd;';
                            }
                        }
                        
                        // 確保所有內容都被保留，包括空格和特殊字符
                        const safeValue = String(cellValue)
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;')
                            .replace(/"/g, '&quot;')
                            .replace(/'/g, '&#39;')
                            .replace(/\n/g, '<br>');
                        
                        html += `<td style="${cellStyle}" title="${safeValue}">${safeValue || '&nbsp;'}</td>`;
                    }
                    html += '</tr>';
                }
                
                html += '</table>';
                
                // 顯示工作表統計信息
                html += `
                    <div style="background: #f8f9fa; padding: 8px; border-radius: 4px; font-size: 11px; color: #495057;">
                        📊 統計: 共 ${sheetCells} 個有內容的儲存格，顯示範圍 A1:${String.fromCharCode(65 + maxCol - 1)}${maxRow}
                    </div>
                `;
                
            } else {
                html += '<div style="text-align: center; color: #6c757d; padding: 20px;">📄 此工作表沒有內容</div>';
            }
            
            html += '</div></div>';
        });
        
        // 更新信息顯示
        if (infoElement) {
            infoElement.textContent = `${totalSheets} 個工作表，共 ${totalCells} 個儲存格`;
        }
        
        tableContainer.innerHTML = html;
        
        console.log(`完整顯示 Excel 內容: ${totalSheets} 個工作表，${totalCells} 個儲存格`);
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
                        <span class="import-case-name">${caseData.case_name}</span>
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
                    }))
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
});