// 價金滾算 JavaScript 邏輯

// 載入 Excel 預設值
async function loadExcelDefaults() {
    try {
        const response = await fetch('/api/get_excel_defaults');
        const data = await response.json();
        
        if (data.status === 'success') {
            const defaults = data.defaults;
            
            // 設定基本參數預設值
            document.getElementById('priceEquipmentCost').value = defaults.equipment_cost || 30000;
            document.getElementById('priceProfitRate').value = defaults.profit_rate || 0.2;
            document.getElementById('priceDevelopmentFee').value = defaults.development_fee || 0;
            document.getElementById('priceBoundary').value = defaults.boundary || 20000;
            
            // 設定各模式參數預設值
            document.getElementById('priceCashStep').value = defaults.cash_step || 2000;
            document.getElementById('priceRatioStep').value = defaults.ratio_step || 0.05;
            
            // 條件模式預設值
            document.getElementById('priceMaxValue').value = defaults.max_value || 50000;
            document.getElementById('priceMinValue').value = defaults.min_value || 30000;
            document.getElementById('priceCondStep1').value = defaults.cond_step_1 || 2000;
            document.getElementById('priceCondStep2').value = defaults.cond_step_2 || 1000;
            document.getElementById('priceCondStep3').value = defaults.cond_step_3 || 500;
            
            // 自訂模式預設值
            document.getElementById('priceAdjustTimes').value = defaults.adjust_times || 5;
            
            console.log('Excel 預設值載入成功:', defaults);
        } else {
            console.warn('載入 Excel 預設值失敗:', data.message);
            // 使用備用預設值
            setFallbackDefaults();
        }
    } catch (error) {
        console.error('載入 Excel 預設值時發生錯誤:', error);
        // 使用備用預設值
        setFallbackDefaults();
    }
}

// 設定備用預設值
function setFallbackDefaults() {
    document.getElementById('priceEquipmentCost').value = 30000;
    document.getElementById('priceProfitRate').value = 0.2;
    document.getElementById('priceDevelopmentFee').value = 0;
    document.getElementById('priceBoundary').value = 20000;
    document.getElementById('priceCashStep').value = 2000;
    document.getElementById('priceRatioStep').value = 0.05;
    document.getElementById('priceMaxValue').value = 50000;
    document.getElementById('priceMinValue').value = 30000;
    document.getElementById('priceCondStep1').value = 2000;
    document.getElementById('priceCondStep2').value = 1000;
    document.getElementById('priceCondStep3').value = 500;
    document.getElementById('priceAdjustTimes').value = 5;
}

// 頁面載入時初始化
document.addEventListener('DOMContentLoaded', function() {
    loadExcelDefaults();
});

// 打開對話框
function openPriceDialog() {
    document.getElementById('priceDialogOverlay').classList.add('active');
    initializePriceDialogDrag();
    // 每次打開對話框時重新載入預設值
    loadExcelDefaults();
}

// 關閉對話框
function closePriceDialog() {
    document.getElementById('priceDialogOverlay').classList.remove('active');
    // 重置對話框位置，避免下次打開時閃爍
    const dialog = document.getElementById('priceDialog');
    dialog.style.transform = 'translate(0, 0)';
    // 重置拖動偏移量
    if (window.priceDialogDragState) {
        window.priceDialogDragState.xOffset = 0;
        window.priceDialogDragState.yOffset = 0;
    }
}

// 關閉結果
function closePriceResult() {
    document.getElementById('priceResultSection').classList.remove('active');
}

// 拖動功能初始化
function initializePriceDialogDrag() {
    const dialog = document.getElementById('priceDialog');
    const header = document.getElementById('priceDialogHeader');

    // 初始化全局狀態（如果不存在）
    if (!window.priceDialogDragState) {
        window.priceDialogDragState = {
            xOffset: 0,
            yOffset: 0,
            isDragging: false
        };
    }

    let currentX;
    let currentY;
    let initialX;
    let initialY;

    // 移除舊的事件監聽器（避免重複綁定）
    header.removeEventListener('mousedown', header.dragStartHandler);
    document.removeEventListener('mousemove', header.dragHandler);
    document.removeEventListener('mouseup', header.dragEndHandler);

    function dragStart(e) {
        if (e.target.classList.contains('price-close-btn')) return;

        initialX = e.clientX - window.priceDialogDragState.xOffset;
        initialY = e.clientY - window.priceDialogDragState.yOffset;

        if (e.target === header || header.contains(e.target)) {
            window.priceDialogDragState.isDragging = true;
        }
    }

    function drag(e) {
        if (window.priceDialogDragState.isDragging) {
            e.preventDefault();

            currentX = e.clientX - initialX;
            currentY = e.clientY - initialY;

            window.priceDialogDragState.xOffset = currentX;
            window.priceDialogDragState.yOffset = currentY;

            setTranslate(currentX, currentY, dialog);
        }
    }

    function dragEnd(e) {
        initialX = currentX;
        initialY = currentY;
        window.priceDialogDragState.isDragging = false;
    }

    function setTranslate(xPos, yPos, el) {
        el.style.transform = `translate(${xPos}px, ${yPos}px)`;
    }

    // 保存事件處理器引用以便之後移除
    header.dragStartHandler = dragStart;
    header.dragHandler = drag;
    header.dragEndHandler = dragEnd;

    header.addEventListener('mousedown', dragStart);
    document.addEventListener('mousemove', drag);
    document.addEventListener('mouseup', dragEnd);
}

// 選擇模式
function selectPriceMode(mode) {
    // 取消所有選項的選中狀態
    document.querySelectorAll('.price-mode-option').forEach(opt => {
        opt.classList.remove('selected');
    });
    document.querySelectorAll('.price-mode-params').forEach(param => {
        param.classList.remove('active');
    });

    // 選中當前模式
    const radio = document.getElementById('priceMode' + mode.charAt(0).toUpperCase() + mode.slice(1));
    radio.checked = true;
    radio.closest('.price-mode-option').classList.add('selected');
    document.getElementById('price' + mode.charAt(0).toUpperCase() + mode.slice(1) + 'Params').classList.add('active');

    // 如果選擇自訂模式，檢查是否需要生成輸入框
    if (mode === 'customize') {
        generatePriceStepInputs();
    }
}

// 生成自訂 Step 輸入框
function generatePriceStepInputs() {
    const times = parseInt(document.getElementById('priceAdjustTimes').value);
    const container = document.getElementById('priceStepInputs');
    const manualContainer = document.getElementById('priceManualStepsContainer');
    const autoCheckbox = document.getElementById('priceAutoGenerate');

    if (!times || times < 1) {
        manualContainer.style.display = 'none';
        return;
    }

    // 如果勾選自動生成，不顯示手動輸入
    if (autoCheckbox.checked) {
        manualContainer.style.display = 'none';
        return;
    }

    manualContainer.style.display = 'block';
    container.innerHTML = '';

    for (let i = 1; i <= times; i++) {
        const div = document.createElement('div');
        div.className = 'price-form-group';
        div.innerHTML = `
            <label>第 ${i} 次 Step</label>
            <input type="number" id="priceManualStep${i}" class="price-manual-step-input" placeholder="輸入金額">
        `;
        container.appendChild(div);
    }
}

// 處理自動生成勾選框
function handlePriceAutoGenerate() {
    const autoCheckbox = document.getElementById('priceAutoGenerate');
    const manualContainer = document.getElementById('priceManualStepsContainer');
    const autoDisplay = document.getElementById('priceAutoStepsDisplay');

    if (autoCheckbox.checked) {
        manualContainer.style.display = 'none';
        autoDisplay.style.display = 'none';
    } else {
        generatePriceStepInputs();
        autoDisplay.style.display = 'none';
    }
}

// 提交表單
function submitPriceForm() {
    const form = document.getElementById('priceCostForm');
    if (!form.checkValidity()) {
        alert('請填寫所有必填欄位！');
        return;
    }

    const selectedMode = document.querySelector('input[name="priceMode"]:checked');
    if (!selectedMode) {
        alert('請選擇一個計算模式！');
        return;
    }

    // 取得基本參數
    const equipmentCost = parseInt(document.getElementById('priceEquipmentCost').value);
    const profitRate = parseFloat(document.getElementById('priceProfitRate').value);
    const developmentFeeInput = document.getElementById('priceDevelopmentFee').value;
    const developmentFee = developmentFeeInput && developmentFeeInput.trim() !== '' ? parseInt(developmentFeeInput) : null;
    const boundary = parseInt(document.getElementById('priceBoundary').value);

    // 準備 API 請求數據
    const requestData = {
        mode: selectedMode.value === 'cash' ? 'CashMode' :
              selectedMode.value === 'ratio' ? 'RatioMode' :
              selectedMode.value === 'conditional' ? 'ConditionalMode' :
              'CustomizeMode',
        equipment_cost: equipmentCost,
        profit_rate: profitRate,
        boundary: boundary
    };

    // 只有當用戶輸入了開發費時才傳送
    if (developmentFee !== null) {
        requestData.development_fee = developmentFee;
    }

    // 獲取當前案場信息
    const currentCase = getCurrentCaseInfo();
    if (currentCase) {
        requestData.case_name = currentCase.case_name;
        requestData.original_filename = currentCase.original_filename;
        requestData.sheet_name = currentCase.sheet_name;
    }

    // 根據選擇的模式添加特定參數
    let validationError = null;

    switch(selectedMode.value) {
        case 'cash':
            const cashStep = parseInt(document.getElementById('priceCashStep').value);
            if (!cashStep) {
                validationError = '請輸入調整 Step！';
            } else {
                requestData.step = cashStep;
            }
            break;

        case 'ratio':
            const ratioStep = parseFloat(document.getElementById('priceRatioStep').value);
            if (!ratioStep) {
                validationError = '請輸入調整比例！';
            } else {
                requestData.step = ratioStep;
            }
            break;

        case 'conditional':
            const maxValue = parseInt(document.getElementById('priceMaxValue').value);
            const minValue = parseInt(document.getElementById('priceMinValue').value);
            const step1 = parseInt(document.getElementById('priceCondStep1').value);
            const step2 = parseInt(document.getElementById('priceCondStep2').value);
            const step3 = parseInt(document.getElementById('priceCondStep3').value);

            if (!maxValue || !minValue || !step1 || !step2 || !step3) {
                validationError = '請填寫所有條件模式參數！';
            } else {
                requestData.max_value = maxValue;
                requestData.min_value = minValue;
                requestData.step1 = step1;
                requestData.step2 = step2;
                requestData.step3 = step3;
            }
            break;

        case 'customize':
            const adjustTimes = parseInt(document.getElementById('priceAdjustTimes').value);
            const isAutoGenerate = document.getElementById('priceAutoGenerate').checked;

            if (!adjustTimes) {
                validationError = '請輸入調整次數！';
            } else {
                requestData.adjust_times = adjustTimes;

                if (!isAutoGenerate) {
                    // 收集手動輸入的 Steps
                    let customSteps = [];
                    for (let i = 1; i <= adjustTimes; i++) {
                        const stepInput = document.getElementById(`priceManualStep${i}`);
                        if (!stepInput || !stepInput.value) {
                            validationError = `請輸入第 ${i} 次的 Step！`;
                            break;
                        }
                        customSteps.push(parseInt(stepInput.value));
                    }
                    if (!validationError) {
                        requestData.steps = customSteps;
                    }
                }
            }
            break;
    }

    if (validationError) {
        alert(validationError);
        return;
    }

    // 關閉對話框
    closePriceDialog();

    // 顯示加載提示
    addMessageToChatBox('系統', '正在計算價金滾算，請稍候...', 'assistant');

    // 調用後端 API
    fetch('/api/calculate_price_rolling', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestData)
    })
    .then(response => response.json())
    .then(data => {
        // 移除加載提示
        removeLastMessage();

        if (data.status === 'success') {
            // 將 Agent 回應插入聊天框
            addMessageToChatBox('AI 助手', data.response, 'assistant');
        } else {
            // 顯示錯誤
            addMessageToChatBox('系統', `計算失敗: ${data.error}`, 'error');
        }
    })
    .catch(error => {
        console.error('API 調用錯誤:', error);
        removeLastMessage();
        addMessageToChatBox('系統', `發生錯誤: ${error.message}`, 'error');
    });
}

// CashMode 計算
function calculateCashMode(equipmentCost, boundary, step) {
    const record = [equipmentCost];
    let adjustableRange = equipmentCost - boundary;

    while (equipmentCost > boundary && adjustableRange >= step) {
        equipmentCost -= step;
        adjustableRange = equipmentCost - boundary;
        record.push(equipmentCost);
    }

    return record;
}

// RatioMode 計算
function calculateRatioMode(equipmentCost, boundary, step) {
    const record = [equipmentCost];

    while (true) {
        const adjustedCost = equipmentCost * (1 - step);

        if (adjustedCost < boundary) {
            break;
        }

        equipmentCost = adjustedCost;
        record.push(Math.round(equipmentCost));
    }

    return record;
}

// ConditionalMode 計算
function calculateConditionalMode(equipmentCost, boundary, maxValue, minValue, step1, step2, step3) {
    const record = [equipmentCost];

    while (true) {
        let step;

        // 判斷使用哪個 step
        if (equipmentCost > maxValue) {
            step = step1;
        } else if (equipmentCost >= minValue && equipmentCost <= maxValue) {
            step = step2;
        } else {
            step = step3;
        }

        // 先檢查邊界
        const adjustableRange = equipmentCost - boundary;
        if (equipmentCost <= boundary || adjustableRange < step) {
            break;
        }

        // 再減去 step
        equipmentCost -= step;
        record.push(equipmentCost);
    }

    return record;
}

// 自動生成 Step
function automaticConfiguration(equipmentCost, boundary, adjustmentTimes) {
    const steps = [];
    let adjustableRange = equipmentCost - boundary;

    adjustmentTimes = Math.min(adjustmentTimes, 10);
    const averageStep = Math.floor(adjustableRange / adjustmentTimes);

    for (let i = 0; i < adjustmentTimes; i++) {
        // 在平均值 ±20% 範圍內生成隨機 step
        const step = Math.floor(Math.random() * (averageStep * 0.4) + averageStep * 0.8);

        // 確保這一次的step不會超過可調整範圍
        const finalStep = Math.min(step, adjustableRange);

        steps.push(finalStep);
        adjustableRange -= finalStep;

        if (adjustableRange <= 0) {
            break;
        }
    }

    return steps;
}

// CustomizeMode 計算
function calculateCustomizeMode(equipmentCost, steps) {
    const record = [equipmentCost];

    for (const step of steps) {
        equipmentCost -= step;
        record.push(equipmentCost);
    }

    return record;
}

// 顯示結果
function displayPriceResults(prices, profitList, profitRate) {
    const table = document.getElementById('priceResultTable');
    let html = `
        <tr>
            <th>次數</th>
            <th>價金 / kW</th>
            <th>信邦利潤率</th>
            <th>信邦利潤 / kW</th>
        </tr>
    `;

    prices.forEach((price, idx) => {
        html += `
            <tr>
                <td>${idx + 1}</td>
                <td>${price}</td>
                <td>${(profitRate * 100).toFixed(2)}%</td>
                <td>${profitList[idx]}</td>
            </tr>
        `;
    });

    table.innerHTML = html;
    document.getElementById('priceResultSection').classList.add('active');
}

// 獲取當前案場信息
function getCurrentCaseInfo() {
    // 從 localStorage 或全局變量中獲取當前案場信息
    // 這個函數需要與 LLMweb.js 中的案場管理邏輯配合

    // 嘗試從 window 對象獲取（如果 LLMweb.js 有設置全局變量）
    if (window.currentCase) {
        return {
            case_name: window.currentCase.name,
            original_filename: window.currentCase.original_filename,
            sheet_name: window.currentCase.sheet_name || null
        };
    }

    // 嘗試從 localStorage 獲取
    const storedCase = localStorage.getItem('currentCase');
    if (storedCase) {
        try {
            const caseData = JSON.parse(storedCase);
            return {
                case_name: caseData.name,
                original_filename: caseData.original_filename,
                sheet_name: caseData.sheet_name || null
            };
        } catch (e) {
            console.error('解析案場信息失敗:', e);
        }
    }

    return null;
}

// 將消息添加到聊天框
function addMessageToChatBox(sender, message, type = 'assistant') {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) {
        console.error('找不到聊天框元素');
        return;
    }

    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${type}`;

    // 創建消息內容
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';

    // 如果消息包含 Markdown，使用 marked 庫渲染（如果可用）
    if (typeof marked !== 'undefined') {
        contentDiv.innerHTML = marked.parse(message);
    } else {
        // 簡單的換行處理
        contentDiv.innerHTML = message.replace(/\n/g, '<br>');
    }

    messageDiv.appendChild(contentDiv);
    chatMessages.appendChild(messageDiv);

    // 滾動到底部
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// 移除最後一條消息
function removeLastMessage() {
    const chatMessages = document.getElementById('chatMessages');
    if (!chatMessages) return;

    const messages = chatMessages.getElementsByClassName('message');
    if (messages.length > 0) {
        const lastMessage = messages[messages.length - 1];
        lastMessage.remove();
    }
}

// 匯出當前聊天室的 Excel 表格
function exportCurrentExcel() {
    // 獲取當前案場信息
    const currentCase = getCurrentCaseInfo();

    if (!currentCase || !currentCase.case_name) {
        alert('請先上傳 Excel 檔案！');
        return;
    }

    // 構建下載 URL，包含案場名稱和原始檔名以確保路徑正確
    const downloadUrl = `/api/download_excel?case_name=${encodeURIComponent(currentCase.case_name)}&original_filename=${encodeURIComponent(currentCase.original_filename || '')}`;

    // 直接使用瀏覽器下載（保留原始檔案格式，包括公式）
    window.location.href = downloadUrl;
}

// 點擊遮罩層關閉對話框
document.addEventListener('DOMContentLoaded', function() {
    // 綁定價金滾算按鈕
    const priceRollingBtn = document.getElementById('priceRollingBtn');
    if (priceRollingBtn) {
        priceRollingBtn.addEventListener('click', openPriceDialog);
    }

    // 點擊遮罩層關閉
    document.getElementById('priceDialogOverlay').addEventListener('click', function(e) {
        if (e.target === this) {
            closePriceDialog();
        }
    });
});
