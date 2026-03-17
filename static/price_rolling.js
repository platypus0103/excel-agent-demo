// 價金滾算 JavaScript 邏輯

// 頁面載入時初始化（無預填值）
document.addEventListener('DOMContentLoaded', function() {
    // 不預填任何數值，讓使用者自行輸入
});

// 打開對話框
// 金額千分位格式化（四捨五入取整）
function fmtMoney(v) {
    if (v === null || v === undefined || v === 'N/A') return 'N/A';
    const n = Math.round(Number(v));
    return isNaN(n) ? String(v) : n.toLocaleString('zh-TW');
}

function openPriceDialog() {
    document.getElementById('priceDialogOverlay').classList.add('active');
    initializePriceDialogDrag();

    // 從 Excel 讀取 C16(初始價金)、C17(利潤率)、C18(開發費) 預填至表單
    const caseInfo = getCurrentCaseInfo();
    let url = '/api/get_excel_defaults';
    if (caseInfo && caseInfo.original_filename) {
        url += `?case_id=${encodeURIComponent(caseInfo.case_id || '')}&case_name=${encodeURIComponent(caseInfo.case_name || '')}&original_filename=${encodeURIComponent(caseInfo.original_filename)}`;
    }

    fetch(url)
        .then(r => r.json())
        .then(data => {
            if (data.status === 'success' && data.defaults) {
                const d = data.defaults;
                const costEl = document.getElementById('priceEquipmentCost');
                const rateEl = document.getElementById('priceProfitRate');
                const feeEl  = document.getElementById('priceDevelopmentFee');
                if (costEl && d.equipment_cost != null) costEl.value = d.equipment_cost;
                if (rateEl && d.profit_rate   != null) rateEl.value  = d.profit_rate;
                if (feeEl  && d.development_fee != null) feeEl.value  = d.development_fee;
            }
        })
        .catch(err => console.warn('無法讀取 Excel 預設值:', err));
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

// 防止價金滾算表單輸入負值（統一處理所有 price-dialog 內的 number 輸入）
function clampNonNegative(input) {
    if (parseFloat(input.value) < 0) {
        input.value = 0;
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
            <input type="number" id="priceManualStep${i}" class="price-manual-step-input" min="0" placeholder="輸入金額" oninput="clampNonNegative(this)">
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

    // 防止負值
    if (equipmentCost < 0 || profitRate < 0 || boundary < 0 || (developmentFee !== null && developmentFee < 0)) {
        alert('所有數值不得為負數！');
        return;
    }

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
        requestData.case_id = currentCase.case_id;
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
                requestData.step = ratioStep / 100;  // 前端輸入 % 值，轉成小數傳給後端
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

    // 在發送前鎖定案場本地 key，避免切換聊天室後回應跑到錯的地方
    // 注意：必須用 getActiveCaseId() 取本地 key，而非 getCurrentCaseInfo().case_id（那是 DB ID）
    const senderCaseId = (typeof window.getActiveCaseId === 'function') ? window.getActiveCaseId() : null;

    // 顯示加載提示（用唯一 ID，切換聊天室後仍可移除）
    const loadingId = 'pr-loading-' + Date.now();
    const loadingDiv = document.createElement('div');
    loadingDiv.id = loadingId;
    loadingDiv.className = 'message assistant';
    loadingDiv.textContent = '正在計算價金滾算，請稍候...';
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
        chatMessages.appendChild(loadingDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

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
        const loadingEl = document.getElementById(loadingId);
        if (loadingEl) loadingEl.remove();

        if (data.status === 'success') {
            if (typeof window.appendMessageToCase === 'function') {
                window.appendMessageToCase(senderCaseId, 'bot', data.response);
            }
        } else {
            if (typeof window.appendMessageToCase === 'function') {
                window.appendMessageToCase(senderCaseId, 'bot', `計算失敗: ${data.error}`);
            }
        }
    })
    .catch(error => {
        console.error('API 調用錯誤:', error);
        const loadingEl = document.getElementById(loadingId);
        if (loadingEl) loadingEl.remove();
        if (typeof window.appendMessageToCase === 'function') {
            window.appendMessageToCase(senderCaseId, 'bot', `發生錯誤: ${error.message}`);
        }
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
                <td>${fmtMoney(price)}</td>
                <td>${(profitRate * 100).toFixed(2)}%</td>
                <td>${fmtMoney(profitList[idx])}</td>
            </tr>
        `;
    });

    table.innerHTML = html;
    document.getElementById('priceResultSection').classList.add('active');
}

// 獲取當前案場信息
function getCurrentCaseInfo() {
    if (window.currentCase) {
        return {
            case_id:           window.currentCase.id || '',
            case_name:         window.currentCase.name || '',
            original_filename: window.currentCase.original_filename || '',
            sheet_name:        window.currentCase.sheet_name || null
        };
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

    // 使用 LLMweb.js 的 renderMarkdown 渲染（bot/assistant 訊息）
    if (type === 'assistant' || type === 'bot') {
        if (typeof window._renderMarkdown === 'function') {
            messageDiv.innerHTML = window._renderMarkdown(message);
        } else {
            messageDiv.innerHTML = message.replace(/\n/g, '<br>');
        }
    } else {
        messageDiv.textContent = message;
    }
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

// 下載 Excel Generic Template 中的公版（第三個工作表，含公式）
function downloadTemplate() {
    window.location.href = '/api/download_template';
}

// 匯出當前聊天室的 Excel 表格
function exportCurrentExcel() {
    // 獲取當前案場信息
    const currentCase = getCurrentCaseInfo();

    if (!currentCase || !currentCase.case_name) {
        alert('請先上傳 Excel 檔案！');
        return;
    }

    // 構建下載 URL，包含 case_id / 案場名稱 / 原始檔名以確保路徑正確
    const downloadUrl = `/api/download_excel?case_id=${encodeURIComponent(currentCase.case_id || '')}&case_name=${encodeURIComponent(currentCase.case_name)}&original_filename=${encodeURIComponent(currentCase.original_filename || '')}`;

    // 直接下載服務器上的原始 Excel 檔案（保留公式）
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

    // 對話框內所有靜態 number input 統一防止輸入負值
    const priceDialog = document.getElementById('priceDialog');
    if (priceDialog) {
        priceDialog.querySelectorAll('input[type="number"]').forEach(function(input) {
            input.addEventListener('input', function() { clampNonNegative(this); });
        });
    }
});
