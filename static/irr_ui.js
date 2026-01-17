/**
 * IRR 計算器 UI 交互邏輯
 * 保留原有的界面功能，移除計算邏輯
 */

// --------------------------------------------------
// 用戶界面交互函數
// --------------------------------------------------

/**
 * 切換輸入模式（每年不同 vs 年份範圍攤平 vs 以KW計算）
 * @param {string} type - 模式類型（income, interest, rent等）
 */
function toggle_mode(type) {
    const range_div = document.getElementById(`${type}_range`);
    const kw_div = document.getElementById(`${type}_kw_based`);
    const decay_div = document.getElementById(`${type}_decay`);
    const area_div = document.getElementById(`${type}_area_based`);
    const new_project_div = document.getElementById(`${type}_new_project`); 
    const selected_mode = document.querySelector(`input[name="${type}_mode"]:checked`).value;
    const electricity_div = document.getElementById(`${type}_electricity_based`); // 【新增】取得電費輸入區塊
    console.log(`切換 ${type} 模式到: ${selected_mode}`);

    // 隱藏所有模式
    if (range_div) range_div.classList.add('hidden');
    if (kw_div) kw_div.classList.add('hidden');
    if (decay_div) decay_div.classList.add('hidden');
    if (area_div) area_div.classList.add('hidden');
    if (new_project_div) new_project_div.classList.add('hidden'); 
    if (electricity_div) electricity_div.classList.add('hidden');

    // 顯示選中的模式
    if (selected_mode === 'range' && range_div) {
        range_div.classList.remove('hidden');
        initializeRangeSliders(type);
    } else if (selected_mode === 'kw_based' && kw_div) {
        kw_div.classList.remove('hidden');
        initializeRangeSliders(`${type}_kw`);
    } else if (selected_mode === 'decay' && decay_div) {
        decay_div.classList.remove('hidden');
        initializeRangeSliders(`${type}_decay`);
    } else if (selected_mode === 'area_based' && area_div) {
        area_div.classList.remove('hidden');
        initializeRangeSliders(`${type}_area`);
    } else if (selected_mode === 'new_project' && new_project_div) { // 【新增】 顯示新面板的邏輯
        new_project_div.classList.remove('hidden');
        initializeRangeSliders(`${type}_new`); // 匹配 'generation_new' 前綴
    }else if (selected_mode === 'electricity_based' && electricity_div) {
        electricity_div.classList.remove('hidden');
        // 注意：這裡傳入 `${type}_elec`，因為 HTML 中的 ID 是 maintenance_elec_start_range
        initializeRangeSliders(`${type}_elec`);
    }
}
/**
 * 初始化年份範圍調整桿
 * @param {string} prefix - 前綴名稱
 */
function initializeRangeSliders(prefix) {
    const startYear = parseInt(document.getElementById('start_year').value) || 2025;
    const endYear = parseInt(document.getElementById('end_year').value) || 2030;

    // 更新調整桿的最小最大值
    const startRange = document.getElementById(`${prefix}_start_range`);
    const endRange = document.getElementById(`${prefix}_end_range`);
    const startInput = document.getElementById(`${prefix}_start_year`);
    const endInput = document.getElementById(`${prefix}_end_year`);

    if (startRange && endRange && startInput && endInput) {
        [startRange, endRange].forEach(slider => {
            slider.min = startYear;
            slider.max = endYear;
            slider.value = slider.id.includes('start') ? startYear : endYear;
        });

        startInput.min = startYear;
        startInput.max = endYear;
        startInput.value = startYear;

        endInput.min = startYear;
        endInput.max = endYear;
        endInput.value = endYear;
    }
}

/**
 * 更新範圍顯示（調整桿改變時）
 * @param {string} prefix - 前綴名稱
 */
function updateRangeDisplay(prefix) {
    const startRange = document.getElementById(`${prefix}_start_range`);
    const endRange = document.getElementById(`${prefix}_end_range`);
    const startInput = document.getElementById(`${prefix}_start_year`);
    const endInput = document.getElementById(`${prefix}_end_year`);

    if (startRange && endRange && startInput && endInput) {
        let startVal = parseInt(startRange.value);
        let endVal = parseInt(endRange.value);

        // 確保開始年份不超過結束年份
        if (startVal > endVal) {
            if (startRange === document.activeElement) {
                endVal = startVal;
                endRange.value = startVal;
            } else {
                startVal = endVal;
                startRange.value = endVal;
            }
        }

        // 更新輸入框
        startInput.value = startVal;
        endInput.value = endVal;

        console.log(`${prefix} 年份範圍: ${startVal} - ${endVal}`);
    }
}

/**
 * 從輸入框更新調整桿（輸入框改變時）
 * @param {string} prefix - 前綴名稱
 */
function updateRangeFromInput(prefix) {
    const startRange = document.getElementById(`${prefix}_start_range`);
    const endRange = document.getElementById(`${prefix}_end_range`);
    const startInput = document.getElementById(`${prefix}_start_year`);
    const endInput = document.getElementById(`${prefix}_end_year`);

    if (startRange && endRange && startInput && endInput) {
        let startVal = parseInt(startInput.value);
        let endVal = parseInt(endInput.value);

        const minYear = parseInt(startInput.min);
        const maxYear = parseInt(startInput.max);

        // 驗證輸入範圍
        startVal = Math.max(minYear, Math.min(maxYear, startVal || minYear));
        endVal = Math.max(minYear, Math.min(maxYear, endVal || maxYear));

        // 確保開始年份不超過結束年份
        if (startVal > endVal) {
            if (startInput === document.activeElement) {
                endVal = startVal;
                endInput.value = startVal;
            } else {
                startVal = endVal;
                startInput.value = endVal;
            }
        }

        // 更新調整桿
        startRange.value = startVal;
        endRange.value = endVal;

        console.log(`${prefix} 年份範圍更新: ${startVal} - ${endVal}`);
    }
}
/**
 * 收集運維費用設定數據
 */
function get_maintenance_inputs() {
    const mode = document.querySelector('input[name="maintenance_mode"]:checked').value;
    
    let result = {
        mode: mode,
        start_year: null,
        end_year: null,
        value: 0,
        extra_params: {}
    };

    if (mode === 'range') {
        // ... (維持原樣)
        result.start_year = parseInt(document.getElementById('maintenance_start_year').value);
        result.end_year = parseInt(document.getElementById('maintenance_end_year').value);
        result.value = parseFloat(document.getElementById('maintenance_annual_amount').value) || 0;
    } 
    else if (mode === 'kw_based') {
        // ... (維持原樣)
        result.start_year = parseInt(document.getElementById('maintenance_kw_start_year').value);
        result.end_year = parseInt(document.getElementById('maintenance_kw_end_year').value);
        result.value = parseFloat(document.getElementById('maintenance_price_per_kw').value) || 0;
    }
    // ... (在 get_maintenance_inputs 函數內)
    else if (mode === 'electricity_based') {
        // 
        const start_year = parseInt(document.getElementById('maintenance_elec_start_year').value);
        const end_year = parseInt(document.getElementById('maintenance_elec_end_year').value);
        const percent = parseFloat(document.getElementById('maintenance_elec_percent').value) || 0;

        //  修正：將數據封裝為 Pydantic 期望的 electricity_based_data 結構
        result.electricity_based_data = {
            revenue_percent: percent,
            start_year: start_year,
            end_year: end_year
        };
        
        // Pydantic 模型不需要頂層的 start/end year for this mode，所以我們只使用子物件的數據
        result.start_year = start_year;
        result.end_year = end_year;
        result.value = 0; 

        delete result.extra_params; 
    }

    return result;
}
/**
 * 自動計算保險費
 *  公式：保險費 = 建置容量 (kW) × 每kW價格 × 保險費率
 */
async function calculateInsuranceFeeAuto() {
    // 讀取所需參數
    const capacity = parseFloat(document.getElementById('capacity').value) || 0;
    const pricePerKw = parseFloat(document.getElementById('price_per_kw').value) || 0;
    const insuranceRate = parseFloat(document.getElementById('insurance_rate').value) || 0;

    const resultElement = document.getElementById('insurance_calculation_result');
    const yearlyInput = document.getElementById('insurance_yearly_input');
    const annualAmountInput = document.getElementById('insurance_annual_amount');

    if (capacity > 0 && pricePerKw > 0 && insuranceRate > 0) {
        // 完整公式：保險費 = 建置容量 × 每kW價格 × 保險費率
        const annualInsuranceFee = capacity * pricePerKw * insuranceRate;
        resultElement.textContent = `每年 ${annualInsuranceFee.toLocaleString('en-US', { maximumFractionDigits: 0 })} 元`;

        // 更新隱藏的 yearly input
        const startYear = parseInt(document.getElementById('insurance_start_year').value);
        const endYear = parseInt(document.getElementById('insurance_end_year').value);
        const projectStartYear = parseInt(document.getElementById('start_year').value);
        const projectEndYear = parseInt(document.getElementById('end_year').value);

        let yearlyValues = [];
        for (let year = projectStartYear; year <= projectEndYear; year++) {
            if (year >= startYear && year <= endYear) {
                yearlyValues.push(annualInsuranceFee);
            } else {
                yearlyValues.push(0);
            }
        }
        yearlyInput.value = yearlyValues.join(':');

        //  修復：更新 annual_amount 欄位供 collectDataByMode 使用
        annualAmountInput.value = annualInsuranceFee;

    } else {
        resultElement.textContent = '請先輸入建置容量、每kW價格和保險費率';
        yearlyInput.value = '';
        annualAmountInput.value = '';
    }
}

// --------------------------------------------------
// 頁面初始化和事件處理
// --------------------------------------------------

/**
 * 頁面載入完成後的初始化函數
 */
document.addEventListener('DOMContentLoaded', function() {
    // 初始化設備費用計算
    calculate_equipment_cost();

    // 初始化保險費年份範圍拉桿並執行初始計算
    initializeRangeSliders('insurance');
    calculateInsuranceFeeAuto();

    console.log('IRR 計算器載入完成 - API 版本');

    // 檢查 API 連接狀態
    checkAPIConnection();
});

/**
 * 檢查 API 連接狀態
 */
async function checkAPIConnection() {
    try {
        await apiClient.healthCheck();
        console.log('✅ Flask API 連接正常');
    } catch (error) {
        console.warn(' Flask API 連接失敗:', error);

        // 在界面上顯示警告
        const result_container = document.getElementById('result_container');
        result_container.innerHTML = `
            <div class="error">
                <h3> 後端連接異常</h3>
                <p>請確保 Flask 服務器正在運行 (http://localhost:5000)</p>
                <p>錯誤詳情: ${error.message}</p>
            </div>
        `;
    }
}

// --------------------------------------------------
// 調試和輔助函數
// --------------------------------------------------

/**
 * 輸出當前所有輸入數據（用於調試）
 */
function debug_print_all_inputs() {
    console.log('=== 當前輸入數據 ===');
    console.log('年度範圍:', {
        start: document.getElementById('start_year').value,
        end: document.getElementById('end_year').value
    });
    console.log('設備費用參數:', {
        capacity: document.getElementById('capacity').value,
        price_per_kw: document.getElementById('price_per_kw').value,
        profit_rate: document.getElementById('profit_rate').value,
        development_fee: document.getElementById('development_fee').value
    });
    console.log('所得稅率:', document.getElementById('tax_rate').value);
    console.log('==================');
}

/**
 * 測試 API 連接
 */
async function test_api_connection() {
    try {
        const response = await apiClient.healthCheck();
        console.log('API 健康檢查回應:', response);
        alert('API 連接正常！');
    } catch (error) {
        console.error('API 連接測試失敗:', error);
        alert(`API 連接失敗: ${error.message}`);
    }
}

/**
 * 切換利息計算模式（無利息 vs 銀行貸款）
 */
function toggle_interest_mode() {
    const checkbox = document.getElementById('no_interest_checkbox');
    const bankLoanParams = document.getElementById('bank_loan_params');

    if (checkbox.checked) {
        // 選擇無利息，隱藏銀行貸款參數
        bankLoanParams.style.display = 'none';
        console.log('利息模式: 無利息');
    } else {
        // 取消無利息，顯示銀行貸款參數
        bankLoanParams.style.display = 'block';
        console.log('利息模式: 銀行貸款');
    }
}

/**
 * 面積單位換算
 * @param {string} sourceUnit - 觸發換算的單位
 */
function convertArea(sourceUnit) {
    const hectare_to_sq_meter = 10000;
    const jia_to_sq_meter = 9699.17;
    const ping_to_sq_meter = 3.30579;

    const inputs = {
        hectare: document.getElementById('conv_hectare'),
        jia: document.getElementById('conv_jia'),
        ping: document.getElementById('conv_ping'),
        sq_meter: document.getElementById('conv_sq_meter')
    };

    const sourceValue = parseFloat(inputs[sourceUnit].value);
    if (isNaN(sourceValue)) {
        // 如果輸入無效，清空所有欄位
        for (const unit in inputs) {
            if (unit !== sourceUnit) {
                inputs[unit].value = '';
            }
        }
        return;
    }

    let valueInSqMeters;
    switch (sourceUnit) {
        case 'hectare':
            valueInSqMeters = sourceValue * hectare_to_sq_meter;
            break;
        case 'jia':
            valueInSqMeters = sourceValue * jia_to_sq_meter;
            break;
        case 'ping':
            valueInSqMeters = sourceValue * ping_to_sq_meter;
            break;
        case 'sq_meter':
            valueInSqMeters = sourceValue;
            break;
    }

    inputs.hectare.value = (valueInSqMeters / hectare_to_sq_meter).toFixed(6);
    inputs.jia.value = (valueInSqMeters / jia_to_sq_meter).toFixed(6);
    inputs.ping.value = (valueInSqMeters / ping_to_sq_meter).toFixed(4);
    inputs.sq_meter.value = valueInSqMeters.toFixed(2);
}


/**
 * 更新範圍輸入框（輸入框改變時）
 * @param {string} prefix - 前綴名稱
 */
function updateRangeInputs(prefix) {
    const startRange = document.getElementById(`${prefix}_start_range`);
    const endRange = document.getElementById(`${prefix}_end_range`);
    const startInput = document.getElementById(`${prefix}_start_year`);
    const endInput = document.getElementById(`${prefix}_end_year`);

    if (startRange && endRange && startInput && endInput) {
        let startVal = parseInt(startRange.value);
        let endVal = parseInt(endRange.value);

        // 確保開始年份不超過結束年份
        if (startVal > endVal) {
            if (startRange === document.activeElement) {
                endVal = startVal;
                endRange.value = startVal;
            } else {
                startVal = endVal;
                startRange.value = endVal;
            }
        }

        // 更新輸入框
        startInput.value = startVal;
        endInput.value = endVal;

        console.log(`${prefix} 年份範圍更新: ${startVal} - ${endVal}`);
    }
}

// 在全域添加測試函數，方便開發時調用
window.debug_print_all_inputs = debug_print_all_inputs;
window.test_api_connection = test_api_connection;