/**
 * 根據模式收集數據 (收入/支出/自定義)
 * @param {string} type - 欄位類型 (income, rent, maintenance, insurance, recycling, custom)
 */
function collectDataByMode(type) {
    const mode_radio = document.querySelector(`input[name="${type}_mode"]:checked`);
    if (!mode_radio) {
        // 如果沒有預選模式，預設為 range
        const range_div = document.getElementById(`${type}_range`);
        if (range_div) {
            const annual_amount = parseFloat(document.getElementById(`${type}_annual_amount`).value) || 0;
            const start_year = parseInt(document.getElementById(`${type}_start_year`).value) || 2025;
            const end_year = parseInt(document.getElementById(`${type}_end_year`).value) || 2030;
            return {
                mode: 'range',
                yearly_data: null,
                range_data: { annual_amount, start_year, end_year },
                kw_based_data: null,
                decay_data: null,
                area_based_data: null,
                new_project_data: null,
                electricity_based_data: null
            };
        }
        // 如果連 range 都沒有，就回傳空值
        return {
            mode: 'range',
            yearly_data: null,
            range_data: { annual_amount: 0, start_year: 0, end_year: 0 },
            kw_based_data: null,
            decay_data: null,
            area_based_data: null,
            new_project_data: null,
            electricity_based_data: null
        };
    }
    const mode = mode_radio.value;

    if (mode === 'range') {
        const annual_amount = parseFloat(document.getElementById(`${type}_annual_amount`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_end_year`).value) || 2030;
        return {
            mode: 'range',
            yearly_data: null,
            range_data: { annual_amount, start_year, end_year },
            kw_based_data: null,
            decay_data: null,
            area_based_data: null,
            new_project_data: null,
            electricity_based_data: null
        };
    } else if (mode === 'kw_based') {
        const price_per_kw = parseFloat(document.getElementById(`${type}_price_per_kw`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_kw_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_kw_end_year`).value) || 2030;
        return {
            mode: 'kw_based',
            yearly_data: null,
            range_data: null,
            kw_based_data: { price_per_kw, start_year, end_year },
            decay_data: null,
            area_based_data: null,
            new_project_data: null,
            electricity_based_data: null
        };
    } else if (mode === 'decay') {
        const first_year_amount = parseFloat(document.getElementById(`${type}_first_year_amount`).value) || 0;
        const decay_rate = parseFloat(document.getElementById(`${type}_decay_rate`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_decay_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_decay_end_year`).value) || 2030;
        return {
            mode: 'decay',
            yearly_data: null,
            range_data: null,
            kw_based_data: null,
            decay_data: { first_year_amount, decay_rate, start_year, end_year },
            area_based_data: null,
            new_project_data: null,
            electricity_based_data: null
        };
    } else if (mode === 'area_based') {
        const area = parseFloat(document.getElementById(`${type}_area_amount`).value) || 0;
        const unit = document.getElementById(`${type}_area_unit`).value;
        const price_per_area = parseFloat(document.getElementById(`${type}_price_per_area`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_area_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_area_end_year`).value) || 2030;
        return {
            mode: 'area_based',
            yearly_data: null,
            range_data: null,
            kw_based_data: null,
            decay_data: null,
            area_based_data: { area, unit, price_per_area, start_year, end_year },
            new_project_data: null,
            electricity_based_data: null
        };
    } else if (mode === 'new_project') {
        const first_year_est = parseFloat(document.getElementById(`${type}_new_first_year_est`).value) || 0;
        const first_year_decay = parseFloat(document.getElementById(`${type}_new_first_year_decay`).value) || 0;
        const remaining_years_decay = parseFloat(document.getElementById(`${type}_new_remaining_years_decay`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_new_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_new_end_year`).value) || 2030;

        return {
            mode: 'new_project',
            yearly_data: null,
            range_data: null,
            kw_based_data: null,
            decay_data: null,
            area_based_data: null,
            new_project_data: {
                first_year_est,
                first_year_decay,
                remaining_years_decay,
                start_year,
                end_year
            },
            electricity_based_data: null
        };
    } else if (mode === 'electricity_based') {
        const revenue_percent = parseFloat(document.getElementById(`${type}_elec_percent`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_elec_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_elec_end_year`).value) || 2030;
        return {
            mode: 'electricity_based',
            yearly_data: null,
            range_data: null,
            kw_based_data: null,
            decay_data: null,
            area_based_data: null,
            new_project_data: null,
            electricity_based_data: {
                revenue_percent,
                start_year,
                end_year
            }
        };
    } else {
        throw new Error(`不支援的模式: ${mode}`);
    }
}
/**
 * API 客戶端 - 處理與 Flask 後端的通信
 * 替換原本的本地計算邏輯
 */

// API 基礎 URL 
 const API_BASE_URL = '/api';

/**
 * API 調用的統一錯誤處理
 */
class APIClient {
    constructor() {
        this.baseURL = API_BASE_URL;
    }

    /**
     * 發送 HTTP 請求的通用方法
     */
    async makeRequest(url, options = {}) {
        try {
            const response = await fetch(`${this.baseURL}${url}`, {
                headers: {
                    'Content-Type': 'application/json',
                    ...options.headers
                },
                ...options
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || `HTTP ${response.status}`);
            }

            return data;
        } catch (error) {
            console.error('API 請求錯誤:', error);
            throw error;
        }
    }

    /**
     * 計算設備費用
     */
    async calculateEquipmentCost(params) {
        return await this.makeRequest('/irr/equipment-cost', {
            method: 'POST',
            body: JSON.stringify(params)
        });
    }

    /**
     * 計算 IRR
     */
    async calculateIRR(requestData) {
        return await this.makeRequest('/irr/calculate', {
            method: 'POST',
            body: JSON.stringify(requestData)
        });
    }

    /**
     * 健康檢查
     */
    async healthCheck() {
        return await this.makeRequest('/irr/health');
    }
}

// 創建全局 API 客戶端實例
const apiClient = new APIClient();

/**
 * 替換原本的設備費用計算函數
 */
async function calculate_equipment_cost() {
    const capacity = parseFloat(document.getElementById('capacity').value) || 0;
    const price_per_kw = parseFloat(document.getElementById('price_per_kw').value) || 0;
    const profit_rate = parseFloat(document.getElementById('profit_rate').value) || 0;
    const development_fee = parseFloat(document.getElementById('development_fee').value) || 0;

    const result_element = document.getElementById('equipment_cost_result');

    // 前端驗證
    if (capacity <= 0 || price_per_kw <= 0 || profit_rate < 0 || profit_rate >= 100) {
        result_element.textContent = '請輸入有效參數';
        result_element.style.color = '#dc2626';
        
        // 同時更新保險費計算
        if (typeof calculateInsuranceFeeAuto === 'function') {
            calculateInsuranceFeeAuto();
        }
        
        return 0;
    }

    try {
        // 調用 Flask API
        const response = await apiClient.calculateEquipmentCost({
            capacity,
            price_per_kw,
            profit_rate,
            development_fee
        });

        if (response.success) {
            result_element.textContent = response.formatted_cost;
            result_element.style.color = '#1e40af';
            
            // 同時更新保險費計算
            if (typeof calculateInsuranceFeeAuto === 'function') {
                calculateInsuranceFeeAuto();
            }
            
            return response.equipment_cost;
        } else {
            throw new Error(response.error);
        }
    } catch (error) {
        console.error('設備費用計算錯誤:', error);
        result_element.textContent = '計算錯誤';
        result_element.style.color = '#dc2626';
        
        // 同時更新保險費計算
        if (typeof calculateInsuranceFeeAuto === 'function') {
            calculateInsuranceFeeAuto();
        }
        
        return 0;
    }
}

/**
 * 獲取設備費用（對外接口）
 * 優先從 UI 讀取，如果沒有則重新計算
 */
async function get_equipment_cost() {
    const result_element = document.getElementById('equipment_cost_result');
    const textContent = result_element.textContent || '';

    // 從顯示的文本中解析數字，例如 "NT$ 5,843,137"
    if (textContent.includes('NT$')) {
        const numberString = textContent.replace(/NT\$|\s|,/g, '');
        const cost = parseFloat(numberString);
        if (!isNaN(cost)) {
            return cost;
        }
    }
    
    // 如果無法從 UI 解析，則重新計算
    return await calculate_equipment_cost();
}

/**
 * 主要 IRR 計算函數 - 調用 Flask API
 */
async function calculate_irr_main() {
    const result_container = document.getElementById('result_container');

    try {
        // 顯示加載狀態
        result_container.innerHTML = '<div class="loading">正在收集數據...</div>';

        // 收集前端數據
        const requestData = await collectFormData();

        // 顯示計算狀態
        result_container.innerHTML = '<div class="loading">正在計算IRR...</div>';

        // 調用 Flask API
        const response = await apiClient.calculateIRR(requestData);

        // 顯示結果
        displayAPIResults(response);

    } catch (error) {
        console.error('IRR 計算錯誤:', error);
        result_container.innerHTML = `
            <div class="error">
                <h3>計算錯誤</h3>
                <p>${error.message}</p>
            </div>
        `;
    }
}

/**
 * 收集表單數據並格式化為 API 請求格式
 */
async function collectFormData() {
    // 年度範圍
    const start_year = parseInt(document.getElementById('start_year').value) || 2025;
    const end_year = parseInt(document.getElementById('end_year').value) || 2030;

    // 驗證年度範圍
    if (start_year >= end_year) {
        throw new Error('結束年度必須大於起始年度');
    }
    if (start_year < 1900 || end_year < 1900) {
        throw new Error('年度必須大於1900');
    }

    // 設備參數
    const equipment_params = {
        capacity: parseFloat(document.getElementById('capacity').value) || 0,
        price_per_kw: parseFloat(document.getElementById('price_per_kw').value) || 0,
        profit_rate: parseFloat(document.getElementById('profit_rate').value) || 0,
        development_fee: parseFloat(document.getElementById('development_fee').value) || 0
    };


    // 發電度數數據
    const generation = collectDataByMode('generation');
    // 最終躉售費率
    const final_tariff = parseFloat(document.getElementById('final_tariff').value) || 0;
    // income 組合為 { mode, yearly_data, range_data, kw_based_data, final_tariff }
    const income = { ...generation, final_tariff };
    // 利息數據 (特殊處理)
    const interest = collectInterestData();
    const rent = collectDataByMode('rent');
    const maintenance = collectDataByMode('maintenance');
    const insurance = collectDataByMode('insurance');
    const recycling = collectDataByMode('recycling');
    const custom = collectDataByMode('custom');
    const custom_label = document.getElementById('custom_label')?.value || '自定義欄位';
    const custom_show = document.getElementById('custom_show')?.checked;
    // 收集顯示設定
    const show_cols = {};
    document.querySelectorAll('.show-col').forEach(cb => {
        show_cols[cb.dataset.col] = cb.checked;
    });

    // 所得稅率
    const tax_rate = parseFloat(document.getElementById('tax_rate').value) || 0;

    // 現金流量表參數
    const dividend_ratio = parseFloat(document.getElementById('dividend_ratio').value) || 0;
    const capital_reduction_period = parseInt(document.getElementById('capital_reduction_period').value) || 1;

    return {
        start_year,
        end_year,
        equipment_params,
    income,
        interest,
        rent,
        maintenance,
        insurance,
        recycling,
        custom,
        custom_label,
        custom_show,
        show_cols,
        tax_rate,
        cash_flow_params: {
            dividend_ratio,
            capital_reduction_period
        }
    };
    const mode_radio = document.querySelector(`input[name="${type}_mode"]:checked`);
    if (!mode_radio) {
        throw new Error(`請選擇 ${type} 的計算模式`);
    }
    const mode = mode_radio.value;

    if (mode === 'yearly') {
        const input = document.getElementById(`${type}_yearly_input`).value;
        if (!input || input.trim() === '') {
            throw new Error(`請輸入 ${type} 的年度數據`);
        }
        const yearly_values = input.split(':').map(x => {
            const value = parseFloat(x);
            return isNaN(value) ? 0 : value;
        });

        return {
            mode: 'yearly',
            yearly_data: {
                yearly_values
            },
            range_data: null,
            kw_based_data: null
        };
    } else if (mode === 'range') {
        const total_amount = parseFloat(document.getElementById(`${type}_total_amount`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_end_year`).value) || 2030;

        return {
            mode: 'range',
            yearly_data: null,
            range_data: {
                total_amount,
                start_year,
                end_year
            },
            kw_based_data: null
        };
    } else if (mode === 'kw_based') {
        const price_per_kw = parseFloat(document.getElementById(`${type}_price_per_kw`).value) || 0;
        const start_year = parseInt(document.getElementById(`${type}_kw_start_year`).value) || 2025;
        const end_year = parseInt(document.getElementById(`${type}_kw_end_year`).value) || 2030;

        return {
            mode: 'kw_based',
            yearly_data: null,
            range_data: null,
            kw_based_data: {
                price_per_kw,
                start_year,
                end_year
            }
        };
    } else {
        throw new Error(`不支援的模式: ${mode}`);
    }
}

/**
 * 收集利息數據 (特殊處理)
 */
function collectInterestData() {
    const noInterestCheckbox = document.getElementById('no_interest_checkbox');

    if (noInterestCheckbox.checked) {
        // 無利息模式
        return {
            no_interest: true,
            bank_loan_data: null
        };
    } else {
        // 銀行貸款模式
        const loanRatio = parseFloat(document.getElementById('loan_ratio').value);
        const bankRate = parseFloat(document.getElementById('bank_rate').value);
        const repaymentPeriod = parseInt(document.getElementById('repayment_period').value);

        // 驗證必要欄位
        if (isNaN(loanRatio) || loanRatio <= 0 || loanRatio > 100) {
            throw new Error('請輸入有效的貸款成數 (0-100%)');
        }
        if (isNaN(bankRate) || bankRate <= 0) {
            throw new Error('請輸入有效的銀行利率');
        }
        if (isNaN(repaymentPeriod) || repaymentPeriod <= 0) {
            throw new Error('請輸入有效的攤還期數');
        }

        return {
            no_interest: false,
            bank_loan_data: {
                loan_ratio: loanRatio,
                bank_rate: bankRate,
                repayment_period: repaymentPeriod
            }
        };
    }
}

/**
 * 顯示 API 回傳的結果
 */
function displayAPIResults(response) {
    const result_container = document.getElementById('result_container');

    let html = '';

    if (response.success && response.irr !== null && !isNaN(response.irr)) {
        const irr_color = response.irr > 0 ? '#22543d' : '#c53030';
        html = `
            <div class="success">✅ 計算成功完成！</div>
            <div class="result-display">
                <div class="result-value" style="color: ${irr_color}">
                    ${response.irr.toFixed(4)}%
                </div>
                <p style="text-align: center; font-size: 1.2em; color: #4a5568;">
                    專案法IRR
                </p>
            </div>
        `;
    } else {
        html = `
            <div class="error">
                無法計算IRR<br>
                ${response.error || '未知錯誤'}
            </div>
        `;
    }

    // 生成現金流表格
    if (response.cash_flows && response.cash_flows.length > 0) {
        html += generateCashFlowTable(response);
    }

    // 生成現金流量表
    if (response.cash_flow_statement && response.cash_flow_statement.length > 0) {
        html += generateCashFlowStatementTable(response);
    }

    result_container.innerHTML = html;
}

/**
 * 生成現金流表格 HTML
 */
function generateCashFlowTable(response) {
    const { cash_flows, years, equipment_cost } = response;

    let html = `
        <div class="result-display" style="margin-top: 20px;">
            <!-- 初始投資顯示區域 -->
            <div style="background: #f7fafc; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
                <div style="font-size: 1.1em; color: #2d3748; margin-bottom: 5px;">初始投資 (重資打包賣價)</div>
                <div style="font-size: 1.4em; font-weight: bold; color: #1a202c;">
                    NT$ ${Math.round(equipment_cost).toLocaleString('zh-TW')}
                </div>
            </div>

            <h3 style="margin-bottom: 15px; color: #2d3748;">綜合損益表</h3>
            <div class="table-container">
                <table class="cash-flow-table">
                    <thead>
                        <tr>
                            <th style="text-align: left;">項目\\年度</th>
    `;

    // 添加年度標題
    years.forEach(year => {
        html += `<th>${year}年</th>`;
    });

    html += `</tr></thead><tbody>`;

    // 各項目行

    const customLabel = response.custom_label || '自定義欄位';
    const customShow = typeof response.custom_show === 'boolean' ? response.custom_show : true;
    const items = [
        { key: 'income', label: '電費收入' },
        { key: 'equipment_depreciation', label: '設備折舊' },
        { key: 'interest', label: '利息費用' },
        { key: 'rent', label: '設備放置區域租金' },
        { key: 'maintenance', label: '運維費用' },
        { key: 'insurance', label: '保險費' },
        { key: 'recycling', label: '模組回收費' }
    ];
    if (customShow) {
        items.push({ key: 'custom', label: customLabel });
    }

    items.forEach(item => {
        html += `<tr><td class=\"row-header\">${item.label}</td>`;
        cash_flows.forEach(cf => {
            html += `<td>${Math.round(cf[item.key]||0).toLocaleString('zh-TW')}</td>`;
        });
        html += `</tr>`;
    });

    // 稅前淨利
    html += `<tr class="net-cash-flow-row"><td class="row-header">稅前淨利</td>`;
    cash_flows.forEach(cf => {
        const color = cf.net_cash_flow >= 0 ? '#22543d' : '#c53030';
        html += `<td style="color: ${color};">${Math.round(cf.net_cash_flow).toLocaleString('zh-TW')}</td>`;
    });
    html += `</tr>`;

    // 所得稅
    html += `<tr class="tax-row"><td class="row-header">所得稅</td>`;
    cash_flows.forEach(cf => {
        html += `<td style="color: #d69e2e;">${Math.round(cf.tax_amount).toLocaleString('zh-TW')}</td>`;
    });
    html += `</tr>`;

    // 稅後淨利
    html += `<tr class="after-tax-row"><td class="row-header">稅後淨利</td>`;
    cash_flows.forEach(cf => {
        const color = cf.after_tax_cash_flow >= 0 ? '#22543d' : '#c53030';
        html += `<td style="color: ${color};">${Math.round(cf.after_tax_cash_flow).toLocaleString('zh-TW')}</td>`;
    });
    html += `</tr>`;

    html += `</tbody></table></div></div>`;

    return html;
}

/**
 * 生成現金流量表 HTML
 */
function generateCashFlowStatementTable(response) {
    const { cash_flow_statement, irr_analysis, years } = response;

    // 為現金流量表創建擴展的年度列表（多一年用於IRR顯示，但最後一年不顯示標題）
    const extendedYears = [...years, years[years.length - 1] + 1];

    let html = `
        <div class="result-display" style="margin-top: 30px;">
            <h3 style="margin-bottom: 15px; color: #2d3748;">現金流量表</h3>
            <div class="table-container">
                <table class="cash-flow-table">
                    <thead>
                        <tr>
                            <th style="text-align: left; width: 200px;">項目\\年度</th>
    `;

    // 添加年度標題（最後一年不顯示標題，只顯示空格）
    years.forEach(year => {
        html += `<th>${year}年</th>`;
    });
    // 最後一年不顯示年度，只顯示空白標題
    html += `<th style="width: 80px;"></th>`;

    html += `</tr></thead><tbody>`;

    // 區塊1：營運活動
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">營運活動</td></tr>`;

    const operatingItems = [
        { key: 'aftertax_net_profit', label: '稅後淨利' },
        { key: 'equipment_depreciation', label: '設備折舊' },
        { key: 'operating_cash_flow', label: '營運活動現金流量', isTotal: true }
    ];

    operatingItems.forEach(item => {
        html += `<tr ${item.isTotal ? 'class="total-row"' : ''}><td class="row-header">${item.label}</td>`;
        cash_flow_statement.forEach(cf => {
            const color = item.isTotal ? (cf[item.key] >= 0 ? '#22543d' : '#c53030') : '';
            const style = item.isTotal ? `style="color: ${color}; font-weight: bold;"` : '';
            html += `<td ${style}>${Math.round(cf[item.key]).toLocaleString('zh-TW')}</td>`;
        });
        // 為最後一年添加空白格
        html += `<td>-</td>`;
        html += `</tr>`;
    });

    // 區塊2：投資活動
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">投資活動</td></tr>`;
    html += `<tr><td class="row-header">投資活動-設備支出</td>`;
    cash_flow_statement.forEach(cf => {
        const color = cf.equipment_expenditure < 0 ? '#c53030' : '#22543d';
        html += `<td style="color: ${color};">${Math.round(cf.equipment_expenditure).toLocaleString('zh-TW')}</td>`;
    });
    html += `<td>-</td></tr>`;

    // 區塊3：理財活動
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">理財活動</td></tr>`;

    const financingItems = [
        { key: 'loan_financing', label: '理財活動-借款(銀行貸款)' },
        { key: 'loan_repayment', label: '理財活動-還款' },
        { key: 'cash_capital_increase', label: '理財活動-現金增資' },
        { key: 'cash_dividend', label: '現金股利' },
        { key: 'capital_reduction', label: '年底減資' }
    ];

    financingItems.forEach(item => {
        html += `<tr><td class="row-header">${item.label}</td>`;
        cash_flow_statement.forEach(cf => {
            const value = cf[item.key];
            const color = value >= 0 ? '#22543d' : '#c53030';
            html += `<td style="color: ${color};">${Math.round(value).toLocaleString('zh-TW')}</td>`;
        });
        html += `<td>-</td></tr>`;
    });

    // 區塊4：現金流匯總
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">現金流匯總</td></tr>`;

    const summaryItems = [
        { key: 'net_cash_inflow', label: '淨現金流入(出)' },
        { key: 'opening_cash_flow', label: '期初淨現金流' },
        { key: 'closing_cash_flow', label: '期末淨現金流' }
    ];

    summaryItems.forEach(item => {
        html += `<tr class="total-row"><td class="row-header">${item.label}</td>`;
        cash_flow_statement.forEach(cf => {
            const color = cf[item.key] >= 0 ? '#22543d' : '#c53030';
            html += `<td style="color: ${color}; font-weight: bold;">${Math.round(cf[item.key]).toLocaleString('zh-TW')}</td>`;
        });
        html += `<td>-</td></tr>`;
    });

    // 區塊5：IRR分析 (顯示實際IRR現金流數據)
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">IRR分析</td></tr>`;

    // 成本法現金流（包含IRR計算用的所有數據，比權益法多一年）
    html += `<tr><td class="row-header">成本法實際現金流</td>`;
    irr_analysis.cost_method_cash_flows.forEach(cashFlow => {
        const color = cashFlow >= 0 ? '#22543d' : '#c53030';
        html += `<td style="color: ${color};">${Math.round(cashFlow).toLocaleString('zh-TW')}</td>`;
    });
    html += `</tr>`;

    // 成本法IRR
    html += `<tr><td class="row-header">成本法IRR</td>`;
    html += `<td colspan="${extendedYears.length}" style="text-align: left; font-weight: bold; color: #1a202c; padding-left: 10px;">`;
    if (irr_analysis.cost_method_irr !== null) {
        html += `${irr_analysis.cost_method_irr.toFixed(2)}%`;
    } else {
        html += '無法計算';
    }
    html += `</td></tr>`;

    // 權益法現金流（包含IRR計算用的所有數據，比成本法少一年）
    html += `<tr><td class="row-header">權益法實際現金流</td>`;
    irr_analysis.equity_method_cash_flows.forEach(cashFlow => {
        const color = cashFlow >= 0 ? '#22543d' : '#c53030';
        html += `<td style="color: ${color};">${Math.round(cashFlow).toLocaleString('zh-TW')}</td>`;
    });
    // 權益法少一年，補空白格
    html += `<td>-</td></tr>`;

    // 權益法IRR
    html += `<tr><td class="row-header">權益法IRR</td>`;
    html += `<td colspan="${extendedYears.length}" style="text-align: left; font-weight: bold; color: #1a202c; padding-left: 10px;">`;
    if (irr_analysis.equity_method_irr !== null) {
        html += `${irr_analysis.equity_method_irr.toFixed(2)}%`;
    } else {
        html += '無法計算';
    }
    html += `</td></tr>`;

    // 區塊6：借款狀況
    html += `<tr class="section-header"><td colspan="${extendedYears.length + 1}" style="background: #e2e8f0; font-weight: bold; text-align: center;">借款狀況</td></tr>`;
    html += `<tr><td class="row-header">借款餘額</td>`;
    cash_flow_statement.forEach(cf => {
        html += `<td style="color: #d69e2e;">${Math.round(cf.loan_balance).toLocaleString('zh-TW')}</td>`;
    });
    html += `<td>-</td></tr>`;

    html += `</tbody></table></div></div>`;

    return html;
}