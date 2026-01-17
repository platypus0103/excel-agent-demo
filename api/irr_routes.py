"""
IRR 計算相關的 API 路由
"""
from flask import Blueprint, request, jsonify
from pydantic import ValidationError
from models.irr_models_v2 import IRRCalculationRequest, EquipmentCostParams
from services.irr_calculator import IRRCalculatorService


# 創建藍圖
irr_bp = Blueprint('irr', __name__, url_prefix='/api/irr')


@irr_bp.route('/calculate', methods=['POST'])
def calculate_irr():
    """
    計算 IRR 的主要 API 端點
    """
    try:
        # 獲取請求數據
        data = request.get_json()

        if not data:
            return jsonify({
                'success': False,
                'error': '請求數據為空'
            }), 400

        # 驗證請求數據
        try:
            irr_request = IRRCalculationRequest(**data)
        except ValidationError as e:
            return jsonify({
                'success': False,
                'error': f'數據驗證失敗: {str(e)}'
            }), 400

        # 執行 IRR 計算
        result = IRRCalculatorService.calculate_irr_full(irr_request)

        # 轉換為字典格式回傳，並加上 custom_label、顯示設定
        result_dict = result.dict()
        result_dict['custom_label'] = data.get('custom_label', '自定義欄位')
        result_dict['custom_show'] = data.get('custom_show', True)
        result_dict['show_cols'] = data.get('show_cols', {})
        return jsonify(result_dict)

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'伺服器錯誤: {str(e)}'
        }), 500


@irr_bp.route('/equipment-cost', methods=['POST'])
def calculate_equipment_cost():
    """
    單獨計算設備費用的 API 端點
    """
    try:
        data = request.get_json()

        if not data:
            return jsonify({
                'success': False,
                'error': '請求數據為空'
            }), 400

        # 驗證設備參數
        try:
            equipment_params = EquipmentCostParams(**data)
        except ValidationError as e:
            return jsonify({
                'success': False,
                'error': f'設備參數驗證失敗: {str(e)}'
            }), 400

        # 計算設備費用
        equipment_cost = IRRCalculatorService.calculate_equipment_cost(equipment_params)

        return jsonify({
            'success': True,
            'equipment_cost': equipment_cost,
            'formatted_cost': f"NT$ {equipment_cost:,.0f}"
        })

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'計算設備費用時發生錯誤: {str(e)}'
        }), 500


@irr_bp.route('/health', methods=['GET'])
def health_check():
    """
    健康檢查端點
    """
    return jsonify({
        'status': 'healthy',
        'service': 'IRR Calculator API',
        'version': '1.0.0'
    })


# 錯誤處理器
@irr_bp.errorhandler(404)
def not_found(error):
    return jsonify({
        'success': False,
        'error': 'API 端點不存在'
    }), 404


@irr_bp.errorhandler(405)
def method_not_allowed(error):
    return jsonify({
        'success': False,
        'error': 'HTTP 方法不被允許'
    }), 405