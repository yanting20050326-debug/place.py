"""
選址規劃與 AI 敏感度分析系統 (Location Planning System)
這是一個供教學使用的 Flask 網頁應用程式。
提供學生進行工廠選址的加權評分計算，並自動產生 AI 敏感度分析提問，最後紀錄學生的決策思維。
"""

from flask import Flask, request, jsonify, render_template
import csv
import os
import datetime

app = Flask(__name__)

# 常數設定：台灣本島縣市 (排除離島)
TAIWAN_CITIES = [
    "基隆市", "台北市", "新北市", "桃園市", "新竹縣", "新竹市", "苗栗縣",
    "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣"
]

DATA_FILE = 'location_answers.csv'

def calculate_location_scores(weights: dict, ratings: dict) -> tuple:
    """
    計算各地點的加權總分與各因素貢獻度。
    """
    results = {}
    contributions = {}
    best_option = ""
    best_value = -float('inf')

    for loc, factors_rating in ratings.items():
        total_score = 0
        loc_contrib = {}
        for factor, weight in weights.items():
            # 加權總分 = Σ（權重 × 評分）
            score = float(weight) * float(factors_rating.get(factor, 0))
            total_score += score
            loc_contrib[factor] = round(score, 2)
            
        results[loc] = round(total_score, 2)
        contributions[loc] = loc_contrib
        
        if total_score > best_value:
            best_value = total_score
            best_option = loc

    return results, contributions, best_option, round(best_value, 2)

def generate_ai_analysis(best_option: str, weights: dict, contributions: dict) -> dict:
    """
    根據最高分地點與其最高貢獻度因素，產生 AI 輔助分析與敏感度問題。
    """
    if not best_option:
        return {}

    best_loc_contrib = contributions[best_option]
    # 找出對最佳地點貢獻度最高的因素
    top_factor = max(best_loc_contrib, key=best_loc_contrib.get)
    
    analysis = {
        "explanation": f"系統建議的最佳建廠地點為 **{best_option}**。",
        "impact_factors": f"從貢獻度分析來看，**{top_factor}** 是讓 {best_option} 脫穎而出的關鍵因素（加權得分最高）。",
        "sensitivity_prompt": f"敏感度分析提問：「如果企業策略轉向，將『{top_factor}』的權重降低，並提高其他弱勢因素的權重，{best_option} 依然會是最佳選擇嗎？企業的策略（如成本導向或市場導向）會如何影響最終決策？」"
    }
    return analysis

@app.route('/')
def index():
    """渲染系統首頁與預設選項"""
    default_factors = ["土地成本", "勞動力供給", "運輸便利性", "市場接近性", "基礎建設", "政策與法規環境"]
    return render_template('LocationPlanning.html', cities=TAIWAN_CITIES, default_factors=default_factors)

@app.route('/api/calculate', methods=['POST'])
def calculate_api():
    """處理加權評分計算 API"""
    data = request.json
    weights = data.get('weights', {})
    ratings = data.get('ratings', {})
    
    # 權重防呆機制：確保總和為 1.0 (容許極小浮點數誤差)
    try:
        total_weight = sum(float(w) for w in weights.values())
        if abs(total_weight - 1.0) > 0.01:
            return jsonify({
                "status": "error", 
                "message": f"權重總和必須為 1.0！目前總和為 {total_weight:.2f}，請重新調整。"
            }), 400
    except ValueError:
        return jsonify({
            "status": "error", 
            "message": "權重格式錯誤，請確認輸入的都是數字！"
        }), 400
    
    results, contributions, best_option, best_value = calculate_location_scores(weights, ratings)
    ai_analysis = generate_ai_analysis(best_option, weights, contributions)
    
    return jsonify({
        "status": "success",
        "calculated_scores": results,
        "contributions": contributions,
        "recommended_option": best_option,
        "best_value": best_value,
        "ai_analysis": ai_analysis
    })

@app.route('/api/submit', methods=['POST'])
def submit_answer():
    """儲存學生的決策紀錄至 CSV"""
    data = request.json
    file_exists = os.path.isfile(DATA_FILE)
    
    with open(DATA_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['交卷時間', '班級', '學號', '姓名', '選定地點數量', '最佳推薦地點', '系統提問(敏感度分析)', '學生回答(決策思維)'])
        
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([
            timestamp, 
            data.get('studentClass', ''), 
            data.get('studentId', ''), 
            data.get('studentName', ''), 
            data.get('locationCount', 0),
            data.get('bestLocation', ''),
            data.get('aiQuestion', ''),
            data.get('q2Answer', '')
        ])
        
    return jsonify({"status": "success", "message": "分析結果與決策已成功儲存！"})

@app.route('/admin')
def admin_view():
    """老師專用後台：檢視所有學生的作答紀錄"""
    if not os.path.isfile(DATA_FILE): 
        return "<h2>還沒有學生提交喔！</h2>"
    
    html = """
    <html>
    <head>
        <meta charset='utf-8'>
        <title>老師專用後台</title>
        <style>
            body {font-family: Arial, sans-serif; margin: 20px;}
            table {width: 100%; border-collapse: collapse; margin-top: 20px;}
            th, td {border: 1px solid #ddd; padding: 10px; text-align: left;}
            th {background: #007bff; color: white;}
            tr:nth-child(even) {background-color: #f2f2f2;}
        </style>
    </head>
    <body>
        <h2>👨‍🏫 選址規劃系統 - 老師專用後台</h2>
        <table>
    """
    
    with open(DATA_FILE, mode='r', encoding='utf-8-sig') as f:
        for i, row in enumerate(csv.reader(f)):
            html += "<tr>"
            for col in row:
                html += f"<th>{col}</th>" if i == 0 else f"<td>{col}</td>"
            html += "</tr>"
            
    html += "</table></body></html>"
    return html

if __name__ == '__main__':
    # 建議將 debug 模式留在環境變數控制，這裡預設開啟方便開發
    app.run(debug=True, port=5001)
