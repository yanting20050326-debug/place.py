"""
選址規劃與 AI 敏感度分析系統 (Google Sheets 雲端版)
功能：將學生選址數據與決策思維直接同步至 Google 試算表。
"""

from flask import Flask, request, jsonify, render_template
import datetime
import gspread

app = Flask(__name__)

# 常數設定：台灣本島縣市
TAIWAN_CITIES = [
    "基隆市", "台北市", "新北市", "桃園市", "新竹縣", "新竹市", "苗栗縣",
    "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣"
]

# ----------------- Google Sheets 初始化 -----------------
try:
    # 讀取 credentials.json 並連線 (Render 上會從 Secret Files 讀取)
    gc = gspread.service_account(filename='credentials.json')
    # 請確保試算表名稱與你在雲端硬碟建立的一致
    sh = gc.open('LocationPlanningDB').sheet1
except Exception as e:
    print(f"⚠️ Google Sheets 連線失敗：{e}")
    sh = None
# --------------------------------------------------------

def calculate_location_scores(weights: dict, ratings: dict) -> tuple:
    """計算各地點的加權總分與各因素貢獻度"""
    results = {}
    contributions = {}
    best_option = ""
    best_value = -float('inf')

    for loc, factors_rating in ratings.items():
        total_score = 0
        loc_contrib = {}
        for factor, weight in weights.items():
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
    """產生 AI 輔助分析與敏感度問題"""
    if not best_option: return {}
    best_loc_contrib = contributions[best_option]
    top_factor = max(best_loc_contrib, key=best_loc_contrib.get)
    
    return {
        "explanation": f"系統建議的最佳建廠地點為 **{best_option}**。",
        "impact_factors": f"從貢獻度分析來看，**{top_factor}** 是讓 {best_option} 脫穎而出的關鍵因素。",
        "sensitivity_prompt": f"敏感度分析提問：「如果企業策略轉向，將『{top_factor}』的權重降低，並提高其他弱勢因素的權重，{best_option} 依然會是最佳選擇嗎？」"
    }

@app.route('/')
def index():
    default_factors = ["土地成本", "勞動力供給", "運輸便利性", "市場接近性", "基礎建設", "政策與法規環境"]
    return render_template('LocationPlanning.html', cities=TAIWAN_CITIES, default_factors=default_factors)

@app.route('/api/calculate', methods=['POST'])
def calculate_api():
    data = request.json
    weights = data.get('weights', {})
    ratings = data.get('ratings', {})
    
    try:
        total_weight = sum(float(w) for w in weights.values())
        if abs(total_weight - 1.0) > 0.01:
            return jsonify({"status": "error", "message": f"權重總和必須為 1.0！目前總和為 {total_weight:.2f}"}), 400
    except ValueError:
        return jsonify({"status": "error", "message": "權重格式錯誤，請確認輸入的都是數字！"}), 400
    
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
    """將資料寫入 Google Sheets，徹底取代 CSV"""
    if sh is None:
        return jsonify({"status": "error", "message": "雲端資料庫連線失敗，請檢查 API 設定。"}), 500

    data = request.json
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    row_data = [
        timestamp, 
        data.get('studentClass', ''), 
        data.get('studentId', ''), 
        data.get('studentName', ''), 
        data.get('locationCount', 0),
        data.get('bestLocation', ''),
        data.get('aiQuestion', ''),
        data.get('q2Answer', '')
    ]
    
    try:
        sh.append_row(row_data)
        return jsonify({"status": "success", "message": "分析結果已成功同步至 Google 雲端試算表！"})
    except Exception as e:
        return jsonify({"status": "error", "message": f"寫入雲端失敗：{e}"}), 500

@app.route('/admin')
def admin_view():
    """從 Google Sheets 抓取即時數據顯示"""
    if sh is None: return "<h2>雲端試算表尚未連線。</h2>"
    
    try:
        records = sh.get_all_values()
        if not records or len(records) <= 1: 
            return "<h2>目前雲端試算表還沒有學生提交紀錄。</h2>"
        
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
            <h2>👨‍🏫 選址規劃系統 - 雲端即時後台</h2>
            <table>
        """
        for i, row in enumerate(records):
            html += "<tr>"
            for col in row:
                html += f"<th>{col}</th>" if i == 0 else f"<td>{col}</td>"
            html += "</tr>"
        return html + "</table></body></html>"
    except Exception as e:
        return f"<h2>讀取雲端資料時發生錯誤：{e}</h2>"

if __name__ == '__main__':
    app.run(debug=True, port=5001)
