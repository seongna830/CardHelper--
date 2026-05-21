from flask import Flask, render_template, request, url_for
import pandas as pd
from thefuzz import process  # 引入模糊搜尋功能
import os

app = Flask(__name__)

def get_card_benefits(store_query):
    try:
        # === 1. 雙重相容讀檔機制 (完全修復縮進與解碼問題) ===
        excel_path = 'card.xlsx'
        df = pd.DataFrame() # 建立一個空的備用表格
        
        # 優先嘗試讀取 Excel 檔案
        if os.path.exists(excel_path):
            try:
                df = pd.read_excel(excel_path)
            except Exception:
                df = pd.DataFrame()
            
        # 如果 Excel 讀出來是空的，或者檔案不存在，自動嘗試讀取 CSV 檔案
        if df.empty:
            # 定義所有可能出現的檔案名稱
            possible_files = [excel_path, 'card.xlsx - 工作表1.csv']
            # 定義台灣最常使用的編碼清單
            encodings = ['utf-8', 'big5', 'utf-8-sig', 'gbk']
            
            for file_name in possible_files:
                if df.empty and os.path.exists(file_name):
                    for enc in encodings:
                        try:
                            df = pd.read_csv(file_name, encoding=enc)
                            if not df.empty:
                                break # 成功讀取到資料，跳出編碼迴圈
                        except Exception:
                            continue # 這個編碼失敗，嘗試下一種
                            
        # 保障機制：如果經歷了上面的層層讀取，最後表格依然是空的，防呆返回
        if df.empty:
            print("【系統警告】無法成功解碼並讀取任何資料庫檔案！請檢查 card.xlsx 是否存在。")
            return []

        # 清洗全表所有空值，一網打盡所有 NaN 亂碼問題
        df = df.fillna("")  

        # 將 DataFrame 轉成純粹的 Python 字典清單 (沿用 logic.py 經典結構)
        data_list = df.to_dict('records')
        
        # 處理使用者輸入：去除空白、轉小寫
        query = str(store_query).strip().lower()
        if not query:
            return []

        # === 2. 階段一：精準/包含比對 ===
        exact_matches = []
        for row in data_list:
            store_name = str(row.get('店家', '')).strip().lower()
            plan_name = str(row.get('方案', '')).strip().lower()
            tag_name = str(row.get('搜尋標籤', '')).strip().lower()
            
            # 只要 店家、方案 或 搜尋標籤 包含使用者輸入的字
            if query in store_name or query in plan_name or query in tag_name:
                if store_name == query:
                    row['match_score'] = 100  
                elif query in store_name:
                    row['match_score'] = 90   
                elif query in tag_name:
                    row['match_score'] = 85   
                else:
                    row['match_score'] = 80
                    
                exact_matches.append(row)

        # === 3. 階段二：模糊比對 ===
        if not exact_matches:
            search_pool = []
            for row in data_list:
                if row.get('店家'): search_pool.append(str(row['店家']).strip())
                if row.get('方案'): search_pool.append(str(row['方案']).strip())
                if row.get('搜尋標籤'):
                    tags = str(row['搜尋標籤']).replace(',', ' ').replace('，', ' ').replace('/', ' ').split()
                    search_pool.extend(tags)
            search_pool = list(set(search_pool))
            
            matches = process.extractBests(query, search_pool, score_cutoff=65, limit=10)
            
            if matches:
                matched_keywords = [match[0].lower() for match in matches]
                score_dict = {match[0].lower(): match[1] for match in matches}
                
                for row in data_list:
                    store_name = str(row.get('店家', '')).strip().lower()
                    plan_name = str(row.get('方案', '')).strip().lower()
                    tag_name = str(row.get('搜尋標籤', '')).strip().lower()
                    
                    hit_store = store_name in matched_keywords
                    hit_plan = plan_name in matched_keywords
                    hit_tag = any(kw in tag_name for kw in matched_keywords)
                    
                    if hit_store or hit_plan or hit_tag:
                        s_score = score_dict.get(store_name, 0)
                        p_score = score_dict.get(plan_name, 0)
                        t_score = max([score_dict.get(kw, 0) for kw in score_dict if kw in tag_name] + [0])
                        row['match_score'] = max(s_score, p_score, t_score)
                        exact_matches.append(row)

        # === 4. 雙重條件排序 ===
        for row in exact_matches:
            try:
                row['回饋率_num'] = float(row.get('回饋率', 0.0))
            except:
                row['回饋率_num'] = 0.0

        exact_matches.sort(key=lambda x: (x.get('match_score', 0), x.get('回饋率_num', 0.0)), reverse=True)

        # === 5. 資料轉包與格式化 ===
        results = []
        for row in exact_matches:
            if str(row.get('店家')) == "非特定消費基本回饋" and query != "基本回饋":
                if len(exact_matches) > 1:
                    continue
                    
            try:
                rate_percent = float(row.get('回饋率', 0.0)) * 100
                if rate_percent.is_integer():
                    rate_display = f"{int(rate_percent)}%"
                else:
                    rate_display = f"{rate_percent:.1f}%"
            except:
                rate_display = "0%"
                
            # 安全尋找網址欄位：遍歷所有欄位名稱，只要帶有「連結」或「url」就把網址提取出來
            card_url = "#"
            for k, v in row.items():
                if "連結" in str(k) or "url" in str(k).lower():
                    if str(v).strip() and str(v).strip().lower() != 'nan':
                        card_url = str(v).strip()
                        break
            
            # 備用保險：如果依照名字都沒找到，預設直接抓整行的最後一欄資料
            if card_url == "#" and len(row) > 0:
                card_url = str(list(row.values())[-1]).strip()

            if card_url.lower() == 'nan' or not card_url:
                card_url = '#'

            results.append({
                'store': row.get('店家', ''),
                'card': row.get('信用卡', ''),
                'plan': row.get('方案', ''),
                'rate': rate_display,
                'url': card_url  
            })

        # === 6. 兜底保障機制 ===
        if not results:
            fallback_cards = [row for row in data_list if str(row.get('店家', '')) == "非特定消費基本回饋"]
            fallback_cards.sort(key=lambda x: float(x.get('回饋率', 0.0)), reverse=True)
            for row in fallback_cards:
                card_url = "#"
                for k, v in row.items():
                    if "連結" in str(k) or "url" in str(k).lower():
                        if str(v).strip() and str(v).strip().lower() != 'nan':
                            card_url = str(v).strip()
                            break
                if card_url == "#" and len(row) > 0:
                    card_url = str(list(row.values())[-1]).strip()

                if card_url.lower() == 'nan' or not card_url:
                    card_url = '#'

                results.append({
                    'store': row.get('店家', ''),
                    'card': row.get('信用卡', ''),
                    'plan': row.get('方案', ''),
                    # 使用 round 四捨五入到小數點後 2 位，再去掉末尾沒用的 0
                    'rate': f"{round(float(row.get('回饋率', 0.0)) * 100, 2)}%",
                    'url': card_url
                })
                
        return results
        
    except Exception as e:
        print(f"程式執行發生錯誤: {e}")
        return []

# ==============================================================================
# Flask 網頁路由控制區
# ==============================================================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    user_input = request.form.get('store_name', '')
    search_results = get_card_benefits(user_input)
    return render_template('result.html', results=search_results, query=user_input)

if __name__ == '__main__':
    app.run(debug=True)