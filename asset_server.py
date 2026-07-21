import sys
import os

def check_requirements():
    import subprocess
    required = ["flask", "flask_cors", "yfinance", "pandas", "xlsxwriter", "requests", "matplotlib", "Pillow"]
    missing = []
    for lib in required:
        try:
            if lib == "Pillow":
                __import__("PIL.Image")
            else:
                __import__(lib.replace('_', '-'))
        except ImportError:
            missing.append(lib.replace('_', '-'))

    if missing:
        print(f"[*] 필수 라이브러리 설치 중... {missing}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
            print("[*] 설치 완료!")
        except Exception as e:
            print(f"[!] 설치 실패. 수동으로 명령어 실행 요망: pip install {' '.join(missing)}")
            sys.exit(1)

check_requirements()

import json
from flask import Flask, request, jsonify
from flask_cors import CORS
import yfinance as yf
import pandas as pd
import threading
import concurrent.futures
import webbrowser
import xlsxwriter
import datetime
import os
import io
import tempfile
import random
import requests
from PIL import Image, ImageDraw, ImageFont
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg

# 정적 파일 경로(index.html이 있는 폴더)를 현재 디렉토리로 설정
app = Flask(__name__, static_folder=os.path.abspath(os.path.dirname(__file__)))
CORS(app)

@app.route('/')
def index():
    # 현재 폴더에 있는 index.html을 서빙합니다.
    index_path = os.path.join(app.static_folder, 'index.html')
    if os.path.exists(index_path):
        with open(index_path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        return "Error: index.html 파일이 현재 폴더에 없습니다.", 404

@app.route('/api/batch', methods=['GET'])
def get_batch_prices():
    tickers_param = request.args.get('tickers', '')
    if not tickers_param:
        return jsonify({"error": "No tickers provided"}), 400
        
    tickers = [t.strip().upper() for t in tickers_param.split(',') if t.strip()]
    if not tickers:
        return jsonify({"error": "Invalid tickers format"}), 400
        
    print(f"\n[*] 부스터 가동: {len(tickers)}개 종목 시세/지표 즉시 조회 중...")
    
    try:
        data = yf.download(tickers, period="5d", progress=False)
        results = {}
        
        # 기본 가격 정보 파싱
        if len(tickers) == 1:
            ticker = tickers[0]
            if not data.empty and 'Close' in data.columns:
                close_prices = data['Close'].dropna()
                if len(close_prices) >= 1:
                    results[ticker] = {
                        "price": float(close_prices.iloc[-1]),
                        "prevClose": float(close_prices.iloc[-2]) if len(close_prices) >= 2 else float(close_prices.iloc[-1])
                    }
        else:
            if not data.empty and 'Close' in data.columns:
                for ticker in tickers:
                    if ticker in data['Close'].columns:
                        close_prices = data['Close'][ticker].dropna()
                        if len(close_prices) >= 1:
                            results[ticker] = {
                                "price": float(close_prices.iloc[-1]),
                                "prevClose": float(close_prices.iloc[-2]) if len(close_prices) >= 2 else float(close_prices.iloc[-1])
                            }

        # 펀더멘털 지표 병렬 수집
        def fetch_info(t):
            try:
                info = yf.Ticker(t).info
                return t, {
                    "marketCap": info.get('marketCap', 0),
                    "trailingPE": info.get('trailingPE', 0),
                    "priceToBook": info.get('priceToBook', 0),
                    "trailingEps": info.get('trailingEps', 0)
                }
            except:
                return t, {}

        with concurrent.futures.ThreadPoolExecutor(max_workers=min(20, len(tickers))) as executor:
            future_to_ticker = {executor.submit(fetch_info, t): t for t in results.keys()}
            for future in concurrent.futures.as_completed(future_to_ticker):
                t, info_data = future.result()
                results[t].update(info_data)
                            
        print(f"[*] 부스터 완료: {len(results)}개 종목 성공")
        return jsonify({"status": "success", "data": results})
        
    except Exception as e:
        print(f"[!] 에러 발생: {e}")
        return jsonify({"error": str(e)}), 500

def generate_chart_img(data, is_up):
    try:
        fig = plt.Figure(figsize=(2, 0.5), dpi=80, facecolor='white')
        ax = fig.add_subplot(111)
        c = '#10B981' if is_up else '#EF4444'
        ax.plot(data, color=c, lw=1.5)
        ax.fill_between(range(len(data)), data, min(data), color=c, alpha=0.1)
        ax.axis('off')
        fig.tight_layout(pad=0)
        
        buf = io.BytesIO()
        FigureCanvasAgg(fig).print_png(buf)
        buf.seek(0)
        return buf
    except:
        return None

def create_text_logo(ticker, path):
    colors = [(37, 99, 235), (220, 38, 38), (5, 150, 105), (124, 58, 237), (219, 39, 119)]
    img = Image.new('RGB', (60, 60), color=random.choice(colors))
    d = ImageDraw.Draw(img)
    try: font = ImageFont.truetype("arialbd.ttf", 16)
    except: font = ImageFont.load_default()
    
    text = ticker[:4].upper()
    bbox = d.textbbox((0,0), text, font=font)
    w, h = bbox[2]-bbox[0], bbox[3]-bbox[1]
    d.text(((60-w)/2, (60-h)/2), text, fill='white', font=font)
    img.save(path, "PNG")
    return path

def get_logo_forced(ticker, logo_dir):
    path = os.path.join(logo_dir, f"{ticker}.png")
    if os.path.exists(path): return path

    clean = ticker.replace('-', '').replace('.', '').lower()
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = [
        f"https://logo.clearbit.com/{clean}.com",
        f"https://logos.stockanalysis.com/{clean}.svg",
        f"https://eodhistoricaldata.com/img/logos/US/{ticker.upper()}.png"
    ]
    for u in urls:
        try:
            r = requests.get(u, headers=headers, timeout=1.5)
            if r.status_code == 200 and len(r.content) > 200:
                img = Image.open(io.BytesIO(r.content))
                if img.mode != 'RGB': img = img.convert('RGB')
                img.thumbnail((60, 60))
                img.save(path, "PNG")
                return path
        except: continue
    return create_text_logo(ticker, path)

@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        data = request.json
        if not data or 'portfolios' not in data:
            return jsonify({"error": "Invalid data format"}), 400

        # 티커 수집 및 1년치 데이터 다운로드
        all_tickers = set()
        for pdata in data['portfolios'].values():
            for h in pdata.get('holdings', []):
                t = h.get('yahoo') or h.get('ticker')
                if t: all_tickers.add(t)
        
        print(f"[*] 엑셀 추출: {len(all_tickers)}개 종목 1년치 차트 데이터 수집 중...")
        hist_data = None
        if all_tickers:
            hist_data = yf.download(list(all_tickers), period="1y", progress=False)

        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        if not os.path.exists(desktop_path):
            desktop_path = os.path.join(os.path.expanduser('~'), 'OneDrive', '바탕 화면')
            if not os.path.exists(desktop_path):
                desktop_path = os.path.dirname(os.path.abspath(__file__))
                
        # 로고 영구 캐시 폴더 설정 (바탕화면 숨김 폴더 또는 프로젝트 내부)
        logo_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo_cache')
        if not os.path.exists(logo_dir):
            os.makedirs(logo_dir)
                
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = os.path.join(desktop_path, f"LProject_AssetReport_{timestamp}.xlsx")

        wb = xlsxwriter.Workbook(fname, {'nan_inf_to_errors': True})
        
        f_hd = wb.add_format({'bold': 1, 'bg_color': '#1E293B', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'bottom': 2})
        f_c = wb.add_format({'align': 'center', 'valign': 'vcenter'})
        f_num = wb.add_format({'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
        f_num_dec = wb.add_format({'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter'})
        f_g = wb.add_format({'bg_color': '#DCFCE7', 'font_color': '#166534', 'num_format': '+0.0%', 'align': 'center', 'valign': 'vcenter'})
        f_r = wb.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'num_format': '0.0%', 'align': 'center', 'valign': 'vcenter'})

        for pid, pdata in data['portfolios'].items():
            name_map = {'owner': '본인', 'wife': '와이프', 'son': '자녀'}
            sheet_name = name_map.get(pid, pid)
            
            ws = wb.add_worksheet(f"투자현황_{sheet_name}")
            ws.set_zoom(90)
            ws.hide_gridlines(2)
            
            ws.write(0, 0, f"📊 {sheet_name} - 주식/ETF 자산", wb.add_format({'bold': 1, 'font_size': 14}))
            
            headers = ["로고", "티커", "종목명", "보유수량", "매수평단가", "현재가", "총 매수금액", "현재 평가금액", "수익률", "시가총액", "PER", "PBR", "EPS", "1년 흐름"]
            ws.write_row(2, 0, headers, f_hd)
            
            ws.set_column('A:A', 8)
            ws.set_column('B:B', 12)
            ws.set_column('C:C', 25)
            ws.set_column('D:F', 12)
            ws.set_column('G:I', 14)
            ws.set_column('J:J', 16) # Market Cap
            ws.set_column('K:M', 10) # PER, PBR, EPS
            ws.set_column('N:N', 27) # Chart
            
            row = 3
            for h in pdata.get('holdings', []):
                ws.set_row(row, 45)
                
                def _n(k):
                    v = h.get(k)
                    return float(v) if v is not None else 0.0

                ticker = h.get('ticker', '')
                yahoo_ticker = h.get('yahoo') or ticker
                
                # 1. 로고 삽입
                if yahoo_ticker:
                    logo_path = get_logo_forced(yahoo_ticker, logo_dir)
                    if logo_path and os.path.exists(logo_path):
                        ws.insert_image(row, 0, logo_path, {'x_scale': 0.7, 'y_scale': 0.7, 'x_offset': 5, 'y_offset': 5})
                
                ws.write(row, 1, ticker, f_c)
                ws.write(row, 2, h.get('name', ''), f_c)
                ws.write(row, 3, _n('shares'), f_num_dec)
                ws.write(row, 4, _n('avgCost'), f_num_dec)
                ws.write(row, 5, _n('currentPrice'), f_num_dec)
                
                cost = _n('shares') * _n('avgCost')
                val = _n('shares') * _n('currentPrice')
                ws.write(row, 6, cost, f_num_dec)
                ws.write(row, 7, val, f_num_dec)
                
                rate = (val / cost - 1) if cost > 0 else 0
                ws.write(row, 8, rate, f_g if rate >= 0 else f_r)

                def large_fmt(val, curr):
                    if not val: return "-"
                    v = float(val)
                    if curr == "KRW":
                        if v >= 1e12: return f"{v/1e12:.1f}조"
                        if v >= 1e8: return f"{v/1e8:.1f}억"
                        return f"{v/1e4:.0f}만"
                    else:
                        if v >= 1e9: return f"${v/1e9:.1f}B"
                        if v >= 1e6: return f"${v/1e6:.1f}M"
                        return f"${v:.0f}"

                curr = h.get('currency', 'KRW')
                ws.write(row, 9, large_fmt(h.get('marketCap'), curr), f_c)
                ws.write(row, 10, h.get('trailingPE', '-'), f_c)
                ws.write(row, 11, h.get('priceToBook', '-'), f_c)
                ws.write(row, 12, h.get('trailingEps', '-'), f_c)
                
                # 2. 1년 차트 삽입
                if yahoo_ticker and hist_data is not None and not hist_data.empty and 'Close' in hist_data.columns:
                    if len(all_tickers) == 1:
                        close_prices = hist_data['Close'].dropna()
                    else:
                        if yahoo_ticker in hist_data['Close'].columns:
                            close_prices = hist_data['Close'][yahoo_ticker].dropna()
                        else:
                            close_prices = None
                            
                    if close_prices is not None and not close_prices.empty:
                        c_vals = close_prices.squeeze().tolist()
                        if len(c_vals) > 100: c_vals = c_vals[::len(c_vals)//100+1]
                        
                        is_up = c_vals[-1] >= c_vals[0]
                        chart_buf = generate_chart_img(c_vals, is_up)
                        if chart_buf:
                            ws.insert_image(row, 13, "chart.png", {
                                'image_data': chart_buf,
                                'x_scale': 0.95, 'y_scale': 0.95,
                                'x_offset': 15, 'y_offset': 3
                            })
                
                row += 1
                
        wb.close()
        os.startfile(fname)
        
        print(f"[*] 엑셀 추출 완료: {fname}")
        return jsonify({"status": "success", "filepath": fname})
        
    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        try:
            with open(os.path.join(desktop_path, 'LProject_Error_Log.txt'), 'w', encoding='utf-8') as f:
                f.write(err_msg)
        except: pass
        print(f"[!] 엑셀 추출 에러: {e}")
        return jsonify({"error": str(e)}), 500

def open_browser():
    print("\n" + "="*50)
    print("🚀 [L-Project 부스터 + 로컬 저장소]가 켜졌습니다!")
    print("웹 브라우저가 자동으로 열립니다.")
    print("이제 이 주소(http://127.0.0.1:5000)를 사용하시면 데이터가 영구 저장되며,")
    print("시세 조회도 0.1초 만에 완료됩니다.")
    print("="*50 + "\n")
    webbrowser.open('http://127.0.0.1:5000/')

if __name__ == '__main__':
    print("[*] 파이썬 서버가 성공적으로 시작되었습니다!")
    # 서버 실행 0.5초 뒤 자동으로 브라우저 팝업
    threading.Timer(0.5, open_browser).start()
    app.run(host='127.0.0.1', port=5000, debug=False)
