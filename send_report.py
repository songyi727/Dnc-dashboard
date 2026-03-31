"""
DNC 매출 현황 대시보드 — 자동 메일 발송 스크립트
=====================================================
GitHub Actions에서 자동 실행됩니다.
report_data.json 파일을 읽어서 메일을 발송합니다.
(엑셀 파일 불필요 — update_dashboard.py 실행 시 자동 생성)
"""
 
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
 
# ============================================================
# 설정 — GitHub Secrets에서 자동으로 읽어옴
# ============================================================
GMAIL_USER     = os.environ.get('GMAIL_USER', '')
GMAIL_PASSWORD = os.environ.get('GMAIL_PASSWORD', '')
RECV_EMAIL     = os.environ.get('RECV_EMAIL', '')
DASHBOARD_URL  = os.environ.get('DASHBOARD_URL', 'https://songyi727.github.io/Dnc-dashboard/')
DATA_FILE      = 'report_data.json'
# ============================================================
 
def fs(v):
    if v >= 1e8: return f"{v/1e8:.1f}억원"
    if v >= 1e4: return f"{v/1e4:,.0f}만원"
    return f"{v:,.0f}원"
 
def chg_color(v):
    if v is None: return '#888888'
    return '#1D9E75' if v >= 0 else '#E24B4A'
 
def chg_arrow(v):
    if v is None: return '-'
    return f"{'▲' if v >= 0 else '▼'}{abs(v):.1f}%"
 
def rate_color(r):
    if r is None: return '#888888'
    if r >= 100: return '#1D9E75'
    if r >= 90:  return '#e8a838'
    return '#E24B4A'
 
def build_html(d):
    today = datetime.now().strftime('%Y년 %m월 %d일')
    mr = d['m_rate']
    ar = d['a_rate']
    fr_rate = d['fcst_rate']
 
    # 품목 행
    item_rows = ''
    for it in d['item_data']:
        item_rows += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;font-weight:500;color:#1a1a1a">{it['item']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:600;color:#1a3a6b">{fs(it['val'])}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;text-align:center;color:{chg_color(it['chg_mom'])};font-weight:500">{chg_arrow(it['chg_mom'])}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #f0f0f0;text-align:center;color:{chg_color(it['chg_avg'])};font-weight:500">{chg_arrow(it['chg_avg'])}</td>
        </tr>"""
 
    # 달성 상태 메시지
    if mr is None:
        diag_bg, diag_color, diag_msg = '#f5f5f3', '#666', 'KPI 데이터 없음'
    elif mr >= 100:
        diag_bg, diag_color, diag_msg = '#EDFAF4', '#0a5c3e', f'🎉 KPI 초과 달성! ({mr:.1f}%)'
    elif mr >= 90:
        diag_bg, diag_color, diag_msg = '#FFF8E5', '#7a4f00', f'⚡ KPI 근접 달성 ({mr:.1f}%) — 목표까지 {fs(d["mKPI"]-d["cur_sales"])} 남음'
    else:
        diag_bg, diag_color, diag_msg = '#FEF0F0', '#c0392b', f'⚠️ KPI 미달 ({mr:.1f}%) — 목표 대비 {fs(d["mKPI"]-d["cur_sales"])} 부족'
 
    # 예측 상태
    if fr_rate is None:
        fcst_diag, fcst_bg, fcst_tc = 'KPI 데이터 없음', '#f5f5f3', '#666'
    elif fr_rate >= 100:
        fcst_diag, fcst_bg, fcst_tc = f'예측 기준 KPI 초과 달성 가능 ({fr_rate:.1f}%)', '#EDFAF4', '#0a5c3e'
    elif fr_rate >= 90:
        fcst_diag, fcst_bg, fcst_tc = f'예측 기준 KPI 근접 달성 예상 ({fr_rate:.1f}%)', '#FFF8E5', '#7a4f00'
    else:
        fcst_diag, fcst_bg, fcst_tc = f'예측 기준 KPI 미달 예상 ({fr_rate:.1f}%) — {fs(max(0, d["mKPI"]-d["fcst"]))} 추가 필요', '#FEF0F0', '#c0392b'
 
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DNC 매출 현황 리포트</title>
</head>
<body style="margin:0;padding:0;background:#f0f0ee;font-family:-apple-system,BlinkMacSystemFont,'Noto Sans KR',sans-serif">
<div style="max-width:620px;margin:24px auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
 
  <!-- 헤더 -->
  <div style="background:#1a3a6b;padding:28px 32px;text-align:center">
    <div style="font-size:22px;font-weight:700;color:#fff;letter-spacing:0.03em">DNC 매출 현황 리포트</div>
    <div style="font-size:13px;color:rgba(255,255,255,0.65);margin-top:6px">{today} 기준 · DA_RPM사업부</div>
    <div style="margin-top:14px">
      <a href="{DASHBOARD_URL}" style="display:inline-block;background:#fff;color:#1a3a6b;padding:8px 22px;border-radius:20px;font-size:12px;font-weight:600;text-decoration:none">📊 대시보드 바로가기</a>
    </div>
  </div>
 
  <div style="padding:28px 32px">
  <p style="font-size:13px;color:#555;margin-bottom:20px">
  안녕하세요.<br>
  DNC 매출 현황 자동 리포트 전달드립니다.
  </p>

    <!-- 업데이트 기준 -->
    <div style="background:#f5f5f3;border-radius:8px;padding:8px 14px;font-size:11px;color:#888;margin-bottom:24px;text-align:center">
      📅 데이터 기준: {d['cy']}년 {d['cm']}월 (업데이트: {d['max_date']})
    </div>
 
    <!-- 핵심 지표 -->
    <div style="font-size:11px;font-weight:600;color:#888;letter-spacing:.07em;text-transform:uppercase;margin-bottom:12px">핵심 지표</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px">
      <tr>
        <td width="33%" style="padding-right:6px">
          <div style="background:#f5f5f3;border-radius:10px;padding:14px">
            <div style="font-size:10px;color:#888;margin-bottom:4px">당월 매출</div>
            <div style="font-size:18px;font-weight:700;color:#1a3a6b">{fs(d['cur_sales'])}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(d['mom_r'])}">전월대비 {chg_arrow(d['mom_r'])}</div>
            <div style="font-size:10px;margin-top:2px;color:{chg_color(d['avg3_r'])}">직3평균 {chg_arrow(d['avg3_r'])}</div>
          </div>
        </td>
        <td width="33%" style="padding:0 3px">
          <div style="background:#f5f5f3;border-radius:10px;padding:14px">
            <div style="font-size:10px;color:#888;margin-bottom:4px">당월 KPI 달성률</div>
            <div style="font-size:18px;font-weight:700;color:{rate_color(mr)}">{f"{mr:.1f}%" if mr else "-"}</div>
            <div style="font-size:10px;margin-top:4px;color:#888">목표 {fs(d['mKPI']) if d['mKPI'] else '-'}</div>
            <div style="font-size:10px;margin-top:2px;color:{rate_color(ar)}">연누적 {f"{ar:.1f}%" if ar else "-"}</div>
          </div>
        </td>
        <td width="33%" style="padding-left:6px">
          <div style="background:#f5f5f3;border-radius:10px;padding:14px">
            <div style="font-size:10px;color:#888;margin-bottom:4px">거래처 수</div>
            <div style="font-size:18px;font-weight:700;color:#1a3a6b">{d['cur_clients']:,}개처</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(d['cl_diff'])}">전월대비 {'▲' if d['cl_diff']>=0 else '▼'}{abs(d['cl_diff'])}개처</div>
          </div>
        </td>
      </tr>
    </table>
    <div style="background:{diag_bg};border-radius:8px;padding:10px 14px;font-size:12px;color:{diag_color};font-weight:500;margin-bottom:24px">{diag_msg}</div>
 
    <!-- 당월 예측 마감 -->
    <div style="font-size:11px;font-weight:600;color:#888;letter-spacing:.07em;text-transform:uppercase;margin-bottom:12px">⚡ 당월 예측 마감</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px">
      <tr>
        <td width="50%" style="padding-right:6px">
          <div style="background:#f5f5f3;border-radius:10px;padding:14px">
            <div style="font-size:10px;color:#888;margin-bottom:4px">예측 마감 매출</div>
            <div style="font-size:18px;font-weight:700;color:{rate_color(fr_rate)}">{fs(d['fcst'])}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(d['fcst_vs_avg'])}">직3평균 대비 {chg_arrow(d['fcst_vs_avg'])}</div>
          </div>
        </td>
        <td width="50%" style="padding-left:6px">
          <div style="background:#f5f5f3;border-radius:10px;padding:14px">
            <div style="font-size:10px;color:#888;margin-bottom:4px">예측 KPI 달성률</div>
            <div style="font-size:18px;font-weight:700;color:{rate_color(fr_rate)}">{f"{fr_rate:.1f}%" if fr_rate else "-"}</div>
            <div style="font-size:10px;margin-top:4px;color:#888">추가 필요: {fs(max(0, d['mKPI']-d['fcst'])) if d['mKPI'] else '-'}</div>
          </div>
        </td>
      </tr>
    </table>
    <div style="background:{fcst_bg};border-radius:8px;padding:10px 14px;font-size:12px;color:{fcst_tc};font-weight:500;margin-bottom:24px">{fcst_diag}</div>
 
    <!-- 주요 품목별 실적 -->
    <div style="font-size:11px;font-weight:600;color:#888;letter-spacing:.07em;text-transform:uppercase;margin-bottom:12px">🏆 주요 품목 실적</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #eee;border-radius:10px;overflow:hidden;margin-bottom:24px">
      <thead>
        <tr style="background:#f5f5f3">
          <th style="padding:8px 12px;text-align:left;font-size:10px;color:#888;font-weight:500">품목</th>
          <th style="padding:8px 12px;text-align:right;font-size:10px;color:#888;font-weight:500">매출</th>
          <th style="padding:8px 12px;text-align:center;font-size:10px;color:#888;font-weight:500">전월대비</th>
          <th style="padding:8px 12px;text-align:center;font-size:10px;color:#888;font-weight:500">직3평균대비</th>
        </tr>
      </thead>
      <tbody>{item_rows}</tbody>
    </table>
 
    <!-- 대시보드 링크 -->
    <div style="text-align:center;margin-bottom:8px">
      <a href="{DASHBOARD_URL}" style="display:inline-block;background:#1a3a6b;color:#fff;padding:12px 32px;border-radius:10px;font-size:13px;font-weight:600;text-decoration:none">📊 전체 대시보드 보기</a>
    </div>
 
  </div>
 
  <!-- 푸터 -->
  <div style="background:#f5f5f3;padding:16px 32px;text-align:center;font-size:10px;color:#aaa">
    DNC AESTHETICS · DA_RPM사업부 매출 현황 자동 리포트<br>본 메일은 자동 발송됩니다.
  </div>
 
</div>
</body>
</html>"""
    return html
 
def send_email(subject, html_body):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From']    = GMAIL_USER
    msg['To']      = RECV_EMAIL
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, RECV_EMAIL, msg.as_string())
    print(f"✅ 메일 발송 완료 → {RECV_EMAIL}")
 
if __name__ == '__main__':
    print("📊 report_data.json 읽는 중...")
    if not os.path.exists(DATA_FILE):
        print(f"❌ {DATA_FILE} 없음! update_dashboard.py 먼저 실행하세요.")
        exit(1)
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        d = json.load(f)
    subject = f"[DNC] {d['cy']}년 {d['cm']}월 매출 현황 리포트 ({d['max_date']} 기준)"
    print(f"📧 메일 발송 중: {subject}")
    html = build_html(d)
    send_email(subject, html)
