"""
DNC 매출 현황 대시보드 — 자동 메일 발송 스크립트
"""

import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

GMAIL_USER     = os.environ.get('GMAIL_USER', '')
GMAIL_PASSWORD = os.environ.get('GMAIL_PASSWORD', '')
RECV_EMAIL     = os.environ.get('RECV_EMAIL', '')
DASHBOARD_URL  = os.environ.get('DASHBOARD_URL', 'https://dnc-dashboard.kr/')
DATA_FILE      = 'report_data.json'

def fs(v):
    if v is None: return '-'
    if v >= 1e8: return f"{v/1e8:.1f}억원"
    if v >= 1e4: return f"{v/1e4:,.0f}만원"
    return f"{v:,.0f}원"

def rate_str(r):
    if r is None: return '-'
    return f"{r:.1f}%"

def chg_color(v):
    if v is None: return '#888888'
    return '#1D9E75' if v >= 0 else '#E24B4A'

def chg_arrow(v):
    if v is None: return '-'
    return f"{'▲' if v >= 0 else '▼'}{abs(v):.1f}%"

def diff_str(v, unit=''):
    if v is None: return '-'
    sign = '+' if v >= 0 else ''
    return f"{sign}{int(v)}{unit}"

def rate_color(r):
    if r is None: return '#888888'
    if r >= 100: return '#1D9E75'
    if r >= 90:  return '#e8a838'
    return '#E24B4A'

TD  = "padding:7px 10px;border-bottom:0.5px solid #eee;text-align:right;font-size:11px;color:#1a1a1a"
TDL = "padding:7px 10px;border-bottom:0.5px solid #eee;text-align:left;font-size:11px;font-weight:500;color:#1a1a1a"
TH  = "padding:6px 10px;border-bottom:0.5px solid #eee;text-align:right;font-size:10px;color:#888;font-weight:400"
THL = "padding:6px 10px;border-bottom:0.5px solid #eee;text-align:left;font-size:10px;color:#888;font-weight:400"
TOT = "padding:7px 10px;text-align:right;font-size:11px;font-weight:500;color:#1a1a1a;background:#f5f5f3"
TOTL= "padding:7px 10px;text-align:left;font-size:11px;font-weight:500;color:#1a1a1a;background:#f5f5f3"

def build_html(d):
    today = datetime.now().strftime('%Y년 %m월 %d일')
    cy, cm = d['cy'], d['cm']
    mr  = d.get('m_rate')
    ar  = d.get('a_rate')
    fr_rate = d.get('fcst_rate')
    max_date = d.get('max_date', '')

    if mr is None:
        diag_bg, diag_tc, diag_msg = '#f5f5f3', '#666', 'KPI 데이터 없음'
    elif mr >= 100:
        diag_bg, diag_tc, diag_msg = '#e8f4ea', '#0F6E56', f'🎉 KPI 초과 달성! ({mr:.1f}%)'
    elif mr >= 90:
        diag_bg, diag_tc, diag_msg = '#FFF8E5', '#7a4f00', f'⚡ KPI 근접 달성 ({mr:.1f}%) — 목표까지 {fs(d["mKPI"]-d["cur_sales"])} 남음'
    else:
        diag_bg, diag_tc, diag_msg = '#fde8e8', '#A32D2D', f'⚠️ KPI 미달 ({mr:.1f}%) — 목표 대비 {fs(d["mKPI"]-d["cur_sales"])} 부족'

    if fr_rate is None:
        fcst_bg, fcst_tc, fcst_msg = '#f5f5f3', '#666', 'KPI 데이터 없음'
    elif fr_rate >= 100:
        fcst_bg, fcst_tc, fcst_msg = '#e8f4ea', '#0F6E56', f'✅ 예측 기준 KPI 초과 달성 가능 ({fr_rate:.1f}%)'
    elif fr_rate >= 90:
        fcst_bg, fcst_tc, fcst_msg = '#e8f4ea', '#0F6E56', f'✅ 예측 기준 KPI 근접 달성 예상 ({fr_rate:.1f}%)'
    else:
        fcst_bg, fcst_tc, fcst_msg = '#fde8e8', '#A32D2D', f'⚠️ 예측 기준 KPI 미달 예상 ({fr_rate:.1f}%) — {fs(max(0, d["mKPI"]-d["fcst"]))} 추가 필요'

    cur_new    = d.get('cur_new', 0)
    cur_exist  = d.get('cur_exist', 0)
    new_diff   = d.get('new_diff', 0)
    exist_diff = d.get('exist_diff', 0)
    cl_diff    = d.get('cl_diff', 0)

    # 품목별 당월 매출 행
    item_rows = ''
    for it in d.get('item_data', []):
        item_rows += f"""
        <tr>
          <td style="{TDL}">{it['item']}</td>
          <td style="{TD};font-weight:500">{fs(it['val'])}</td>
          <td style="{TD};color:{rate_color(it.get('rate'))}">{rate_str(it.get('rate'))}</td>
          <td style="{TD};color:{chg_color(it.get('chg_mom'))}">{chg_arrow(it.get('chg_mom'))}</td>
          <td style="{TD};color:{chg_color(it.get('chg_avg'))}">{chg_arrow(it.get('chg_avg'))}</td>
        </tr>"""
    item_total = sum(it['val'] for it in d.get('item_data', []))
    item_rows += f"""
        <tr>
          <td style="{TOTL}">합계</td>
          <td style="{TOT}">{fs(item_total)}</td>
          <td style="{TOT};color:{rate_color(mr)}">{rate_str(mr)}</td>
          <td style="{TOT};color:{chg_color(d.get('mom_r'))}">{chg_arrow(d.get('mom_r'))}</td>
          <td style="{TOT};color:{chg_color(d.get('avg3_r'))}">{chg_arrow(d.get('avg3_r'))}</td>
        </tr>"""

    # 팀별 당월 매출 행
    team_rows = ''
    team_total = 0
    for t in d.get('team_data', []):
        if t.get('val', 0) == 0: continue
        team_total += t.get('val', 0)
        team_rows += f"""
        <tr>
          <td style="{TDL}">{t['team']}</td>
          <td style="{TD};font-weight:500">{fs(t['val'])}</td>
          <td style="{TD};color:{rate_color(t.get('rate'))}">{rate_str(t.get('rate'))}</td>
          <td style="{TD};color:{chg_color(t.get('chg_mom'))}">{chg_arrow(t.get('chg_mom'))}</td>
          <td style="{TD};color:{chg_color(t.get('chg_avg'))}">{chg_arrow(t.get('chg_avg'))}</td>
        </tr>"""
    team_rows += f"""
        <tr>
          <td style="{TOTL}">합계</td>
          <td style="{TOT}">{fs(team_total)}</td>
          <td style="{TOT};color:{rate_color(mr)}">{rate_str(mr)}</td>
          <td style="{TOT};color:{chg_color(d.get('mom_r'))}">{chg_arrow(d.get('mom_r'))}</td>
          <td style="{TOT};color:{chg_color(d.get('avg3_r'))}">{chg_arrow(d.get('avg3_r'))}</td>
        </tr>"""

    # 주요 품목 누적 매출 행 (매출+달성률만)
    acc_rows = ''
    acc_total = 0
    for it in d.get('item_data', []):
        acc_total += it['val']
        acc_rows += f"""
        <tr>
          <td style="{TDL}">{it['item']}</td>
          <td style="{TD};font-weight:500">{fs(it['val'])}</td>
          <td style="{TD};color:{rate_color(it.get('rate'))}">{rate_str(it.get('rate'))}</td>
        </tr>"""
    acc_rows += f"""
        <tr>
          <td style="{TOTL}">합계</td>
          <td style="{TOT}">{fs(acc_total)}</td>
          <td style="{TOT};color:{rate_color(mr)}">{rate_str(mr)}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DNC 매출 현황 리포트</title>
</head>
<body style="margin:0;padding:0;background:#f0f0ee;font-family:-apple-system,BlinkMacSystemFont,'Noto Sans KR',sans-serif">
<div style="max-width:620px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;border:0.5px solid #ddd">

  <div style="background:#1a2a4a;padding:28px 32px;text-align:center">
    <div style="font-size:18px;font-weight:500;color:#fff;margin-bottom:4px">DNC 매출 현황 리포트</div>
    <div style="font-size:12px;color:#aac4ff;margin-bottom:14px">{today} 기준 · DA_RPM사업부</div>
    <a href="{DASHBOARD_URL}" style="display:inline-flex;align-items:center;gap:6px;padding:8px 18px;background:#2557a0;color:#fff;border-radius:8px;font-size:12px;font-weight:500;text-decoration:none">🔲 대시보드 바로가기</a>
  </div>

  <div style="padding:1rem 1.5rem;font-size:13px;color:#1a1a1a;border-bottom:0.5px solid #eee">
    안녕하세요.<br>DNC 매출 현황 자동 리포트 전달드립니다.
    <div style="margin-top:8px;padding:5px 10px;background:#f5f5f3;border-radius:6px;font-size:11px;color:#888;display:inline-block">
      📅 데이터 기준: {cy}년 {cm}월 (업데이트: {max_date})
    </div>
  </div>

  <div style="padding:1.25rem 1.5rem;border-bottom:0.5px solid #eee">
    <div style="font-size:12px;font-weight:500;color:#1a1a1a;margin-bottom:12px">핵심 지표</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px">
      <tr>
        <td width="50%" style="padding-right:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">당월 매출</div>
            <div style="font-size:20px;font-weight:500;color:#1a1a1a">{fs(d['cur_sales'])}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(d.get('mom_r'))}">전월대비 {chg_arrow(d.get('mom_r'))}</div>
            <div style="font-size:10px;margin-top:2px;color:#888">직3평균 {chg_arrow(d.get('avg3_r'))}</div>
          </div>
        </td>
        <td width="50%" style="padding-left:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">당월 KPI 달성률</div>
            <div style="font-size:20px;font-weight:500;color:{rate_color(mr)}">{rate_str(mr)}</div>
            <div style="font-size:10px;margin-top:4px;color:#888">목표 {fs(d.get('mKPI', 0))}</div>
            <div style="font-size:10px;margin-top:2px;color:#888">연누적 {rate_str(ar)}</div>
          </div>
        </td>
      </tr>
    </table>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px">
      <tr>
        <td width="33%" style="padding-right:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">거래처 수</div>
            <div style="font-size:20px;font-weight:500;color:#1a1a1a">{d.get('cur_clients', 0):,}개처</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(cl_diff)}">전월대비 {diff_str(cl_diff)}</div>
          </div>
        </td>
        <td width="33%" style="padding:0 2px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">신규처</div>
            <div style="font-size:20px;font-weight:500;color:#1a1a1a">{cur_new}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(new_diff)}">전월대비 {diff_str(new_diff)}</div>
          </div>
        </td>
        <td width="33%" style="padding-left:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">기존처</div>
            <div style="font-size:20px;font-weight:500;color:#1a1a1a">{cur_exist}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(exist_diff)}">전월대비 {diff_str(exist_diff)}</div>
          </div>
        </td>
      </tr>
    </table>
    <div style="padding:8px 12px;border-radius:6px;font-size:11px;background:{diag_bg};color:{diag_tc}">{diag_msg}</div>
  </div>

  <div style="padding:1.25rem 1.5rem;border-bottom:0.5px solid #eee">
    <div style="font-size:12px;font-weight:500;color:#1a1a1a;margin-bottom:12px">⚡ 당월 예측 마감</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px">
      <tr>
        <td width="50%" style="padding-right:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">예측 마감 매출</div>
            <div style="font-size:20px;font-weight:500;color:{rate_color(fr_rate)}">{fs(d.get('fcst', 0))}</div>
            <div style="font-size:10px;margin-top:4px;color:{chg_color(d.get('fcst_vs_avg'))}">직3평균 대비 {chg_arrow(d.get('fcst_vs_avg'))}</div>
          </div>
        </td>
        <td width="50%" style="padding-left:4px">
          <div style="background:#f5f5f3;border-radius:8px;padding:12px 14px">
            <div style="font-size:10px;color:#888;margin-bottom:6px">예측 KPI 달성률</div>
            <div style="font-size:20px;font-weight:500;color:{rate_color(fr_rate)}">{rate_str(fr_rate)}</div>
            <div style="font-size:10px;margin-top:4px;color:#888">추가 필요: {fs(max(0, d.get('mKPI', 0)-d.get('fcst', 0)))}</div>
          </div>
        </td>
      </tr>
    </table>
    <div style="padding:8px 12px;border-radius:6px;font-size:11px;background:{fcst_bg};color:{fcst_tc}">{fcst_msg}</div>
  </div>

  <div style="padding:1.25rem 1.5rem;border-bottom:0.5px solid #eee">
    <div style="font-size:12px;font-weight:500;color:#1a1a1a;margin-bottom:12px">🏷 품목별 당월 매출</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:0.5px solid #eee;border-radius:8px;overflow:hidden">
      <thead><tr style="background:#f5f5f3">
        <th style="{THL}">품목</th><th style="{TH}">매출액</th><th style="{TH}">달성률</th><th style="{TH}">전월대비</th><th style="{TH}">직3평균대비</th>
      </tr></thead>
      <tbody>{item_rows}</tbody>
    </table>
  </div>

  <div style="padding:1.25rem 1.5rem;border-bottom:0.5px solid #eee">
    <div style="font-size:12px;font-weight:500;color:#1a1a1a;margin-bottom:12px">👥 팀별 당월 매출</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:0.5px solid #eee;border-radius:8px;overflow:hidden">
      <thead><tr style="background:#f5f5f3">
        <th style="{THL}">팀</th><th style="{TH}">매출</th><th style="{TH}">달성률</th><th style="{TH}">전월대비</th><th style="{TH}">직3평균대비</th>
      </tr></thead>
      <tbody>{team_rows}</tbody>
    </table>
  </div>

  <div style="padding:1.25rem 1.5rem;border-bottom:0.5px solid #eee">
    <div style="font-size:12px;font-weight:500;color:#1a1a1a;margin-bottom:12px">🏆 주요 품목 누적 매출 <span style="font-size:10px;color:#888;font-weight:400">1~{cm}월 누적</span></div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:0.5px solid #eee;border-radius:8px;overflow:hidden">
      <thead><tr style="background:#f5f5f3">
        <th style="{THL}">품목</th><th style="{TH}">매출</th><th style="{TH}">달성률</th>
      </tr></thead>
      <tbody>{acc_rows}</tbody>
    </table>
  </div>

  <div style="padding:1.25rem 1.5rem;text-align:center">
    <a href="{DASHBOARD_URL}" style="display:inline-flex;align-items:center;gap:6px;padding:10px 24px;background:#185FA5;color:#fff;border-radius:8px;font-size:13px;font-weight:500;text-decoration:none;margin-bottom:16px">🔲 전체 대시보드 보기</a>
    <div style="font-size:11px;color:#888;line-height:1.7">DNC AESTHETICS · DA_RPM사업부 매출 현황 자동 리포트<br>본 메일은 자동 발송됩니다.</div>
  </div>

</div>
</body>
</html>"""
    return html


def send_email(subject, html_body):
    recipients = [r.strip() for r in RECV_EMAIL.split(',')]
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From']    = GMAIL_USER
    msg['To']      = ', '.join(recipients)
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, recipients, msg.as_string())
    print(f"✅ 메일 발송 완료 → {', '.join(recipients)}")


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
