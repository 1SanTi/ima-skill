#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教师问卷在线答题与数据汇总系统（零第三方依赖，仅用 Python 标准库）

功能：
  /               手机端答题页（老师扫码打开）
  /submit         接收答卷并入库（SQLite）
  /thanks         提交成功页
  /admin          后台数据汇总看板（统计图表 + 明细 + 导出 + 二维码）
  /api/stats      统计数据（JSON）
  /api/responses  全部答卷明细（JSON）
  /api/export.csv 导出 CSV（Excel 可直接打开）
  /api/qr.svg     答题链接二维码（SVG）

启动：python3 server.py            （默认端口 8000，可用环境变量 PORT 修改）
后台：浏览器打开 http://本机IP:端口/admin
"""
import json
import os
import sys
import csv
import io
import html
import socket
import sqlite3
import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE_DIR, "vendor"))

SPEC_PATH = os.path.join(BASE_DIR, "questions.json")
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "responses.db")
PORT = int(os.environ.get("PORT", "8000"))
ADMIN_PWD = os.environ.get("ADMIN_PWD", "").strip()
HOSTNAME = socket.gethostname()

with open(SPEC_PATH, encoding="utf-8") as f:
    SPEC = json.load(f)
SECTIONS = SPEC["sections"]
QUESTIONS = [q for s in SECTIONS for q in s["questions"]]
QMAP = {q["id"]: q for q in QUESTIONS}
SCALE = SPEC.get("scale_options", ["非常清楚", "比较清楚", "一般", "不太清楚", "完全不清楚"])


# ---------------------------------------------------------------- 数据库
def db_conn():
    os.makedirs(DATA_DIR, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.execute(
        "CREATE TABLE IF NOT EXISTS responses ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, submitted_at TEXT, ip TEXT, ua TEXT, data TEXT)"
    )
    return con


def init_db():
    con = db_conn()
    con.commit()
    con.close()


def get_lan_ip():
    ip = "127.0.0.1"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        try:
            ip = socket.gethostbyname(HOSTNAME)
        except Exception:
            pass
    return ip


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- 工具
def q_options(q):
    if q["type"] == "scale" or q.get("options_scale"):
        return SCALE
    if q["type"] == "rating":
        return [str(i) for i in range(q.get("min", 1), q.get("max", 10) + 1)]
    return q.get("options", [])


def esc(s):
    return html.escape(str(s), quote=True)


# ---------------------------------------------------------------- 答题页
PAGE_CSS = """
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;
 background:#f2f4f8;color:#1f2733;line-height:1.7;font-size:16px}
.wrap{max-width:680px;margin:0 auto;padding:0 0 90px}
.head{background:linear-gradient(135deg,#2b6cff,#4f8cff);color:#fff;padding:26px 20px 22px;border-radius:0 0 18px 18px}
.head h1{margin:0 0 6px;font-size:21px;font-weight:700}
.head .sub{opacity:.9;font-size:13px}
.intro{background:#fff;margin:-10px 12px 0;padding:16px 16px;border-radius:14px;
 box-shadow:0 4px 16px rgba(30,60,120,.08);font-size:14px;color:#3a4657;white-space:pre-line}
.sec{margin:20px 12px 6px;font-size:15px;font-weight:700;color:#2b6cff;display:flex;align-items:center;gap:8px}
.sec::before{content:"";width:4px;height:16px;background:#2b6cff;border-radius:2px}
.card{background:#fff;margin:10px 12px;padding:16px;border-radius:14px;box-shadow:0 2px 10px rgba(30,60,120,.06)}
.qtitle{font-weight:600;font-size:15.5px;margin-bottom:12px}
.qtitle .no{color:#2b6cff}
.req{color:#e5484d;font-weight:400;font-size:13px}
.opts{display:flex;flex-direction:column;gap:8px}
.opt{display:flex;align-items:flex-start;gap:10px;padding:11px 12px;border:1.5px solid #e6ebf2;
 border-radius:11px;cursor:pointer;transition:.15s;font-size:15px}
.opt:active{transform:scale(.99)}
.opt input{margin:3px 0 0;accent-color:#2b6cff;width:18px;height:18px;flex:none}
.opt.sel{border-color:#2b6cff;background:#eef4ff}
.scale{display:flex;flex-direction:column;gap:8px}
.opts.rate{flex-direction:row;flex-wrap:wrap;gap:8px}
.opts.rate .opt{padding:10px 12px;flex:0 0 auto;min-width:50px;justify-content:center}
.opts.rate .opt span{min-width:14px;text-align:center}
.ratehint{font-size:12.5px;color:#8a95a5;margin-top:10px;display:flex;justify-content:space-between}
textarea{width:100%;min-height:96px;border:1.5px solid #e6ebf2;border-radius:11px;padding:11px 12px;
 font-size:15px;font-family:inherit;resize:vertical;outline:none}
textarea:focus{border-color:#2b6cff}
.other-in{border:1.5px solid #e6ebf2;border-radius:10px;padding:8px 10px;font-size:15px;width:100%;
 font-family:inherit;outline:none;margin-top:-2px}
.other-in:focus{border-color:#2b6cff}
.bar{position:fixed;left:0;right:0;bottom:0;background:#fff;padding:12px 16px;
 box-shadow:0 -4px 18px rgba(30,60,120,.12);display:flex;justify-content:center}
.bar .inner{max-width:680px;width:100%}
.btn{width:100%;border:0;border-radius:12px;padding:15px;font-size:17px;font-weight:700;color:#fff;
 background:linear-gradient(135deg,#2b6cff,#4f8cff);cursor:pointer}
.btn:active{opacity:.9}
.err{background:#fff1f1;border:1.5px solid #ffd0d0;color:#c0272d;margin:10px 12px;padding:12px 14px;
 border-radius:12px;font-size:14px}
.missing{border-color:#ffb3b3!important;background:#fff7f7}
.center{text-align:center;padding:60px 24px}
.center .ic{font-size:56px}
.center h2{margin:14px 0 8px;color:#18a058}
.center p{color:#6b7787}
.done-btn{display:inline-block;margin-top:18px;padding:12px 26px;border-radius:12px;background:#eef4ff;
 color:#2b6cff;text-decoration:none;font-weight:600}
"""

PAGE_JS = """
document.querySelectorAll('.opt').forEach(function(l){
  var inp=l.querySelector('input');
  function sync(){
    var name=inp.name;
    if(inp.type==='radio'){
      document.querySelectorAll('input[name="'+name+'"]').forEach(function(r){
        var lb=r.closest('.opt'); if(lb) lb.classList.toggle('sel', r.checked);
      });
    } else {
      l.classList.toggle('sel', inp.checked);
    }
    var other=document.querySelector('input[name="'+name+'"][value="__other__"]');
    var t=document.getElementById(name+'__other');
    if(t){
      var on = other ? other.checked : false;
      t.style.display = on ? 'block' : 'none';
      if(on && inp===other) t.focus();
    }
  }
  inp.addEventListener('change',sync); sync();
});
"""


def _opt_html(q, opt, values, missing):
    qid = q["id"]
    is_other = opt == "__other__"
    label = "其他（请注明）" if is_other else opt
    if q["type"] == "multi":
        checked = (opt in values) if not is_other else (any(str(v).startswith("其他") for v in values))
        inp = f'<input type="checkbox" name="{qid}" value="{esc(opt)}" {"checked" if checked else ""}>'
        text = ""
        if is_other:
            other_txt = ""
            for v in values:
                if str(v).startswith("其他："):
                    other_txt = v.split("：", 1)[1]
            text = (f'<input class="other-in" id="{qid}__other" name="{qid}__other" '
                    f'placeholder="请填写" value="{esc(other_txt)}" style="display:{"block" if checked else "none"}">')
        return f'<label class="opt {"sel" if checked else ""}">{inp}<span>{esc(label)}</span></label>{text}'
    else:
        cur = values.get(qid, "")
        checked = (cur == "__other__") if is_other else (cur == opt)
        inp = f'<input type="radio" name="{qid}" value="{esc(opt)}" {"checked" if checked else ""}>'
        text = ""
        if is_other:
            other_txt = ""
            if str(cur).startswith("其他："):
                other_txt = cur.split("：", 1)[1]
            text = (f'<input class="other-in" id="{qid}__other" name="{qid}__other" '
                    f'placeholder="请填写" value="{esc(other_txt)}" style="display:{"block" if checked else "none"}">')
        return f'<label class="opt {"sel" if checked else ""}">{inp}<span>{esc(label)}</span></label>{text}'


def render_form(values=None, missing=None):
    values = values or {}
    missing = set(missing or [])
    parts = []
    parts.append('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">')
    parts.append('<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">')
    parts.append(f'<title>{esc(SPEC.get("title","调查问卷"))}</title>')
    parts.append(f'<style>{PAGE_CSS}</style></head><body><div class="wrap">')
    parts.append('<div class="head"><h1>%s</h1><div class="sub">%s</div></div>' % (
        esc(SPEC.get("title", "调查问卷")), esc(SPEC.get("subtitle", ""))))
    parts.append('<div class="intro">%s</div>' % esc(SPEC.get("intro", "")))
    if missing:
        parts.append('<div class="err">还有 %d 道必答题未完成，请补全后再提交（下方红色边框处）。</div>' % len(missing))
    parts.append('<form method="post" action="/submit">')
    n = 0
    for sec in SECTIONS:
        parts.append('<div class="sec">%s</div>' % esc(sec["title"]))
        for q in sec["questions"]:
            n += 1
            qid = q["id"]
            badge = '<span class="req">（必答）</span>' if q.get("required") else ''
            cls = "card missing" if qid in missing else "card"
            parts.append('<div class="%s" id="card_%s">' % (cls, qid))
            parts.append('<div class="qtitle"><span class="no">%d.</span> %s %s</div>' % (n, esc(q["title"]), badge))
            if q["type"] == "text":
                v = values.get(qid, "")
                parts.append('<textarea name="%s" placeholder="%s">%s</textarea>' % (
                    qid, esc(q.get("placeholder", "请填写")), esc(v)))
            else:
                is_rate = q["type"] == "rating"
                parts.append('<div class="opts %s">' % ("rate" if is_rate else "scale"))
                for opt in q_options(q):
                    parts.append(_opt_html(q, opt, values, missing))
                if q.get("allow_other"):
                    parts.append(_opt_html(q, "__other__", values, missing))
                parts.append('</div>')
                if is_rate:
                    parts.append('<div class="ratehint"><span>%s</span><span>%s</span></div>'
                                 % (esc(q.get("low_label", "")), esc(q.get("high_label", ""))))
            parts.append('</div>')
    parts.append('<div class="bar"><div class="inner"><button class="btn" type="submit">提 交 问 卷</button></div></div>')
    parts.append('</form></div>')
    parts.append('<script>%s</script>' % PAGE_JS)
    parts.append('</body></html>')
    return "".join(parts)


THANKS_HTML = """<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>提交成功</title><style>%s</style></head><body>
<div class="center"><div class="ic">✅</div><h2>提交成功，感谢您的参与！</h2>
<p>您的答卷已记录。<br>本问卷为匿名填写，感谢您对学校工作的支持。</p>
<a class="done-btn" href="/">再填一份</a></div></body></html>""" % PAGE_CSS


# ---------------------------------------------------------------- 后台看板
ADMIN_HTML = r"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>问卷数据后台</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#eef1f6;color:#1f2733}
.top{background:#1f2733;color:#fff;padding:16px 22px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.top h1{font-size:17px;margin:0;font-weight:600}
.top .sp{flex:1}
.btn{border:0;border-radius:9px;padding:9px 16px;font-size:14px;cursor:pointer;background:#2b6cff;color:#fff;font-weight:600}
.btn.gray{background:#3a4657}
.btn.green{background:#18a058}
.cards{display:flex;gap:14px;padding:18px 22px 0;flex-wrap:wrap}
.kpi{background:#fff;border-radius:14px;padding:16px 22px;min-width:150px;box-shadow:0 2px 10px rgba(30,60,120,.07)}
.kpi .v{font-size:30px;font-weight:800;color:#2b6cff}
.kpi .l{color:#6b7787;font-size:13px;margin-top:2px}
.grid{padding:18px 22px;display:grid;grid-template-columns:repeat(auto-fill,minmax(430px,1fr));gap:16px}
.q{background:#fff;border-radius:14px;padding:16px 18px;box-shadow:0 2px 10px rgba(30,60,120,.07)}
.q h3{margin:0 0 4px;font-size:15px}
.q .tag{font-size:12px;color:#8a95a5;margin-bottom:10px}
.row{display:flex;align-items:center;gap:10px;margin:7px 0;font-size:13.5px}
.row .lab{width:150px;flex:none;color:#3a4657;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.row .barwrap{flex:1;background:#f0f3f8;border-radius:7px;height:20px;position:relative;overflow:hidden}
.row .bar{height:100%;background:linear-gradient(90deg,#4f8cff,#2b6cff);border-radius:7px;transition:width .4s}
.row .num{width:86px;flex:none;color:#6b7787;font-size:12.5px}
.txtitem{background:#f7f9fc;border-radius:9px;padding:9px 12px;margin:7px 0;font-size:14px;color:#3a4657;white-space:pre-wrap}
.tablewrap{padding:0 22px 30px}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;font-size:13px;box-shadow:0 2px 10px rgba(30,60,120,.07)}
th,td{padding:9px 10px;border-bottom:1px solid #eef1f6;text-align:left;vertical-align:top}
th{background:#f6f8fc;position:sticky;top:0;font-weight:600;white-space:nowrap}
td{max-width:260px}
.qrbox{background:#fff;border-radius:14px;padding:16px;box-shadow:0 2px 10px rgba(30,60,120,.07);display:flex;gap:16px;align-items:center;flex-wrap:wrap}
.qrbox img{width:150px;height:150px;border:1px solid #eef1f6;border-radius:8px}
.qrbox .info{font-size:13px;color:#3a4657;word-break:break-all;max-width:340px}
.qrbox input{border:1.5px solid #dbe2ec;border-radius:9px;padding:8px 10px;font-size:13px;width:300px;font-family:inherit}
.section-t{margin:26px 22px 10px;font-size:16px;font-weight:700;color:#1f2733}
.link{color:#2b6cff;text-decoration:none}
</style></head><body>
<div class="top">
  <h1 id="ttl">问卷数据后台</h1><div class="sp"></div>
  <button class="btn gray" onclick="loadAll()">刷新数据</button>
  <a class="btn green" href="/api/export.csv">导出 Excel(CSV)</a>
</div>
<div class="cards" id="kpis"></div>
<div class="section-t">📱 答题入口</div>
<div style="padding:0 22px"><div class="qrbox">
  <img id="qr" src="/api/qr.svg" alt="二维码">
  <div class="info">
    <div style="font-weight:600;margin-bottom:6px">老师用手机扫码即可答题</div>
    <div>答题网址：</div>
    <input id="url" value="" onchange="document.getElementById('qr').src='/api/qr.svg?url='+encodeURIComponent(this.value)">
    <div style="margin-top:8px;color:#8a95a5">把网址改成对外公布的地址后可生成对应二维码</div>
  </div>
</div></div>
<div class="section-t">📊 各题统计</div>
<div class="grid" id="stats"></div>
<div class="section-t">📋 答卷明细（<span id="cnt">0</span> 份）</div>
<div class="tablewrap"><div style="overflow:auto;max-height:560px"><table id="tbl"><thead id="thead"></thead><tbody id="tbody"></tbody></table></div></div>
<script>
function pct(a,b){return b?Math.round(a*1000/b)/10:0}
function loadAll(){
  fetch('/api/stats').then(function(r){return r.json()}).then(renderStats);
  fetch('/api/responses').then(function(r){return r.json()}).then(renderTable);
}
function renderStats(d){
  document.getElementById('ttl').textContent=d.title+' — 数据后台';
  document.getElementById('url').value=d.url;
  document.getElementById('qr').src='/api/qr.svg?url='+encodeURIComponent(d.url)+'&t='+Date.now();
  var k=document.getElementById('kpis');
  k.innerHTML='';
  [['答卷总数',d.total],['今日新增',d.today],['最早提交',d.first||'—'],['最近提交',d.last||'—']].forEach(function(x){
    var el=document.createElement('div');el.className='kpi';
    el.innerHTML='<div class="v" style="'+(typeof x[1]==='number'?'':'font-size:16px')+'">'+x[1]+'</div><div class="l">'+x[0]+'</div>';
    k.appendChild(el);
  });
  var box=document.getElementById('stats');box.innerHTML='';
  d.questions.forEach(function(q){
    var el=document.createElement('div');el.className='q';
    var tag={single:'单选题',scale:'量表题',rating:'打分题',multi:'多选题',text:'填空题'}[q.type]||'';
    var extra=(q.type==='rating'&&q.average!=null)?'　<span style="color:#2b6cff;font-weight:700">平均 '+q.average+' 分</span>':'';
    var h='<h3>'+q.no+'. '+q.title+'</h3><div class="tag">'+tag+'　有效作答 '+q.answered+' 份'+extra+'</div>';
    if(q.type==='text'){
      if(q.texts.length===0){h+='<div class="tag" style="color:#c0c7d1">（暂无填写）</div>';}
      q.texts.forEach(function(t){h+='<div class="txtitem">'+t+'</div>';});
    }else{
      h+='<div>';
      q.items.forEach(function(it){
        h+='<div class="row"><div class="lab" title="'+it.label+'">'+it.label+'</div>'+
           '<div class="barwrap"><div class="bar" style="width:'+Math.max(it.pct,it.count?2:0)+'%"></div></div>'+
           '<div class="num">'+it.count+' 人 · '+it.pct+'%</div></div>';
      });
      h+='</div>';
    }
    el.innerHTML=h;box.appendChild(el);
  });
}
function renderTable(d){
  document.getElementById('cnt').textContent=d.rows.length;
  var cols=d.columns,head='<tr><th>#</th><th>提交时间</th>';
  cols.forEach(function(c){head+='<th>'+c.no+'. '+c.title+'</th>';});
  head+='</tr>';
  document.getElementById('thead').innerHTML=head;
  var tb=document.getElementById('tbody');tb.innerHTML='';
  d.rows.slice().reverse().forEach(function(r,i){
    var tr='<tr><td>'+(d.rows.length-i)+'</td><td style="white-space:nowrap">'+r.time+'</td>';
    cols.forEach(function(c){tr+='<td>'+r.answers[c.id]+'</td>';});
    tr+='</tr>';
    tb.insertAdjacentHTML('beforeend',tr);
  });
}
loadAll();setInterval(loadAll,30000);
</script></body></html>"""


# ---------------------------------------------------------------- 统计
def compute_stats(base_url=None):
    con = db_conn()
    rows = con.execute("SELECT id,submitted_at,data FROM responses ORDER BY id").fetchall()
    con.close()
    recs = [{"id": r[0], "time": r[1], "data": json.loads(r[2])} for r in rows]
    total = len(recs)
    today = datetime.date.today().strftime("%Y-%m-%d")
    today_n = sum(1 for r in recs if r["time"].startswith(today))
    first = recs[0]["time"] if recs else ""
    last = recs[-1]["time"] if recs else ""
    qstats = []
    n = 0
    for q in QUESTIONS:
        n += 1
        qid = q["id"]
        item = {"id": qid, "no": n, "title": q["title"], "type": q["type"]}
        if q["type"] == "text":
            texts = [r["data"].get(qid, "") for r in recs if r["data"].get(qid)]
            item["texts"] = texts
            item["answered"] = len(texts)
            qstats.append(item)
            continue
        opts = q_options(q)
        denom = total
        counts = {}
        unanswered = 0
        if q["type"] == "multi":
            for o in opts:
                counts[o] = 0
            others = []
            for r in recs:
                vals = r["data"].get(qid) or []
                if not vals:
                    unanswered += 1
                for v in vals:
                    if v in counts:
                        counts[v] += 1
                    else:
                        others.append(v)
            items = [{"label": o, "count": counts[o]} for o in opts]
            if others:
                from collections import Counter
                for v, c in Counter(others).items():
                    items.append({"label": v, "count": c})
            for it in items:
                it["pct"] = pct_calc(it["count"], denom)
            item["items"] = items
            item["answered"] = total - unanswered
        else:
            for o in opts:
                counts[o] = 0
            counts["（未作答）"] = 0
            for r in recs:
                v = r["data"].get(qid, "")
                if v in counts:
                    counts[v] += 1
                elif v == "":
                    counts["（未作答）"] += 1
                else:
                    counts[v] = counts.get(v, 0) + 1
            items = [{"label": o, "count": counts[o]} for o in opts] + \
                    [{"label": "（未作答）", "count": counts["（未作答）"]}]
            for it in items:
                it["pct"] = pct_calc(it["count"], denom)
            item["items"] = items
            item["answered"] = total - counts["（未作答）"]
            if q["type"] == "rating":
                nums = []
                for r in recs:
                    v = str(r["data"].get(qid, ""))
                    if v.isdigit():
                        nums.append(int(v))
                item["average"] = round(sum(nums) / len(nums), 2) if nums else None
        qstats.append(item)
    return {
        "title": SPEC.get("title", "调查问卷"),
        "url": base_url or ("http://%s:%d/" % (get_lan_ip(), PORT)),
        "total": total, "today": today_n, "first": first, "last": last,
        "questions": qstats,
    }


def pct_calc(a, b):
    return round(a * 1000 / b) / 10 if b else 0


def responses_payload():
    con = db_conn()
    rows = con.execute("SELECT id,submitted_at,data FROM responses ORDER BY id").fetchall()
    con.close()
    cols = []
    n = 0
    for q in QUESTIONS:
        n += 1
        cols.append({"id": q["id"], "no": n, "title": q["title"], "type": q["type"]})
    out = []
    for r in rows:
        data = json.loads(r[2])
        answers = {}
        for q in QUESTIONS:
            v = data.get(q["id"], "")
            if isinstance(v, list):
                answers[q["id"]] = "；".join([str(x) for x in v]) if v else ""
            else:
                answers[q["id"]] = str(v)
        out.append({"id": r[0], "time": r[1], "answers": answers})
    return {"columns": cols, "rows": out}


# ---------------------------------------------------------------- 二维码
def qr_svg(data, box=8, border=2):
    try:
        sys.path.insert(0, os.path.join(BASE_DIR, "vendor"))
        import qrcode
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, border=border, box_size=1)
        qr.add_data(data)
        qr.make(fit=True)
        m = qr.get_matrix()
    except Exception as e:
        return '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="60"><text x="10" y="34" font-size="12" fill="#c0272d">二维码生成失败: %s</text></svg>' % esc(e)
    n = len(m)
    size = n * box
    r = []
    for y, rowm in enumerate(m):
        for x, val in enumerate(rowm):
            if val:
                r.append('<rect x="%d" y="%d" width="%d" height="%d"/>' % (x * box, y * box, box, box))
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
            '<rect width="%d" height="%d" fill="#ffffff"/><g fill="#000000">%s</g></svg>'
            % (size, size, size, size, size, size, "".join(r)))


# ---------------------------------------------------------------- 请求处理
class Handler(BaseHTTPRequestHandler):
    server_version = "SurveyServer/1.0"

    def log_message(self, fmt, *args):
        pass  # 静默

    def _base_url(self):
        """答题网址：优先用环境变量 PUBLIC_URL，否则用当前访问的域名/Host 自动推断。
        这样无论部署在公网服务器、内网穿透域名还是本机，后台二维码都会自动指向正确的地址。"""
        env = os.environ.get("PUBLIC_URL", "").strip()
        if env:
            return env if env.endswith("/") else env + "/"
        host = self.headers.get("Host") or ("%s:%d" % (get_lan_ip(), PORT))
        proto = self.headers.get("X-Forwarded-Proto", "http")
        return "%s://%s/" % (proto, host)

    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

    def _admin_ok(self, query):
        if not ADMIN_PWD:
            return True
        return query.get("pwd", [""])[0] == ADMIN_PWD

    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        query = parse_qs(u.query)
        if path in ("/", "/index.html"):
            self._send(200, render_form())
        elif path == "/thanks":
            self._send(200, THANKS_HTML)
        elif path == "/admin":
            if not self._admin_ok(query):
                self._send(200, '<meta charset="utf-8"><div style="font-family:sans-serif;padding:40px">'
                                '需要访问密码：请在网址后加 ?pwd=你的密码</div>')
            else:
                self._send(200, ADMIN_HTML)
        elif path == "/api/stats":
            self._json(compute_stats(self._base_url()))
        elif path == "/api/responses":
            self._json(responses_payload())
        elif path == "/api/qr.svg":
            url = query.get("url", [""])[0] or self._base_url()
            self._send(200, qr_svg(url), "image/svg+xml; charset=utf-8")
        elif path == "/api/export.csv":
            self._send(200, self._csv(), "text/csv; charset=utf-8",
                       {"Content-Disposition": "attachment; filename*=UTF-8''%s"
                        % quote("问卷数据_%s.csv" % datetime.date.today().strftime("%Y%m%d"))})
        elif path == "/health":
            self._send(200, "OK", "text/plain; charset=utf-8")
        elif path == "/favicon.ico":
            self._send(204, b"", "image/x-icon")
        else:
            self._send(404, "<meta charset='utf-8'>404 Not Found")

    def _csv(self):
        con = db_conn()
        rows = con.execute("SELECT id,submitted_at,data FROM responses ORDER BY id").fetchall()
        con.close()
        buf = io.StringIO()
        w = csv.writer(buf)
        header = ["序号", "提交时间"] + ["%d.%s" % (i + 1, q["title"]) for i, q in enumerate(QUESTIONS)]
        w.writerow(header)
        for i, r in enumerate(rows, 1):
            data = json.loads(r[2])
            line = [i, r[1]]
            for q in QUESTIONS:
                v = data.get(q["id"], "")
                line.append("；".join(map(str, v)) if isinstance(v, list) else v)
            w.writerow(line)
        return "\ufeff" + buf.getvalue()  # BOM，便于 Excel 识别

    def do_POST(self):
        u = urlparse(self.path)
        if u.path != "/submit":
            self._send(404, "404")
            return
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8")
        qs = parse_qs(body, keep_blank_values=True)
        data = {}
        missing = []
        for q in QUESTIONS:
            qid = q["id"]
            if q["type"] == "multi":
                raw = qs.get(qid, [])
                vals = [v for v in raw if v != "__other__"]
                other = qs.get(qid + "__other", [""])[0].strip()
                if "__other__" in raw and other:
                    vals.append("其他：" + other)
                data[qid] = vals
                if q.get("required") and not vals:
                    missing.append(qid)
            elif q["type"] == "text":
                data[qid] = qs.get(qid, [""])[0].strip()
            else:
                v = qs.get(qid, [""])[0]
                if v == "__other__":
                    other = qs.get(qid + "__other", [""])[0].strip()
                    v = ("其他：" + other) if other else "其他"
                data[qid] = v
                if q.get("required") and not v:
                    missing.append(qid)
        if missing:
            self._send(200, render_form(values=data, missing=missing))
            return
        con = db_conn()
        con.execute("INSERT INTO responses (submitted_at,ip,ua,data) VALUES (?,?,?,?)",
                    (now_str(), self.client_address[0],
                     self.headers.get("User-Agent", ""), json.dumps(data, ensure_ascii=False)))
        con.commit()
        con.close()
        self.send_response(303)
        self.send_header("Location", "/thanks")
        self.end_headers()


def main():
    init_db()
    ip = get_lan_ip()
    print("=" * 60)
    print("  教师问卷在线答题与数据汇总系统  已启动")
    print("-" * 60)
    print("  老师答题：  http://%s:%d/" % (ip, PORT))
    print("  数据后台：  http://%s:%d/admin" % (ip, PORT))
    print("  本机访问：  http://127.0.0.1:%d/" % PORT)
    if ADMIN_PWD:
        print("  后台密码：  已启用（网址后加 ?pwd=你的密码）")
    print("=" * 60)
    print("  按 Ctrl+C 停止服务")
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")


if __name__ == "__main__":
    main()
