"""
server.py — локальный веб-сервер для Market Research
Запуск:
    pip install flask httpx
    python server.py
Открой: http://localhost:5000
"""

import asyncio
import json
import httpx
from flask import Flask, request, jsonify, Response

app = Flask(__name__)

PERPLEXITY_API_KEY = "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"
PERPLEXITY_URL     = "https://api.perplexity.ai/chat/completions"

SYS = """Ты — эксперт по маркетинговым исследованиям. Отвечай на русском языке.
Для российских рынков используй рубли (₽ млрд), для остальных — доллары ($B).
Всегда приводи конкретные цифры, источники и год данных."""

HTML = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Market Research — Perplexity</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f8f9fa;color:#1a1a1a;min-height:100vh}
.container{max-width:860px;margin:0 auto;padding:24px 16px}
.header{display:flex;align-items:center;gap:12px;margin-bottom:24px}
.logo{width:36px;height:36px;border-radius:50%;background:#1D9E75;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#fff;flex-shrink:0}
h1{font-size:20px;font-weight:600;color:#111}
.subtitle{font-size:13px;color:#888;margin-top:2px}
.card{background:#fff;border:1px solid #e8e8e8;border-radius:12px;padding:20px;margin-bottom:16px}
.form-row{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.form-group{display:flex;flex-direction:column;gap:5px;flex:1;min-width:160px}
label{font-size:12px;font-weight:500;color:#555}
input{padding:9px 12px;font-size:14px;border:1px solid #ddd;border-radius:8px;background:#fff;color:#111;transition:border-color .15s}
input:focus{outline:none;border-color:#1D9E75}
.btn{width:100%;padding:11px;background:#1D9E75;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;transition:opacity .15s}
.btn:hover{opacity:.88}
.btn:disabled{opacity:.4;cursor:not-allowed}

.steps{display:flex;align-items:flex-start;gap:0;margin-bottom:20px}
.step{display:flex;flex-direction:column;align-items:center;gap:4px;flex:1}
.step-dot{width:22px;height:22px;border-radius:50%;border:2px solid #ddd;background:#f8f9fa;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:#aaa;transition:all .3s}
.step-dot.active{border-color:#1D9E75;background:#E1F5EE;color:#0F6E56}
.step-dot.done{border-color:#1D9E75;background:#1D9E75;color:#fff}
.step-lbl{font-size:9px;color:#aaa;white-space:nowrap;text-align:center}
.step-lbl.done{color:#1D9E75;font-weight:600}
.step-line{flex:1;height:2px;background:#e8e8e8;margin-top:-12px;transition:background .3s}
.step-line.done{background:#9FE1CB}
@keyframes spin{to{transform:rotate(360deg)}}
.spinner{width:10px;height:10px;border:2px solid #9FE1CB;border-top-color:#1D9E75;border-radius:50%;animation:spin .7s linear infinite}

.status{padding:10px 14px;border-radius:8px;font-size:13px;display:flex;align-items:center;gap:8px;margin-bottom:12px}
.status.loading{background:#E1F5EE;color:#0F6E56}
.status.error{background:#FCEBEB;color:#A32D2D}

.result-block{background:#fff;border:1px solid #e8e8e8;border-radius:10px;margin-bottom:10px;overflow:hidden}
.result-head{display:flex;align-items:center;gap:10px;padding:10px 14px;background:#f8f9fa;border-bottom:1px solid #e8e8e8;cursor:pointer;user-select:none}
.result-num{width:22px;height:22px;border-radius:50%;background:#1D9E75;color:#fff;font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.result-title{font-size:13px;font-weight:600;flex:1}
.result-toggle{font-size:12px;color:#aaa}
.result-body{padding:14px;font-size:13px;line-height:1.7;color:#333;white-space:pre-wrap;word-break:break-word;max-height:400px;overflow-y:auto;display:block}
.result-body.hidden{display:none}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <div class="logo">MR</div>
    <div>
      <h1>Market Research</h1>
      <div class="subtitle">Powered by Perplexity sonar-pro · локальный сервер</div>
    </div>
  </div>

  <div class="card">
    <div class="form-row">
      <div class="form-group" style="flex:3">
        <label>Рынок</label>
        <input id="market" type="text" placeholder="рынок автомобилей в России" value="рынок автомобилей в России">
      </div>
      <div class="form-group" style="max-width:120px">
        <label>Год прогноза</label>
        <input id="year" type="number" value="2030" min="2025" max="2040">
      </div>
    </div>
    <button class="btn" id="btnRun" onclick="run()">Начать исследование →</button>
  </div>

  <div class="card" id="progressCard" style="display:none">
    <div class="steps" id="stepsBar">
      <div class="step"><div class="step-dot" id="d0">1</div><div class="step-lbl" id="l0">Объём</div></div>
      <div class="step-line" id="ln0"></div>
      <div class="step"><div class="step-dot" id="d1">2</div><div class="step-lbl" id="l1">Игроки</div></div>
      <div class="step-line" id="ln1"></div>
      <div class="step"><div class="step-dot" id="d2">3</div><div class="step-lbl" id="l2">Прогноз</div></div>
      <div class="step-line" id="ln2"></div>
      <div class="step"><div class="step-dot" id="d3">4</div><div class="step-lbl" id="l3">Проблемы</div></div>
      <div class="step-line" id="ln3"></div>
      <div class="step"><div class="step-dot" id="d4">5</div><div class="step-lbl" id="l4">Отчёт</div></div>
    </div>
    <div id="statusBox" class="status loading" style="display:none"></div>
  </div>

  <div id="results"></div>
</div>

<script>
const TITLES = ["Объём рынка","Топ игроки","Прогноз роста","Проблемы отрасли","Итоговый отчёт"];
let running = false;

function setStep(i, state) {
  const dot = document.getElementById('d'+i);
  const lbl = document.getElementById('l'+i);
  dot.className = 'step-dot ' + state;
  lbl.className = 'step-lbl ' + (state==='done'?'done':'');
  if (state==='active') dot.innerHTML = '<div class="spinner"></div>';
  else if (state==='done') { dot.textContent='✓'; if(i<4) document.getElementById('ln'+i).className='step-line done'; }
  else dot.textContent = i+1;
}

function showStatus(msg, isErr) {
  const el = document.getElementById('statusBox');
  el.style.display='flex'; el.className='status '+(isErr?'error':'loading');
  el.innerHTML = isErr ? '⚠ '+msg : '<div class="spinner"></div><span>'+msg+'</span>';
}

function addResult(i, content) {
  const div = document.createElement('div'); div.className='result-block';
  div.innerHTML = `<div class="result-head" onclick="tog(${i})">
    <div class="result-num">${i+1}</div>
    <div class="result-title">${TITLES[i]}</div>
    <div class="result-toggle" id="t${i}">▲</div>
  </div><div class="result-body" id="b${i}">${content.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>`;
  document.getElementById('results').appendChild(div);
}

function tog(i) {
  const b=document.getElementById('b'+i), t=document.getElementById('t'+i);
  b.classList.toggle('hidden'); t.textContent=b.classList.contains('hidden')?'▼':'▲';
}

async function run() {
  if (running) return;
  const market = document.getElementById('market').value.trim() || 'рынок автомобилей в России';
  const year   = document.getElementById('year').value || '2030';
  running=true;
  document.getElementById('btnRun').disabled=true;
  document.getElementById('results').innerHTML='';
  document.getElementById('progressCard').style.display='block';
  for(let i=0;i<5;i++) setStep(i,'');

  const isRu = /россия|russia|рф|российск|снг/i.test(market);
  const cur  = isRu ? '₽ млрд' : '$B';
  const geo  = isRu ? 'Россия' : 'Global';

  const prompts = [
    `Проведи анализ объёма рынка: ${market} (${geo}).\n\n## Текущий объём рынка\nТAM, SAM с источниками и годом. Диапазон оценок. Валюта: ${cur}.\n\n## Confidence Level\nHigh/Medium/Low — почему?\n\n## Топ источники\n3-5 авторитетных источников с годом.`,
    `Топ-10 компаний рынка: ${market} (${geo}).\nДля каждой: место, название, доля рынка %, сегмент, регион, ключевое преимущество.\nДанные 2023-2025. В конце — инсайт по структуре конкуренции.`,
    `Прогноз роста рынка: ${market} (${geo}) до ${year}.\nТри сценария с CAGR и итоговым размером рынка:\n- Пессимистичный: CAGR X%, Y ${cur}. Причины.\n- Базовый: CAGR X%, Y ${cur}.\n- Оптимистичный: CAGR X%, Y ${cur}.\nТоп-5 драйверов и топ-3 риска.`,
    `Ключевые проблемы рынка: ${market} (${geo}).\nМинимум 8 проблем по категориям: Технологические, Регуляторные, Рыночные, Операционные, Финансовые.\nДля каждой: описание, частота (High/Medium/Low), источники.\nВ конце: возможности для новых игроков.`,
    `Итоговый отчёт по рынку: ${market} (${geo}), прогноз до ${year}.\n\n## TL;DR\n3-4 ключевых вывода.\n\n## Обзор рынка\n## Размер рынка (${cur})\n## Конкурентный ландшафт\n## Прогноз роста\n## Ключевые проблемы\n## Стратегический вывод\nИДТИ / НЕ ИДТИ — с обоснованием и описанием конкретной возможности.`
  ];

  for (let i=0; i<prompts.length; i++) {
    setStep(i,'active');
    showStatus(`Шаг ${i+1}/5 — ${TITLES[i]}...`);
    try {
      const resp = await fetch('/research', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ prompt: prompts[i] })
      });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      setStep(i,'done');
      addResult(i, data.result);
    } catch(e) {
      setStep(i,'');
      showStatus('Ошибка на шаге '+(i+1)+': '+e.message, true);
      running=false; document.getElementById('btnRun').disabled=false; return;
    }
  }

  document.getElementById('statusBox').style.display='none';
  running=false;
  document.getElementById('btnRun').disabled=false;
  document.getElementById('btnRun').textContent='Новое исследование →';
}
</script>
</body>
</html>"""


@app.route("/")
def index():
    return HTML


@app.route("/research", methods=["POST"])
def research():
    data   = request.json
    prompt = data.get("prompt", "")

    async def call():
        async with httpx.AsyncClient(timeout=60) as client:
            resp = await client.post(
                PERPLEXITY_URL,
                headers={
                    "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": "sonar-pro",
                    "messages": [
                        {"role": "system", "content": SYS},
                        {"role": "user",   "content": prompt},
                    ],
                    "max_tokens": 2000,
                    "temperature": 0.2,
                },
            )
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"]

    try:
        result = asyncio.run(call())
        return jsonify({"result": result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  Market Research Server → http://localhost:{port}\n")
    app.run(debug=False, host="0.0.0.0", port=port)
