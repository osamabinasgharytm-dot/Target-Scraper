"""Generate a self-contained HTML dashboard from exported workbook sheets."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger("scrapper")


def _json_records(rows: list[dict]) -> str:
    """Serialize row data safely for embedding in an HTML script element."""
    return (
        pd.DataFrame(rows)
        .to_json(orient="records", force_ascii=False, date_format="iso")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
    )


def write_dashboard(
    product_rows: list[dict],
    review_rows: list[dict],
    path: Path,
) -> Path:
    """Write a standalone dashboard using the Products and Reviews sheet data."""
    path.parent.mkdir(parents=True, exist_ok=True)
    products_json = _json_records(product_rows)
    reviews_json = _json_records(review_rows)
    html = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Target Product Insights</title>
  <style>
    :root{color-scheme:light;--page:#f4f6f3;--surface:#fff;--surface-muted:#f8faf8;--text:#17211d;--muted:#758079;--border:#e8ece8;--line:#edf0ed;--control-border:#dce2dc;--pill:#eff2ef;--chart-text:#536059;--chart-track:#edf1ed;--shadow:#18271c08;font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--text);background:var(--page);font-synthesis:none}
    :root[data-theme="dark"]{color-scheme:dark;--page:#151a17;--surface:#202723;--surface-muted:#272f2a;--text:#edf3ee;--muted:#a0aca4;--border:#343e37;--line:#343e37;--control-border:#465249;--pill:#343e37;--chart-text:#c0cbc2;--chart-track:#3a453d;--shadow:#0004}
    *{box-sizing:border-box}body{margin:0;background:var(--page);color:var(--text);transition:background-color .2s,color .2s}.shell{max-width:1440px;margin:auto;padding:32px}
    header{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:28px}
    h1{font-size:28px;margin:0 0 5px;letter-spacing:-.5px}.subtle,.muted{color:var(--muted);font-size:13px}
    .header-actions{display:flex;align-items:center;gap:10px}.theme-toggle{border:1px solid var(--control-border);border-radius:20px;padding:8px 13px;background:var(--surface);color:var(--text);font:inherit;font-size:12px;font-weight:650;cursor:pointer}.theme-toggle:hover{background:var(--surface-muted)}.theme-toggle:focus-visible{outline:2px solid #31865a;outline-offset:2px}
    .badge{border-radius:20px;padding:8px 13px;background:#e5f3e9;color:#237345;font-size:12px;font-weight:700}
    .cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:18px}
    .card,.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:0 3px 12px var(--shadow)}
    .card{padding:19px}.label{font-size:12px;color:var(--muted);font-weight:600}.value{font-size:29px;font-weight:750;margin-top:10px;letter-spacing:-.5px}
    .charts{display:grid;grid-template-columns:1fr 1.5fr;gap:16px;margin-bottom:18px}
    .panel{padding:20px;margin-bottom:18px}.charts .panel{margin:0;min-height:250px}
    h2{font-size:16px;margin:0 0 18px}canvas{width:100%;height:190px}
    .section-heading{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:16px}
    input,select{border:1px solid var(--control-border);border-radius:8px;padding:10px 12px;font:inherit;color:var(--text);background:var(--surface)}
    input{min-width:220px}.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
    th{text-align:left;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.05em;background:var(--surface-muted)}
    td,th{padding:12px 13px;border-bottom:1px solid var(--line)}tbody tr:hover{background:var(--surface-muted)}
    .pill{border-radius:20px;padding:4px 8px;background:var(--pill);font-size:11px}
    .positive{color:#237345}.negative{color:#b44940}.empty{text-align:center;padding:24px;color:#758079}
    .pagination{display:flex;align-items:center;justify-content:flex-end;gap:12px;margin-top:14px}.pagination button{border:1px solid var(--control-border);border-radius:8px;padding:8px 12px;background:var(--surface);color:var(--text);font:inherit;cursor:pointer}.pagination button:disabled{opacity:.45;cursor:not-allowed}.pagination button:not(:disabled):hover{background:var(--surface-muted)}
    @media(max-width:850px){.shell{padding:20px}.cards{grid-template-columns:repeat(2,minmax(0,1fr))}.charts{grid-template-columns:1fr}}
    @media(max-width:520px){.shell{padding:14px}header,.section-heading{align-items:flex-start;flex-direction:column}.cards{gap:10px}.value{font-size:24px}input{width:100%}}
  </style>
</head>
<body>
  <main class="shell">
    <header>
      <div><h1>Target Product Insights</h1><div class="subtle">Product performance and customer voice from your latest export</div></div>
      <div class="header-actions"><button class="theme-toggle" id="theme-toggle" type="button" aria-label="Switch to dark mode" aria-pressed="false">Dark mode</button><span class="badge" id="export-date"></span></div>
    </header>
    <section class="cards" aria-label="Summary">
      <article class="card"><div class="label">Products</div><div class="value" id="product-count">—</div></article>
      <article class="card"><div class="label">Customer reviews</div><div class="value" id="review-count">—</div></article>
      <article class="card"><div class="label">Average rating</div><div class="value" id="average-rating">—</div></article>
      <article class="card"><div class="label">Positive reviews</div><div class="value" id="positive-share">—</div></article>
    </section>
    <section class="charts">
      <article class="panel"><h2>Review sentiment</h2><canvas id="sentiment-chart" aria-label="Review sentiment chart"></canvas></article>
      <article class="panel"><h2>Highest-rated products</h2><canvas id="rating-chart" aria-label="Highest rated products chart"></canvas></article>
    </section>
    <section class="panel">
      <div class="section-heading"><div><h2 style="margin:0 0 4px">Products</h2><span class="muted" id="product-caption"></span></div><input id="product-search" type="search" placeholder="Search products, brands, or DPCI" aria-label="Search products"></div>
      <div class="table-wrap"><table><thead><tr><th>DPCI</th><th>Product</th><th>Brand</th><th>Rating</th><th>Positive</th><th>Satisfaction</th><th>Top remark</th><th>Library status</th><th>Website status</th></tr></thead><tbody id="products-body"></tbody></table></div>
    </section>
    <section class="panel">
      <div class="section-heading"><div><h2 style="margin:0 0 4px">Customer reviews</h2><span class="muted" id="review-caption"></span></div><div><select id="sentiment-filter" aria-label="Filter by sentiment"><option value="">All sentiment</option><option>Positive</option><option>Neutral</option><option>Negative</option></select> <input id="review-search" type="search" placeholder="Search reviews or DPCI" aria-label="Search reviews"></div></div>
      <div class="table-wrap"><table><thead><tr><th>DPCI</th><th>Rating</th><th>Sentiment</th><th>Review</th><th>Date</th><th>Defect</th></tr></thead><tbody id="reviews-body"></tbody></table></div>
      <nav class="pagination" aria-label="Customer review pages"><button id="reviews-previous" type="button" disabled>Previous</button><span id="reviews-page-info" class="muted" aria-live="polite"></span><button id="reviews-next" type="button" disabled>Next</button></nav>
    </section>
    <footer class="subtle">Generated from the Products and Reviews sheets. Re-run the scraper to refresh this dashboard.</footer>
  </main>
  <script>
    const products = __PRODUCTS__;
    const reviews = __REVIEWS__;
    const reviewsPerPage = 25;
    let reviewPage = 1;
    const themeToggle = document.getElementById("theme-toggle");
    function setTheme(theme) {
      const isDark = theme === "dark";
      document.documentElement.dataset.theme = isDark ? "dark" : "light";
      themeToggle.textContent = isDark ? "Light mode" : "Dark mode";
      themeToggle.setAttribute("aria-label", `Switch to ${isDark ? "light" : "dark"} mode`);
      themeToggle.setAttribute("aria-pressed", String(isDark));
      drawSentiment();
      drawRatings();
    }
    let initialTheme = "light";
    try {
      const savedTheme = localStorage.getItem("target-dashboard-theme");
      if (savedTheme === "dark" || savedTheme === "light") initialTheme = savedTheme;
    } catch (error) {
      console.warn("Could not read the saved dashboard theme.", error);
    }
    const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
    const text = value => value == null || value === "nan" ? "" : String(value);
    const safe = value => text(value).replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
    const average = values => values.length ? values.reduce((a,b) => a+b,0) / values.length : null;
    const rated = products.map(p => number(p["Avg. Rating"])).filter(n => n > 0);
    const sentimentCounts = reviews.reduce((counts, review) => {
      const key = text(review.Sentiment) || "Neutral"; counts[key] = (counts[key] || 0) + 1; return counts;
    }, {Positive:0,Neutral:0,Negative:0});
    setTheme(initialTheme);
    themeToggle.addEventListener("click", () => {
      const theme = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
      setTheme(theme);
      try {
        localStorage.setItem("target-dashboard-theme", theme);
      } catch (error) {
        console.warn("Could not save the dashboard theme.", error);
      }
    });
    const overallPositive = reviews.length ? sentimentCounts.Positive / reviews.length * 100 :
      (products.length ? average(products.map(p => number(p["Positive %"]))) : null);
    document.getElementById("product-count").textContent = products.length.toLocaleString();
    document.getElementById("review-count").textContent = reviews.length.toLocaleString();
    document.getElementById("average-rating").textContent = average(rated) == null ? "—" : average(rated).toFixed(1) + " / 5";
    document.getElementById("positive-share").textContent = overallPositive == null ? "—" : overallPositive.toFixed(1) + "%";
    document.getElementById("export-date").textContent = "Exported " + new Date().toLocaleDateString(undefined,{year:"numeric",month:"short",day:"numeric"});
    document.getElementById("product-caption").textContent = products.length.toLocaleString() + " products";
    document.getElementById("review-caption").textContent = reviews.length.toLocaleString() + " reviews";

    function setupCanvas(canvas) {
      const ratio = window.devicePixelRatio || 1, bounds = canvas.getBoundingClientRect();
      canvas.width = bounds.width * ratio; canvas.height = bounds.height * ratio;
      const ctx = canvas.getContext("2d"); ctx.scale(ratio,ratio); return {ctx,width:bounds.width,height:bounds.height};
    }
    function drawSentiment() {
      const {ctx,width,height} = setupCanvas(document.getElementById("sentiment-chart"));
      const themeStyles = getComputedStyle(document.documentElement);
      const chartText = themeStyles.getPropertyValue("--chart-text").trim();
      const total = Object.values(sentimentCounts).reduce((a,b) => a+b,0);
      if (!total) {ctx.fillStyle="#758079";ctx.font="14px sans-serif";ctx.fillText("No review sentiment data",12,30);return;}
      const colors={Positive:"#31865a",Neutral:"#e2ad46",Negative:"#d6675d"}, cx=Math.min(width*.34,100), cy=height/2, radius=Math.min(74,height*.38);
      let start=-Math.PI/2;
      Object.entries(sentimentCounts).forEach(([label,count]) => {
        if (!count) return; const end=start+count/total*Math.PI*2;
        ctx.beginPath();ctx.arc(cx,cy,radius,start,end);ctx.arc(cx,cy,radius*.62,end,start,true);ctx.closePath();ctx.fillStyle=colors[label]||"#758079";ctx.fill();start=end;
      });
      ctx.textAlign="center";ctx.fillStyle=themeStyles.getPropertyValue("--text").trim();ctx.font="700 19px sans-serif";ctx.fillText(total,cx,cy+5);
      ctx.textAlign="left";ctx.font="12px sans-serif";
      Object.entries(sentimentCounts).forEach(([label,count],i) => {
        const y=height/2-28+i*25;ctx.fillStyle=colors[label]||"#758079";ctx.fillRect(Math.min(width*.62,190),y-9,9,9);
        ctx.fillStyle=chartText;ctx.fillText(`${label}  ${count} (${(count/total*100).toFixed(1)}%)`,Math.min(width*.62,205),y);
      });
    }
    function drawRatings() {
      const {ctx,width,height}=setupCanvas(document.getElementById("rating-chart"));
      const themeStyles = getComputedStyle(document.documentElement);
      const top=products.filter(p=>number(p["Avg. Rating"])>0).sort((a,b)=>number(b["Avg. Rating"])-number(a["Avg. Rating"])).slice(0,5);
      if (!top.length) {ctx.fillStyle="#758079";ctx.font="14px sans-serif";ctx.fillText("No product ratings available",12,30);return;}
      const left=Math.min(220,width*.52), row=height/top.length, barWidth=Math.max(width-left-55,20);
      top.forEach((p,i)=>{const y=i*row+row*.28,label=text(p["Item Description"]||p.DPCI);ctx.fillStyle=themeStyles.getPropertyValue("--chart-text").trim();ctx.font="12px sans-serif";ctx.textAlign="right";ctx.fillText(label.length>32?label.slice(0,29)+"…":label,left-12,y+12);ctx.fillStyle=themeStyles.getPropertyValue("--chart-track").trim();ctx.fillRect(left,y,barWidth,15);ctx.fillStyle="#31865a";ctx.fillRect(left,y,barWidth*Math.min(number(p["Avg. Rating"])/5,1),15);ctx.textAlign="left";ctx.fillStyle=themeStyles.getPropertyValue("--text").trim();ctx.fillText(number(p["Avg. Rating"]).toFixed(1),left+barWidth+9,y+12);});
    }
    function renderProducts() {
      const query=document.getElementById("product-search").value.toLowerCase();
      const filtered=products.filter(p=>[p.DPCI,p["Item Description"],p.Brand,p.Program,p["Top Remark"]].some(v=>text(v).toLowerCase().includes(query)));
      document.getElementById("products-body").innerHTML=filtered.length?filtered.map(p=>`<tr><td>${safe(p.DPCI)}</td><td>${safe(p["Item Description"])}</td><td>${safe(p.Brand)}</td><td>${number(p["Avg. Rating"]) ? number(p["Avg. Rating"]).toFixed(1) : "—"}</td><td>${number(p["Positive %"]).toFixed(1)}%</td><td>${p["Average Satisfaction %"]==null?"—":number(p["Average Satisfaction %"]).toFixed(1)+"%"}</td><td>${safe(p["Top Remark"])||"—"}</td><td><span class="pill">${safe(p.Status)||"—"}</span></td><td><span class="pill ${p["Website Status"]==="Found"?"positive":p["Website Status"]==="Not Found"?"negative":""}">${safe(p["Website Status"])||"—"}</span></td></tr>`).join(""):'<tr><td class="empty" colspan="9">No matching products</td></tr>';
    }
    function renderReviews() {
      const query=document.getElementById("review-search").value.toLowerCase(), sentiment=document.getElementById("sentiment-filter").value;
      const filtered=reviews.filter(r=>(!sentiment||text(r.Sentiment)===sentiment)&&[r.DPCI,r["Review Title"],r["Review Text"],r.Reason,r.Defect].some(v=>text(v).toLowerCase().includes(query)));
      const pageCount=Math.ceil(filtered.length/reviewsPerPage);
      reviewPage=Math.min(reviewPage,Math.max(pageCount,1));
      const start=(reviewPage-1)*reviewsPerPage;
      document.getElementById("reviews-body").innerHTML=filtered.length?filtered.slice(start,start+reviewsPerPage).map(r=>`<tr><td>${safe(r.DPCI)}</td><td>${r["Review Rating"]==null?"—":safe(r["Review Rating"])+" / 5"}</td><td><span class="pill ${text(r.Sentiment)==="Positive"?"positive":text(r.Sentiment)==="Negative"?"negative":""}">${safe(r.Sentiment)||"—"}</span></td><td title="${safe(r["Review Text"])}">${safe((text(r["Review Text"])||text(r["Review Title"])).slice(0,180))||"—"}</td><td>${safe(r["Date Posted"])||"—"}</td><td>${safe(r.Defect==="-"?"":r.Defect)||safe(r.Reason==="-"?"":r.Reason)||"—"}</td></tr>`).join(""):'<tr><td class="empty" colspan="6">No matching reviews</td></tr>';
      document.getElementById("reviews-page-info").textContent=filtered.length?`Showing ${start+1}–${Math.min(start+reviewsPerPage,filtered.length)} of ${filtered.length} reviews`:"No reviews";
      document.getElementById("reviews-previous").disabled=reviewPage<=1;
      document.getElementById("reviews-next").disabled=reviewPage>=pageCount;
    }
    document.getElementById("product-search").addEventListener("input",renderProducts);
    document.getElementById("review-search").addEventListener("input",()=>{reviewPage=1;renderReviews();});
    document.getElementById("sentiment-filter").addEventListener("change",()=>{reviewPage=1;renderReviews();});
    document.getElementById("reviews-previous").addEventListener("click",()=>{if(reviewPage>1){reviewPage--;renderReviews();}});
    document.getElementById("reviews-next").addEventListener("click",()=>{reviewPage++;renderReviews();});
    window.addEventListener("resize",()=>{drawSentiment();drawRatings();});
    drawSentiment();drawRatings();renderProducts();renderReviews();
  </script>
</body>
</html>
"""
    html = html.replace("__PRODUCTS__", products_json).replace("__REVIEWS__", reviews_json)
    path.write_text(html, encoding="utf-8")
    logger.info("Wrote interactive dashboard -> %s", path)
    return path
