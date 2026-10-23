"""Generate a self-contained HTML dashboard from exported workbook sheets."""

from __future__ import annotations

import logging
from datetime import date
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


def _latest_product_rows(product_rows: list[dict]) -> list[dict]:
    """Keep the most recently appended product row for each DPCI."""
    latest_rows: list[dict] = []
    row_indexes: dict[str, int] = {}
    for row in product_rows:
        dpci = row.get("DPCI")
        dpci_key = "" if dpci is None or pd.isna(dpci) else str(dpci).strip()
        if not dpci_key:
            latest_rows.append(row)
        elif dpci_key in row_indexes:
            latest_rows[row_indexes[dpci_key]] = row
        else:
            row_indexes[dpci_key] = len(latest_rows)
            latest_rows.append(row)
    return latest_rows


def write_dashboard(
    product_rows: list[dict],
    review_rows: list[dict],
    path: Path,
) -> Path:
    """Write a standalone dashboard using the Products and Reviews sheet data."""
    path.parent.mkdir(parents=True, exist_ok=True)
    products_json = _json_records(_latest_product_rows(product_rows))
    reviews_json = _json_records(review_rows)
    html = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Target Product Insights | __CURRENT_MONTH__</title>
  <style>
    :root{color-scheme:light;--page:#f4f6f3;--surface:#fff;--surface-muted:#f8faf8;--text:#17211d;--muted:#758079;--border:#e8ece8;--line:#edf0ed;--control-border:#dce2dc;--pill:#eff2ef;--chart-text:#536059;--chart-track:#edf1ed;--shadow:#18271c08;font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--text);background:var(--page);font-synthesis:none}
    :root[data-theme="dark"]{color-scheme:dark;--page:#151a17;--surface:#202723;--surface-muted:#272f2a;--text:#edf3ee;--muted:#a0aca4;--border:#343e37;--line:#343e37;--control-border:#465249;--pill:#343e37;--chart-text:#c0cbc2;--chart-track:#3a453d;--shadow:#0004}
    *{box-sizing:border-box}body{margin:0;background:var(--page);color:var(--text);transition:background-color .2s,color .2s}.shell{max-width:1920px;margin:auto;padding:32px}
    header{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:28px}
    h1{font-size:28px;margin:0 0 5px;letter-spacing:-.5px}.subtle,.muted{color:var(--muted);font-size:13px; text-align:center}
    .header-actions{display:flex;align-items:center;gap:10px}.theme-toggle{border:1px solid var(--control-border);border-radius:20px;padding:8px 13px;background:var(--surface);color:var(--text);font:inherit;font-size:12px;font-weight:650;cursor:pointer}.theme-toggle:hover{background:var(--surface-muted)}.theme-toggle:focus-visible{outline:2px solid #2ED003;outline-offset:2px}
    .badge{border-radius:20px;padding:8px 13px;background:#e5f3e9;color:#2ED003;font-size:12px;font-weight:700}
    .cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:18px}
    .card,.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:0 3px 12px var(--shadow)}
    .card{padding:19px}.label{font-size:12px;color:var(--muted);font-weight:600}.value{font-size:29px;font-weight:750;margin-top:10px;letter-spacing:-.5px}
    .charts{display:grid;grid-template-columns:1fr 1.5fr;gap:16px;margin-bottom:18px}
    .panel{padding:20px;margin-bottom:18px}.charts .panel{margin:0;min-height:250px}
    h2{font-size:16px;margin:0 0 18px}canvas{width:100%;height:190px}
    .section-heading{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:16px}
    .product-controls{display:flex;align-items:center;justify-content:flex-end;gap:10px;flex-wrap:wrap}
    input,select{border:1px solid var(--control-border);border-radius:8px;padding:10px 12px;font:inherit;color:var(--text);background:var(--surface)}
    input{min-width:220px}.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
    th{text-align:left;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.05em;background:var(--surface-muted)}
    td,th{padding:12px 13px;border-bottom:1px solid var(--line)}tbody tr:hover{background:var(--surface-muted)}
    .pill{border-radius:20px;padding:4px 8px;background:var(--pill);font-size:11px}
    .positive{color:#2ED003}.negative{color:#F7090A}.empty{text-align:center;padding:24px;color:#758079}
    .pagination{display:flex;align-items:center;justify-content:center;gap:10px;margin-top:14px;flex-wrap:wrap}.pagination button{border:1px solid var(--control-border);border-radius:8px;padding:8px 12px;background:var(--surface);color:var(--text);font:inherit;cursor:pointer}.pagination button:disabled{opacity:.45;cursor:not-allowed}.pagination button:not(:disabled):hover{background:var(--surface-muted)}.page-numbers{display:flex;align-items:center;gap:5px}.pagination .page-number{min-width:34px;padding:7px 9px}.pagination .page-number[aria-current="page"]{background:#2ED003;border-color:#2ED003;color:#fff;font-weight:700}.page-ellipsis{padding:0 2px;color:var(--muted)}
    @media(max-width:850px){.shell{padding:20px}.cards{grid-template-columns:repeat(2,minmax(0,1fr))}.charts{grid-template-columns:1fr}}
    @media(max-width:520px){.shell{padding:14px}header,.section-heading{align-items:flex-start;flex-direction:column}.cards{gap:10px}.value{font-size:24px}input{width:100%}.product-controls{width:100%;align-items:stretch;flex-direction:column}.product-controls select{width:100%}}
  </style>
</head>
<body>
  <main class="shell">
    <header>
      <div><h1>Target Product Insights | __CURRENT_MONTH__</h1><div class="subtle">Product performance and customer voice from your latest export</div></div>
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
      <div class="section-heading"><div><h2 style="margin:0 0 4px">Products</h2><span class="muted" id="product-caption"></span></div><div class="product-controls"><select id="library-status-filter" aria-label="Filter products by library status"><option value="">All library statuses</option></select><select id="website-status-filter" aria-label="Filter products by website status"><option value="">All website statuses</option></select><input id="product-search" type="search" placeholder="Search products, brands, or DPCI" aria-label="Search products"></div></div>
      <div class="table-wrap"><table><thead><tr><th>DPCI</th><th>Product</th><th>Brand</th><th>Rating</th><th>Positive</th><th>Satisfaction</th><th>Top remark</th><th>Library status</th><th>Website status</th></tr></thead><tbody id="products-body"></tbody></table></div>
      <nav class="pagination" aria-label="Product pages"><button id="products-previous" type="button" disabled>Previous</button><span class="page-numbers" id="products-page-numbers"></span><span id="products-page-info" class="muted" aria-live="polite"></span><button id="products-next" type="button" disabled>Next</button></nav>
    </section>
    <section class="panel">
      <div class="section-heading"><div><h2 style="margin:0 0 4px">Customer reviews</h2><span class="muted" id="review-caption"></span></div><div><select id="sentiment-filter" aria-label="Filter by sentiment"><option value="">All sentiment</option><option>Positive</option><option>Neutral</option><option>Negative</option></select> <select id="defect-filter" aria-label="Filter by defect"><option value="">All defects</option></select> <input id="review-search" type="search" placeholder="Search reviews or DPCI" aria-label="Search reviews"></div></div>
      <div class="table-wrap"><table><thead><tr><th>DPCI</th><th>Rating</th><th>Sentiment</th><th>Review</th><th>Date</th><th>Defect</th></tr></thead><tbody id="reviews-body"></tbody></table></div>
      <nav class="pagination" aria-label="Customer review pages"><button id="reviews-previous" type="button" disabled>Previous</button><span class="page-numbers" id="reviews-page-numbers"></span><span id="reviews-page-info" class="muted" aria-live="polite"></span><button id="reviews-next" type="button" disabled>Next</button></nav>
    </section>
    <footer class="subtle">Generated from the Products and Reviews sheets. Re-run the scraper to refresh this dashboard.</footer>
  </main>
  <script>
    const products = __PRODUCTS__;
    const reviews = __REVIEWS__;
    const productsPerPage = 15;
    const reviewsPerPage = 15;
    let productPage = 1;
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
      const colors={Positive:"#2ED003",Neutral:"#F7DC00",Negative:"#F7090A"}, cx=Math.min(width*.34,100), cy=height/2, radius=Math.min(74,height*.38);
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
      top.forEach((p,i)=>{const y=i*row+row*.28,label=text(p["Item Description"]||p.DPCI);ctx.fillStyle=themeStyles.getPropertyValue("--chart-text").trim();ctx.font="12px sans-serif";ctx.textAlign="right";ctx.fillText(label.length>32?label.slice(0,29)+"…":label,left-12,y+12);ctx.fillStyle=themeStyles.getPropertyValue("--chart-track").trim();ctx.fillRect(left,y,barWidth,15);ctx.fillStyle="#2ED003";ctx.fillRect(left,y,barWidth*Math.min(number(p["Avg. Rating"])/5,1),15);ctx.textAlign="left";ctx.fillStyle=themeStyles.getPropertyValue("--text").trim();ctx.fillText(number(p["Avg. Rating"]).toFixed(1),left+barWidth+9,y+12);});
    }
    function renderPageNumbers(containerId,currentPage,pageCount) {
      const container=document.getElementById(containerId);
      if (!pageCount) {container.innerHTML="";return;}
      const pages=new Set([1,pageCount]);
      for(let page=Math.max(1,currentPage-2);page<=Math.min(pageCount,currentPage+2);page++) pages.add(page);
      let previousPage=0;
      container.innerHTML=[...pages].sort((a,b)=>a-b).map(page=>{
        const gap=page-previousPage>1?'<span class="page-ellipsis" aria-hidden="true">…</span>':"";
        previousPage=page;
        return `${gap}<button class="page-number" type="button" data-page="${page}" aria-label="Page ${page}" ${page===currentPage?'aria-current="page"':""}>${page}</button>`;
      }).join("");
    }
    function renderProducts() {
      const query=document.getElementById("product-search").value.toLowerCase();
      const libraryStatus=document.getElementById("library-status-filter").value;
      const websiteStatus=document.getElementById("website-status-filter").value;
      const filtered=products.filter(p=>
        (!libraryStatus||text(p.Status)===libraryStatus)&&
        (!websiteStatus||text(p["Website Status"])===websiteStatus)&&
        [p.DPCI,p["Item Description"],p.Brand,p.Program,p["Top Remark"]].some(v=>text(v).toLowerCase().includes(query))
      );
      const pageCount=Math.ceil(filtered.length/productsPerPage);
      productPage=Math.min(productPage,Math.max(pageCount,1));
      const start=(productPage-1)*productsPerPage;
      document.getElementById("products-body").innerHTML=filtered.length?filtered.slice(start,start+productsPerPage).map(p=>`<tr><td>${safe(p.DPCI)}</td><td>${safe(p["Item Description"])}</td><td>${safe(p.Brand)}</td><td>${number(p["Avg. Rating"]) ? number(p["Avg. Rating"]).toFixed(1) : "—"}</td><td>${number(p["Positive %"]).toFixed(1)}%</td><td>${p["Average Satisfaction %"]==null?"—":number(p["Average Satisfaction %"]).toFixed(1)+"%"}</td><td>${safe(p["Top Remark"])||"—"}</td><td><span class="pill">${safe(p.Status)||"—"}</span></td><td><span class="pill ${p["Website Status"]==="Found"?"positive":p["Website Status"]==="Not Found"?"negative":""}">${safe(p["Website Status"])||"—"}</span></td></tr>`).join(""):'<tr><td class="empty" colspan="9">No matching products</td></tr>';
      document.getElementById("products-page-info").textContent=filtered.length?`Showing ${start+1}–${Math.min(start+productsPerPage,filtered.length)} of ${filtered.length} products`:"No products";
      document.getElementById("products-previous").disabled=productPage<=1;
      document.getElementById("products-next").disabled=productPage>=pageCount;
      renderPageNumbers("products-page-numbers",productPage,pageCount);
    }
    function populateStatusFilter(id, column, label, defaultValue="") {
      const select=document.getElementById(id);
      const values=[...new Set(products.map(product=>text(product[column]).trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
      select.innerHTML=`<option value="">All ${label}</option>`+values.map(value=>`<option value="${safe(value)}">${safe(value)}</option>`).join("");
      if (values.includes(defaultValue)) select.value=defaultValue;
    }
    function reviewDefect(review) {
      const defect=text(review.Defect).trim();
      if (defect && defect!=="-") return defect;
      const reason=text(review.Reason).trim();
      return reason==="-" ? "" : reason;
    }
    function populateDefectFilter() {
      const select=document.getElementById("defect-filter");
      const defects=[...new Set(reviews.map(reviewDefect).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
      select.innerHTML='<option value="">All defects</option>'+defects.map(defect=>`<option value="${safe(defect)}">${safe(defect)}</option>`).join("");
    }
    function renderReviews() {
      const query=document.getElementById("review-search").value.toLowerCase(), sentiment=document.getElementById("sentiment-filter").value, defect=document.getElementById("defect-filter").value;
      const filtered=reviews.filter(r=>(!sentiment||text(r.Sentiment)===sentiment)&&(!defect||reviewDefect(r)===defect)&&[r.DPCI,r["Review Title"],r["Review Text"],r.Reason,r.Defect].some(v=>text(v).toLowerCase().includes(query)));
      const pageCount=Math.ceil(filtered.length/reviewsPerPage);
      reviewPage=Math.min(reviewPage,Math.max(pageCount,1));
      const start=(reviewPage-1)*reviewsPerPage;
      document.getElementById("reviews-body").innerHTML=filtered.length?filtered.slice(start,start+reviewsPerPage).map(r=>`<tr><td>${safe(r.DPCI)}</td><td>${r["Review Rating"]==null?"—":safe(r["Review Rating"])+" / 5"}</td><td><span class="pill ${text(r.Sentiment)==="Positive"?"positive":text(r.Sentiment)==="Negative"?"negative":""}">${safe(r.Sentiment)||"—"}</span></td><td title="${safe(r["Review Text"])}">${safe((text(r["Review Text"])||text(r["Review Title"])).slice(0,180))||"—"}</td><td>${safe(r["Date Posted"])||"—"}</td><td>${safe(reviewDefect(r))||"—"}</td></tr>`).join(""):'<tr><td class="empty" colspan="6">No matching reviews</td></tr>';
      document.getElementById("reviews-page-info").textContent=filtered.length?`Showing ${start+1}–${Math.min(start+reviewsPerPage,filtered.length)} of ${filtered.length} reviews`:"No reviews";
      document.getElementById("reviews-previous").disabled=reviewPage<=1;
      document.getElementById("reviews-next").disabled=reviewPage>=pageCount;
      renderPageNumbers("reviews-page-numbers",reviewPage,pageCount);
    }
    populateStatusFilter("library-status-filter","Status","library statuses");
    populateStatusFilter("website-status-filter","Website Status","website statuses","Found");
    populateDefectFilter();
    document.getElementById("product-search").addEventListener("input",()=>{productPage=1;renderProducts();});
    document.getElementById("library-status-filter").addEventListener("change",()=>{productPage=1;renderProducts();});
    document.getElementById("website-status-filter").addEventListener("change",()=>{productPage=1;renderProducts();});
    document.getElementById("products-previous").addEventListener("click",()=>{if(productPage>1){productPage--;renderProducts();}});
    document.getElementById("products-next").addEventListener("click",()=>{productPage++;renderProducts();});
    document.getElementById("products-page-numbers").addEventListener("click",event=>{
      const button=event.target.closest("button[data-page]");
      if(button){productPage=Number(button.dataset.page);renderProducts();}
    });
    document.getElementById("review-search").addEventListener("input",()=>{reviewPage=1;renderReviews();});
    document.getElementById("sentiment-filter").addEventListener("change",()=>{reviewPage=1;renderReviews();});
    document.getElementById("defect-filter").addEventListener("change",()=>{reviewPage=1;renderReviews();});
    document.getElementById("reviews-previous").addEventListener("click",()=>{if(reviewPage>1){reviewPage--;renderReviews();}});
    document.getElementById("reviews-next").addEventListener("click",()=>{reviewPage++;renderReviews();});
    document.getElementById("reviews-page-numbers").addEventListener("click",event=>{
      const button=event.target.closest("button[data-page]");
      if(button){reviewPage=Number(button.dataset.page);renderReviews();}
    });
    window.addEventListener("resize",()=>{drawSentiment();drawRatings();});
    drawSentiment();drawRatings();renderProducts();renderReviews();
  </script>
</body>
</html>
"""
    html = (
        html.replace("__PRODUCTS__", products_json)
        .replace("__REVIEWS__", reviews_json)
        .replace("__CURRENT_MONTH__", date.today().strftime("%B %Y"))
    )
    path.write_text(html, encoding="utf-8")
    logger.info("Wrote interactive dashboard -> %s", path)
    return path


def write_dashboard_from_excel(workbook_path: Path) -> Path:
    """Generate a dashboard from an existing combined Products/Reviews workbook."""
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    with pd.ExcelFile(workbook_path) as workbook:
        required_sheets = {"Products", "Reviews"}
        missing_sheets = required_sheets.difference(workbook.sheet_names)
        if missing_sheets:
            missing = ", ".join(sorted(missing_sheets))
            raise ValueError(
                f"Excel workbook must contain Products and Reviews sheets; "
                f"missing: {missing}"
            )
        products = pd.read_excel(workbook, sheet_name="Products")
        reviews = pd.read_excel(workbook, sheet_name="Reviews")

    dashboard_path = workbook_path.with_name(
        f"{workbook_path.stem} Dashboard.html"
    )
    return write_dashboard(
        products.to_dict(orient="records"),
        reviews.to_dict(orient="records"),
        dashboard_path,
    )
