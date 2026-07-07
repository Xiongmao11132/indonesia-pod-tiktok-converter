import { buildTikTokWorkbook, parseBigSellerFile } from "./converter.js";

const sourceInput = document.querySelector("#sourceFile");
const generateButton = document.querySelector("#generateButton");
const productsEl = document.querySelector("#products");
const summaryEl = document.querySelector("#summary");
const statusEl = document.querySelector("#status");
const downloadsEl = document.querySelector("#downloads");
const xlsxDownload = document.querySelector("#xlsxDownload");
const reviewDownload = document.querySelector("#reviewDownload");

let sourceState = null;
let choices = {};
let currentUrls = [];

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b42318" : "#667085";
}

function setDownload(link, blob, filename) {
  const url = URL.createObjectURL(blob);
  currentUrls.push(url);
  link.href = url;
  link.download = filename;
}

function clearDownloads() {
  currentUrls.forEach((url) => URL.revokeObjectURL(url));
  currentUrls = [];
  downloadsEl.hidden = true;
  xlsxDownload.removeAttribute("href");
  reviewDownload.removeAttribute("href");
}

function renderProducts(groups) {
  productsEl.innerHTML = "";
  choices = {};

  for (const group of groups) {
    choices[group.productNo] = {
      side: "P",
      color: group.inferredColor === "WH" ? "WH" : "BK",
    };

    const card = document.createElement("article");
    card.className = "product-card";
    card.innerHTML = `
      <img class="cover" src="${group.image}" alt="">
      <div>
        <h2 class="product-title">${group.productNo}. ${escapeHtml(group.title)}</h2>
        <div class="meta">
          <span>${group.rows.length} 个源 SKU</span>
          <span>${escapeHtml(group.sourceUrl || "无来源链接")}</span>
        </div>
        <div class="thumbs">
          ${group.thumbs.slice(0, 8).map((url) => `<img src="${url}" alt="">`).join("")}
        </div>
      </div>
      <div class="controls">
        <div>
          <div class="field-label">印花面</div>
          <div class="segmented">
            <label><input type="radio" name="side-${group.productNo}" value="P" checked> P</label>
            <label><input type="radio" name="side-${group.productNo}" value="PR"> PR</label>
          </div>
        </div>
        <div>
          <div class="field-label">无颜色选项时默认底色</div>
          <select data-color="${group.productNo}">
            <option value="BK" ${choices[group.productNo].color === "BK" ? "selected" : ""}>BK 黑色</option>
            <option value="WH" ${choices[group.productNo].color === "WH" ? "selected" : ""}>WH 白色</option>
          </select>
        </div>
      </div>
    `;

    card.querySelectorAll(`input[name="side-${group.productNo}"]`).forEach((radio) => {
      radio.addEventListener("change", () => {
        choices[group.productNo].side = radio.value;
      });
    });
    card.querySelector(`[data-color="${group.productNo}"]`).addEventListener("change", (event) => {
      choices[group.productNo].color = event.target.value;
    });
    productsEl.append(card);
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

sourceInput.addEventListener("change", async () => {
  const [file] = sourceInput.files || [];
  clearDownloads();
  productsEl.innerHTML = "";
  generateButton.disabled = true;
  sourceState = null;

  if (!file) return;
  try {
    setStatus("正在读取 Excel...");
    sourceState = await parseBigSellerFile(file);
    renderProducts(sourceState.groups);
    summaryEl.textContent = `已读取 ${sourceState.groups.length} 个商品，SKU 日期 ${sourceState.skuDate}`;
    generateButton.disabled = false;
    setStatus("请确认每个商品的 P/PR。");
  } catch (error) {
    setStatus(error.message || String(error), true);
  }
});

generateButton.addEventListener("click", async () => {
  if (!sourceState) return;
  clearDownloads();
  generateButton.disabled = true;
  try {
    setStatus("正在生成上传表...");
    const result = await buildTikTokWorkbook({
      groups: sourceState.groups,
      choices,
      sourceFilename: sourceState.filename,
    });
    setDownload(xlsxDownload, result.xlsxBlob, result.xlsxName);
    setDownload(reviewDownload, result.reviewBlob, result.reviewName);
    downloadsEl.hidden = false;
    setStatus(`已生成 ${result.outputRows.length} 行 SKU。`);
  } catch (error) {
    setStatus(error.message || String(error), true);
  } finally {
    generateButton.disabled = false;
  }
});
