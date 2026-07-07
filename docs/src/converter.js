const SOURCE_COLUMNS = {
  name: "产品名称",
  long_description: "长描述",
  short_description: "短描述",
  source_url: "产品来源链接",
  variation_name_1: "变种名称 1",
  variation_value_1: "变种选项 1",
  variation_name_2: "变种名称 2",
  variation_value_2: "变种选项 2",
  price: "价格",
  sale_price: "促销价",
  seller_sku: "SKU",
  variant_image_1: "变种图 1",
};

const TEMPLATE_COLUMNS = {
  category: 1,
  brand: 2,
  product_name: 3,
  product_description: 4,
  main_image: 5,
  image_2: 6,
  image_3: 7,
  image_4: 8,
  image_5: 9,
  image_6: 10,
  image_7: 11,
  image_8: 12,
  image_9: 13,
  property_name_1: 14,
  property_value_1: 15,
  property_1_image: 16,
  property_name_2: 17,
  property_value_2: 18,
  parcel_weight: 19,
  parcel_length: 20,
  parcel_width: 21,
  parcel_height: 22,
  delivery: 23,
  price: 24,
  pre_order_time: 25,
  quantity: 26,
  seller_sku: 27,
  minimum_order_quantity: 28,
  size_chart: 29,
  cod: 30,
  shipping_insurance: 31,
  materials: 32,
  pattern: 33,
  neckline: 34,
  sleeve_length: 35,
  season: 36,
  style: 37,
  fit: 38,
  stretch: 39,
  washing_instructions: 40,
  waist_height: 41,
};

const PRODUCT_PROPERTIES = {
  materials: "100%Sorona",
  pattern: "Graphic",
  neckline: "Round Neck",
  sleeve_length: "Short Sleeve",
  season: "All Seasons",
  style: "Casual",
  fit: "Loose-Fitting",
  stretch: "",
  washing_instructions: "Machine Washable",
  waist_height: "",
};

const COLOR_CODES = {
  WHITE: "WH",
  WH: "WH",
  PUTIH: "WH",
  BLACK: "BK",
  BK: "BK",
  HITAM: "BK",
  RED: "RED",
  MERAH: "RED",
  BLUE: "BL",
  BIRU: "BL",
  GREEN: "GR",
  HIJAU: "GR",
  ORANGE: "OG",
  OREN: "OG",
  APRICOT: "AP",
  CREAM: "AP",
  KREM: "AP",
  GREY: "GY",
  GRAY: "GY",
  ABU: "GY",
};

const COLOR_DISPLAY = {
  BK: "Black",
  WH: "White",
  RED: "Red",
  BL: "Blue",
  GR: "Green",
  OG: "Orange",
  AP: "Apricot",
  GY: "Grey",
};

const SIZE_ALIASES = {
  S: "S",
  M: "M",
  L: "L",
  XL: "XL",
  XXL: "XXL",
  "2XL": "XXL",
  XXXL: "XXXL",
  "3XL": "XXXL",
};

const TARGET_SIZES = ["S", "M", "L", "XL", "XXL", "XXXL"];
const DATA_START_ROW = 6;
const DELIVERY_TEXT = "The delivery options for this product are the same as the delivery options for the shop. ";
const EXTRA_VARIATION_VALUE_1 = "Pengiriman kilat 48 jam";
const EXTRA_VARIATION_VALUE_2 = "Jangan beli";

function text(value) {
  return value == null ? "" : String(value).trim();
}

function truncate(value, limit) {
  const clean = text(value);
  return clean.length <= limit ? clean : clean.slice(0, limit);
}

function normalizeSize(value) {
  const clean = text(value).toUpperCase().replace(/[\s_-]+/g, "");
  return SIZE_ALIASES[clean] || "";
}

function normalizeColorCode(value) {
  const clean = text(value).toUpperCase();
  if (!clean) return "";
  const tokens = clean.split(/[^0-9A-Z]+/).filter(Boolean);
  for (const token of tokens) {
    if (COLOR_CODES[token]) return COLOR_CODES[token];
  }
  const compact = clean.replace(/[^0-9A-Z]+/g, "");
  for (const [token, code] of Object.entries(COLOR_CODES)) {
    if (compact.includes(token)) return code;
  }
  return "";
}

function normalizeMaterialText(value) {
  let output = text(value);
  if (!output) return output;
  output = output.replace(/\bbahan\s+(?:100\s*%\s*)?(?:cotton|katun)(?:\s+(?:combed|premium|pe|20s|24s|30s))*\b/gi, "Bahan 100%Sorona");
  output = output.replace(/\b(?:100\s*%\s*)?(?:cotton|katun)(?:\s+(?:combed|premium|pe|20s|24s|30s))*\b/gi, "100%Sorona");
  output = output.replace(/\b(?:semi\s+katun|cotton\s+pe|pe\s+cotton)\b/gi, "100%Sorona");
  return output.replace(/[ \t]{2,}/g, " ");
}

function sourceImages(row) {
  return Array.from({ length: 9 }, (_, index) => text(row[`产品图 ${index + 1}`]));
}

function uniqueUrls(values) {
  const seen = new Set();
  const urls = [];
  for (const value of values) {
    const url = text(value);
    if (!/^https?:\/\//i.test(url) || seen.has(url)) continue;
    seen.add(url);
    urls.push(url);
  }
  return urls;
}

function joinUrl(base, kind, filename) {
  const parts = [base.replace(/\/+$/, "")];
  if (kind) parts.push(encodeURIComponent(kind));
  parts.push(...filename.split("/").filter(Boolean).map((part) => encodeURIComponent(part)));
  return parts.join("/");
}

function assetUrls(manifest, kind) {
  if (Array.isArray(manifest[`${kind}_urls`]) && manifest[`${kind}_urls`].length) {
    return manifest[`${kind}_urls`];
  }
  const files = manifest[`${kind}_files`] || [];
  const base = manifest[`${kind}_url_base`] || manifest.asset_url_base;
  return files.map((filename) => joinUrl(base, kind, filename));
}

function assetUrl(manifest, key, kind) {
  if (manifest[`${key}_url`]) return manifest[`${key}_url`];
  return joinUrl(manifest.asset_url_base, kind, manifest[`${key}_file`]);
}

function skuPrefix(value = "X") {
  return text(value).toUpperCase().replace(/[^0-9A-Z]+/g, "").slice(0, 8) || "X";
}

function skuDateFromFilename(filename) {
  const matches = text(filename).match(/20\d{6}/g) || [];
  for (const value of matches) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6));
    const day = Number(value.slice(6, 8));
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
      return value;
    }
  }
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
}

function detectSizeKey(rows) {
  const candidates = ["variation_value_1", "variation_value_2"];
  const scored = candidates.map((key) => {
    const header = SOURCE_COLUMNS[key];
    return [key, rows.filter((row) => normalizeSize(row[header])).length];
  });
  scored.sort((a, b) => b[1] - a[1] || (a[0] === "variation_value_2" ? -1 : 1));
  return scored[0][1] ? scored[0][0] : "";
}

function pairedVariationKey(valueKey, target) {
  const number = valueKey.split("_").pop();
  return `variation_${target}_${number}`;
}

function otherVariationValueKey(valueKey) {
  return valueKey === "variation_value_1" ? "variation_value_2" : "variation_value_1";
}

function orderedUniqueValues(rows, sourceKey) {
  const seen = new Set();
  const values = [];
  const header = SOURCE_COLUMNS[sourceKey];
  for (const row of rows) {
    const value = text(row[header]);
    if (!value || seen.has(value)) continue;
    seen.add(value);
    values.push(value);
  }
  return values;
}

function colorCodesInRows(rows, valueKey) {
  const header = SOURCE_COLUMNS[valueKey];
  const seen = new Set();
  const codes = [];
  for (const row of rows) {
    const code = normalizeColorCode(row[header]);
    if (!["BK", "WH"].includes(code) || seen.has(code)) continue;
    seen.add(code);
    codes.push(code);
  }
  return codes;
}

function sourceImageColorCodes(rows) {
  const codes = new Set();
  for (const key of ["variation_value_1", "variation_value_2"]) {
    const header = SOURCE_COLUMNS[key];
    for (const row of rows) {
      const code = normalizeColorCode(row[header]);
      if (["BK", "WH"].includes(code) && text(row[SOURCE_COLUMNS.variant_image_1])) {
        codes.add(code);
      }
    }
  }
  return Array.from(codes);
}

function outputColorGroups(selectedColors) {
  const clean = Array.isArray(selectedColors) ? selectedColors : [];
  const unique = clean.filter((code, index) => ["BK", "WH"].includes(code) && clean.indexOf(code) === index);
  if (!unique.length) throw new Error("每个商品至少要选择一个生成底色 BK 或 WH");
  return unique.map((code) => ({ code, label: COLOR_DISPLAY[code] || code }));
}

function expandRowsToTargetSizes(rows, selectedColors, sourceImageColor = "BK") {
  const selectedGroups = outputColorGroups(selectedColors);
  const manualSourceColor = sourceImageColor === "WH" ? "WH" : "BK";
  const sizeKey = detectSizeKey(rows);
  if (!sizeKey) {
    const expanded = [];
    for (const group of selectedGroups) {
      for (const size of TARGET_SIZES) {
        const next = { ...rows[0] };
        if (selectedGroups.length > 1) {
          next[SOURCE_COLUMNS.variation_name_1] = "Warna";
          next[SOURCE_COLUMNS.variation_value_1] = group.label;
          next[SOURCE_COLUMNS.variation_name_2] = "Ukuran";
          next[SOURCE_COLUMNS.variation_value_2] = size;
        } else {
          next[SOURCE_COLUMNS.variation_name_1] = "Ukuran";
          next[SOURCE_COLUMNS.variation_value_1] = size;
          next[SOURCE_COLUMNS.variation_name_2] = "";
          next[SOURCE_COLUMNS.variation_value_2] = "";
        }
        next._sku_size = size;
        next._sku_color_code = group.code;
        next._sku_variant_image_allowed = group.code === manualSourceColor;
        expanded.push(next);
      }
    }
    return expanded;
  }

  const sizeHeader = SOURCE_COLUMNS[sizeKey];
  const sizeNameHeader = SOURCE_COLUMNS[pairedVariationKey(sizeKey, "name")];
  const otherKey = otherVariationValueKey(sizeKey);
  const otherHeader = SOURCE_COLUMNS[otherKey];
  const otherNameHeader = SOURCE_COLUMNS[pairedVariationKey(otherKey, "name")];
  const sourceColorCodes = colorCodesInRows(rows, otherKey);
  const otherIsColor = sourceColorCodes.length > 0;

  const expanded = [];
  for (const group of selectedGroups) {
    const hasSourceColor = otherIsColor && sourceColorCodes.includes(group.code);
    const candidates = hasSourceColor ? rows.filter((row) => normalizeColorCode(row[otherHeader]) === group.code) : rows;
    const validCandidates = candidates.filter((row) => normalizeSize(row[sizeHeader]));
    const fallback = validCandidates[0] || candidates[0] || rows[0];

    for (const size of TARGET_SIZES) {
      const template = validCandidates.find((row) => normalizeSize(row[sizeHeader]) === size) || fallback;
      const next = { ...template };
      next[sizeHeader] = size;
      next[sizeNameHeader] = text(next[sizeNameHeader]) || "Ukuran";
      next._sku_size = size;
      next._sku_color_code = group.code;
      next._sku_variant_image_allowed = hasSourceColor || (!otherIsColor && group.code === manualSourceColor);
      if (otherIsColor || selectedGroups.length > 1) {
        next[otherHeader] = group.label || COLOR_DISPLAY[group.code] || group.code;
        next[otherNameHeader] = "Warna";
      } else {
        next[otherHeader] = "";
        next[otherNameHeader] = "";
      }
      expanded.push(next);
    }
  }
  return expanded;
}

function replaceDescriptionImages(description, detailUrls) {
  const tags = detailUrls.map((url) => `<img src="${url}">`).join("\n");
  const source = text(description);
  if (!source) return tags;
  let replaced = false;
  if (/<img\b[^>]*>/i.test(source)) {
    return source.replace(/<img\b[^>]*>/gi, () => {
      if (replaced) return "";
      replaced = true;
      return tags;
    });
  }
  return source.replace(/<body[^>]*>/i, (match) => `${match}\n${tags}`) || `${tags}\n${source}`;
}

function normalSku(row, skuDate, groupIndex, side) {
  const size = text(row._sku_size) || normalizeSize(row[SOURCE_COLUMNS.variation_value_1]) || normalizeSize(row[SOURCE_COLUMNS.variation_value_2]);
  const color = text(row._sku_color_code) || "BK";
  return `${skuPrefix()}-${skuDate}${String(groupIndex).padStart(3, "0")}-${side}-${color}-${size}`.slice(0, 50);
}

function extraSku(skuDate, groupIndex) {
  return `${skuPrefix()}-${skuDate}${String(groupIndex).padStart(3, "0")}-NB`.slice(0, 50);
}

function templateRow(row, productImages, description, sizeChartUrl, skuDate, groupIndex, side) {
  const variantImage = row._sku_variant_image_allowed === false ? "" : text(row[SOURCE_COLUMNS.variant_image_1]);
  return {
    category: "Men's Tops/T-shirts",
    brand: "No brand",
    product_name: truncate(normalizeMaterialText(row[SOURCE_COLUMNS.name]), 254),
    product_description: normalizeMaterialText(description),
    main_image: productImages[0],
    image_2: productImages[1],
    property_name_1: text(row[SOURCE_COLUMNS.variation_name_1]),
    property_value_1: truncate(row[SOURCE_COLUMNS.variation_value_1], 50),
    property_1_image: variantImage,
    property_name_2: text(row[SOURCE_COLUMNS.variation_name_2]),
    property_value_2: truncate(row[SOURCE_COLUMNS.variation_value_2], 50),
    parcel_weight: 210,
    parcel_length: 25,
    parcel_width: 22,
    parcel_height: 4,
    delivery: DELIVERY_TEXT,
    price: 259000,
    quantity: 999,
    seller_sku: normalSku(row, skuDate, groupIndex, side),
    size_chart: sizeChartUrl,
    shipping_insurance: "Optional",
    ...PRODUCT_PROPERTIES,
  };
}

function extraRow(firstRow, productImages, description, extraImageUrl, sizeChartUrl, skuDate, groupIndex) {
  const variationName1 = text(firstRow[SOURCE_COLUMNS.variation_name_1]);
  const variationName2 = text(firstRow[SOURCE_COLUMNS.variation_name_2]);
  return {
    category: "Men's Tops/T-shirts",
    brand: "No brand",
    product_name: truncate(normalizeMaterialText(firstRow[SOURCE_COLUMNS.name]), 254),
    product_description: normalizeMaterialText(description),
    main_image: extraImageUrl,
    image_2: productImages[1],
    property_name_1: variationName1,
    property_value_1: variationName1 ? EXTRA_VARIATION_VALUE_1 : "",
    property_1_image: extraImageUrl,
    property_name_2: variationName2,
    property_value_2: variationName2 ? EXTRA_VARIATION_VALUE_2 : "",
    parcel_weight: 210,
    parcel_length: 25,
    parcel_width: 22,
    parcel_height: 4,
    delivery: DELIVERY_TEXT,
    price: 141300,
    quantity: 999,
    seller_sku: extraSku(skuDate, groupIndex),
    size_chart: sizeChartUrl,
    shipping_insurance: "Optional",
    ...PRODUCT_PROPERTIES,
  };
}

function setCell(sheet, rowIndex, colIndex, value) {
  const address = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex - 1 });
  if (value === "" || value == null) {
    delete sheet[address];
    return;
  }
  if (typeof value === "number") {
    sheet[address] = { t: "n", v: value };
  } else {
    sheet[address] = { t: "s", v: String(value) };
  }
}

function clearDataArea(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:AO6");
  for (let row = DATA_START_ROW; row <= Math.max(range.e.r + 1, DATA_START_ROW); row += 1) {
    for (let col = 1; col <= 41; col += 1) {
      setCell(sheet, row, col, "");
    }
  }
}

function writeTemplateRow(sheet, rowIndex, data) {
  for (const [key, col] of Object.entries(TEMPLATE_COLUMNS)) {
    setCell(sheet, rowIndex, col, data[key] ?? "");
  }
}

function outputRowsFromGroups(groups, choices, assets, skuDate) {
  const outputRows = [];
  const reviewRows = [];

  groups.forEach((group, index) => {
    const groupIndex = index + 1;
    const productNo = String(groupIndex).padStart(3, "0");
    const firstRow = group.rows[0];
    const choice = choices[productNo] || {};
    const side = choice.side === "PR" ? "PR" : "P";
    const selectedColors = Array.isArray(choice.colors) ? choice.colors : ["BK", "WH"];
    const sourceImageColor = choice.sourceImageColor === "WH" ? "WH" : "BK";
    const originalImages = sourceImages(firstRow);
    if (!originalImages[0]) {
      throw new Error(`商品 ${productNo} 缺少产品图 1`);
    }

    const productImages = [originalImages[0], ...assets.mainUrls];
    const description = replaceDescriptionImages(
      text(firstRow[SOURCE_COLUMNS.long_description]) || text(firstRow[SOURCE_COLUMNS.short_description]),
      assets.detailUrls,
    );
    const skuRows = expandRowsToTargetSizes(group.rows, selectedColors, sourceImageColor);

    reviewRows.push({
      product_no: productNo,
      print_side: side,
      confidence: "manual",
      method: "web",
      evidence: `selected ${side}`,
      product_name: text(firstRow[SOURCE_COLUMNS.name]),
      source_url: text(firstRow[SOURCE_COLUMNS.source_url]),
    });

    for (const row of skuRows) {
      const out = templateRow(row, productImages, description, assets.sizeChartUrl, skuDate, groupIndex, side);
      assets.mainUrls.slice(1).forEach((url, offset) => {
        out[`image_${offset + 3}`] = url;
      });
      outputRows.push(out);
    }

    const extra = extraRow(firstRow, productImages, description, assets.extraUrl, assets.sizeChartUrl, skuDate, groupIndex);
    assets.mainUrls.slice(1).forEach((url, offset) => {
      extra[`image_${offset + 3}`] = url;
    });
    outputRows.push(extra);
  });

  return { outputRows, reviewRows };
}

function groupRows(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = text(row[SOURCE_COLUMNS.source_url]) || text(row[SOURCE_COLUMNS.name]);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return Array.from(map.entries()).map(([key, groupRowsValue], index) => {
    const firstRow = groupRowsValue[0];
    const thumbs = uniqueUrls([
      ...groupRowsValue.map((row) => row[SOURCE_COLUMNS.variant_image_1]),
      ...sourceImages(firstRow),
    ]);
    return {
      key,
      productNo: String(index + 1).padStart(3, "0"),
      rows: groupRowsValue,
      title: text(firstRow[SOURCE_COLUMNS.name]),
      sourceUrl: text(firstRow[SOURCE_COLUMNS.source_url]),
      image: sourceImages(firstRow)[0],
      thumbs,
      sourceImageColors: sourceImageColorCodes(groupRowsValue),
    };
  });
}

export async function parseBigSellerFile(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
  const filtered = rows.filter((row) => text(row[SOURCE_COLUMNS.name]));
  if (!filtered.length) throw new Error("BigSeller 表里没有可转换的数据行");
  return {
    filename: file.name,
    skuDate: skuDateFromFilename(file.name),
    groups: groupRows(filtered),
  };
}

export async function buildTikTokWorkbook({ groups, choices, sourceFilename }) {
  const manifest = await fetch("./assets_r2.json").then((response) => response.json());
  const assets = {
    mainUrls: assetUrls(manifest, "main"),
    detailUrls: assetUrls(manifest, "detail"),
    extraUrl: assetUrl(manifest, "extra", "extra"),
    sizeChartUrl: assetUrl(manifest, "size_chart", "detail"),
  };
  if (assets.mainUrls.length !== 8) throw new Error("R2 主图配置必须正好有 8 张图片");
  if (!assets.detailUrls.length) throw new Error("R2 详情图配置不能为空");
  if (!assets.sizeChartUrl) throw new Error("R2 尺码表配置不能为空");

  const skuDate = skuDateFromFilename(sourceFilename);
  const templateArray = await fetch("./template.xlsx").then((response) => response.arrayBuffer());
  const workbook = XLSX.read(templateArray, { type: "array" });
  const sheet = workbook.Sheets.Template || workbook.Sheets[workbook.SheetNames[0]];
  const { outputRows, reviewRows } = outputRowsFromGroups(groups, choices, assets, skuDate);

  clearDataArea(sheet);
  outputRows.forEach((row, index) => writeTemplateRow(sheet, DATA_START_ROW + index, row));
  const lastRow = DATA_START_ROW + outputRows.length - 1;
  sheet["!ref"] = `A1:AO${Math.max(lastRow, DATA_START_ROW)}`;

  const binary = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return {
    xlsxBlob: new Blob([binary], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    reviewBlob: new Blob([reviewCsv(reviewRows)], { type: "text/csv;charset=utf-8" }),
    xlsxName: `tiktokshop_import_${skuDate}.xlsx`,
    reviewName: `tiktokshop_import_${skuDate}_print_side_review.csv`,
    outputRows,
    reviewRows,
  };
}

function csvValue(value) {
  const clean = text(value).replace(/"/g, '""');
  return /[",\n\r]/.test(clean) ? `"${clean}"` : clean;
}

function reviewCsv(rows) {
  const fields = ["product_no", "print_side", "confidence", "method", "evidence", "product_name", "source_url"];
  const lines = [fields.join(",")];
  for (const row of rows) {
    lines.push(fields.map((field) => csvValue(row[field])).join(","));
  }
  return `\ufeff${lines.join("\n")}\n`;
}
