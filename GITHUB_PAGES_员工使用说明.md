# GitHub Pages 纯前端版使用说明

这个页面不需要服务器。员工上传的 BigSeller Excel 只会在浏览器本地处理，不会上传到 GitHub 或其他服务器。

## 员工怎么用

1. 打开 GitHub Pages 网页。
2. 点击 `选择 Excel`。
3. 选择 BigSeller 导出的 `scraped_product_日期...xlsx` 文件。
4. 页面会列出每个商品。
5. 人工判断每个商品是单面 `P` 还是正反面 `PR`。
6. 如果商品源表没有颜色选项，就在右侧选择默认底色 `BK 黑色` 或 `WH 白色`。
7. 点击 `生成上传表`。
8. 下载 `TikTok 上传表` 和 `PR 复核表`。

## 当前自动规则

- 只保留 BigSeller 的第 1 张主图。
- 主图 2-9 自动替换为 R2 公共链接。
- 主图第 9 张是尺码表。
- 详情图自动替换为 R2 详情图。
- 尺码自动生成 `S, M, L, XL, XXL, XXXL`。
- 正常 SKU 价格 `259000`。
- 不要购买 SKU 价格 `141300`。
- 库存 `999`。
- 重量 `210g`。
- 长宽高 `25 * 22 * 4`。
- 材质自动改为 `100%Sorona`。
- SKU 日期优先从文件名提取，例如 `scraped_product_20260707023136128.xlsx` 会使用 `20260707`。

## 发布网页

仓库里已经有 `.github/workflows/deploy-pages.yml`。

推送到 `main` 后，GitHub 会尝试自动发布 `docs` 目录。

如果私有仓库无法启用 GitHub Pages，说明当前 GitHub 账号套餐不支持私有仓库 Pages。可以选择：

- 把这个网页工具仓库改成公开仓库。
- 或者把 `docs` 文件夹放到公司内网静态服务器上。

## 注意

- 页面不再自动识别 P/PR，必须由员工手动选择。
- 页面不读取 R2 密钥，只使用 R2 公共图片链接。
- 如果以后 R2 素材变了，修改 `docs/assets_r2.json` 和 `config/assets_r2.json`。
