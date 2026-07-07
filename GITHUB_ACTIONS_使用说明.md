# GitHub Actions 版使用说明

这个版本不需要服务器。你把 BigSeller Excel 放进 GitHub 仓库，然后手动运行一次 Actions，GitHub 会生成 TikTok Shop 批量上传表供你下载。

## 仓库里需要保留的文件

- `convert_bigseller_to_tiktok.py`
- `Bothplatforms_Menswear & Underwear_20260706_template.xlsx`
- `requirements.txt`
- `print_side_overrides.csv`，可选；只有需要人工指定某些商品为 P/PR 时才用，默认只保留表头
- `config/assets_r2.json`
- `.github/workflows/convert-bigseller.yml`
- `input/.gitkeep`

## 第一次上传到 GitHub

1. 建议创建一个 Private 私有仓库。
2. 把上面这些文件上传到仓库。
3. 不要把 Cloudflare R2 密钥上传到 GitHub。当前 Actions 版只使用 R2 公共图片 URL，不需要密钥。

## 每次转换怎么做

1. 在 GitHub 仓库里打开 `input` 文件夹。
2. 点击 `Add file`，选择 `Upload files`。
3. 上传 BigSeller 导出的 Excel。
4. 推荐把文件命名为 `bigseller.xlsx`，放在 `input/bigseller.xlsx`。
5. 打开仓库顶部的 `Actions`。
6. 左侧选择 `Convert BigSeller Excel`。
7. 点击 `Run workflow`。
8. `source_file` 填 `input/bigseller.xlsx`。如果你的文件名不同，就填实际路径。
9. `sku_date` 可以留空，也可以填类似 `20260707`。
10. `print_side_override` 默认留空，让脚本自动识别 P/PR。如果某一批要人工指定，再填 `print_side_overrides.csv` 或你的 CSV 路径。
11. 点击绿色的 `Run workflow`。

## 下载结果

1. 等运行状态变成绿色。
2. 点进这次运行记录。
3. 页面下方找到 `Artifacts`。
4. 下载 `tiktokshop-import-数字`。
5. 压缩包里会有：
   - `tiktokshop_import_数字.xlsx`
   - `tiktokshop_import_数字_print_side_review.csv`

## 改素材链接

如果以后 R2 图片文件名或路径变了，只改 `config/assets_r2.json`。

- `main_files` 是商品图 3-9 的 7 张主图。
- `detail_files` 是详情页图片。
- `extra_file` 是不要购买 SKU 的图片。
- `size_chart_file` 是尺码表图片。

## 注意

- GitHub Actions 不能像网页表单那样直接上传 Excel，所以需要先把 Excel 上传到仓库的 `input` 文件夹。
- 建议用私有仓库，避免源商品表被公开。
- 生成文件保留 14 天，过期后需要重新运行。
