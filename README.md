# 老師小工具箱

## 放到 GitHub Pages（一次性）
1. 在 github.com 註冊並登入，按右上角「+」→「New repository」，名稱例如 `teacher-tools`，選 Public，建立。
2. 在新倉庫頁面按「uploading an existing file」，把本資料夾**裡面**的所有檔案與資料夾（index.html、tools.json、.nojekyll、typing 資料夾）拖進去，按「Commit changes」。
3. 到 Settings → Pages，「Source」選 `Deploy from a branch`，Branch 選 `main`、資料夾選 `/ (root)`，按 Save。
4. 約 1～2 分鐘後，網址是 `https://你的帳號.github.io/teacher-tools/`。把這個網址分享給其他老師。

## 新增工具
1. 在根目錄新增資料夾，例如 `spelling/`，放入工具的 `index.html`（和 `guide.html`，如果需要後端）。
2. 在 `tools.json` 加一筆：
   { "name": "工具名稱", "icon": "🧩", "desc": "一句話介紹", "url": "spelling/", "guide": "spelling/guide.html", "tag": "分類" }
3. 上傳（GitHub 網頁上「Add file → Upload files」），首頁會自動出現新卡片。

## 更新工具
直接在 GitHub 上傳同名檔案覆蓋即可。若更動了 `typing/Code.gs`，請把 `VERSION` 與 `typing/index.html` 裡的 `EXPECT` 一起加一，老師登入後台時會收到更新提醒。

## 授權
本專案採用 **CC BY-NC 4.0（姓名標示－非商業性）**：可自由使用、修改、分享，需標示作者，不得商用。詳見 `LICENSE`。

## 公開倉庫注意事項
- `typing/Code.gs` 的 `PW` 請保持「請改成你的老師密碼」，不要把真實密碼或你自己的 /exec 網址存進倉庫。
- 學生名單與成績匯出檔請勿上傳。
