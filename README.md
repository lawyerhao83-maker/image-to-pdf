# 圖片檔轉 PDF

可直接在瀏覽器使用的 JPG 轉 PDF 工具。圖片全程在使用者的瀏覽器本機處理，不會上傳到伺服器。

## 線上使用

[開啟「圖片檔轉 PDF」](https://lawyerhao83-maker.github.io/image-to-pdf/)

## 功能

- 一次選取或拖放多張 JPG／JPEG 圖片
- 直接拖曳圖片縮圖調整 PDF 頁面順序
- 個別旋轉、刪除圖片
- 支援 A4、美國信紙及符合原圖尺寸
- 可設定頁面方向與邊距
- 預設完整保留圖片和文字，不裁切內容
- 在瀏覽器本機產生並下載 PDF

## 本機執行環境

- Node.js 22.13.0 或更新版本
- Chrome、Edge 或其他現代瀏覽器
- 不需要安裝 Python

Windows 使用者可雙擊 `啟動 PictoPDF.bat`。第一次啟動需要網路下載套件。

也可以在終端機執行：

```bash
npm install
npm run dev
```

## GitHub Pages

推送到 `main` 分支後，GitHub Actions 會自動建立並發布 GitHub Pages 版本。
