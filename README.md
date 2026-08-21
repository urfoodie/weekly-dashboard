# 本周核心数据

这是一个可直接部署到 GitHub Pages 的纯静态 HTML 看板。

## 功能

- 浏览器中直接上传 `.xlsx` / `.xls` / `.csv`
- 自动识别周次字段
- 一页展示全部周报模块
- 提供 Excel 填写模板

## 本地预览

直接双击 `index.html` 即可，或用任意静态服务器打开。

## 部署到 GitHub Pages

1. 把本目录内容推送到 GitHub 仓库
2. 进入仓库 `Settings`
3. 打开 `Pages`
4. `Build and deployment` 选择：
   - `Source`: `Deploy from a branch`
   - `Branch`: `main`（或你的发布分支）
   - `Folder`: `/ (root)`
5. 保存后等待 GitHub 发布

## 文件说明

- `index.html`：页面入口
- `app.js`：Excel 解析与单页渲染逻辑
- `styles.css`：页面样式
- `assets/js/xlsx.full.min.js`：浏览器端 Excel 解析库
- `周维度经营看板上传模板.xlsx`：上传模板
