# 亚马逊评论数据收集 Skill (Amazon Review Scraping)

<div align="center">
  <img src="./icon.png" alt="Amazon Review Scraping Icon" width="150" height="150" />
</div>

本 Skill 提供了一整套包含前置策略的高级亚马逊产品评论抓取工作流，同时保留了基础和隐秘的 Playwright 页面抓取功能。

## 🌟 核心能力与适用场景

当您需要进行以下操作时，请使用本 Skill：
- 从各大亚马逊站点（如 `amazon.sg`, `amazon.com`）批量抓取商品评价。
- 支持开启真实的浏览器并进行人工登录验证，复用持久化会话继续抓取任务。
- 提取并下载评论中所附带的买家晒图图片。
- 一键导出包含：评价详情、中文翻译、内置原图展现等功能的**多产品汇总 Excel 表格**。
- 对普通的动态 JS 渲染网页或具有强力反爬措施（如 Cloudflare, 403防护等）的域进行无感知抓取爬透。

## 🚀 主要包含脚本工具

### 1. 通用 Playwright 抓取模式
- `scripts/playwright-simple.js`：适用于普通无强力防抓取设置的 JS 页面。
- `scripts/playwright-stealth.js`：附带隐身 Stealth 插件的高级模式，适用于 Cloudflare 防护或较高强度反爬网页。

### 2. 亚马逊评论数据专属流
- **搜集初稿预检** (`amazon-preflight-workflow.js`)：支持 2~3 个关键词初步提取分类候选、预估评论规模、接受诸如“不搜索：1、8、11”等筛选命令后才开始执行爬虫操作。
- **爬虫实体** (`amazon-review-login-scrape.js`)：负责开启浏览器等待验证、爬取多标签下的真实评价，避免页码陷阱，全量提取 Review ID 和用户附件。
- **转制 Excel** (`amazon-reviews-to-excel.py` / `amazon-competitor-to-excel.py`)：将清洗和拉取后获取的一系列 JSON，汇总为可视化及汉化后的分析用多标签 Excel 文档。

## 💡 安装须知

请在本地 Skill 文件夹下执行以下命令进行初始化依赖安装：
```bash
npm install
npx playwright install chromium
python3 -m pip install openpyxl pillow requests
```

## 🛠️ 最佳实践与提示

1. **面临登录验证时**：务必使用 `HEADLESS=false` （带界面模式）。此时脚本会打开带有评价内容的登录页面并阻塞等候，直至您在真实弹出的窗口内完成通过 Amazon 系统登录验证及检测流程。
2. **搜索范围受控（Preflight Flow）**：为防止无端过载数据，采用此流时需确认生成了带有数字标号的**选择清单初稿**。若不满意部分数据，回复去除后必须输入明确的 `开始执行` 指令，才会真正激活底层爬虫。
3. **页面分页解析逻辑**：无需猜测或解析URL，该爬虫将模仿真实用户，连续点击底部的 `Show more reviews` 后加载真实DOM以抽取节点数据。
