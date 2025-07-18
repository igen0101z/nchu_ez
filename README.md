# 興大學習日誌自動填寫工具

## 介紹
本程式可自動登入國立中興大學 EZ-Come 系統，批量填寫學習日誌，支援日期範圍選擇、校內編號管理、工作內容自動填入，並提供視覺化操作介面。

## 主要功能
- 自動登入 EZ-Come 系統
- 批量填寫學習日誌（每日模式）
- 可視化瀏覽器操作（Selenium）
- 智能表單識別
- 配置檔案自動儲存/載入
- 校內編號管理

## 執行環境
- Windows 作業系統
- Python 3.6 以上
- Chrome 瀏覽器
- chromedriver.exe（需放在程式同目錄）

## 安裝套件
請先安裝必要套件：
```
pip install -r requirements.txt
```

## requirements.txt
```
selenium>=4.0.0
tkcalendar>=1.6.1
```

## config.json 說明
```json
{
  "username": "帳號",
  "password": "密碼",
  "school_ids": ["校內編號1", "校內編號2", "..."],
  "url": "EZ-Come系統網址",
  "work_content": "自動填入的工作內容"
}
```
- username：登入用帳號
- password：登入用密碼
- school_ids：可選多個校內編號
- url：系統登入網址
- work_content：自動填入的日誌內容

## 使用方式
1. 執行 `main.py`
2. 輸入帳號、密碼、選擇校內編號
3. 選擇日期範圍
4. 輸入工作內容
5. 點擊「開始執行」

## 注意事項
- 建議先測試 1-2 天確認無誤
- 配置檔案會保存在程式目錄下
- 工作內容會原樣填入，請確保內容正確
- 預設操作延遲為 1 秒

## 常見問題
- 若無法使用日期選擇器，請安裝 tkcalendar：
  `pip install tkcalendar`
- 若 Selenium 或 chromedriver 有問題，請確認版本相符

