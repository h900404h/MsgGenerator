# 📌 Msg Generator

**Java 自動化產生 Outlook .msg 郵件工具**\
**Java-based automated Outlook .msg generator**\
**Outlook .msg 自動生成ツール（Java）**

------------------------------------------------------------------------

# 📖 目錄 / Table of Contents / 目次

-   [繁體中文](#繁體中文)
-   [English](#english)
-   [日本語](#日本語)

------------------------------------------------------------------------

# 繁體中文

## 📌 專案介紹

本工具使用 **Java + Maven** 開發，可根據 **CSV 名單** 搭配 **table/ 下的
XLSM 明細檔** 自動產生 Outlook 專用的 **.msg 郵件檔案**。

本 repository **包含完整原始碼（src/）、示例 CSV、示例 XLSM 以及可執行
JAR**，可完整展示在作品集或面試中。

------------------------------------------------------------------------

## ✨ 功能特色

-   讀取 `mail_list.csv`
-   自動依照 `filename_suffix` 搜尋 table/ 內附件 (xlsm)
-   每個收件人產生一封獨立 `.msg`
-   CC 多筆支援
-   自動產生 `output_msg/YYYY/MM/`
-   支援全形、半形、多語系檔名
-   可使用 `run.bat` 一鍵執行

------------------------------------------------------------------------

## 📂 專案目錄結構

    demo1/
     ├─ table/
     ├─ mail_list.csv
     ├─ src/main/java/demo/msggenerator/MsgGenerator.java
     ├─ msg-generator-1.0.0-shaded.jar
     ├─ pom.xml
     └─ run.bat

------------------------------------------------------------------------

## 📝 CSV 格式

``` csv
to,cc,filename_suffix
user01@example.com,"staff01@example.com; staff02@example.com",UserA
user02@example.com,"staff01@example.com; staff02@example.com",UserB
sample@example.com,"staff01@example.com; staff02@example.com",Sample
```

------------------------------------------------------------------------

## ▶️ 執行方式

### 方式 1：雙擊 run.bat

``` bat
java -jar msg-generator-1.0.0-shaded.jar
pause > nul
```

### 方式 2：Maven

    mvn clean package
    java -jar target/msg-generator-1.0.0-shaded.jar

------------------------------------------------------------------------

# English

## Overview

This Java tool automatically generates **Outlook .msg files** using a
**CSV recipient list** and **XLSM attachments** stored inside `table/`.

This repository **includes full source code (src/), sample CSV, sample
XLSM files, and the executable shaded JAR**, making it suitable for
portfolio demonstration.

------------------------------------------------------------------------

## Features

-   Reads `mail_list.csv`
-   Retrieves XLSM attachments from `table/`
-   One `.msg` generated per recipient
-   Multiple CC support
-   Auto output folder creation
-   Single-click execution via `run.bat`

------------------------------------------------------------------------

## Run

    java -jar msg-generator-1.0.0-shaded.jar
    pause > nul

------------------------------------------------------------------------

# 日本語

## プロジェクト概要

本ツールは **Java + Maven** を使用し、CSV と `table/` の XLSM
ファイルから Outlook 用 **.msg ファイルを自動生成**します。

本リポジトリには、**完全なソースコード（src/）、サンプル CSV、サンプル
XLSM、実行可能 JAR** が含まれています。

------------------------------------------------------------------------

## 実行方法

    java -jar msg-generator-1.0.0-shaded.jar
    pause > nul

------------------------------------------------------------------------
