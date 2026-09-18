# Paya-Bank-Reconciliation
# 🏦 Bank & PAYA Reconciliation

### Local Deterministic Reconciliation Engine + AI-Assisted Financial Analysis

A Windows desktop application for reconciling Iranian bank statements and PAYA transaction reports using a **local deterministic reconciliation engine**, with an optional **OpenAI-compatible AI analysis layer**.

The core reconciliation logic does **not depend on AI**.
AI is used as a secondary analytical layer for explaining ambiguities, identifying suspicious patterns, and generating an analytical report.

---

## 🇬🇧 English

### ✨ Features

* 🏦 Import Bank Excel workbooks
* 💳 Import PAYA Excel reports
* 🔍 Automatic worksheet detection
* 🧩 Automatic header-row detection
* 🧠 Automatic transaction-column detection
* 🔧 Manual column mapping when automatic detection needs adjustment
* 🔢 Persian and Arabic digit normalization
* 💰 Decimal-based financial amount processing
* 📅 Gregorian and Jalali date handling
* 🔗 Transaction matching using:

  * Transaction/reference IDs
  * Date
  * Amount
  * Transaction direction
  * Description similarity
* ⚠️ Detection of:

  * Bank-only transactions
  * PAYA-only transactions
  * Amount mismatches
  * Potential duplicates
  * Ambiguous matches
* 🧮 Balance arithmetic control
* 📊 Daily balance analysis
* 🤖 Optional AI-assisted analysis
* 📦 Chunked transmission of complete source datasets to an OpenAI-compatible API
* 📑 Multi-sheet Excel report generation
* 🪟 Native Windows GUI using Tkinter
* ⚙️ Local configuration through `config.json`
* 🚫 No Node.js required

---

### 🧠 Architecture

The application follows a deliberately hybrid architecture:

```text
                 ┌─────────────────────┐
                 │   Bank Excel File   │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Excel Parser /      │
                 │ Structure Detection │
                 └──────────┬──────────┘
                            │
                            │
                 ┌──────────▼──────────┐
                 │ Deterministic       │
                 │ Reconciliation      │
                 │ Engine              │
                 └──────────┬──────────┘
                            │
                            ├──────────────► Final Numeric Truth
                            │
                            ▼
                 ┌─────────────────────┐
                 │ AI Analysis Layer   │
                 │ (Optional)          │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Excel Report        │
                 └─────────────────────┘

                 ┌─────────────────────┐
                 │   PAYA Excel File   │
                 └─────────────────────┘
```

The deterministic engine is intentionally kept as the authoritative arithmetic and matching layer.

AI does not replace the reconciliation engine.

---

### 📊 Generated Report

The generated Excel workbook can contain:

| Sheet             | Purpose                                       |
| ----------------- | --------------------------------------------- |
| `Summary`         | Overall reconciliation summary                |
| `Matched`         | Matched bank/PAYA transactions                |
| `Bank Only`       | Transactions found only in the bank statement |
| `PAYA Only`       | Transactions found only in PAYA               |
| `Amount Mismatch` | Matched records with amount differences       |
| `Duplicates`      | Potential duplicate transaction groups        |
| `Daily Balance`   | Daily transaction/balance analysis            |
| `AI Analysis`     | AI-generated analytical report                |
| `Source Bank`     | Imported bank source data                     |
| `Source PAYA`     | Imported PAYA source data                     |
| `Final Control`   | Final arithmetic and reconciliation controls  |

---

### 🔐 Privacy & Security

**Important:** When AI analysis is enabled, the application can transmit the complete imported Bank and PAYA datasets to the configured OpenAI-compatible API in numbered chunks.

Do **not** use the AI mode with confidential financial data unless you have verified that the selected API provider, account, network and organizational policies permit such transmission.

For sensitive environments, the deterministic/local reconciliation engine can be used independently of the AI layer.

---

### ⚙️ Configuration

Default configuration is stored locally in:

```text
config.json
```

Example:

```json
{
  "base_url": "https://api.example.com/v1",
  "api_key": "",
  "model": "gpt-4o",
  "temperature": 0.0,
  "max_tokens": 4096,
  "timeout": 120,
  "send_all_rows": true,
  "ai_enabled": true,
  "chunk_rows": 150
}
```

The application is designed for APIs that expose an OpenAI-compatible:

```text
POST /chat/completions
```

interface.

---

### 🚀 Installation

Python 3.10+ is recommended.

Install dependencies:

```bash
py -m pip install -r requirements.txt
```

Run the application:

```bash
run.bat
```

Or:

```bash
python main.py
```

> Replace `main.py` with the actual application filename if different.

---

### 🪟 Windows EXE

The project can optionally be packaged as a standalone Windows executable.

If `build_exe.bat` is included:

```bash
build_exe.bat
```

The resulting executable can then be distributed without requiring users to install Python manually.

---

### 🧪 Development

Recommended future improvements:

* Unit tests for reconciliation rules
* Regression tests using real-world anonymized Excel samples
* More robust fuzzy matching
* Configurable matching thresholds
* Better duplicate detection
* Improved date normalization
* Persian/Jalali date conversion
* API retry and rate-limit handling
* Secure API-key storage
* Structured AI JSON responses
* Optional local/offline LLM support
* Comprehensive audit logging
* Automated CI tests
* Additional bank/PAYA format profiles

---

### ⚠️ Current Limitations

This project should be considered an **active development project**, not a certified banking or accounting system.

Matching financial transactions is inherently sensitive to the structure and semantics of the source files.

Always review reconciliation exceptions before making accounting, treasury or financial decisions.

---

### 📄 License

Add your preferred license before publishing the repository.

For example:

```text
MIT License
```

if you want to allow broad reuse of the project.

---

# 🇮🇷 فارسی

## 🏦 سامانه مغایرت‌گیری بانک و پایا

یک نرم‌افزار دسکتاپ ویندوزی برای **مغایرت‌گیری صورتحساب بانکی و گزارش تراکنش‌های پایا** که هسته اصلی آن بر پایه پردازش قطعی و محلی داده‌ها طراحی شده و در کنار آن، امکان استفاده از **هوش مصنوعی از طریق APIهای سازگار با OpenAI** را فراهم می‌کند.

نکته کلیدی معماری پروژه این است که:

> **هوش مصنوعی مرجع نهایی محاسبات و مغایرت‌گیری نیست.**

موتور محلی و deterministic مسئول محاسبات، تطبیق تراکنش‌ها و کنترل‌های عددی است و AI صرفاً به‌عنوان یک لایه تحلیلی ثانویه مورد استفاده قرار می‌گیرد.

---

## ✨ امکانات

* دریافت فایل Excel صورتحساب بانک
* دریافت فایل Excel گزارش پایا
* تشخیص خودکار Sheet مناسب
* تشخیص خودکار ردیف Header
* تشخیص خودکار ستون‌های تراکنش
* امکان اصلاح دستی Mapping ستون‌ها
* نرمال‌سازی ارقام فارسی و عربی
* پردازش مبالغ با `Decimal`
* پشتیبانی از تاریخ‌های متنی و میلادی
* تطبیق تراکنش‌ها بر اساس:

  * شماره پیگیری / شناسه
  * تاریخ
  * مبلغ
  * جهت تراکنش
  * شرح تراکنش
* شناسایی:

  * تراکنش‌های فقط موجود در بانک
  * تراکنش‌های فقط موجود در پایا
  * مغایرت مبلغ
  * تراکنش‌های احتمالی تکراری
  * موارد مبهم
* کنترل ریاضی مانده حساب
* تحلیل روزانه مانده
* تحلیل اختیاری با هوش مصنوعی
* ارسال داده‌ها به AI به‌صورت Chunk
* تولید گزارش Excel چندبخشی
* رابط گرافیکی Windows با Tkinter
* ذخیره تنظیمات در `config.json`
* بدون نیاز به Node.js

---

## 🧠 معماری سیستم

معماری پروژه عمداً ترکیبی طراحی شده است:

```text
فایل بانک ─────┐
               │
               ▼
       ┌──────────────────┐
       │ Excel Processing │
       │ & Detection      │
       └────────┬─────────┘
                │
                ▼
       ┌──────────────────┐
       │ موتور قطعی       │
       │ مغایرت‌گیری      │
       └────────┬─────────┘
                │
                ├──────► مرجع نهایی محاسبات
                │
                ▼
       ┌──────────────────┐
       │ لایه تحلیل AI    │
       │ (اختیاری)        │
       └────────┬─────────┘
                │
                ▼
       ┌──────────────────┐
       │ گزارش نهایی Excel│
       └──────────────────┘

فایل پایا ─────┘
```

این تفکیک باعث می‌شود نتیجه عددی سیستم به تصمیم یک مدل زبانی وابسته نباشد.

---

## 📊 ساختار گزارش خروجی

گزارش Excel شامل بخش‌هایی مانند موارد زیر است:

* خلاصه مغایرت‌گیری
* تراکنش‌های تطبیق‌یافته
* تراکنش‌های فقط بانکی
* تراکنش‌های فقط پایا
* مغایرت مبالغ
* تراکنش‌های احتمالی تکراری
* تحلیل روزانه مانده
* تحلیل هوش مصنوعی
* داده خام بانک
* داده خام پایا
* کنترل نهایی حساب

---

## 🔐 حریم خصوصی

**هشدار مهم:**

در صورت فعال بودن قابلیت AI، داده‌های واردشده از فایل بانک و پایا می‌توانند به API انتخاب‌شده ارسال شوند.

بنابراین قبل از استفاده از AI برای اطلاعات واقعی مالی باید سیاست‌های محرمانگی، امنیت و نگهداری اطلاعات سرویس‌دهنده API بررسی شود.

برای محیط‌های حساس، می‌توان از **موتور محلی مغایرت‌گیری بدون فعال‌سازی AI** استفاده کرد.

---

## ⚙️ تنظیمات

تنظیمات برنامه در فایل زیر ذخیره می‌شود:

```text
config.json
```

این تنظیمات شامل مواردی مانند:

* Base URL
* API Key
* Model
* Temperature
* Maximum Tokens
* Timeout
* تعداد ردیف هر Chunk
* فعال/غیرفعال بودن AI

است.

---

## 🚀 نصب

پیش‌نیاز پیشنهادی:

```text
Python 3.10+
```

سپس:

```bash
py -m pip install -r requirements.txt
```

و اجرای برنامه:

```bash
run.bat
```

---

## 🪟 نسخه اجرایی Windows

پروژه قابلیت بسته‌بندی به‌صورت Windows EXE را نیز دارد.

در صورت وجود فایل:

```text
build_exe.bat
```

می‌توان از آن برای ساخت نسخه اجرایی استفاده کرد.

---

## 🧪 مسیر توسعه آینده

برخی قابلیت‌هایی که می‌توان در نسخه‌های بعدی اضافه کرد:

* تست‌های Unit برای موتور مغایرت‌گیری
* مجموعه داده‌های تست anonymized
* الگوریتم‌های دقیق‌تر Fuzzy Matching
* Threshold قابل تنظیم برای تطبیق
* تشخیص پیشرفته‌تر تراکنش‌های تکراری
* تبدیل کامل تاریخ شمسی و میلادی
* Retry و مدیریت Rate Limit API
* ذخیره امن API Key
* خروجی ساختاریافته JSON از AI
* پشتیبانی از مدل‌های زبانی Local/Offline
* Audit Log کامل
* تست خودکار GitHub Actions
* پروفایل اختصاصی برای فرمت بانک‌های مختلف

---

## ⚠️ وضعیت پروژه

این پروژه در حال توسعه است و نباید در وضعیت فعلی به‌عنوان یک **سیستم رسمی بانکی، حسابداری یا سامانه تأییدشده مالی** تلقی شود.

به دلیل حساسیت مغایرت‌گیری مالی، تمام موارد استثنا و مغایرت باید قبل از هرگونه تصمیم حسابداری یا مالی توسط کاربر متخصص بررسی شوند.

---

## 📜 License

پیش از انتشار عمومی، License پروژه مشخص شود.

برای پروژه‌های متن‌باز عمومی، یکی از گزینه‌های رایج:

```text
MIT License
```

است.

---

### 👤 Project

Developed as a practical financial-data automation project combining:

**Python + Excel Processing + Deterministic Reconciliation + AI-Assisted Analysis**

---

## ⭐ هدف پروژه

هدف اصلی این پروژه استفاده از هوش مصنوعی برای **تقویت تحلیل مالی** است، نه جایگزین کردن منطق قطعی و قابل حسابرسی با یک مدل زبانی.

**Local Truth First. AI Second.**

