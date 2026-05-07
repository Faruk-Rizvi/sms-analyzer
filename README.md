# 📱 SMS Forensic Analyzer — Streamlit App

## 🚀 লোকালে চালানো

```bash
git clone https://github.com/yourusername/sms-forensic-analyzer-bd.git
cd sms-forensic-analyzer-bd

pip install -r requirements.txt

streamlit run app.py
```
ব্রাউজারে স্বয়ংক্রিয় খুলবে: **http://localhost:8501**

---

## 🌐 Streamlit Cloud (বিনামূল্যে হোস্টিং)

1. GitHub-এ push করুন
2. [share.streamlit.io](https://share.streamlit.io) → Sign in with GitHub
3. **New app** → repo নির্বাচন → `app.py`
4. **Deploy** → লিংক পাবেন

---

## 📦 ফাইল কাঠামো

```
sms-forensic-analyzer-bd/
├── app.py              ← Streamlit অ্যাপ
├── sms_analyzer.py     ← বিশ্লেষণ ইঞ্জিন
├── requirements.txt
├── .streamlit/
│   └── config.toml     ← থিম ও সেটিংস
└── README.md
```

---

## Word রিপোর্ট

Streamlit Cloud-এ Node.js নেই, তাই Word রিপোর্ট শুধু লোকালে পাবেন:
```bash
npm install -g docx
streamlit run app.py
```
