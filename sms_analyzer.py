#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS Forensic Analyzer v10 — Bangladesh Law Enforcement
তিন স্তরের লেনদেন + ব্যক্তিগত/সেবা SMS
"""
import openpyxl, re, html as hl, sys, os, csv
from datetime import datetime
from collections import Counter

TIER_SMALL = 10_000
TIER_MED   = 50_000
TIER_HIGH  = 100_000

# ── BD নম্বর ────────────────────────────────────────────────
BD_RE = re.compile(r'^\+?8801[3-9]\d{8}$|^01[3-9]\d{8}$')
def is_bd(s): return bool(BD_RE.match(str(s).strip()))

# ── Sender type ────────────────────────────────────────────
MFS_KW    = ['bkash','b-kash','nagad','rocket','dutch-bangla','dbbl','upay','mcash','tap','surecash']
BANK_KW   = ['ibbl','ibbl .','ibbl.','islamibank','islami bank',
             'trustbank','trust bank','trustcard',
             'sonali','janata bank','agrani bank','rupali bank',
             'brac bank','bracbank','brac-bank',
             'city bank','citybank','ebl','eastern bank','prime bank',
             'mutual trust','mtb.','mtb ','\bmtb\b',
             'southeast bank','ab bank','ucb','hsbc','standard chartered',
             'ific','idlc','metlife','dutch-bangla','dbbl','nexuspay','sonali bank']

# Short-code senders that are actually banks (not to be filtered)
BANK_SHORTCODES = {'16216', '16219', '16240', '16232'}

# Promotional / system numbers — সম্পূর্ণ filter
PROMO_NUMBERS = {
    '8801886016460','01886016460',
    '8801782542199','01782542199',
    '8801554702104','01554702104',
    # নতুন প্রমো/সার্ভিস নম্বর
    '8801905186028','01905186028',
    '8801401866071','01401866071',
    '8801927787270','01927787270',
    '8801964123108','01964123108',
    '8801924967582','01924967582',
    '8801577556527','01577556527',
    '8801908568652','01908568652',
    '8801402940991','01402940991',
    '8801934147013','01934147013',
}

SVC_KW    = ['desco','dpdc','breb','titas','wasa','aarong','brac','daraz','chaldal',
             'pathao','steadfast','paperfly','redx','sundarban','ecourier','pandago',
             'foodpanda','air astra','biman','us-bangla','novoair','ivac','ivac_bd',
             'epassportbd','epassport','walton','flexiload']

def stype(sender):
    sl = str(sender).lower().strip()
    s  = str(sender).strip()
    # Promotional numbers — filter out
    if s in PROMO_NUMBERS: return "promo"
    # Bank short-codes — always bank
    if s in BANK_SHORTCODES: return "bank"
    # MTB. specific match
    if re.match(r'^mtb\.?\s*$', sl): return "bank"
    if any(m in sl for m in MFS_KW):  return "mfs"
    if any(b in sl for b in BANK_KW): return "bank"
    if any(s_ in sl for s_ in SVC_KW): return "service"
    if is_bd(sender):                  return "personal"
    return "other"

# ── ফিল্টার ────────────────────────────────────────────────
SC_RE = re.compile(r'^\d{4,8}$')
OP_KW = ['airtel','at_only4u','at_special','at_topdeals','at_1p_sec','at_comeback',
         'at_boishakh','atforlife','monthly_minutes','airtel_data','airteloffer',
         'airtelbonus','combop','limitedtime','ajker_offer','mini_combo','mini_pack',
         'bishalof','bondho_','specialpack','darun_offer','beshi_min','goongoon',
         'secretoffer','tong offer','lucky_hour','best_deal','cashback','myfamily',
         'myrobi','nagad_offer','apnar_offer','gp','gpstar','gp-info','gpalert',
         'grameenphone','robi','banglalink','bl-','mybl','bl offer','bl commerce',
         'blsuperdeal','teletalk','govtinfo','govt info','govt','btrc','news_alert']

SKIP_BODY_RE = re.compile(
    # Messaging apps
    r"whatsapp\.com|Let's chat on WhatsApp|Let's video call"
    r"|imo\.onelink|imo-তে|Let's have a call|meet\.google\.com"
    r"|Can't talk now|I'll be there soon|This invite expires"

    # মিসড কল / কল নোটিফিকেশন
    r"|কলটি গ্রহণ করতে পারবেন|কল মিস|এখন কলটি গ্রহণ"
    r"|missed call|miss call|কলটি মিস|মিস করেছেন"

    # অপারেটর নাম
    r"|Grameenphone|Grameen\s*phone|Robi(?:\s+Axiata)?|Airtel"
    r"|Banglalink|Teletalk|MyBL|MyRobi|MyAirtel"

    # বাংলা প্রমো keywords
    r"|পুরস্কার\s*পান|লটারি.বিজয়ী|জেতা\s*সম্ভব"
    r"|ফ্রি\s*বোনাস\s*পেতে|বোনাসটি\s*অ্যাকাউন্টে\s*জমা"
    r"|প্রিয়\s*গ্রাহক(?!\s*আপনার\s*(?:পার্সেল|অর্ডার|পণ্য))"  # প্রিয় গ্রাহক but NOT courier ones

    # Delivery OTP — বাদ (কিন্তু tracking রাখো)
    r"|ডেলিভারি\s*ওটিপি|delivery\s*otp|অর্ডারটির\s*ডেলিভারি\s*ওটিপি"

    # Gambling / bonus / lottery — সব বাদ
    r"|\bলটারি\b|lottery|বিজয়ী\s*নম্বর"
    r"|নিবন্ধন.*ফ্রি.*বোনাস|ফ্রি\s*\+.*%.*বোনাস|\d+%\s*বোনাস|বোনাস\s*\d+%"
    r"|এখনই\s*নিবন্ধন|ক্যাসিনো|casino|জুয়া|\bbet\b"

    # Short URL spam (ln.run, bit.ly etc — courier tracking URL নয়)
    r"|ln\.run/|bit\.ly/(?!.*(?:parcel|track|order))",
    re.I | re.S)

MISSED_RE = re.compile(r'মিস করেছেন|কল মিস|missed call|miss call|কলটি মিস|কলটি গ্রহণ করতে পারবেন', re.I)

# OTP / verification — সব filter হবে
OTP_BODY_RE = re.compile(
    r'(?:your\s+)?otp\s+(?:is|for)|one.time.(?:password|key|code|pin)[\s:]'
    r'|verification\s+code\s+(?:is|for)'
    r'|do\s+not\s+share\s+(?:this|your|otp)'
    r'|(?:your\s+)?code\s+is\s*:\s*\d'
    r'|is\s+your\s+otp\b'
    r'|as\s+(?:one.time|otp)[\s:]'
    r'|use\s+the\s+(?:number|code|pin|key)\s*:\s*\d'
    r'|authentication\s+code\s+(?:is|for)'
    r'|login\s+otp\b'
    r'|ibas.*use.*(?:number|code|key)\s*:\s*\d'
    r'|please\s+use\s+the\s+(?:number|code|key)\s*:\s*\d',
    re.I)

def is_filtered(sender, body):
    s  = str(sender).strip()
    st = stype(s)
    # Promotional numbers — always filter
    if st == "promo": return True
    # OTP/key মেসেজ — সব sender-এর জন্য বাদ
    if OTP_BODY_RE.search(body): return True
    # Body keywords — সব sender-এর জন্য check করো
    if SKIP_BODY_RE.search(body): return True
    if MISSED_RE.search(body): return True
    # Short-code filter — bank shortcodes ছাড়া
    if SC_RE.match(s) and s not in BANK_SHORTCODES: return True
    # Operator sender keywords
    sl = s.lower()
    return any(f in sl for f in OP_KW)

# ── Multi-part merge ────────────────────────────────────────
def to_dt(d):
    if isinstance(d, datetime): return d
    s = str(d).strip()
    # Format: 20260210121148 (YYYYMMDDHHmmss) — GP operator format
    if re.match(r'^\d{14}$', s):
        try: return datetime.strptime(s, '%Y%m%d%H%M%S')
        except: pass
    # Format: 20260210 (YYYYMMDD only)
    if re.match(r'^\d{8}$', s):
        try: return datetime.strptime(s, '%Y%m%d')
        except: pass
    # Standard formats
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%dT%H:%M:%S','%d/%m/%Y %H:%M'):
        try: return datetime.strptime(s[:19], fmt)
        except: pass
    return datetime.min

def merge(sms_list):
    """শুধু ব্যক্তিগত BD নম্বরের ভাঙা SMS জোড়া লাগায়।
    Bank / MFS / Service SMS কখনো merge হবে না।"""
    items = sorted(sms_list, key=lambda x: to_dt(x["date"]))
    out, i = [], 0
    while i < len(items):
        cur  = dict(items[i])
        st   = stype(cur["sender"])
        # Bank/MFS/Service/Other — merge নয়, সরাসরি রাখো
        if st != "personal":
            cur["parts"] = 1
            out.append(cur)
            i += 1
            continue
        # Personal BD নম্বর — ২ মিনিটের মধ্যে একই sender হলে জোড়া
        dt0  = to_dt(cur["date"])
        sl0  = cur["sender"].strip().lower()
        body = cur["body"]
        j = i + 1
        while j < len(items):
            nx = items[j]
            if (stype(nx["sender"]) == "personal" and
                    nx["sender"].strip().lower() == sl0 and
                    abs((to_dt(nx["date"]) - dt0).total_seconds()) <= 120):
                body += "\n" + nx["body"]
                j += 1
            else:
                break
        cur["body"]  = body
        cur["parts"] = j - i
        out.append(cur)
        i = j
    return out

# ── লেনদেন পরিমাণ ──────────────────────────────────────────
TXN_P = [
    r'received\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+from',
    r'received\s+deposit.*?of\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'sent\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+to',
    r'[Cc]ash\s*[Ii]n\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Cc]ash\s*[Oo]ut\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Mm]oney\s+[Rr]eceived\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Pp]aid\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Aa]mount[:\s]+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Dd]eposit\s+[Aa]mount[:\s]+([\d,]+\.?\d*)',
    r'transfer(?:red)?\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Pp]ayment\s+of\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+has\s+been\s+(?:deposited|withdrawn|transferred|credited|debited)',
    r'[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+(?:withdrawn|credited|deposited)\s+(?:from|to)',
    r'[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+has\s+been\s+credited\s+to',
    r'[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+\(incl',
    r'[Ww]ithdrawn\s+([\d,]+\.?\d*)\s+BDT',
    r'debited\s+by\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'BDT\s*([\d,]+\.?\d*)\s+has\s+been\s+(?:transferred|deposited|debited|credited)',
    r'[Tt]xn\s+(?:of\s+)?[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Pp]urchase\s+(?:BDT|INR|USD|\$)\s*([\d,]+\.?\d*)',
    r'[Dd]ebited\s+for.*?(?:BDT|[Tt][Kk])\.?\s*([\d,]+\.?\d*)',
    # DBBL / 16216 pattern: "debited (ATM Cash Withdrawal) by Tk20,000.00"
    r'debited\s+\(.*?\)\s+by\s+Tk([\d,]+\.?\d*)',
    r'credited\s+\(.*?\)\s+by\s+Tk([\d,]+\.?\d*)',
    # Sonali Bank: "for BDT 7,488.00"
    r'for\s+BDT\s+([\d,]+\.?\d*)',
    # IBBL: "Withdrawn 5000.00 BDT" / "5000.00 BDT"
    r'(?:Withdrawn|Deposited)\s+([\d,]+\.?\d*)\s+BDT',
    r'[Pp]aid[:\s]*([\d,]+)',
]

def get_amount(text):
    # Balance / C/B / Fee বাদ দিই
    c = re.sub(r'C/B\s+Tk[\d,]+\.?\d*','',text)
    c = re.sub(r'Avl\s+Bal[:\s]+[\d,]+\.?\d*','',c)
    c = re.sub(r'A/C\s+Bal\s+BDT\s+[\d,.]+','',c)
    c = re.sub(r'(?:[Aa]vl?\.?\s+)?[Bb]al(?:ance)?[:\s\.]+[Tt]?[Kk]?\.?\s*[\d,]+\.?\d*','',c)
    c = re.sub(r'[Ff]ee[:\s]+[Tt][Kk]?\.?\s*[\d,]+\.?\d*','',c)
    vals = []
    for p in TXN_P:
        for m in re.findall(p,c,re.I):
            try:
                f = float(str(m).replace(',',''))
                if f > 0: vals.append(f)
            except: pass
    return max(vals) if vals else 0.0

def tier(amt):
    if amt >= TIER_HIGH:  return "high"
    if amt >= TIER_MED:   return "med"
    if amt >= TIER_SMALL: return "small"
    return ""

# ── কীওয়ার্ড ──────────────────────────────────────────────
ID_RE = re.compile(
    # ব্যক্তিগত পরিচয়
    r'নাম[:\s]|পিতা[:\s]|মাতা[:\s]|স্বামী|স্ত্রী|বিবাহ|নিকাহ'
    r'|name[:\s]|father[:\s]|mother[:\s]'
    # ঠিকানা — বাংলা
    r'|গ্রাম[:\s]|ডাকঘর[:\s]|উপজেলা|মহল্লা[:\s]|বাড়ি[:\s]|বাসা[:\s]'
    r'|রোড নং|লেন[:\s]|হোল্ডিং|ব্লক[:\s]|ফ্লোর[:\s]'
    r'|পোস্ট অফিস|পোস্টকোড|থানা[:\s]'
    # ঠিকানা — ইংরেজি
    r'|house\s*no|flat\s*no|holding\s*no|road\s*no|block\s*[a-z\d]'
    r'|floor\s*\d|apt\s+\d|sector\s*\d|building\s*name|address'
    r'|village[:\s]|post\s*office|upazila|thana[:\s]|district[:\s]'
    # ব্যাংক তথ্য
    r'|acc\.?\s*no|a/c\s*#|account\s*no|routing\s*no|swift\s*code|bank\s*name'
    , re.I)

LOVE_RE = re.compile(
    r'love\s+you|i\s+love|ভালোবাসি|ভালোবাসা|তোমাকে ভালো|আমি তোমাকে'
    r'|আব্বু|আম্মু|\bবাবা\b|সোনামণি|প্রিয়\s|A\+\s*পাইছে|রেজাল্ট'
    r'|\bজ্বর\b|দুলাভাই|ভাবি|eid mubarak|ঈদ মুবারক', re.I)

TRAVEL_RE = re.compile(
    r'\bpnr\b|flight\s+\d|boarding|e-ticket|ticket\s+amount'
    r'|webfile|ivac|your\s+visa|visa\s+application|passport\s+is\s+ready'
    r'|departure|arrival|check-in|air\s+astra|biman|us-bangla|novoair'
    r'|ওমরা|umrah|\bহজ\b|omra|ফ্লাইট|ভিসা|\bNOC\b|N\.O\.C|ডিরেক্ট\s+ফ্লাইট', re.I)

COURIER_RE = re.compile(
    # প্রতিষ্ঠানের নাম
    r'\bsteadfast\b|paperfly|sundarban\s+courier|\becourier\b|\bredx\b'
    r'|sa\s+paribahan|s\.a\.\s*paribahan|sundarban|continental\s+courier'
    r'|\bpathao\b|\bdaraz\b|\bchaldal\b'
    # ডেলিভারি সংক্রান্ত
    r'|your\s+parcel|track\s+your\s+parcel|your\s+order\s+has'
    r'|out\s+for\s+delivery|delivered\s+to|delivery\s+address'
    r'|order\s+shipped|package\s+delivered|আপনার\s+পার্সেল'
    r'|পার্সেল\s+ডেলিভারি|ডেলিভারি\s+হয়েছে'
    # Tracking
    r'|tracking\s+(no|number|id)|track\s+at\s+https'
    , re.I)

SHOPPING_RE = re.compile(
    r'aarong|invoice|\bvn:|thanks\s+for\s+choosing|receipt|order\s*#|star\s+tech'
    r'|walton|কিস্তি|installment|barcode.*নিবন্ধন'
    r'|desco|dpdc|wasa|titas\s+gas|বিদ্যুৎ\s+বিল', re.I)

OTP_RE = OTP_BODY_RE  # alias — same regex
LOW_BAL_RE = re.compile(r'your\s+balance\s+is\s+(very\s+low|finished)', re.I)

# ── একটি SMS বিশ্লেষণ ──────────────────────────────────────
# Body থেকে bank/MFS service detect করার জন্য keywords
BODY_BANK_RE = re.compile(
    r'\bislami\s+bank\b|\bIBBL\b|\bBRAC\s+Bank\b|\bTrust\s+Bank\b|'
    r'\bSonali\s+Bank\b|\bDutch.Bangla\b|\bDBBL\b|\bmutual\s+trust\b|'
    r'\bCity\s+Bank\b|\bEastern\s+Bank\b|\bEBL\b',
    re.I
)
BODY_MFS_RE = re.compile(
    r'\bbKash\b|\bNagad\b|\bRocket\b|\bUpay\b|\bMcash\b', re.I
)

def detect_service_from_body(body):
    """Body থেকে ব্যাংক/MFS শনাক্ত করে"""
    bl = body.lower()
    if "islami bank" in bl or re.search(r'\bibbl\b', bl):  return "bank"
    if "brac bank" in bl or "bracbank" in bl:               return "bank"
    if "trust bank" in bl or "trustbank" in bl:             return "bank"
    if "sonali bank" in bl:                                  return "bank"
    if "dutch-bangla" in bl or re.search(r'\bdbbl\b', bl): return "bank"
    if "mutual trust" in bl or re.search(r'\bmtb\b', bl) or "mtblbddh" in bl: return "bank"
    if "city bank" in bl:                                   return "bank"
    if "southeast bank" in bl or "sebl" in bl:              return "bank"
    if "ab bank" in bl:                                     return "bank"
    if "prime bank" in bl:                                  return "bank"
    if "ucb" in bl and "bank" in bl:                        return "bank"
    if "eastern bank" in bl or re.search(r'\bebl\b', bl):  return "bank"
    if re.search(r'\bbkash\b', bl):                         return "mfs"
    if re.search(r'\bnagad\b', bl):                         return "mfs"
    if re.search(r'\brocket\b', bl) and "balance" in bl:   return "mfs"
    return ""

def analyze_one(sender, body, date, direction=""):
    st  = stype(sender)
    amt = get_amount(body)
    t   = tier(amt)
    cats = []

    if t: cats.append(f"txn_{t}")

    # Personal নম্বর হলেও body-তে bank/MFS থাকলে সেই type অনুযায়ী চলবে
    body_st = detect_service_from_body(body)
    if st == "personal" and body_st in ("bank","mfs"):
        st = body_st

    # MFS/Bank — OTP ও low balance বাদ
    if st in ("mfs","bank") and not t:
        if OTP_RE.search(body) or LOW_BAL_RE.search(body): return None
        # Balance notification — sidebar-এর জন্য রাখো
        acct = get_acct(sender, body, 0)
        if acct.get("balance",0) > 0:
            cats.append("bank_info")   # special — sidebar only, card-এ দেখাবে না

    if st == "personal":
        if ID_RE.search(body):         cats.append("identity")
        elif LOVE_RE.search(body):     cats.append("love_family")
        elif TRAVEL_RE.search(body):   cats.append("travel")
        elif COURIER_RE.search(body):  cats.append("courier")
        elif SHOPPING_RE.search(body): cats.append("shopping")
        elif len(body.strip()) > 20:   cats.append("personal")
        if not cats: return None

    elif st in ("service","other"):
        if TRAVEL_RE.search(body):        cats.append("travel")
        elif COURIER_RE.search(body):     cats.append("courier")
        elif SHOPPING_RE.search(body):    cats.append("shopping")
        elif ID_RE.search(body):          cats.append("identity")
        if not cats and not t: return None

    if not cats: return None

    def to_month(d):
        if isinstance(d, datetime): return d.strftime('%Y-%m')
        try: return datetime.strptime(str(d)[:10],'%Y-%m-%d').strftime('%Y-%m')
        except: return ""

    acct = {}
    if st in ("mfs","bank") or t or body_st:
        acct = get_acct(sender, body, amt)

    return {"sender":str(sender),"body":body,"date":str(date) if date else "",
            "month":to_month(date),"sender_type":st,"categories":cats,
            "amount":amt,"tier":t,"acct":acct,"direction":direction}

# ── একাউন্ট তথ্য ───────────────────────────────────────────
def get_acct(sender, body, amt):
    sl = str(sender).lower()
    s  = str(sender).strip()
    bl = body.lower()
    info = {"service":"","trx_id":"","direction":"","other_number":"","balance":0,"amount":amt}

    # Short-code bank senders
    if s == "16216": info["service"] = "Dutch-Bangla Bank (DBBL)"
    elif s == "16219": info["service"] = "Mutual Trust Bank (MTB)"
    elif re.match(r'^mtb\.?\s*$', sl): info["service"] = "Mutual Trust Bank (MTB)"

    # Sender keyword detect
    if not info["service"]:
        if   "bkash"    in sl: info["service"]="bKash"
        elif "nagad"    in sl: info["service"]="Nagad"
        elif "rocket"   in sl: info["service"]="Rocket"
        elif "dutch-bangla" in sl or "dbbl" in sl: info["service"]="Dutch-Bangla Bank (DBBL)"
        elif "upay"     in sl: info["service"]="Upay"
        elif "ibbl"     in sl or "islami" in sl: info["service"]="Islami Bank"
        elif "trustbank" in sl or "trust bank" in sl: info["service"]="Trust Bank"
        elif "trustcard" in sl: info["service"]="Trust Card"
        elif "sonali"   in sl: info["service"]="Sonali Bank"
        elif "brac"     in sl: info["service"]="BRAC Bank"
        elif "metlife"  in sl: info["service"]="MetLife"

    # Body fallback — যেকোনো নম্বর থেকে আসলেও bank detect করবে
    if not info["service"]:
        if "islami bank" in bl or re.search(r'\bibbl\b',bl):         info["service"]="Islami Bank"
        elif "brac bank" in bl or "bracbank" in bl:                   info["service"]="BRAC Bank"
        elif "mutual trust" in bl or "mtblbddh" in bl or re.search(r'\bmtb\b',bl):
            info["service"]="Mutual Trust Bank (MTB)"
        elif "trust bank" in bl or "trustbank" in bl:                 info["service"]="Trust Bank"
        elif "sonali bank" in bl:                                      info["service"]="Sonali Bank"
        elif re.search(r'\bbkash\b',bl):                              info["service"]="bKash"
        elif re.search(r'\bnagad\b',bl):                              info["service"]="Nagad"
        elif "dutch-bangla" in bl or re.search(r'\bdbbl\b',bl) or "nexuspay" in bl:
            info["service"]="Dutch-Bangla Bank (DBBL)"
        elif "southeast bank" in bl or "sebl" in bl:                  info["service"]="Southeast Bank"
        elif "prime bank" in bl:                                       info["service"]="Prime Bank"
        elif "eastern bank" in bl or re.search(r'\bebl\b',bl):       info["service"]="Eastern Bank"
        elif "city bank" in bl:                                        info["service"]="City Bank"
        else: info["service"] = s

    trx = re.search(r'(?:[Tt]rx|[Tt]xn)\s*[Ii][Dd]?[:\s]+([A-Z0-9]{8,})', body)
    if trx: info["trx_id"]=trx.group(1)
    nums = re.findall(r'\b(01[3-9]\d{8})\b', body)
    if nums: info["other_number"]=nums[0]

    # Balance — C/B pattern (DBBL), Avl Bal, Balance, BDT balance
    bal = re.search(r'C/B\s+Tk([\d,]+\.?\d*)', body)
    if not bal: bal = re.search(r'Avl\s+Bal[:\s]+([\d,]+\.?\d*)', body)
    if not bal: bal = re.search(r'available\s+balance.*?(?:is\s+)?BDT\s+([\d,]+\.?\d*)', body, re.I)
    if not bal: bal = re.search(r'A/C\s+Bal\s+BDT\s+([\d,.]+)', body)
    if not bal: bal = re.search(r'[Bb]al(?:ance)?[:\s\.]+[Tt]?[Kk]?\.?\s*([\d,]+\.?\d*)', body)
    if bal:
        try: info["balance"]=float(bal.group(1).replace(',',''))
        except: pass

    bl2 = body.lower()
    # Credit = Received (গ্রহণ), Debit = Sent/Withdrawn (প্রদান/উত্তোলন)
    # Priority: explicit credit/debit keywords first
    if re.search(r'\bcredited\b|\breceived\b|\bcash\s*in\b|\bdeposit(?:ed)?\b|\bsalary\b|\bfund\s+transfer.*credit', bl2):
        info["direction"] = "Received"
    elif re.search(r'\bdebited\b|\bwithdrawn\b|\bwithdrawal\b|\bcash\s*out\b|\bpurchase\b|\bpos\b|\batm\b|\btransfer.*debit', bl2):
        info["direction"] = "Sent"
    elif re.search(r'\bsent\b', bl2):
        info["direction"] = "Sent"
    return info

# ── Data load ──────────────────────────────────────────────
def load_xlsx(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    hdr = [str(c).strip().lower() if c else "" for c in rows[0]]
    def ci(ns):
        for n in ns:
            for i,h in enumerate(hdr):
                if n in h: return i
        return None
    # Support both "StartTime"/"Start" and "A Party"/"AParty"/"BParty"/"B Party"
    ti = ci(['starttime','start','date','time'])
    ai = ci(['a party','aparty','sender_a'])
    bi = ci(['b party','bparty','sender','sender_b'])
    si = ci(['sender'])
    mi = ci(['message content','message','body','msg'])
    di = ci(['direction'])
    out=[]
    for r in rows[1:]:
        def g(i): return str(r[i]).strip() if i is not None and i<len(r) and r[i] else ""
        aparty    = g(ai)
        sender    = g(bi) if bi is not None else (g(ai) if ai is not None else g(si))
        body      = g(mi)
        raw_date  = r[ti] if ti is not None and ti<len(r) else ""
        # Convert to datetime for consistent handling
        date      = to_dt(raw_date) if raw_date else ""
        direction = g(di).upper() if di is not None else ""
        if direction in ("SMSMT","MTSMS"): direction = "SMSMT"
        elif direction in ("SMSMO","MOSMS"): direction = "SMSMO"
        if sender and body:
            out.append({"sender":sender,"body":body,"date":date,
                        "aparty":aparty,"direction":direction})
    return out

def load_csv_f(path):
    out=[]
    with open(path,encoding='utf-8') as f:
        for row in csv.DictReader(f):
            s=row.get('BParty') or row.get('sender') or row.get('AParty','')
            b=row.get('Message Content') or row.get('body') or row.get('message','')
            d=row.get('Start') or row.get('date','')
            if s and b: out.append({"sender":s,"body":b,"date":d})
    return out

# ── HTML ───────────────────────────────────────────────────
CAT_META = {
    "txn_high":   {"label":"উচ্চ লেনদেন (১লাখ+)",     "color":"#991b1b","icon":"🔴"},
    "txn_med":    {"label":"মাঝারি লেনদেন (৫০K-১লাখ)", "color":"#dc2626","icon":"🟠"},
    "txn_small":  {"label":"ছোট লেনদেন (১০K-৫০K)",     "color":"#ea580c","icon":"🟡"},
    "identity":   {"label":"পরিচয় / তথ্য",             "color":"#0369a1","icon":"👤"},
    "love_family":{"label":"পারিবারিক / আবেগ",          "color":"#db2777","icon":"❤️"},
    "travel":     {"label":"ভ্রমণ / ভিসা / ওমরা",      "color":"#0284c7","icon":"✈️"},
    "courier":    {"label":"কুরিয়ার / ডেলিভারি",        "color":"#6d28d9","icon":"📦"},
    "shopping":   {"label":"কেনাকাটা / বিল",            "color":"#059669","icon":"🛒"},
    "service":    {"label":"সার্ভিস মেসেজ",              "color":"#475569","icon":"🔔"},
    "personal":   {"label":"ব্যক্তিগত কথোপকথন",        "color":"#64748b","icon":"💬"},
}
TIER_COL = {"high":"#991b1b","med":"#dc2626","small":"#ea580c","":"#94a3b8"}
TIER_BG  = {"high":"#fef2f2","med":"#fff5f5","small":"#fff7ed","":"#f8fafc"}
TIER_LBL = {"high":"🔴 উচ্চ (১লাখ+)","med":"🟠 মাঝারি (৫০K-১লাখ)","small":"🟡 ছোট (১০K-৫০K)"}
MONTH_BN = {"01":"জানুয়ারি","02":"ফেব্রুয়ারি","03":"মার্চ","04":"এপ্রিল",
            "05":"মে","06":"জুন","07":"জুলাই","08":"আগস্ট","09":"সেপ্টেম্বর",
            "10":"অক্টোবর","11":"নভেম্বর","12":"ডিসেম্বর"}

def build_bank_summary(sms_list, extra_balances=None):
    if extra_balances is None: extra_balances = {}
    from collections import defaultdict
    summ = defaultdict(lambda:{
        "service":"","txn_count":0,"total_in":0.0,"total_out":0.0,
        "last_balance":0.0,"last_balance_date":"","max_txn":0.0,"numbers":set()
    })
    SVC_COL = {
        "bKash":"#e11d48","Nagad":"#ea580c","Rocket":"#7c3aed","Upay":"#16a34a",
        "Islami Bank":"#1d4ed8","Trust Bank":"#0369a1","Trust Card":"#7c3aed",
        "Sonali Bank":"#047857","BRAC Bank":"#b45309","MetLife":"#0284c7",
        "Dutch-Bangla Bank (DBBL)":"#6d28d9",
        "Mutual Trust Bank (MTB)":"#0891b2","Mutual Trust Bank":"#0891b2",
        "Southeast Bank":"#7c2d12","Prime Bank":"#1e40af","Eastern Bank":"#065f46",
        "City Bank":"#9f1239",
    }
    for s in sms_list:
        ai = s.get("acct",{})
        if not ai or not ai.get("service"): continue
        svc = ai["service"]; d = summ[svc]
        d["service"] = svc

        if s["tier"]:
            d["txn_count"] += 1
            amt = ai.get("amount",0) or 0
            if "Received" in (ai.get("direction","") or ""): d["total_in"]  += amt
            else:                                             d["total_out"] += amt
            if amt > d["max_txn"]: d["max_txn"] = amt

        # last_balance — তারিখ অনুযায়ী সর্বশেষ balance রাখো
        bal = ai.get("balance",0) or 0
        if bal > 0:
            sms_date = str(s.get("date",""))
            if sms_date > d["last_balance_date"]:
                d["last_balance"] = bal
                d["last_balance_date"] = sms_date

        if ai.get("other_number"):
            d["numbers"].add(ai["other_number"])

    result = {}
    for svc, d in summ.items():
        result[svc] = {**d, "numbers": list(d["numbers"])[:4],
                       "color": SVC_COL.get(svc, "#2563eb")}

    # extra_balances — transaction ছাড়া bank balance (যেমন MTB balance notification)
    for svc, eb in extra_balances.items():
        if svc not in result:
            # নতুন entry — শুধু balance আছে, transaction নেই
            result[svc] = {
                "service": svc, "txn_count": 0, "total_in": 0.0, "total_out": 0.0,
                "last_balance": eb["balance"], "last_balance_date": eb["date"],
                "max_txn": 0.0, "numbers": [],
                "color": SVC_COL.get(svc, "#2563eb")
            }
        else:
            # Already exists — update balance if extra is more recent
            if eb["date"] > result[svc].get("last_balance_date",""):
                result[svc]["last_balance"] = eb["balance"]
                result[svc]["last_balance_date"] = eb["date"]
    return result


def generate_html(sms_list, stats, out_path, tn=""):
    now = datetime.now().strftime("%d %B %Y, %I:%M %p")

    # bank_info — sidebar-এর জন্য, card-এ দেখাবে না
    display_list = [s for s in sms_list if "bank_info" not in s.get("categories",[])]

    months    = sorted({s["month"] for s in display_list if s["month"]})
    mnt_cnt   = Counter(s["month"] for s in display_list)
    tier_cnt  = Counter(s["tier"] for s in display_list)
    cat_cnt   = Counter(c for s in display_list for c in s["categories"])
    bank_summ = build_bank_summary(sms_list)
    smsmt_count = sum(1 for s in display_list if s.get("direction") == "SMSMT")
    smsmo_count = sum(1 for s in display_list if s.get("direction") == "SMSMO")
    MONTH_EN = {"01":"January","02":"February","03":"March","04":"April","05":"May",
                "06":"June","07":"July","08":"August","09":"September",
                "10":"October","11":"November","12":"December"}
    CAT_EN = {"txn_high":"High Transaction (100K+)","txn_med":"Medium Transaction (50K–100K)",
              "txn_small":"Small Transaction (10K–50K)","identity":"Identity / Profile Info",
              "love_family":"Personal / Family","travel":"Travel / Visa / Umrah",
              "courier":"Courier / Delivery","shopping":"Shopping / Billing",
              "service":"Service Messages","personal":"Personal Conversation"}
    def esc(t): return hl.escape(str(t))
    def cat_tags(cats):
        o=""
        for c in cats:
            m=CAT_META.get(c,{}); col=m.get("color","#555")
            lbl=CAT_EN.get(c,m.get("label",c))
            o+=(f'<span class="ctag" style="background:{col}12;border:1px solid {col}35;color:{col}">'
                f'{m.get("icon","🔍")} {lbl}</span>')
        return o
    def hl_body(body, amt):
        r=esc(body)
        if amt>=TIER_SMALL:
            for v in [f'{int(amt):,}', str(int(amt)), f'{amt:,.2f}']:
                try: r=re.sub(re.escape(v),f'<mark>{v}</mark>',r,count=1,flags=re.I)
                except: pass
        return r
    # Cards
    cards=""
    for i,s in enumerate(display_list,1):
        t=s["tier"] or ""; bdr=TIER_COL.get(t,"#cbd5e1"); bg=TIER_BG.get(t,"#f8fafc")
        ai=s.get("acct",{}); direction=s.get("direction","")
        TIER_EN_LBL={"high":"🔴 High (100K+)","med":"🟠 Medium (50K–100K)","small":"🟡 Small (10K–50K)"}
        amt_tag=(f'<span class="amt-tag" style="background:{bdr}">💰 Tk {s["amount"]:,.2f}</span>'
                 if s["amount"]>=TIER_SMALL else "")
        tier_badge=(f'<span class="rbadge" style="background:{bdr}">{TIER_EN_LBL.get(t,"")}</span>' if t else "")
        dir_badge=""
        if direction=="SMSMT": dir_badge='<span class="dir-in">● IN</span>'
        elif direction=="SMSMO": dir_badge='<span class="dir-out">● OUT</span>'
        acct_mini=""
        if ai and ai.get("service"):
            pts=[]
            if ai.get("direction"):    pts.append(f'<b>{ai["direction"]}</b>')
            if ai.get("other_number"): pts.append(f'📱 {ai["other_number"]}')
            if ai.get("trx_id"):       pts.append(f'TrxID <code>{ai["trx_id"]}</code>')
            if ai.get("balance"):      pts.append(f'Balance: Tk {ai["balance"]:,.2f}')
            if pts: acct_mini=f'<div class="acct-mini">{" &nbsp;·&nbsp; ".join(pts)}</div>'
        merged_tag=(f'<span class="merged-tag">🔗 {s["parts"]} parts</span>' if s.get("parts",1)>1 else "")
        cards+=f'''
        <div class="scard" data-tier="{t}" data-month="{s["month"]}" data-cats="{" ".join(s["categories"])}" data-dir="{direction}">
          <div class="sc-head">
            <div class="sc-left">
              <span class="snum">#{i}</span>
              <div>
                <div class="snm">{dir_badge} 📱 {esc(s["sender"])}</div>
                <div class="sdate">🕐 {esc(s["date"])} {merged_tag}</div>
              </div>
            </div>
            <div class="sc-right">{tier_badge}{amt_tag}</div>
          </div>
          {acct_mini}
          <div class="sbody" style="background:{bg}">{hl_body(s["body"],s["amount"])}</div>
          <div class="sc-foot">{cat_tags(s["categories"])}</div>
        </div>'''
    if not cards: cards='<div class="no-result"><div style="font-size:48px">✅</div><div>No SMS records found</div></div>'
    # Sidebar
    cat_sb=""
    for ck,meta in CAT_META.items():
        cnt=cat_cnt.get(ck,0)
        if not cnt: continue
        col=meta["color"]; lbl=CAT_EN.get(ck,meta.get("label",ck))
        pct=int(cnt/len(display_list)*100) if display_list else 0
        cat_sb+=(f'<div class="stat-row"><span>{meta["icon"]}</span>'
                 f'<span class="slbl">{lbl}</span>'
                 f'<div class="sbar-w"><div class="sbar" style="width:{pct}%;background:{col}"></div></div>'
                 f'<span class="scnt">{cnt}</span></div>')
    bank_sb=""
    for svc,d in bank_summ.items():
        col=d["color"]; nums="".join(f'<div class="acct-num">📱 {n}</div>' for n in d["numbers"])
        # txn_count=0 হলেও সব ব্যাংক দেখাবে যদি balance থাকে
        if d["txn_count"] == 0 and d["last_balance"] == 0: continue
        bal_html = f'<div class="ag-bal">💰 Last Balance: Tk {d["last_balance"]:,.2f}</div>' if d["last_balance"] else ""
        bank_sb+=f'''<div class="acct-card" style="border-left:3px solid {col}">
          <div class="acct-svc" style="color:{col}">💳 {esc(svc)}</div>
          <div class="acct-grid4">
            <div class="ag-c"><div class="ag-v">{d["txn_count"]}</div><div class="ag-l">Transactions</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#16a34a;font-size:11px">৳{d["total_in"]:,.0f}</div><div class="ag-l">Received</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#dc2626;font-size:11px">৳{d["total_out"]:,.0f}</div><div class="ag-l">Sent</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#d97706;font-size:10px">৳{d["max_txn"]:,.0f}</div><div class="ag-l">Highest</div></div>
          </div>
          {bal_html}
          {nums}</div>'''
    if not bank_sb: bank_sb='<div style="color:#9ca3af;font-size:11px;padding:4px 0">No MFS/Bank data</div>'
    def fb(lbl,js,col=None):
        st=f' style="border-color:{col}"' if col else ""
        return f'<button class="fbtn"{st} onclick="{js}">{lbl}</button>'
    fbtns =f'<button class="fbtn active" onclick="fAll(this)">All ({len(display_list)})</button>'
    fbtns+=fb(f'🔴 High ({tier_cnt.get("high",0)})',"fTier(this,'high')","#991b1b")
    fbtns+=fb(f'🟠 Medium ({tier_cnt.get("med",0)})',"fTier(this,'med')","#dc2626")
    fbtns+=fb(f'🟡 Small ({tier_cnt.get("small",0)})',"fTier(this,'small')","#ea580c")
    if smsmt_count: pass  # SMSMT shown in summary only, not as filter button
    if smsmo_count: pass  # SMSMO shown in summary only, not as filter button
    for ck in ["identity","love_family","travel","courier","shopping","personal"]:
        c=cat_cnt.get(ck,0)
        if not c: continue
        meta=CAT_META.get(ck,{}); lbl=CAT_EN.get(ck,meta.get("label",ck))
        fbtns+=fb(f'{meta.get("icon","")} {lbl} ({c})',f"fCat(this,'{ck}')",meta.get("color"))
    mbtns="".join(f'<button class="mbtn" onclick="fMonth(this,\'{m}\')">'
                  f'{MONTH_EN.get(m.split("-")[1],m)} {m.split("-")[0]} ({mnt_cnt[m]})</button>' for m in months)
    tnb=f'<div class="tnbadge">📱 {esc(tn)}</div>' if tn else ""

    html=f'''<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SMS Intelligence Report — {esc(tn) if tn else "All Numbers"}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Bengali:wght@400;600;700&family=JetBrains+Mono:wght@500&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#f4f5f7;color:#1a1a2e;font-family:'Outfit',sans-serif;font-size:14px}}
.topbar{{background:linear-gradient(135deg,#1a0505 0%,#7b0d0d 45%,#c0392b 100%);
  padding:0 28px;height:66px;display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:200;box-shadow:0 2px 12px rgba(0,0,0,.2)}}
.tb-left{{display:flex;align-items:center;gap:12px}}
.tb-logo{{width:40px;height:40px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);
  border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:20px}}
.tb-h1{{font-size:17px;font-weight:800;color:#fff;letter-spacing:-.3px}}
.tb-sub{{font-size:10px;color:rgba(255,255,255,.6);margin-top:1px}}
.tb-right{{display:flex;align-items:center;gap:10px}}
.cfbadge{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:#fff;
  font-size:9px;font-weight:700;letter-spacing:2px;padding:3px 10px;border-radius:20px}}
.tnbadge{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:#fff;
  font-size:11px;padding:3px 10px;border-radius:20px;font-family:'JetBrains Mono',monospace}}
.dl-top{{display:flex;align-items:center;gap:6px;background:#fff;color:#c0392b;
  font-size:12px;font-weight:700;padding:7px 16px;border-radius:20px;cursor:pointer;
  border:none;font-family:'Outfit',sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.15);transition:all .2s;
  text-decoration:none}}
.dl-top:hover{{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.2)}}
.summary-sheet{{background:linear-gradient(135deg,#0f172a,#1e293b);
  padding:28px 28px 24px;border-bottom:3px solid #c0392b}}
.ss-header{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:20px;gap:20px;flex-wrap:wrap}}
.ss-title{{font-size:22px;font-weight:800;color:#fff;letter-spacing:-.3px;margin-bottom:4px}}
.ss-sub{{font-size:12px;color:rgba(255,255,255,.5)}}
.ss-meta{{text-align:right;font-size:11px;color:rgba(255,255,255,.4);font-family:'JetBrains Mono',monospace;line-height:1.8}}
.ss-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}}
.ss-box{{background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:12px;text-align:center}}
.ss-num{{font-size:26px;font-weight:800;font-family:'JetBrains Mono',monospace;line-height:1;margin-bottom:3px}}
.ss-lbl{{font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.5px}}
.ss-box.red{{border-color:rgba(192,57,43,.5);background:rgba(192,57,43,.15)}}
.ss-box.grn{{border-color:rgba(22,163,74,.5);background:rgba(22,163,74,.1)}}
.ss-box.blu{{border-color:rgba(37,99,235,.5);background:rgba(37,99,235,.1)}}
.layout{{display:grid;grid-template-columns:280px 1fr;min-height:calc(100vh - 200px)}}
.sidebar{{background:#fff;border-right:1px solid #e8eaed;padding:16px 14px;
  position:sticky;top:66px;height:calc(100vh - 66px);overflow-y:auto;
  scrollbar-width:thin;scrollbar-color:#cbd5e1 transparent}}
.stitle{{font-size:9px;font-weight:700;letter-spacing:1.5px;color:#9ca3af;
  text-transform:uppercase;margin:16px 0 8px;padding-top:12px;border-top:1px solid #f1f5f9}}
.stitle:first-child{{margin-top:0;padding-top:0;border-top:none}}
.sgrid{{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:6px}}
.sc2{{background:#f8fafc;border:1px solid #e8eaed;border-radius:8px;padding:10px;text-align:center}}
.sc2.full{{grid-column:1/-1}}
.sc2-num{{font-size:22px;font-weight:700;font-family:'JetBrains Mono',monospace;line-height:1;margin-bottom:2px}}
.sc2-lbl{{font-size:9px;color:#9ca3af;text-transform:uppercase;letter-spacing:.3px}}
.tinfo{{background:#fffbeb;border:1px solid #fde68a;border-radius:7px;padding:8px 10px;font-size:11px;color:#78350f;line-height:1.7;margin-bottom:8px}}
.tinfo b{{color:#92400e}}
.stat-row{{display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid #f1f5f9}}
.stat-row:last-child{{border-bottom:none}}
.slbl{{font-size:10px;color:#475569;flex:1;line-height:1.4}}
.sbar-w{{width:48px;height:3px;background:#f1f5f9;border-radius:2px;overflow:hidden}}
.sbar{{height:100%;border-radius:2px}}
.scnt{{font-size:10px;color:#9ca3af;font-family:'JetBrains Mono',monospace;min-width:18px;text-align:right}}
.acct-card{{background:#f8fafc;border:1px solid #e8eaed;border-radius:8px;padding:10px 11px;margin-bottom:7px}}
.acct-svc{{font-size:13px;font-weight:700;margin-bottom:7px}}
.acct-grid4{{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:6px}}
.ag-c{{background:#fff;border:1px solid #e8eaed;border-radius:5px;padding:5px 4px;text-align:center}}
.ag-v{{font-size:13px;font-weight:700;font-family:'JetBrains Mono',monospace;line-height:1}}
.ag-l{{font-size:9px;color:#9ca3af;margin-top:2px}}
.ag-bal{{font-size:10px;color:#16a34a;font-family:'JetBrains Mono',monospace;
  background:#f0fdf4;border:1px solid #bbf7d0;border-radius:5px;padding:4px 7px;margin-bottom:4px}}
.acct-num{{font-size:10px;color:#2563eb;font-family:'JetBrains Mono',monospace;padding:2px 0;border-bottom:1px solid #f1f5f9}}
.acct-num:last-child{{border-bottom:none}}
.fbar{{background:#fff;border-bottom:1px solid #e8eaed;padding:10px 22px;
  display:flex;align-items:center;gap:7px;flex-wrap:wrap;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.flbl{{font-size:11px;color:#9ca3af;white-space:nowrap}}
.fbtn,.mbtn{{padding:4px 11px;border-radius:20px;border:1.5px solid #e2e8f0;
  background:#fff;color:#475569;font-size:11px;cursor:pointer;font-family:inherit;font-weight:500}}
.fbtn:hover,.mbtn:hover{{border-color:#c0392b;color:#c0392b;background:#fef2f2}}
.fbtn.active{{background:#1a1a2e;border-color:#1a1a2e;color:#fff;font-weight:600}}
.mbtn.active{{background:#c0392b;border-color:#c0392b;color:#fff;font-weight:600}}
.fsep{{width:1px;height:20px;background:#e2e8f0;margin:0 3px}}
.sbox{{margin-left:auto;padding:5px 12px;background:#f8fafc;border:1.5px solid #e2e8f0;
  border-radius:20px;color:#1a1a2e;font-size:11px;outline:none;width:210px}}
.sbox:focus{{border-color:#c0392b;background:#fff}}
.content{{padding:18px 22px;background:#f4f5f7}}
.rh{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}}
.rinfo{{font-size:12px;color:#64748b}}.rinfo span{{color:#c0392b;font-weight:600}}
.scards{{display:flex;flex-direction:column;gap:9px}}
.scard{{background:#fff;border:1px solid #e8eaed;border-left:5px solid #cbd5e1;
  border-radius:10px;padding:13px 16px;transition:all .2s}}
.scard:hover{{box-shadow:0 4px 14px rgba(0,0,0,.07);transform:translateY(-1px)}}
.sc-head{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:9px;gap:10px}}
.sc-left{{display:flex;align-items:center;gap:9px;min-width:0}}
.snum{{background:#f1f5f9;border:1px solid #e2e8f0;color:#9ca3af;
  font-family:'JetBrains Mono',monospace;font-size:10px;padding:2px 6px;border-radius:4px;flex-shrink:0}}
.snm{{font-family:'JetBrains Mono',monospace;font-size:13px;color:#2563eb;font-weight:600;
  display:flex;align-items:center;gap:6px;flex-wrap:wrap}}
.sdate{{font-size:10px;color:#9ca3af;margin-top:2px}}
.sc-right{{display:flex;gap:5px;align-items:center;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0}}
.rbadge{{font-size:9px;font-weight:700;padding:3px 9px;border-radius:20px;color:#fff;letter-spacing:.5px}}
.amt-tag{{color:#fff;font-size:11px;font-family:'JetBrains Mono',monospace;padding:3px 10px;border-radius:20px;font-weight:700}}
.merged-tag{{background:#dbeafe;color:#1e40af;font-size:9px;padding:2px 7px;border-radius:10px;font-weight:600;margin-left:5px}}
.dir-in{{background:#dcfce7;color:#15803d;font-size:9px;font-weight:700;padding:2px 7px;border-radius:10px;letter-spacing:.5px}}
.dir-out{{background:#dbeafe;color:#1d4ed8;font-size:9px;font-weight:700;padding:2px 7px;border-radius:10px;letter-spacing:.5px}}
.acct-mini{{background:#f0f9ff;border:1px solid #bae6fd;border-radius:6px;
  padding:6px 10px;margin-bottom:8px;font-size:11px;color:#0369a1;
  font-family:'JetBrains Mono',monospace;line-height:1.7}}
.acct-mini code{{background:#dbeafe;color:#1e40af;padding:1px 4px;border-radius:3px}}
.sbody{{font-size:13px;line-height:1.9;color:#374151;font-family:'Noto Serif Bengali',serif;
  padding:10px 13px;border-radius:7px;border:1px solid #e2e8f0;white-space:pre-wrap;word-break:break-word}}
.sbody mark{{background:#fef08a;color:#713f12;padding:1px 4px;border-radius:3px;
  font-weight:700;font-family:'JetBrains Mono',monospace}}
.sc-foot{{display:flex;flex-wrap:wrap;gap:5px;margin-top:8px}}
.ctag{{font-size:10px;padding:3px 9px;border-radius:20px;font-weight:500}}
.no-result{{text-align:center;padding:60px;color:#9ca3af}}
.footer{{background:#1a1a2e;padding:16px 28px;
  display:flex;justify-content:space-between;font-size:10px;color:rgba(255,255,255,.35)}}
@media print{{.topbar{{position:static}}.summary-sheet{{position:static}}
  .sidebar{{position:static;height:auto}}.layout{{grid-template-columns:1fr}}
  .fbar{{display:none}}.scard{{break-inside:avoid}}}}
</style></head><body>

<div class="topbar">
  <div class="tb-left">
    <div class="tb-logo">🔍</div>
    <div>
      <div class="tb-h1">SMS Intelligence Report</div>
      <div class="tb-sub">CDR Forensic Analysis · Bangladesh Law Enforcement</div>
    </div>
  </div>
  <div class="tb-right">
    <div class="cfbadge">CONFIDENTIAL</div>
    {tnb}
    <button class="dl-top" onclick="downloadHTML()">⬇️ Download HTML</button>
    <button class="dl-top" onclick="window.print()" style="margin-left:0">🖨 Print</button>
  </div>
</div>

<div class="summary-sheet">
  <div class="ss-header">
    <div>
      <div class="ss-title">Forensic SMS Analysis Report</div>
      <div class="ss-sub">SMS Intelligence Dashboard &nbsp;·&nbsp; {"Target: " + esc(tn) if tn else "All Numbers"}</div>
    </div>
    <div class="ss-meta">Generated: {now}<br>{"Target Number: " + esc(tn) + "<br>" if tn else ""}Total Records in File: {stats["total"]}</div>
  </div>
  <div class="ss-grid">
    <div class="ss-box red" style="grid-column:1/-1">
      <div class="ss-num" style="color:#f87171;font-size:32px">{len(display_list)}</div>
      <div class="ss-lbl">Total Identified SMS</div>
    </div>
    <div class="ss-box red"><div class="ss-num" style="color:#fca5a5">{tier_cnt.get("high",0)}</div><div class="ss-lbl">High Value (100K+)</div></div>
    <div class="ss-box red"><div class="ss-num" style="color:#fdba74">{tier_cnt.get("med",0)}</div><div class="ss-lbl">Medium (50K–100K)</div></div>
    <div class="ss-box red"><div class="ss-num" style="color:#fde047">{tier_cnt.get("small",0)}</div><div class="ss-lbl">Small (10K–50K)</div></div>
    <div class="ss-box grn"><div class="ss-num" style="color:#86efac">{smsmt_count}</div><div class="ss-lbl">SMSMT (Incoming)</div></div>
    <div class="ss-box blu"><div class="ss-num" style="color:#93c5fd">{smsmo_count}</div><div class="ss-lbl">SMSMO (Outgoing)</div></div>
    <div class="ss-box"><div class="ss-num" style="color:#cbd5e1">{stats["filtered"]}</div><div class="ss-lbl">Operator Filtered</div></div>
  </div>
</div>

<div class="layout">
<div class="sidebar">
  <div class="stitle">Thresholds</div>
  <div class="tinfo">🔴 <b>100K+</b> → High<br>🟠 <b>50K–100K</b> → Medium<br>🟡 <b>10K–50K</b> → Small</div>
  <div class="stitle">Overview</div>
  <div class="sgrid">
    <div class="sc2 full"><div class="sc2-num" style="color:#c0392b">{len(display_list)}</div><div class="sc2-lbl">Total Identified</div></div>
    <div class="sc2"><div class="sc2-num" style="color:#991b1b">{tier_cnt.get("high",0)}</div><div class="sc2-lbl">High (100K+)</div></div>
    <div class="sc2"><div class="sc2-num" style="color:#dc2626">{tier_cnt.get("med",0)}</div><div class="sc2-lbl">Medium</div></div>
    <div class="sc2"><div class="sc2-num" style="color:#ea580c">{tier_cnt.get("small",0)}</div><div class="sc2-lbl">Small (10K–50K)</div></div>
    <div class="sc2"><div class="sc2-num" style="color:#16a34a">{smsmt_count}</div><div class="sc2-lbl">SMSMT In</div></div>
    <div class="sc2"><div class="sc2-num" style="color:#2563eb">{smsmo_count}</div><div class="sc2-lbl">SMSMO Out</div></div>
  </div>
  <div class="stitle">Categories</div>
  {cat_sb or '<div style="color:#9ca3af;font-size:11px;padding:4px 0">No data</div>'}
  <div class="stitle">Bank / MFS Summary</div>
  {bank_sb}
</div>
<div style="display:flex;flex-direction:column;">
  <div class="fbar">
    <span class="flbl">Filter:</span>
    {fbtns}
    <div class="fsep"></div>
    <span class="flbl">Month:</span>
    <button class="mbtn active" onclick="fMonthAll(this)">All</button>
    {mbtns}
    <input class="sbox" type="text" placeholder="🔍 Search messages..." oninput="fSearch(this.value)">
  </div>
  <div class="content">
    <div class="rh"><div class="rinfo">Showing: <span id="vc">{len(display_list)}</span> SMS records</div></div>
    <div class="scards" id="cards">{cards}</div>
  </div>
</div>
</div>
<div class="footer">
  <span>🔒 Confidential — For Law Enforcement Use Only</span>
  <span>{"Target: " + esc(tn) + " | " if tn else ""}Generated: {now} | SMS Intelligence Dashboard</span>
</div>
<script>
const allC=document.querySelectorAll('.scard');
let aT='all',aCat='all',aM='all',aD='all',sQ='';
function apply(){{let v=0;allC.forEach(c=>{{const ok=(aT==='all'||c.dataset.tier===aT)&&(aCat==='all'||(c.dataset.cats||'').includes(aCat))&&(aM==='all'||c.dataset.month===aM)&&(aD==='all'||c.dataset.dir===aD)&&(!sQ||c.innerText.toLowerCase().includes(sQ.toLowerCase()));c.style.display=ok?'':'none';if(ok)v++;}});document.getElementById('vc').textContent=v;}}
function act(b,sel){{document.querySelectorAll(sel).forEach(x=>x.classList.remove('active'));b.classList.add('active');}}
function fAll(b){{aT='all';aCat='all';aD='all';act(b,'.fbtn');apply();}}
function fTier(b,t){{aT=t;aCat='all';aD='all';act(b,'.fbtn');apply();}}
function fCat(b,c){{aCat=c;aT='all';aD='all';act(b,'.fbtn');apply();}}
function fDir(b,d){{aD=d;aT='all';aCat='all';act(b,'.fbtn');apply();}}
function fMonthAll(b){{aM='all';act(b,'.mbtn');apply();}}
function fMonth(b,m){{aM=m;act(b,'.mbtn');apply();}}
function fSearch(v){{sQ=v;apply();}}
function downloadHTML(){{const blob=new Blob([document.documentElement.outerHTML],{{type:'text/html'}});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='SMS_Intelligence_Report_{esc(tn) if tn else "report"}.html';a.click();}}
</script>
</body></html>'''
    with open(out_path,'w',encoding='utf-8') as f: f.write(html)

def process_file(input_path, html_out, docx_out=None, target_number=""):
    ext = os.path.splitext(input_path)[1].lower()
    raw = load_xlsx(input_path) if ext!='.csv' else load_csv_f(input_path)

    # target_number দিলে সেই নম্বরের সাথে সংশ্লিষ্ট SMS ফিল্টার
    if target_number:
        tn_clean = re.sub(r'[^\d]','', target_number.strip())
        # variants: 01XXXXXXXXX ↔ 8801XXXXXXXXX
        tn_variants = {tn_clean}
        if tn_clean.startswith('01'): tn_variants.add('880'+tn_clean)
        if tn_clean.startswith('880') and len(tn_clean) >= 13: tn_variants.add(tn_clean[3:])

        def is_related(item):
            # sender (BParty) বা শুধু যদি body-তে নম্বর থাকে
            sender_digits = re.sub(r'[^\d]','', str(item.get("sender","")))
            # AParty match (ফাইলে AParty = মালিকের নম্বর) — raw_row এ আছে
            aparty_digits = re.sub(r'[^\d]','', str(item.get("aparty","")))
            return (any(v in sender_digits for v in tn_variants)
                    or any(v in aparty_digits for v in tn_variants)
                    or any(v in item.get("body","") for v in tn_variants))
        raw = [r for r in raw if is_related(r)]

    merged_ = merge(raw)
    filtered_cnt = 0
    results = []
    seen = set()

    for item in merged_:
        if not item["body"].strip(): continue
        if is_filtered(item["sender"], item["body"]): filtered_cnt+=1; continue
        body_key  = item["body"].strip()
        dt_str    = str(item["date"])[:13]   # YYYY-MM-DD HH — ১ ঘন্টার মধ্যে same body একটাই
        dedup_key = (str(item["sender"]).strip().lower(), body_key, dt_str)
        if dedup_key in seen: filtered_cnt+=1; continue
        seen.add(dedup_key)
        res = analyze_one(item["sender"], item["body"], item["date"],
                          item.get("direction",""))
        if res:
            res["parts"] = item.get("parts",1)
            results.append(res)

    stats = {"total":len(raw),"after_merge":len(merged_),"filtered":filtered_cnt}

    # Target number — filename থেকে বা parameter থেকে
    tn = re.sub(r'[^\d]','', target_number) if target_number else ""
    if not tn:
        base = os.path.splitext(os.path.basename(input_path))[0]
        m = re.search(r'(\d{10,})', base)
        tn = m.group(1) if m else ""
    if tn.startswith('880') and len(tn)>11: tn='0'+tn[3:]

    generate_html(results, stats, html_out, tn)
    tc = Counter(s["tier"] for s in results)
    display_count = sum(1 for s in results if "bank_info" not in s.get("categories",[]))
    return {"target_number":tn,"total_sms":len(raw),"after_merge":len(merged_),
            "filtered":filtered_cnt,"total_shown":display_count,
            "high":tc.get("high",0),"med":tc.get("med",0),"small":tc.get("small",0)}

def main():
    if len(sys.argv)<2: print("python3 sms_analyzer.py input.xlsx [output.html]"); sys.exit(1)
    inp=sys.argv[1]; out=sys.argv[2] if len(sys.argv)>2 else re.sub(r'\.(xlsx?|csv)$','.html',inp)
    r=process_file(inp,out)
    print(f"✅ মোট: {r['total_shown']} | 🔴উচ্চ:{r['high']} | 🟠মাঝারি:{r['med']} | 🟡ছোট:{r['small']}")
    print(f"📄 {os.path.abspath(out)}")

if __name__=="__main__": main()
