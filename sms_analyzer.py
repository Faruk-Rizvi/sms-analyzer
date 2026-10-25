#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS Forensic Analyzer v12 — Bangladesh Law Enforcement
তিন স্তরের লেনদেন + ব্যক্তিগত/সেবা SMS + বিদেশী ব্যাংক SMS
"""
import openpyxl, re, html as hl, sys, os, csv, io
from datetime import datetime
from collections import Counter

TIER_SMALL = 10_000
TIER_MED   = 50_000
TIER_HIGH  = 100_000

# ── BD নম্বর ────────────────────────────────────────────────
# CDR export-এ "+880 17XXXXXXXX" বা "017-12345678" ফরম্যাটেও আসতে পারে
BD_RE = re.compile(r'^\+?8801[3-9]\d{8}$|^01[3-9]\d{8}$')
def is_bd(s):
    # spaces ও hyphens সরিয়ে normalize করো, তারপর match করো
    normalized = re.sub(r'[\s\-]', '', str(s).strip())
    return bool(BD_RE.match(normalized))

# ── Sender type ────────────────────────────────────────────
# ── বিদেশী ব্যাংক sender names ────────────────────────────
# AlRajhi Malaysia, CIMB, Maybank, etc.
FOREIGN_BANK_KW = [
    'arj', 'alrajhi', 'al rajhi', 'al-rajhi',   # AlRajhi Bank Malaysia
    'maybank', 'cimb', 'rhb bank', 'rhb',         # Malaysia banks
    'public bank', 'hong leong',                   # Malaysia banks
    'dbs bank', 'ocbc', 'uob',                     # Singapore banks
    'hdfc', 'icici', 'sbi', 'axis bank',          # India banks
    'emirates nbd', 'adcb', 'fab bank',           # UAE banks
]

FOREIGN_BANK_BODY_RE = re.compile(
    r'\balrajhi\b|\bal\s*rajhi\b'
    r'|maybank(?:\s+\w+)?:'
    r'|cimb(?:\s+bank)?:'
    r'|your\s+card\s+ending\s+\d{4}\s+(?:was\s+used|has\s+been)'
    r'|(?:MYR|SGD|AED|INR|USD|GBP|EUR)\s*[\d,]+\.\d{2}\s+(?:was\s+made|spent|debited|charged)'
    r'|spending\s+(?:MYR|SGD|AED|INR)\s*[\d,]+',
    re.I
)

def is_foreign_bank(sender, body=''):
    sl = str(sender).lower().strip()
    # SBIBNK with BDT = SBI Bangladesh branch — NOT foreign
    if sl == 'sbibnk':
        if body and re.search(r'\bbdt\b', body, re.I):
            return False
        return False  # SBIBNK সবসময় বাংলাদেশের SBI শাখা হিসেবে ধরা হবে
    if any(k in sl for k in FOREIGN_BANK_KW): return True
    if body and FOREIGN_BANK_BODY_RE.search(body): return True
    return False

# ── বিদেশী মুদ্রা amount parser ───────────────────────────
FOREIGN_AMT_RE = re.compile(
    r'(?:MYR|SGD|AED|INR|USD|GBP|EUR)\s*([\d,]+\.?\d*)'
    r'|spending\s+(?:MYR|SGD|AED|INR|USD)\s*([\d,]+\.?\d*)',
    re.I
)
FOREIGN_CCY_RE = re.compile(r'(?<![A-Za-z])(MYR|SGD|AED|INR|USD|GBP|EUR)(?![A-Za-z])', re.I)

def get_foreign_amount(body):
    """বিদেশী মুদ্রার পরিমাণ বের করো।"""
    for m in FOREIGN_AMT_RE.finditer(body):
        v = m.group(1) or m.group(2)
        try:
            f = float(v.replace(',',''))
            if f > 0: return f
        except (ValueError, TypeError): pass
    return 0.0

def get_foreign_currency(body):
    m = FOREIGN_CCY_RE.search(body)
    return m.group(1).upper() if m else ''


MFS_KW    = ['bkash','b-kash','nagad','rocket','upay','mcash','tap','surecash',
             'binimoy','binimoypay']   # Bangladesh Bank interoperable MFS (চালু ২০২১)
BANK_KW   = ['ibbl','ibbl .','ibbl.','islamibank','islami bank',
             'trustbank','trust bank','trustcard',
             'sonali','janata bank','agrani bank','rupali bank',
             'brac bank','bracbank','brac-bank',
             'city bank','citybank','ebl','eastern bank','prime bank',
             'mutual trust','mtb.','mtb ','\bmtb\b',
             'southeast bank','ab bank','ucb','hsbc','standard chartered',
             'ific','idlc','metlife','dutch-bangla','dutch bangla','dbbl',
             'nexuspay','nexus pay','sonali bank',
             'standardbnk','standard bank','standard islami',
             'dutchbangla',                                    # NexusPay/DBBL no-separator variant
             'stanchart','stancharbd','stanchartbd',           # Standard Chartered BD
             # ── মিসিং ইসলামি ব্যাংক (বাগ ৭) ──────────────────────
             'al-arafah','al arafah','alarafah',            # Al-Arafah Islami Bank
             'shahjalal','sjibl',                           # Shahjalal Islami Bank
             'exim bank','eximbank','exim-bank',            # EXIM Bank (Islamic mode)
             'first security','fsibl','fsib',               # First Security Islami Bank
             'union bank','unionbank',                      # Union Bank (Islamic)
             'global islami','globalislami',                # Global Islami Bank
             # ── মিসিং বেসরকারি ব্যাংক (বাগ ৮) ───────────────────
             'mercantile bank','mercanbank','mercentilebank','mercantile', # Mercantile Bank
             'one bank','onebank',                          # ONE Bank
             'nrb bank','nrbbank','nrb commercial',         # NRB Bank / NRB Commercial
             'midland bank','midlandbank',                  # Midland Bank
             'community bank','communitybank',              # Community Bank Bangladesh
             'padma bank','padmabank',                      # Padma Bank
             'bangladesh commerce','bdcommerce',            # Bangladesh Commerce Bank
             'jamuna bank','jamunabank',                    # Jamuna Bank
             'modhumoti bank','modhumotibank',              # Modhumoti Bank
             'shimanto bank','shimanto',                    # Shimanto Bank
             'bengal commercial',                           # Bengal Commercial Bank
             'sbibnk','sbi bangladesh','state bank of india', # SBI Bangladesh branch
             ]

# Short-code senders that are actually banks (not to be filtered)
BANK_SHORTCODES = {
    # Bank shortcodes
    '16216','16219','16240','16232',
    # MFS shortcodes — bKash, Nagad, Rocket
    '16247',   # bKash
    '16167',   # Nagad
    '16262',   # Rocket/DBBL
    '26969',   # Upay
}

# Promotional / system numbers — সম্পূর্ণ filter
PROMO_NUMBERS = {
    '8801886016460','01886016460',
    '8801782542199','01782542199',
    '8801554702104','01554702104',
    '8801905186028','01905186028',
    '8801401866071','01401866071',
    '8801927787270','01927787270',
    '8801964123108','01964123108',
    '8801924967582','01924967582',
    '8801577556527','01577556527',
    '8801908568652','01908568652',
    '8801402940991','01402940991',
    '8801934147013','01934147013',
    # নতুন block
    '8801961931606','01961931606',
    # mygov OTP sender
    '8801552146224','01552146224',
    # Microsite OTP sender
    '8801847348716','01847348716',
    # Udvash coaching promo number
    '8809601775566','9601775566',
    # ── v12 এ চিহ্নিত bulk BD promo/spam senders ──────────────
    '8801575200164','01575200164',   # Phoenix Finance FDR promo
    '8801511313043','01511313043',   # সোফা/পানির ট্যাংক/পোকামাকড় সার্ভিস
    '8801771953770','01771953770',   # Zihad Security Services
    '8801518915313','01518915313',   # IDLC FDR promo
    '8801540651328','01540651328',   # Tutor/Tuition promo
    '8801847348715','01847348715',   # রাজউক plot বিক্রয়
    '8801897659713','01897659713',   # Nishat Linen sale
    '8801969908192','01969908192',   # B2 clothing promo
    '8801614298997','01614298997',   # Pista Choco Krunch food promo
    '8801938838845','01938838845',   # Passport/citizenship promo
    '8801819313319','01819313319',   # Electrical substation promo
    '8801969900913','01969900913',   # Engineering promo
    '8801974090800','01974090800',   # CVC Finance FDR promo
    '8801969901227','01969901227',   # B2 clothing promo variant
    '8801941642267','01941642267',   # Tour packages promo
    '8801607707665','01607707665',   # Tourist Visa promo
    '8801844160132','01844160132',   # Citizenship/immigration promo
    '8801738771299','01738771299',   # IDLC FDR promo (personal number)
    '8801715952838','01715952838',   # BiyeMedia matrimony
    '8801871012645','01871012645',   # BiyeMedia matrimony variant
    '8801979313319','01979313319',   # Electrical substation promo
    '8801969906275','01969906275',   # Political campaign SMS
    '8801552202766','01552202766',   # Political/religious appeal
    '8809617611614','9617611614',    # Real estate (Wellcare Jolshiri)
    '8809617611781','9617611781',    # FIFA/Visa promo
    '8809601017265','9601017265',    # Purbachal plot promo
    '8809648904099','9648904099',    # Restaurant promo (Teriyaki Boy)
    # ── sample SMS থেকে চিহ্নিত ────────────────────────────────
    '8801819399999','01819399999',   # StanChart sharetrip discount promo
    '8801613322341','01613322341',   # Property management brokerage
    '8809640910929','9640910929',    # Wellcare Greencity plot promo
    '8809601017006','9601017006',    # Purbachal Wellcare plot promo (variant)
    '8801540400015','01540400015',   # ৫তারকা হোটেল শেয়ার বিক্রয়
    '8801537275871','01537275871',   # Land/property listing (=Katha>Purbachal)
    '8801713002610','01713002610',   # UMCH property/land listing
    '8801847050242','01847050242',   # ব্যক্তিগত নম্বর — ব্লক করা হয়েছে
    '8801969910636','01969910636',   # Chaldal promo
}

# OTP sender names — always filter
OTP_SENDERS = {
    'gp otp','gpotp','boesl','zipper',
    'nid otp','mygov otp','mygov',
}

# Alpha/named senders যেগুলো সবসময় promo/spam — number নয়, নাম
NAMED_PROMO_SENDERS = {
    'diamndworld', 'diamond world', 'diamondworld',  # jewelry promo
    'navana cond', 'navana',                          # real estate
    'umch',                                           # property listing
    'country',                                        # cattle/livestock
    'prems clton', 'prems collection',                # fashion brand
    'artisan',                                        # real estate brand
    'mirror', 'mirrorfashion',                        # fashion brand
    'best deal', 'best_deal', 'dhamakadeal',          # telecom promo
    'aryan city', 'aryancity',                       # real estate
    'tuitionmate', 'tutorclub', 'tutorhunt',          # tutor platforms
    'manyavar bd', 'manyavar',                        # fashion brand
    'goongoon', 'jhotpotloan',                        # telecom promo
    'udvash',                                         # coaching promo (not OTP)
    'apanjewelrs', 'apanjwewlers', 'apan jewelrs', 'apan jewellers',  # jewelry promo
}

SVC_KW    = ['desco','dpdc','breb','titas','wasa','aarong','brac','daraz','chaldal',
             'pathao','steadfast','paperfly','redx','sundarban','ecourier','pandago',
             'foodpanda','air astra','biman','us-bangla','novoair','ivac','ivac_bd',
             'epassportbd','epassport','walton','flexiload']

def stype(sender, body=''):
    sl = str(sender).lower().strip()
    s  = str(sender).strip()
    # Promotional numbers — filter out
    if s in PROMO_NUMBERS: return "promo"
    # Bank short-codes — always bank
    if s in BANK_SHORTCODES: return "bank"
    # NexusPay/DBBL — "[ ! ]", "[!]", "[ !]", "! " ইত্যাদি সব variant
    if re.match(r'^\[?\s*!\s*\]?$', s): return "bank"
    # NexusPay newer sender name variants — NEXUSPAY, DutchBangla (no separator)
    if re.match(r'^nexuspay$|^nexus\s*pay$|^dutchbangla$', sl): return "bank"
    # Foreign bank senders (AlRajhi MY, Maybank, CIMB, etc.)
    if is_foreign_bank(sender, body): return "foreign_bank"
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
         'airtelbonus','airtelvoice','airteloffers','airtel best','airtel_hub',
         'airtel_loan','airtel rebate','at rebate','at_rebate',
         'combop','limitedtime','ajker_offer','mini_combo','mini_pack',
         'bishalof','bondho_','specialpack','darun_offer','beshi_min','goongoon',
         'secretoffer','tong offer','lucky_hour','best_deal','cashback','myfamily',
         # Airtel promo sender names
         'combopac','best_voice','best voice','iphone14','at_250min','at_200min',
         '1paisa call','1paisa','beshi meyad','beshi_meyad','newyroffer','newyr',
         'voice_60day','beshi_kotha','shera offer','75min_50tk','super_deal',
         'at_500min','at_50min','at_30day','30min_20tk','at_300min','at_100min',
         'airtel_loan','airtelbonus','airteldata','beshi min','beshimin',
         # Robi promo senders
         'myrobi','robi','robiminutes','robi deal','robi data','robi elite',
         'robioffer','robispecial','robi_offer','robi 30days','robi_mfs',
         'robi 70tk','super_offer','superof','extra offer','dhamakadeal',
         'jhotpotloan','jhotpot',   # Robi emergency loan — not real transaction
         'mini_bundle','robi_loan',
         # GP, BL, Teletalk promo
         'nagad_offer','apnar_offer','gp','gpstar','gp-info','gpalert',
         'grameenphone','banglalink','bl-','mybl','bl offer','bl commerce',
         'blsuperdeal','teletalk','govtinfo','govt info','govt','btrc','news_alert',
         'gpautoren','gp minute','gp data','gpminute','gp internet','gp loan',
         'gp info','flexiload']

# Airtel promo body keywords — sender নাম না পেলেও body দেখে filter
AIRTEL_PROMO_BODY_RE = re.compile(
    r'এয়ারটেলে\s+(?:আজ|বেশি|মাস|৩০দিন|\d+জিবি|\d+মিনিট|\d+দিন|সেরা|অফার|রিচার্জ)'
    r'|\*4\*\d+\*1#'        # Airtel USSD code pattern
    r'|\*২১২\*\d+#'         # Airtel Bengali USSD
    r'|airtel.*(?:\d+gb|\d+জিবি|\d+min|\d+মিনিট)',
    re.I)

SKIP_BODY_RE = re.compile(
    # Messaging apps
    r"whatsapp\.com|Let's chat on WhatsApp|Let's video call"
    r"|imo\.onelink|imo-তে|Let's have a call|meet\.google\.com"
    r"|Can't talk now|I'll be there soon|This invite expires"

    # মিসড কল / কল নোটিফিকেশন
    r"|কলটি গ্রহণ করতে পারবেন|কল মিস|এখন কলটি গ্রহণ"
    r"|missed call|miss call|কলটি মিস|মিস করেছেন"
    r"|মিসড\s*কল|missed\s*call\s*alert|মিসড\s*কল\s*এলার্ট"
    r"|টি মিসড কল আছে|টি\s*মিসড\s*কল|মিসড কল আছে"
    r"|থেকে আপনার \d+\s*টি|আপনার \d+\s*টি মিস"
    r"|দয়া করে কল করুন"

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

    # Eid / religious greeting — শুভেচ্ছা ও প্রমো মেসেজ বাদ
    r"|eid\s*mubarak|ঈদ\s*মুবারক|ঈদ\s*মোবারক"
    r"|ঈদুল\s*ফিতর|ঈদুল\s*আজহা|eid\s*ul\s*fitr|eid\s*ul\s*adha"
    r"|পবিত্র\s*ঈদ.*শুভেচ্ছা|ঈদের\s*শুভেচ্ছা|ঈদ\s*কালেকশন|ঈদ\s*অফার|ঈদ\s*সেল"
    r"|happy\s*eid|eid\s*greet|wishing.*eid|eid.*wishing|eid.*collection|eid.*offer"
    r"|শুভ\s*নববর্ষ|শুভ\s*নব\s*বর্ষ|happy\s*new\s*year|নববর্ষের\s*শুভেচ্ছা"
    r"|ফ্রি\s*সার্ভিস|ফ্রি\s*ডেলিভারি|free\s*service|free\s*delivery"
    r"|ঈদ\s*গিফট|ঈদ\s*স্পেশাল|eid\s*special|eid\s*gift"

    # Short URL spam — কিন্তু bank-এর নিজস্ব link বাদ নয়
    r"|ln\.run/|bit\.ly/(?!.*(?:parcel|track|order|nexuspay|nexus))"

    # সরকারি তথ্য / ৯৯৯ — বাদ
    r"|\b999\b|৯৯৯|জাতীয়\s*তথ্য\s*বাতায়ন"
    r"|government\s*(?:info|notice|alert)|সরকারি\s*বিজ্ঞপ্তি"
    r"|জাতীয়\s*হেল্পলাইন|national\s*helpline"

    # NGO / Insurance — বাদ
    r"|ngo\s*(?:সেবা|service)|এনজিও|microcredit|ক্ষুদ্রঋণ"
    r"|insurance\s*(?:premium|policy|renewal|due)|বীমা\s*(?:প্রিমিয়াম|পলিসি|নবায়ন)"
    r"|life\s*insurance|স্বাস্থ্য\s*বীমা|জীবন\s*বীমা|মেয়াদী\s*বীমা"
    r"|delta\s*life|metlife.*premium|sunlife|pragati\s*life|national\s*life.*insurance"

    # Home Tutor / Coaching centre — প্রমো বাদ
    r"|home\s*tutor(?:\s*available|\s*needed|\s*wanted|\s*provide)"
    r"|need\s*(?:home\s*)?tutors?|experienced.*(?:male|female).*tutor"
    r"|du/buet/medical.*tutor|tutor.*du/buet|coaching\s*(?:centre|center|class)"
    r"|কোচিং\s*(?:সেন্টার|ক্লাস|করুন)|হোম\s*টিউটর\s*(?:পাওয়া|দেওয়া|প্রয়োজন|চাই)"
    r"|free\s*(?:trial|class).*tutor|two\s*trial.*free\s*class"

    # ডাক্তার সেবা প্রমো / হেলথ প্যাকেজ
    r"|ডাক্তার\s*সেবা\s*মাত্র|সম্পূর্ণ\s*ডাক্তার\s*সেবা"
    r"|মাসের\s*সম্পূর্ণ.*সেবা.*টাকায়"
    r"|doctor\s*service.*tk|health\s*package.*taka"

    # মিসড কল এলার্ট সহ প্রমো (Mobil oil, কুপন, শর্ট লিঙ্ক)
    r"|মিসড\s*কল\s*এলার্ট.*(?:কিনলে|ছাড়|কুপন|tiger|t\.ly|nk/|link)"
    r"|missed\s*call\s*alert.*(?:discount|coupon|buy|link|oil)"

    # SMS subscription service promo (Life Style, RT, STS type)
    r"|start\s+(?:rt|sts|sms|st)\s+লিখে.*(?:\d{4,5}|নম্বরে|পাঠিয়ে)"
    r"|লিখে\s+sms\s+কর|লিখে\s+পাঠিয়ে\s*দিন\s*\d{4,5}"
    r"|চার্জ\s*প্রতিদিন\s*মাত্র|অটো\s*রিনিউ\s*প্রযোজ্য"
    r"|স্ট্যাটাস\s*পেতে\s*start|ক্যাপশন.*start.*sms"

    # Mobil / lubricant oil promo + buymobil domain
    r"|mobil\s+for\s+hybrid|mobil\s+(?:oil|engine|hybrid|lube)"
    r"|hybrid\s+engine.*পারফরম্যান্|হাইব্রিড\s+গাড়ির.*পারফরম্যান্"
    r"|buymobil\.mjlbl\.com|mjlbl\.com/campaign"

    # প্লট বুকিং / আবাসন প্রমো (specific)
    r"|প্রবাসী\s*পল্লীতে\s*প্লট|প্রবাসী\s*পল্লী.*বুকিং"
    r"|পূর্বাচল.*প্লট.*বুকিং|জাইকার\s*বিপরীতে"
    r"|প্লট\s*বুকিং\s*চলছে"

    # Tower/lottery scam + mobile বিজ্ঞাপন
    r"|tower\s+বানিয়ে\s+পেয়ে|tower.*স্মার্টফ|টাওয়ার.*স্মার্টফ"
    r"|কাউন্টডাউন.*tower|উৎসবের.*tower"

    # স্বাস্থ্য সেবা প্রমো + missed call combo
    r"|আপনার\s+স্বাস্থ্য.*আমাদের\s+দায়িত্ব"
    r"|স্বাস্থ্য.*দায়িত্ব.*যোগাযোগ"
    r"|এখন\s+আমার\s+সাথে\s+যোগাযোগ\s+করতে\s+পারবেন"

    # SMSC delivery receipt — NIP/MOne CDR format, body নেই
    r"|id:\d+\s+submit\s+date:\d+\s+done\s+date:\d+\s+stat:(?:DELIVRD|UNDELIV|EXPIRED|DELETED)"

    # ── BD promo number থেকে আসা spam body patterns (v12) ─────────────────
    r"|(?:fdr|fixed\s*deposit|fdm).*(?:rate|profit|return|interest)"
    r"|(?:attractive|competitive|best)\s*(?:profit|interest|rate).*(?:fdr|deposit|investment)"
    r"|security\s*guard.*(?:provide|service|rate|tk|taka|\d+)"
    r"|(?:cleaning|pest\s*control|rodent|termite|mosquito|cockroach).*(?:service|call|contact)"
    r"|সোফা.*ক্লিন|পোকামাকড়.*দমন|তেলাপোকা.*সার্ভিস|উইপোকা.*দমন"
    r"|(?:tour\s*package|air\s*ticket|hotel\s*booking).*(?:call|contact|\d{10,11})"
    r"|citizenship\s*program|visa.free.*travel|passport.*program"
    r"|(?:tutor|tuition|coaching).*(?:available|needed|call|contact|01\d{9})"
    r"|(?:electrical|substation|switchgear|ht\s*panel|lt\s*panel).*(?:supply|provide|contact)"
    r"|(?:plot|flat|land).*(?:rajuk|purbachal|bashundhara|uttara).*(?:sale|tk|লক্ষ)"
    r"|(?:winter|eid|season)\s*(?:collection|sale|offer|discount).*(?:call|visit|01\d{9,10})"
    r"|(?:linen|garments|clothing|fashion)\s*(?:sale|offer|collection|discount).*(?:%|tk|taka)"

    # System maintenance notice (ব্যাংক planned maintenance)
    r"|for\s+system\s+maintenance.*(?:will\s+be|temporarily)\s+unavailable"
    r"|(?:ibanking|mobile\s*app|internet\s*banking).*temporarily\s*unavailable"
    r"|planned\s+maintenance.*(?:services?|unavailable|suspended)"

    # Missed call alert service (auto-reply SMS)
    r"|মিসড\s*কল\s*এলার্ট.*কল\s*গ্রহণ\s*করতে\s*পারছিনা"
    r"|missed\s*call\s*alert.*unable\s*to\s*(?:receive|answer)"

    # Emergency SOS (Google/phone auto-SMS)
    r"|emergency\s*sos\s*i\s*called\s*emergency\s*services"
    r"|called\s+emergency\s+services\s+from\s+this\s+approximate\s+location"

    # Matrimony / marriage media
    r"|biyemedia|পাত্র\s*পাত্রীর\s*সন্ধান|পাত্র\s*পাত্রীর\s*খোঁজ|বিবাহ\s*মিডিয়া"

    # Immigration / visa service promo
    r"|(?:tourist|student|business)\s*visa\s*processing.*(?:canada|uk|usa|australia|schengen)"
    r"|citizenship\s*(?:in|program).*(?:new\s*zealand|turkey|portugal|malta|caribbean)"
    r"|immigration\s*lawyer.*dhaka|meet.*immigration.*(?:lawyer|consultant)"

    # Political / association election campaign — ভোট চাওয়ার আবেদন (SCBA, DBA, DTBA, ইত্যাদি)
    r"|নির্বাচিত\s*হলে.*(?:আমানত|জীবন|দেশ|সেবা)|ভোট\s*দিন|ভোট\s*চাই|ভোট\s*ও\s*দোয়া"
    r"|ভোটও\s*দোয়া|ভোট\s*প্রার্থী|ভোটপ্রার্থী|ভোট\s*প্রার্থনা|ভোট\s*কামনা"
    r"|ভোট\s*ও\s*সমর্থন|আপনার\s*ভোট|মূল্যবান\s*ভোট|ভোট\s*ও\s*আশীর্বাদ"
    r"|পদপ্রার্থী|পদ\s*প্রার্থী|প্রার্থী.*(?:নির্বাচন|ব্যালট)|ব্যালট.*প্রার্থী"
    r"|নির্বাচনে.*(?:সদস্য|সম্পাদক|সভাপতি|সহ-সভাপতি|ট্রেজারার|কোষাধ্যক্ষ).*পদে"
    r"|(?:support|doa|vote).*(?:election|নির্বাচন)|seeking\s*your\s*kind\s*support.*(?:me\s*for|election)"
    r"|cast\s*your\s*(?:valuable\s*)?vote|valuable\s*vote|kind\s*support.*vote"
    r"|(?:candidate|ballot)\s*(?:no\.?|#|number)?\s*\d*.*(?:election|নির্বাচন)"
    r"|election.*(?:candidate|ballot)|ballot\s*(?:no\.?|#)\s*\d+"
    r"|(?:secretary|president|treasurer|member|vice-president)\s*candidate"
    r"|seeking\s*your\s*(?:valuable\s*)?vote|humbly\s*(?:seek|request|solicit)"
    r"|solicit\s*(?:your|ur)\s*(?:blessing|vote|support)"
    r"|(?:scba|dba|dtba)\s*(?:election|নির্বাচন)|(?:election|নির্বাচন).*(?:scba|dba|dtba)"
    r"|voter\s*(?:no|number|id)[:\s]*\d+|আপনার\s*ভোটাধিকার"
    r"|elected\s*as\s*(?:a|an)?\s*(?:member|secretary|president)"
    r"|(?:thanks|gratitude).*(?:vote|support).*(?:election|elect)"
    r"|(?:highest|valuable)\s*vote(?:s)?\s*(?:of|for|in)|voted\s*for\s*me"
    r"|trust,?\s*blessings\s*&\s*electing|electing\s*me\s*as"
    r"|your\s*voter\s*no[:\.\s]*\d+|voter\s*no\.?\s*(?:is)?[:\s-]*\d+"
    r"|ভোটার\s*তালিকা|ভোটার\s*নম্বর|নাম\s*কর্তন\s*এড়াতে"
    r"|ভোটে\s*নির্বাচিত|নির্বাচিত\s*করায়\s*কৃতজ্ঞ|সর্বোচ্চ\s*ভোটে"
    r"|seek.*support.*(?:and|&)\s*vote|seeking.*support.*(?:and|&)\s*vote"
    r"|\(?dba\s*news\)?[:\s]|dba\s*news\b"
    r"|আঞ্জুমান|যাকাত\s*ও\s*সাদাকা|যাকাত.*সাদাকা|সাদাকা.*যাকাত"
    r"|wish(?:es)?\s*you\s*(?:and\s*(?:your\s*)?family)?(?:[\s,.-]+(?:a\s+)?(?:happy|prosperous|healthy|blessed|joyful)|\s*(?:good\s*luck|good\s*health|happiness|peace|prosperity))"
    r"|wish\s*you\s*and\s*(?:your\s*)?family|wishing\s*you\s*(?:and\s*(?:your\s*)?family)?"

    # AGM / shareholder notices
    r"|(?:\d+(?:st|nd|rd|th)\s+agm|annual\s+general\s+meeting).*(?:plc|ltd|limited|company)"
    r"|dear\s+valued\s+shareholder"

    # StanChart generic security alerts (non-transaction)
    r"|স্ট্যানচার্ট\s*ব্যাংক\s*কখনোই.*(?:কার্ড\s*নম্বর|pin|otp|cvv)"
    r"|stanChart.*(?:never\s*ask|won't\s*ask|do\s*not\s*share).*(?:pin|otp|password|cvv)"

    # Trustee / vehicle certificate renewal
    r"|trustee\s+board\s+certificate.*(?:expire|expiry|renew|pay)"
    r"|expiry\s+of\s+trustee.*metro.*(?:expire|renew|pay)"

    # ── sample SMS থেকে চিহ্নিত নতুন patterns ──────────────────────────────
    # StanChart card discount promo (Bengali) — sharetrip, ছাড়, ১৬%
    r"|স্ট্যানচার্ট.*(?:কার্ড|card).*(?:ছাড়|discount|cashback|offer)"
    r"|শেয়ারট্রিপ.*(?:ছাড়|discount|off)|sharetrip.*(?:ছাড়|discount)"
    # Short promo URL (t.ly, tinyurl) — ব্যাংক/MFS কখনো এগুলো ব্যবহার করে না
    r"|(?<!\w)t\.ly/\w+"
    # Property management/brokerage service
    r"|প্রপার্টি.*(?:ব্যবস্থাপনা|ক্রয়.*বিক্রয়|management).*\d{11}"
    # Wellcare/Jolshiri/Purbachal plot promo
    r"|ওয়েলকেয়ার.*(?:গ্রিনসিটি|greencity|প্লট|plot)"
    r"|wellcare.*(?:গ্রিনসিটি|greencity|plot|প্লট)"
    r"|জলসিড়ি.*(?:প্লট|plot|পূর্বাচল)"
    # Hotel share / timeshare sale
    r"|হোটেল\s*শেয়ার\s*বিক্রয়|hotel\s*share\s*(?:sale|sell)"
    # Property listing (=Flat =Land >Bashundhara/Purbachal/Gulshan format)
    r"|=\s*(?:land|flat|hospital|plot|katha|factory).*>\s*(?:bashundhara|purbachal|gulshan|moghbazar|mirpur|uttara)"
    # Maternity / delivery package promo from medical college
    r"|মেডিকেল\s*কলেজ.*(?:প্যাকেজ|package|delivery|প্রসূতি|পেইনলেস)"
    r"|(?:normal|c-section|পেইনলেস)\s*delivery.*package"
    r"|(?:maternity|প্রসূতি|প্রসব).*(?:package|প্যাকেজ).*(?:call|কল|01\d{9})"

    # ── Club / Society / Org mass SMS (এই CDR থেকে চিহ্নিত) ─────────────────
    # Golf tournament
    r"|golf\s*(?:tournament|cup|championship|house\s*corporate)"
    r"|kgcdhaka|armygolfclub|agcdhaka"
    # Club committee election / ballot
    r"|(?:ballot|ব্যালট)\s*(?:no|নং|number|#|নম্বর)[\s–\-]*\d+"
    r"|কার্যনির্বাহী\s*কমিটি.*(?:নির্বাচন|ভোট|সমর্থন)"
    r"|(?:executive\s*committee|ec\s*election).*(?:vote|support|candidate)"
    # AGM / shareholder (any variant)
    r"|annual\s*general\s*meeting"
    # Janaza / condolence notice (social club mass SMS)
    r"|নামাজে\s*জানাজা|জানাজার\s*নামাজ"
    # Iftar party invite (mass SMS)
    r"|ইফতার\s*(?:পার্টি|মাহফিল|অনুষ্ঠান).*আয়োজন"
    # Doctor promo (NDC / Pain clinic)
    r"|(?:consultant|কনসালটেন্ট).*(?:pain|পেইন).*(?:ndc|hospital|হাসপাতাল|মঙ্গলবার|বুধবার)"
    # Jewelry / diamond sale
    r"|(?:diamond|ডায়মন্ড).*(?:\d+%|ছাড়|sale|off|মেলা)"
    # House of Bengal / jewelry shop promo
    r"|(?:হাউজ\s*অব\s*বেঙ্গল|house\s*of\s*bengal)"
    r"|(?:সোনা|হীরা).*(?:গহনা|সম্ভার).*(?:বসুন্ধরা|gulshan|দোকান|shop)"
    # Card statement notification (not transaction)
    r"|(?:statement|স্টেটমেন্ট).*(?:credit\s*card|closing\s*balance).*(?:sent|পাঠানো)"
    # Tax deduction certificate email
    r"|income\s*tax.*(?:deduction|certificate).*(?:sent|email|registered)"
    # Fashion brand with social media link
    r"|(?:fashion|brand|collection|পোশাক\s*ব্র্যান্ড).*(?:fb:|facebook\.com|instagram|rb\.gy)"
    # Political/party mass SMS
    r"|(?:খালেদা\s*জিয়া|বিএনপি|আওয়ামী).*(?:গণতন্ত্র|মুক্তি|আন্দোলন)"

    # ── এই CDR থেকে চিহ্নিত নতুন patterns ──────────────────────────────────
    # মাসিক কিস্তিতে প্লট/ফ্ল্যাট promo
    r"|মাসিক\s+কিস্তিতে\s+(?:প্লট|ফ্ল্যাট|plot|flat)"
    r"|[\d,]+\s*টাকা\s+মাসিক\s+কিস্তি"
    r"|কিস্তি\s+দিলেই\s+বুকিং|ডাউনপেমেন্টের\s+নেই\s+ভয়"
    # Tutor promo — Bengali & English variants
    r"|(?:টিউটর|tutor)\s+পেতে\s+(?:কল|call)\s+করুন"
    r"|(?:হোম|home)\s+টিউটর.*(?:কলঃ|কল:|01\d{9})"
    r"|(?:home|হোম)\s+tutor\s+পেতে\s+(?:call|কল)"
    r"|do\s+you\s+need\s+(?:home|hom)\s+tutor"
    r"|ফ্রি\s+ক্লাস.*(?:টিউটর|tutor)"
    r"|অনলাইন.*(?:যেকোনো\s+বিষয়|বিষয়ের).*(?:অভিজ্ঞ|শিক্ষক).*(?:কল|01\d{9})"
    r"|tutorclubsite|tuitionmate|tutorhunt"
    # Matrimony promo
    r"|পাত্র.*পাত্রী.*(?:খুজুন|খুঁজুন|সন্ধান)"
    r"|পাত্র\s+পাত্রি\s+(?:খুজুন|খুঁজুন)"
    r"|(?:বিশ্বস্ত|প্রাইভেসি\s+বজায়).*পাত্র.*(?:যোগাযোগ|contact)"
    # Gems govt portal OTP (government employee system)
    r"|gems\s+password\s+reset\s+otp|gems\.gov\.bd"
    # FDR/investment personal solicitation
    r"|any\s+new\s+personal\s+investment|invest.*now.*fixed\s+deposit"
    r"|dear\s+sir.*opportunity\s+to\s+invest.*fixed\s+dep"
    # Fogger/lawnmower/equipment promo
    r"|fogger\s+machine|lawnmower|carwasher.*(?:chainsaw|chemical)"
    r"|chainsaw.*(?:chemical|contact|01\d{9})"
    # MRF steel / building systems solicitation
    r"|mrf\s+steel|building\s+systems.*(?:rcc|steel).*(?:develop|construct)"
    # DTBA/club death notice (mass SMS)
    r"|dtba\s+(?:news|member).*(?:passed\s+away|ইন্তেকাল)"
    # cutt.ly / short promo links (personal BD number থেকে channel/content promo)
    r"|(?<!\w)cutt\.ly/\w+|rb\.gy/\w+"
    # টিউটর খুঁজছেন (present tense — different from পেতে কল করুন)
    r"|(?:টিউটর|tutor).*(?:খুঁজছেন|খুজছেন|looking\s+for\s+tutor)"
    r"|(?:সন্তানের\s+জন্য\s+টিউটর|আপনার\s+সন্তানের.*টিউটর)"
    # নিবেদিতপ্রাণ চ্যানেল (dedicated YouTube/social channel promo)
    r"|নিবেদিতপ্রাণ.*চ্যানেল|dedicated.*channel.*(?:cutt\.ly|bit\.ly|subscribe)"
    # Club/org event invitation (mass SMS repeated from personal BD number)
    r"|সম্মানিত\s+সদস্য.*(?:উপলক্ষে|অনুষ্ঠান|আমন্ত্রিত|কুপন|অংশগ্রহণ)"
    r"|(?:গরু|ইলিশ|মিনিকেট|চিনি)\s*\d+\*"
    # ARTISAN / real estate brand promo
    r"|\[ARTISAN\]|artisan.*(?:flat|ফ্ল্যাট).*(?:%|ছাড়|discount)"
    # FDR % offer from personal number
    r"|(?:fdr|fixed\s*deposit).*\d+[\.,]\d+%.*(?:double|triple|years|বছর)"
    # Industrial equipment supply from personal number
    r"|(?:generator|ups|sub.station|solar\s*power).*\b01\d{9}\b"
    r"|we\s+provide.*(?:generator|ups|solar|battery).*\d{11}"
    # AC/Fridge/appliance maintenance service promo
    r"|(?:air.condition|fridge|washing\s*machine).*(?:maintenance|repair|servic).*(?:contact|call|01\d{9})"
    # Interior design promo
    r"|(?:interior\s*design|ইন্টেরিয়র\s*ডিজাইন).*(?:call|whatsapp|01\d{9,10})"
    # Apartment/flat promo from personal number
    r"|(?:bedroom|bhk|apartment|flat).*(?:ready|ongoing|available).*(?:visit|contact|call|01\d{9})"
    # Flat with area (sft/sqft) = real estate promo
    r"|(?:flat|ফ্ল্যাট)\s+\d+\s*(?:sft|sqft|sq\.?\s*ft)"
    # Property listing =Flat =Land format (personal number)
    r"|=(?:flat|house|land|katha|sft)"
    # RJSC share transfer appointment = legal/business solicitation
    r"|(?:rjsc|share\s*transfer).*(?:appointment|re.signature|premises)"
    # Insurance / financial institution solicitation (DMD/CFO etc.)
    r"|(?:dmd|cfo|ceo)\s*(?:&|and)?\s*(?:cfo|ceo|md)?.*(?:insurance|finance).*(?:appointment|meet)"
    # ATM/app service maintenance from personal number
    r"|এটিএম.*পরিষেবা\s*বন্ধ|atm.*service.*(?:unavailab|closed|bnd).*(?:16269|\d{5})"
    # USSD promo (*XXXXX# format)
    r"|(?:ফ্রি|free).*(?:সার্ভিস|service).*\*\d{4,7}#"
    r"|\*\d{4,7}#.*(?:ডায়াল|dial|চালু)"
    # Cattle/livestock Eid promo
    r"|(?:bull|গরু|cattle).*(?:grass.fed|natural|eid).*(?:delivery|01\d{9})"
    # Emergency SOS location notification (auto-generated phone message)
    r"|emergency\s*sos.*(?:approximate\s*location|location\s*has\s*changed)"
    r"|you\s*are\s*receiving\s*this.*emergency\s*contact"
    # Brand new year / Eid greeting from shop (brand name + শুভেচ্ছা)
    r"|(?:কালেকশন|collection|fashion|boutique).*(?:এর\s*পক্ষ|পক্ষ\s*থেকে).*(?:শুভেচ্ছা|wishes)"
    # Ramadan interior / home decor seasonal promo
    r"|(?:ramadan|eid).*(?:special|package|offer).*(?:interior|design|home\s*decor)"
    # Jewelry (ইস্কাটনে ফ্ল্যাট ও land share)
    r"|ইস্কাটনে.*(?:ফ্ল্যাট|land|flat).*(?:চলছে|sale|01\d{9})"

    # ── নতুন (এই CDR থেকে) ──────────────────────────────────────────────────
    # নেসকো / BREB বিদ্যুৎ বন্ধ নোটিশ
    r"|নেসকো|breb.*(?:বিদ্যুৎ|electric)|বিদ্যুৎ.*(?:বন্ধ|সরবরাহ\s*বন্ধ)"
    r"|(?:হরিশপুর|গ্রিড|উপকেন্দ্র).*(?:রক্ষণাবেক্ষণ|বিদ্যুৎ|বন্ধ)"
    r"|দপ্তরের\s*আওতাধীন\s*সকল\s*এলাকায়\s*বিদ্যুৎ"

    # Focus promo (restaurant / UAC / Farmgate)
    r"|focus\s*(?:farmgate|uac|restaurant|01\d{9})"
    r"|ll\.\s*focus\s*farmgate"

    # সচিবালয় প্রবেশ OTP
    r"|(?:বাংলাদেশ\s*সচিবালয়).*(?:ওটিপি|otp|প্রবেশ)"
    r"|সচিবালয়.*প্রবেশের\s*জন্য.*(?:ওটিপি|otp|\d{4,6})"

    # Pure greeting only — body শুধু greeting (আর কিছু নেই)
    # আসসালামু আলাইকুম স্যার (only) — no substantive content after
    # r"|^[^।!]+আস+সালাম\w*\s+\w{0,8}[।\s]*$"  ← too risky, use is_greeting_only() below

    # Happy New Year from brands/orgs (আগে থেকেই আংশিক আছে, শক্তিশালী করছি)
    r"|happy\s+new\s+year\s+202[0-9]|happy\s+new\s+year.*(?:family|contact|call|visit)"
    r"|শুভ\s+নববর্ষ\s+202[0-9]|নববর্ষের\s+(?:শুভেচ্ছা|আন্তরিক|উষ্ণ)"

    # Fit Elegance টেইলরিং / Fit Basics ফ্যাব্রিক্স promo
    r"|fit\s*elegance|ফিট\s*এলিগ্যান্স|fit\s*basics|ফিট\s*বেসিক্স"
    # bKash ক্যাশব্যাক / অ্যাপ হালনাগাদ promo ("৳৫০ ক্যাশব্যাক পেতে এখনই বিকাশ অ্যাপে")
    r"|ক্যাশব্যাক\s*পেতে|বিকাশ\s*অ্যাপে\s*তথ্য\s*হালনাগাদ|তথ্য\s*হালনাগাদ\s*কর"
    # Grievance submission notification (tracking number) — সিস্টেম নোটিফিকেশন, ইন্টেলিজেন্স ভ্যালু নেই
    r"|a\s*new\s*grievance\s*is\s*submitted|grievance\s*is\s*submitted\s*with\s*tracking"
    r"|অভিযোগ\s*দাখিল.*ট্র্যাকিং|নতুন\s*অভিযোগ\s*দাখিল"
    ,
    re.I | re.S)

# সব regex pattern-এ বাংলা string NFC normalize করি
# কারণ: xlsx থেকে আসা বাংলা text-এ "য়" U+09DF (single char)
# কিন্তু source code-এ "য়" = U+09AF + U+09BC (two chars)
# regex engine এ দুটো আলাদা — NFC normalize করলে দুটোই U+09DF হয়
import unicodedata as _ud_init
def _nfc_re(pattern_str, flags=0):
    return re.compile(_ud_init.normalize('NFC', pattern_str), flags)

SKIP_BODY_RE    = _nfc_re(SKIP_BODY_RE.pattern, SKIP_BODY_RE.flags)

MISSED_RE = _nfc_re(
    r'মিস করেছেন|কল মিস|missed call|miss call|কলটি মিস|কলটি গ্রহণ করতে পারবেন'
    r'|মিসড\s*কল|missed\s*call\s*alert|মিসড\s*কল\s*এলার্ট'
    r'|টি মিসড কল আছে|টি\s*মিসড\s*কল|মিসড কল আছে'
    r'|দয়া করে কল করুন',
    re.I)

# OTP / verification — সব filter হবে
OTP_BODY_RE = re.compile(
    r'(?:your\s+)?otp\s+(?:is|for|to\s+proceed)|one.time.(?:password|key|code|pin)[\s:]'
    r'|verification\s+code\s+(?:is|for)'
    r'|do\s+not\s+share\s+(?:this|your|otp)'
    r'|(?:your\s+)?code\s+is\s*:\s*\d'
    r'|is\s+your\s+otp\b'
    r'|as\s+(?:one.time|otp)[\s:]'
    r'|use\s+the\s+(?:number|code|pin|key)\s*:\s*\d'
    r'|authentication\s+code\s+(?:is|for)'
    r'|login\s+otp\b'
    r'|ibas.*use.*(?:number|code|key)\s*:\s*\d'
    r'|please\s+use\s+the\s+(?:number|code|key)\s*:\s*\d'
    r'|সক্রিয়করণ\s*কোড[\s:]+\d'
    r'|ticket\s+purchase.*is\s*:\s*\d'
    r'|use\s+otp\s*:\s*\d'                  # BRAC: "Use OTP: 846941 to proceed"
    r'|otp\s*:\s*\d{4,8}\s+to\s+proceed'   # generic OTP: XXXXXX to proceed
    # NPSB/RTGS/interbank — শুধু তখনই OTP যদি কাছাকাছি ৪-৮ digit standalone code থাকে
    # (TrxID/Ref alphanumeric তাই \b\d{4,8}\b দিয়ে আলাদা করা যায়)
    r'|\b\d{4,8}\b[^\n]{0,60}proceed\s+with\s+(?:npsb|rtgs|within\s+bank|other\s+bank)'
    r'|proceed\s+with\s+(?:npsb|rtgs|within\s+bank|other\s+bank)[^\n]{0,60}\b\d{4,8}\b',
    re.I)

# ── Transaction guard — আসল লেনদেন কখনো ফিল্টার হবে না ──────────────────────
# অপারেটরের নাম (Grameenphone/Banglalink) merchant payment-এ থাকলেও যেন
# প্রোমো ভেবে বাদ না পড়ে — confirmed money-movement মেসেজ সবসময় রাখি।
TXN_GUARD_RE = re.compile(
    r'(?:payment\s+of\s+tk\s*[\d,]+.*?(?:successful|complete)'
    r'|received\s+tk\s*[\d,]+\s+from'
    r'|cash\s*(?:in|out)\s+tk\s*[\d,]+'
    r'|cash\s*in\s+received.*?amount:\s*tk\s*[\d,]+'
    r'|money\s+received.*?amount:\s*tk\s*[\d,]+'
    r'|payment\s+to\s+.{1,60}?\s+is\s+successful.*?amount:\s*tk\s*[\d,]+'
    r'|(?:credited|debited)\s+(?:with|from|for)\s+bdt\s*[\d,]+'
    r'|transaction\s+of\s+bdt\s*[\d,]+\s+(?:credited|debited)'
    r'|loan\s+(?:repayment|disbursement).*?tk\s*[\d,]+'
    r'|received\s+(?:digital\s+)?loan)',
    re.I | re.S)

# ── কোচিং / ভর্তি বিজ্ঞাপন ──
EDU_BRAND_RE = re.compile(
    r'উদ্ভাস|udvash|উন্মেষ|unmesh|রেটিনা|\bretina\b|সাইফুর[\'’]?স|saifurs|অগ্রণী\s*কোচিং|মেডিকো\s*কোচিং|'
    r'10\s*minute\s*school|টেন\s*মিনিট\s*স্কুল|shikho|শিখো|প্যারাগন\s*কোচিং|ইউসিসি\b|\bUCC\b\s*কোচিং|'
    r'ফোকাস\s*(?:কোচিং|এডমিশন)|উত্তরণ\s*কোচিং|e-?sho\s*shikhi', re.I)
EDU_PROMO_RE = re.compile(
    r'(?:ভর্তি|এডমিশন|admission|মডেল\s*টেস্ট|model\s*test|ব্যাচ|batch|কোর্স|course|রিভিশন|ওরিয়েন্টেশন|কোচিং)'
    r'.{0,80}?(?:ছাড়|শেষ\s*দিন|চলছে|শুরু|discount|offer|সীমিত\s*সময়|[\d০-৯,]+\s*/-)'
    r'|(?:ছাড়|discount)[^।]{0,40}?ভর্তি', re.I | re.S)
TUTOR_AD_RE = re.compile(
    r'(?:যোগ্য|অভিজ্ঞ|দক্ষ|ভেরিফাইড|মেধাবী|verified|experienced)[^।\n]{0,40}?টিউটর[^।\n]{0,30}?(?:পেতে|খুঁজে|কল\s*করুন|যোগাযোগ)'
    r'|টিউটর\s*(?:পেতে|খুঁজে\s*পেতে)[^।\n]{0,20}?(?:কল|যোগাযোগ)|home\s*tutor[^.\n]{0,30}?(?:call|contact)', re.I)
# ── রিয়েল এস্টেট বিজ্ঞাপন ──
REALESTATE_BRAND_RE = re.compile(
    r'শেলটেক|sheltech|\bbti\b|building\s+technology|rangs\s+properties|navana\s+real|concord\s+real|'
    r'assure\s+group|anwar\s+landmark|\bjcx\b|shanta\s+holdings|principal\s+group|edison\s+real|'
    r'amin\s+mohammad|asset\s+developments|bashundhara\s+(?:plot|land)|পূর্বাচল\s*আমেরিকান\s*সিটি', re.I)
REALESTATE_AD_RE = re.compile(
    r'(?:ফ্ল্যাট|flat|apartment|অ্যাপার্টমেন্ট|প্লট|plot|কমার্শিয়াল\s*স্পেস|commercial\s*space|office\s*space|দোকান|প্রপার্টি|property)'
    r'.{0,120}?(?:বুকিং\s*(?:চলছে|করুন)|বিক্রয়\s*চলছে|for\s*sale|কিস্তি\s*সুবিধা|luxury|south\s*facing|\bSFT\b|'
    r'বিনিয়োগ\s*করুন|নিশ্চিত\s*রিটার্ন|হ্যান্ডওভার|handover)', re.I | re.S)
# ── কুরিয়ার: ঠিকানা/লিংক ছাড়া রুটিন নোটিফিকেশন বাদ ──
COURIER_ANY_RE = re.compile(r'pathao|redx|steadfast|sundarban|ecourier|paperfly|sa\s*paribahan|delivery\s*tiger|carrybee|'
                        r'janani\s*express|korotoa|সুন্দরবন\s*কুরিয়ার|কুরিয়ার|courier', re.I)
COURIER_MSG_RE = re.compile(r'পার্সেল|parcel|\bCOD\b|ডেলিভারি|delivery|পণ্য\s*সংগ্রহ|consignment|shipment', re.I)
COURIER_KEEP_RE = re.compile(r'https?://|www\.|\b[\w\-]+\.(?:app|com|net|org|bd|ly|me)/\S+|ready\s+at|ঠিকানা|address|'
                             r'office|অফিস|শাখা|branch|point|পয়েন্ট', re.I)

def _is_personal_sender(snd):
    d = re.sub(r'\D', '', str(snd))
    return bool(re.fullmatch(r'(?:880|0)?1[3-9]\d{8}', d))

# প্রিপেইড বিদ্যুৎ মিটার টোকেন মেসেজ (মূল মেসেজ + আলাদা হয়ে যাওয়া খণ্ডাংশ)
METER_TOKEN_RE = re.compile(
    r'prepaid\s+token\s+\d{4}[\s\-]?\d{4}'                          # "DPDC Prepaid Token 3413-0050-..."
    r'|\btoken\s*(?:no\.?|number)?\s*[:#]?\s*\d{4}-\d{4}-\d{4}-\d{4}'  # "Token: 1234-5678-9012-3456"
    r'|energy\s+cost\s*:.*?demand\s+charge\s*:.*?meter\s+rent'          # টোকেনের খণ্ডাংশ
    r'|মিটার\s*টোকেন|প্রিপেইড\s*টোকেন',
    re.I | re.S)

# ব্যাংক লোন/প্রোডাক্ট বিজ্ঞাপন (সেলস পিচ, আসল লেনদেন নয়) — বাদ
BANK_AD_RE = re.compile(
    r'(?:personal|car|home|auto)\s+loan[!\s].*?(?:up\s+to|lac|lacs|lakh|crore|emi|approval|contact)'
    r'|up\s+to\s+bdt\s*[\d,]+\s*(?:lac|lacs|lakh|crore)'
    r'|(?:easy|flexible)\s+emi.*?(?:approval|contact|offer)'
    r'|🔹\s*(?:car|home|personal)\s+loan'
    r'|loan\s+offer[!\s].*?(?:lac|emi|approval)'
    r'|(?:fdr|dps)\b.*?(?:easy|attractive|processing)'
    r'|if\s+you\s+want\s+to\s+take\s+a\s+(?:personal|home|car)\s+loan'
    r'|(?:personal|home|car)\s+loan.*?(?:takeover|for\s+details,?\s+contact)',
    re.I | re.S)

def is_filtered(sender, body):
    s  = str(sender).strip()
    sl = s.lower()
    st = stype(s, body)
    # আসল লেনদেন হলে কখনো ফিল্টার করো না (OTP ছাড়া — OTP আলাদা হ্যান্ডল হয়)
    if TXN_GUARD_RE.search(body) and not OTP_BODY_RE.search(body):
        return False
    # ব্যাংক লোন/প্রোডাক্ট বিজ্ঞাপন — ইন্টেলিজেন্স ভ্যালু নেই, বাদ
    if BANK_AD_RE.search(body):
        return True
    # প্রিপেইড মিটার টোকেন (DPDC/DESCO/NESCO ইত্যাদি) — টোকেন নম্বর ও বিলের খুঁটিনাটি, বাদ।
    # (bKash/ব্যাংকের "Bill successfully paid" আসল টাকার লেনদেন — সেটা উপরের guard-এ রক্ষা পায়;
    #  "Tax-Token" গাড়ির ট্যাক্স টোকেন — মিটার টোকেন নয়, তাই রাখা হয়)
    if METER_TOKEN_RE.search(body) or sl in ("meter token", "metertoken"):
        return True
    # বাংলা নামের অদৃশ্য অক্ষর (ZWNJ/ZWJ) সরিয়ে মেলাই — যেমন "শেল্‌টেক্‌"
    _bn = re.sub(r'[\u200c\u200d\u09cd]', '', body)
    _sn = re.sub(r'[\u200c\u200d\u09cd]', '', s)
    # OTP (অতিরিক্ত ফরম্যাট)
    if OTP_EXTRA_RE.search(body): return True
    # কোচিং/ভর্তি বিজ্ঞাপন: ব্র্যান্ড নাম থাকলে সবসময়; সাধারণ বিজ্ঞাপন-ভাষা হলে শুধু বাল্ক/ব্র্যান্ড sender থেকে
    if any(EDU_BRAND_RE.search(x) for x in (body, _bn, s, _sn)): return True
    if EDU_PROMO_RE.search(body) and not _is_personal_sender(s): return True
    if TUTOR_AD_RE.search(body): return True
    # রিয়েল এস্টেট বিজ্ঞাপন: ডেভেলপারের নাম, বা ফ্ল্যাট/প্লট + বিজ্ঞাপনের ভাষা
    if any(REALESTATE_BRAND_RE.search(x) for x in (body, _bn, s, _sn)):
        return True
    if REALESTATE_AD_RE.search(body): return True
    # কুরিয়ার: লিংক বা ঠিকানা/পিকআপ-স্থান না থাকলে বাদ (প্রোফাইল-সূত্রে আগেই সংগ্রহ হয়)
    if (COURIER_ANY_RE.search(body) or COURIER_ANY_RE.search(s)) and COURIER_MSG_RE.search(body) \
            and not COURIER_KEEP_RE.search(body):
        return True
    # OTP sender names — always filter
    if sl in OTP_SENDERS: return True
    # Named promo/brand senders — always filter
    if sl in NAMED_PROMO_SENDERS: return True
    # Promotional numbers — always filter
    if st == "promo": return True
    # OTP/key মেসেজ — সব sender-এর জন্য বাদ
    if OTP_BODY_RE.search(body): return True
    # Courier pickup/OTP/cancelled — বাদ
    if COURIER_FILTER_RE.search(body): return True
    # Real estate promo — বাদ
    if REAL_ESTATE_FILTER_RE.search(body): return True
    # Body keywords — সব sender-এর জন্য check করো
    if SKIP_BODY_RE.search(body): return True
    if MISSED_RE.search(body): return True
    # Pure greeting only — no substantive content beyond salutation
    # "আসসালামু আলাইকুম স্যার" alone / "ভাই আসসালামু আলাইকুম।" alone → বাদ
    _b = body.strip().rstrip('।.!?-')
    if re.match(
        r'^(?:ভাই\s+|ছোট\s+ভাই\s+|বড়\s+ভাই\s+)?'
        r'আস+সালাম(?:ু\s*[আ]লাইকুম|ুয়ালাইকুম|\s*ওয়ালাইকুম)'
        r'(?:\s+ওয়া\s+রহমতুল্লাহ)?'
        r'(?:\s+(?:স্যার|ভাই|ছোট\s*ভাই|বড়\s*ভাই|ওয়ালাইকুম))?$',
        _b, re.I
    ): return True
    # Happy New Year from org/hospital with contact info or family wishes
    if re.search(r'happy\s+(?:new\s+year|birthday).*(?:family|contact|call|visit|ph:|01\d{9})', body, re.I): return True
    # Short-code filter — bank shortcodes ছাড়া
    if SC_RE.match(s) and s not in BANK_SHORTCODES: return True
    # Operator sender keywords
    sl = s.lower()
    if any(f in sl for f in OP_KW): return True
    # Airtel promo body — যেকোনো sender থেকে হলেও বাদ
    if AIRTEL_PROMO_BODY_RE.search(body): return True
    return False

# ── Multi-part merge ────────────────────────────────────────
def to_dt(d):
    if isinstance(d, datetime): return d
    s = str(d).strip()
    # Format: 20260210121148 (YYYYMMDDHHmmss) — GP operator format
    if re.match(r'^\d{14}$', s):
        try: return datetime.strptime(s, '%Y%m%d%H%M%S')
        except ValueError: pass
    # Format: 20260210 (YYYYMMDD only)
    if re.match(r'^\d{8}$', s):
        try: return datetime.strptime(s, '%Y%m%d')
        except ValueError: pass
    # Standard formats
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%dT%H:%M:%S','%d/%m/%Y %H:%M'):
        try: return datetime.strptime(s[:19], fmt)
        except ValueError: pass
    return datetime.min

def merge(sms_list):
    """ভাঙা (concatenated) SMS জোড়া লাগায়।
    - Personal BD নম্বর: ৪৫ সেকেন্ডের মধ্যে একই sender → merge
      (শর্ত: current body ≤ 153 char — single SMS part limit; বড় body = complete SMS)
    - MFS/Bank/Service: ৩০ সেকেন্ডের মধ্যে একই sender → merge
      (শর্ত: current body-তে balance/TrxID নেই — থাকলে complete SMS, আলাদা রাখো)
    """
    _PERSONAL_MERGE_WINDOW = 45    # seconds
    _PERSONAL_BODY_MAX     = 153   # char — এর বেশি = complete personal SMS

    # Bank/MFS "complete" SMS detector:
    # balance / TrxID / helpline থাকলে → নিজেই complete, পরেরটার সাথে merge করব না
    # এটাই EBL batch transaction সমস্যার মূল fix
    BANK_COMPLETE_RE = re.compile(
        r'[Bb]al(?:ance)?\s*(?:is\s*)?(?:BDT|[Tt][Kk])?\s*[\d,]+'  # Balance is BDT / Bal Tk
        r'|Avail(?:able)?\s+Bal'           # Avail Bal
        r'|C/B\s+Tk'                       # DBBL C/B pattern
        r'|Helpline\s+\d{4,5}'            # EBL Helpline 16230
        r'|TrxID\s+\w{6,}'                # bKash TrxID
        r'|Txn\s*ID\s*[:\s]\w{6,}'
        r'|নতুন\s+ব্যালেন্স'               # Bangla balance
        r'|বর্তমান\s+ব্যালেন্স',
        re.I
    )

    items = sorted(sms_list, key=lambda x: to_dt(x["date"]))
    out, i = [], 0
    while i < len(items):
        cur  = dict(items[i])
        st   = stype(cur["sender"], cur.get("body",""))
        dt0  = to_dt(cur["date"])
        sl0  = cur["sender"].strip().lower()
        body = cur["body"]

        if st == "personal":
            # Personal: ৪৫ সেকেন্ড window এবং body ছোট হলেই merge
            j = i + 1
            while j < len(items):
                nx = items[j]
                time_diff = abs((to_dt(nx["date"]) - dt0).total_seconds())
                if (stype(nx["sender"]) == "personal"
                        and nx["sender"].strip().lower() == sl0
                        and time_diff <= _PERSONAL_MERGE_WINDOW
                        and len(body) <= _PERSONAL_BODY_MAX):
                    body += " " + nx["body"]
                    j += 1
                else:
                    break
        else:
            # MFS/Bank/Service/Other: ৩০ সেকেন্ড window
            j = i + 1
            while j < len(items):
                nx = items[j]
                time_diff = abs((to_dt(nx["date"]) - dt0).total_seconds())
                if not (nx["sender"].strip().lower() == sl0 and time_diff <= 30):
                    break

                cur_body = body  # accumulated so far
                nx_body  = nx["body"]

                # ★ FIX 1: OTP body → কখনো merge করা যাবে না
                # প্রতিটা OTP আলাদা — একই sender থেকে একই সময়ে ভিন্ন OTP আসতে পারে
                cur_is_otp = bool(OTP_BODY_RE.search(cur_body))
                nx_is_otp  = bool(OTP_BODY_RE.search(nx_body))
                if cur_is_otp or nx_is_otp:
                    break

                # ★ FIX 2: current body complete হলে পরেরটা আলাদা SMS
                if BANK_COMPLETE_RE.search(cur_body):
                    break

                # ★ FIX 3: SCHOLASTICA/service SMS — একটা complete part-এর শেষ চিহ্ন
                # যদি cur_body শেষ হয় ।/!/? দিয়ে → complete sentence → merge না
                # কিন্তু multi-part SMS-এ প্রথম part মাঝখানে কেটে যায়
                # সঠিক heuristic: যদি cur_body ≥ 100 char AND শেষ char পরিচিত ending দিয়ে শেষ
                # তাহলে complete SMS → merge করো না
                if (len(cur_body.strip()) >= 100
                        and cur_body.strip()[-1] in '.!?।'
                        and not cur_body.strip().endswith('...')):
                    break

                # ★ FIX 4: next body নতুন SMS শুরু হওয়ার স্পষ্ট চিহ্ন দিলে merge করো না
                # "প্রিয়/সম্মানিত/Dear/আমাদের" দিয়ে শুরু = নতুন independent SMS
                if re.match(
                    r'^(?:প্রিয়|সম্মানিত|আপনার|আমাদের|বিজ্ঞপ্তি|'
                    r'dear\s|urgent|notice|important\s|greetings)',
                    nx_body.strip(), re.I
                ):
                    break

                body += " " + nx_body
                j += 1

        cur["body"]  = body.strip()
        cur["parts"] = j - i
        out.append(cur)
        i = j
    return out

# ── লেনদেন পরিমাণ ──────────────────────────────────────────
TXN_P = [
    # Cash Out ভ্যারিয়েন্ট — "cashed out", হাইফেন "Cash-Out", বাংলা "ক্যাশ আউট"
    r'cashed\s+out\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Cc]ash[\s\-]*[Oo]ut\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'ক্যাশ\s*আউট\s+[Tt][Kk]\.?\s*([\d,\u09e6-\u09ef]+\.?\d*)',
    r'ক্যাশ\s*আউট\s+([\d,\u09e6-\u09ef]+\.?\d*)\s*টাকা',
    r'received\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+from',
    r'received\s+deposit.*?of\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'received\s+a\s+deposit\s+of\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',  # bKash from bank
    r'received\s+remittance.*?[Tt]otal[:\s]+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'remittance.*?[Tt]otal[:\s]+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'sent\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+to',
    r'[Cc]ash\s*[Ii]n\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Cc]ash\s*[Oo]ut\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Ss]end\s+[Mm]oney\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',   # bKash Send Money
    r'[Mm]oney\s+[Rr]eceived\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Pp]aid\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Aa]mount[:\s]+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Dd]eposit\s+[Aa]mount[:\s]+([\d,]+\.?\d*)',
    r'transfer(?:red)?\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)',
    r'[Pp]ayment\s+(?:of\s+)?[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+to',
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
    # DBBL / NexusPay pattern
    r'debited\s+(?:\(.*?\)|[A-Z\s]+)\s+by\s+Tk-?([\d,]+\.?\d*)',
    r'credited\s*(?:\(.*?\)|[A-Z\s]+)\s+by\s+Tk([\d,]+\.?\d*)',
    # Sonali Bank
    r'for\s+BDT\s+([\d,]+\.?\d*)',
    # IBBL: "Withdrawn 5000.00 BDT" / "5000.00 BDT"
    r'(?:Withdrawn|Deposited)\s+([\d,]+\.?\d*)\s+BDT',
    # IBBL clearing: "Tk.80399.00 has been clearing withdrawn/deposited"
    #               "Tk 1153134.00 has been clearing deposited"
    r'[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+has\s+been\s+clearing\s+(?:withdrawn|deposited)',
    # Standard Bank / Standard Islami
    r'BDT\s+([\d,]+\.?\d*)\s+(?:credited|debited)',
    # IBBL "Deposited X.XX BDT for Card"
    r'Deposited\s+([\d,]+\.?\d*)\s+BDT\s+for\s+Card',
    # ATM credited: "ATM Tk 145039 has been credited"
    r'ATM\s+[Tt][Kk]\.?\s*([\d,]+\.?\d*)\s+has\s+been\s+credited',
    # EBL / bank CHEQUE CLEARING: "AC XXXX is credited with BDT XXXXX as CHEQUE CLEARING"
    r'is\s+credited\s+with\s+BDT\s+([\d,]+\.?\d*)',
    # EBL debit: "AC XXXX is debited with BDT XXXXX as TRANSFER"
    r'is\s+debited\s+with\s+BDT\s+([\d,]+\.?\d*)',
    r'Payment\s+of\s+BDT\s+([\d,]+\.?\d*)\s+has\s+been\s+credited',
    # IBBL CellFin: "Received 200000.00Tk"
    r'[Rr]eceived\s+([\d,]+\.?\d*)\s*[Tt]k\b',
    # IBBL CellFin / IBBPLC: "Deposit Amount: 100,000.00"
    r'[Dd]eposit\s+[Aa]mount\s*:\s*([\d,]+\.?\d*)',
    # bKash Interest / savings return
    r'bkash.e\s+[Tt]k\s*([\d,]+\.?\d*)\s+interest\s+peyechhen',
    # MetLife: "BDT** 11377.00" (asterisks between BDT and amount)
    r'BDT\s*\*+\s*([\d,]+\.?\d*)',
    # Bangla taka sign ৳ — MetLife IDRAMETLIFE: "৳১১৩৭৭.০০"
    r'৳\s*([\d,]+\.?\d*)',
    # Bangla MFS patterns
    r'আপনি\s+([\d,]+\.?\d*)\s*টাকা\s+(?:পেয়েছেন|ক্যাশ\s*ইন)',
    r'([\d,]+\.?\d*)\s*টাকা\s+(?:পেয়েছেন|জমা\s*হয়েছে)',
    r'([\d,]+\.?\d*)\s*টাকা\s+ক্যাশ\s*আউট',
    r'[Pp]aid[:\s]*([\d,]+)',
    # ── v13: পূর্বে বাদ-পড়া ফরম্যাট (Janata/AB Bank/bKash/Nagad/ATM) ──────────
    # Janata Bank PLC: "account *7637 debited with BDT  64000.00 from"
    r'(?:credited|debited)\s+with\s+BDT\s*([\d,]+\.?\d*)',
    # Janata interest line: "credited for interest: BDT2.13"
    r'(?:credited|debited)\s+for\s+interest:\s*BDT\s*([\d,]+\.?\d*)',
    # AB Bank i-banking: "Income Tax Transaction of BDT 900 debited from A/C"
    r'Transaction\s+of\s+BDT\s*([\d,]+\.?\d*)\s+(?:credited|debited)',
    # bKash: "received Tk 5,000.00 from" (Tk mixed-case, not only TK)
    r'received\s+[Tt]k\s*([\d,]+\.?\d*)\s+from',
    # bKash: "Payment of Tk 1,410.00 to / is being reserved for"
    r'[Pp]ayment\s+of\s+[Tt]k\s*([\d,]+\.?\d*)\s+(?:to|is\s+being\s+reserved)',
    # bKash: "received deposit from iBanking of Tk 5,000.00 from Janata Bank"
    r'received\s+deposit\s+from\s+iBanking\s+of\s+[Tt]k\s*([\d,]+\.?\d*)',
    # AB credit card: "BDT 500.00 received by CASH to your credit card"
    r'BDT\s*([\d,]+\.?\d*)\s+received\s+by',
    # AB Bank card payment: "AB Bank: Payment BDT 28137.00 from ABDIRECT"
    r'Payment\s+BDT\s*([\d,]+\.?\d*)\s+from',
    # Generic bank ATM/interest: "debited (ATM Cash Withdrawal) by Tk3,500.00"
    r'debited\s+\([^)]*\)\s+by\s+[Tt]k\s*([\d,]+\.?\d*)',
    r'credited\s+\([^)]*\)\s+by\s+[Tt]k\s*([\d,]+\.?\d*)',
    # AB Card Bill: "Total Due: 17067.36BDT"
    r'[Tt]otal\s+Due:\s*([\d,]+\.?\d*)\s*BDT',
    # Islami Bank Hospital receipt: "Received with thanks total Tk.  17000-"
    r'Received\s+with\s+thanks\s+total\s+[Tt]k\.?\s*([\d,]+)',
    # ── v14: Sonali Bank + DNCC holding tax ফরম্যাট ──────────────────────────
    # Sonali: "credited/debited from [Branch] for BDT 78,106.00 on" (lakh grouping 4,00,000)
    r'(?:credited|debited)\s+from\s+.+?\s+for\s+BDT\s*([\d,]+\.?\d*)\s+on',
    # Sonali: "with [Branch] has been credited for BDT 7,78,819.00 on"
    r'has\s+been\s+credited\s+for\s+BDT\s*([\d,]+\.?\d*)\s+on',
    # Sonali interest line: "Credited for Interest: BDT 68.00"
    r'[Cc]redited\s+for\s+Interest:\s*BDT\s*([\d,]+\.?\d*)',
    # Sonali charge line: "Debited for Maintenance Charge: BDT 200.00"
    r'[Dd]ebited\s+for\s+(?:Maintenance|SMS|TDS|VAT|Charge).*?BDT\s*([\d,]+\.?\d*)',
    # DNCC / municipal tax: "Tk5,132.16 paid to DNCC Holding Tax"
    r'[Tt]k\s*([\d,]+\.?\d*)\s+paid\s+to',
    # ── v15: bKash Digital Loan (City Bank) + payment-successful (no 'to') ──
    # "Digital Loan Repayment of Tk 1,009.24 is successful"
    r'Loan\s+Repayment\s+of\s+[Tt]k\s*([\d,]+\.?\d*)',
    # "received Loan of Tk 2,983" / "received Digital Loan Tk 2,000"
    r'received\s+(?:Digital\s+)?Loan\s+(?:of\s+)?[Tt]k\s*([\d,]+\.?\d*)',
    # "Payment of Tk 1,410.00 is successful" (no 'to' / 'reserved')
    r'[Pp]ayment\s+of\s+[Tt]k\s*([\d,]+\.?\d*)\s+is\s+successful',
    # Nagad/bKash bill: "Amount: Tk 2000.00" (append LAST — most generic)
    r'Amount:\s*[Tt]k\s*([\d,]+\.?\d*)',
]

def get_amount(text):
    # ── ধাপ ০: বাংলা সংখ্যা → ASCII digit convert (MetLife IDRAMETLIFE format) ──
    bn_map = str.maketrans('০১২৩৪৫৬৭৮৯', '0123456789')
    text   = text.translate(bn_map)

    # ── ধাপ ১: balance / fee / A/C number জাতীয় সংখ্যা বাদ দিই ──────────────
    c = re.sub(r'C/B\s+Tk[\d,]+\.?\d*','',text)
    c = re.sub(r'Avl\s+Bal[:\s]+[\d,]+\.?\d*','',c)
    c = re.sub(r'Avail\s+Bal[:\s]+BDT\s+[\d,]+\.?\d*','',c,flags=re.I)
    c = re.sub(r'A/C\s+Bal\s+BDT\s+[\d,.]+','',c)
    # "Balance TK X" / "Balance: TK X" / "Balance is BDT X"
    c = re.sub(r'[Bb]alance[:\s\.]+[Tt]?[Kk]?\.?\s*[\d,]+\.?\d*','',c)
    c = re.sub(r'(?:[Aa]vl?\.?\s+)?[Bb]al(?:ance)?[:\s\.]+[Tt]?[Kk]?\.?\s*[\d,]+\.?\d*','',c)
    c = re.sub(r'[Ff]ee[:\s]+[Tt][Kk]?\.?\s*[\d,]+\.?\d*','',c)
    # "Ref X" — bKash Ref number বাদ
    c = re.sub(r'\bRef\s+\d+\b','',c,flags=re.I)
    # বাংলা ব্যালেন্স — "নতুন ব্যালেন্স: X টাকা"
    c = re.sub(r'(?:নতুন|বর্তমান)\s+ব্যালেন্স[:\s]+[\d,]+\.?\d*\s*টাকা','',c)
    # MetLife Policy number — "Policy-102268" / "পলিসি **৬৮৭৮" (masked) → বাদ
    c = re.sub(r'[Pp]olicy[-\s\*]+[\d\*]+','',c,flags=re.I)
    c = re.sub(r'পলিসি\s*[\*\d]+','',c)
    # A/C number — masked বাদ
    c = re.sub(r'A/C\s*#?[\d\*]+','',c,flags=re.I)
    # ── ধাপ ২: TXN_P patterns দিয়ে amount খোঁজো ─────────────────────────────
    # Priority: প্রথম match-এর priority সবচেয়ে বেশি (most specific pattern আগে)
    # শুধু tie-break (কোনো match না হলে) ক্ষেত্রে max() ব্যবহার হয়
    first_val = None
    all_vals  = []
    for p in TXN_P:
        for m in re.findall(p, c, re.I):
            try:
                f = float(str(m).replace(',', ''))
                if f > 0:
                    if first_val is None: first_val = f  # first specific match
                    all_vals.append(f)
            except (ValueError, TypeError): pass
    if first_val is not None: return first_val
    return 0.0

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
    # Tracking link/ID সহ courier SMS — সবসময় রাখো
    r'track\s+your\s+parcel|steadfast\.com\.bd/t/|ptho\.app/'
    # Sundarban / SA Paribahan ইত্যাদি: "Track here: http://…", পিকআপ-অফিস ("ready at … Office"), বুকিং কনফার্ম
    r'|track\s+here\s*:?\s*https?://|scsapps\.\w+'
    r'|(?:parcel|fruits?|product|poly|consignment|goods|item|mango|lichi)[\w /]{0,30}?\s+ready\s+at\s+.{3,80}?office'
    r'|booking\s+confirmed\s+from\s+.{2,40}?(?:https?://|track)'
    r'|track\s+at\s+https|parcel\s+id.*track|tracking\s+id'
    r'|redx\.com\.bd/\S+|ecourier\.com\.bd/\S+|paperfly\.com\.bd/\S+'
    r'|track(?:ing)?\s+(?:id|no|code)\s*[:#]?\s*\w{5,}'
    r'|bus\s+point.*dhaka|dhaka.*bus\s+point|seat.*reserved.*from.*to'  # bus ticket
    # Delivery status — শুধু tracking link/domain থাকলে রাখো
    r'|(?:parcel|shipment|order).*(?:picked\s+up|out\s+for\s+delivery|delivered\s+successfully|being\s+returned).*(?:https?://|steadfast|ptho|redx|track)'
    r'|rider.*(?:will\s+contact|is\s+coming|is\s+nearby|call\s+rider)'
    , re.I | re.S)

# Courier messages to FILTER (OTP, pickup, cancelled)
COURIER_FILTER_RE = re.compile(
    r'ডেলিভারি\s*ওটিপি|delivery\s*otp|অর্ডারটির\s*ডেলিভারি'
    r'|আমরা.*থেকে\s+আপনার\s+পণ্য.*সংগ্রহ'  # pickup notification
    r'|cash\s+on\s+delivery.*যাচাই|verification\s+code.*delivery'
    # Parcel status notifications with Parcel ID (cancelled/picked up etc.)
    r'|parcel\s+id\s*#?\d+.*(?:cancelled|canceled|picked|dispatched|delivered)'
    r'|parcel.*#\d+.*(?:cancelled|canceled)'
    , re.I | re.S)

SHOPPING_RE = re.compile(
    r'aarong|invoice|\bvn:|thanks\s+for\s+choosing|receipt|order\s*#|star\s+tech'
    r'|walton|কিস্তি|installment|barcode.*নিবন্ধন'
    r'|desco|dpdc|wasa|titas\s+gas|বিদ্যুৎ\s+বিল', re.I)

# শিক্ষা প্রতিষ্ঠান — ভর্তি / পেমেন্ট / রেজাল্ট
EDU_RE = re.compile(
    r'admission|ভর্তি|semester\s+fee|tuition\s+fee|semester\s+payment'
    r'|student\s+id|roll\s+no.*exam|result\s+published|পরীক্ষার\s+ফলাফল'
    r'|university|college|school|বিশ্ববিদ্যালয়|মহাবিদ্যালয়|বিদ্যালয়'
    # hsc/ssc — context দেখো, shop abbreviation বাদ দাও
    r'|\bhsc\s+(?:exam|result|admission|fee|form|pass|fail|batch|২০\d\d)\b'
    r'|\bhsc\b(?!\s*(?:kuala|kl|kual|malaysia|my|street|road|tower|mall|plaza|building|avenue))'
    r'|\bssc\s+(?:exam|result|admission|fee|form|pass|fail|batch|২০\d\d)\b'
    r'|\bjsc\b|\bpsc\b'
    r'|\bdiu\b|\bnsu\b|\bbrac\s+university\b|\buiu\b'
    r'|আইইউবি|নর্থসাউথ|ব্র্যাক\s+বিশ্ববিদ্যালয়'
    r'|hall\s+fee|hall\s+due|hostel\s+fee|registration\s+fee.*university'
    r'|payment.*received.*university|university.*payment'
    , re.I)

# সরকারি চালান / ট্রেজারি / কোর্ট ফি — রাখতে হবে
CHALAN_RE = re.compile(
    r'challan|chalan|চালান|treasury\s*challan|সরকারি\s*চালান'
    r'|court\s*fee|কোর্ট\s*ফি|court\s*challan'
    r'|ibas\+\+|iBAS|treasury.*payment|government.*challan'
    r'|revenue\s*stamp|রাজস্ব|tax\s*payment.*treasury'
    r'|brtc.*fee|brta.*fee|driving\s*license.*fee|vehicle.*registration.*fee'
    r'|passport\s*fee|জাতীয়\s*রাজস্ব|national\s*board\s*of\s*revenue'
    , re.I)

# BO Account (Beneficiary Owner — শেয়ার মার্কেট)
BO_RE = re.compile(
    r'\bBO\s+(?:account|a/c|no|id)\b|\bbeneficiary\s+owner\b'
    r'|\bCDBC\b|\bCDSL\b|\bBSEC\b|stock\s+exchange|share\s+market'
    r'|\bDSE\b|\bCSE\b|bo\s+account|বিও\s+একাউন্ট'
    r'|demat\s+account|portfolio.*shares?'
    , re.I)

# অতিরিক্ত OTP ফরম্যাট: "Submit OTP :257267", "OTP: 1234", "OTP-123456" ইত্যাদি
OTP_EXTRA_RE = re.compile(r'\bOTP\s*[:：\-]?\s*(?:is\s*)?[:：]?\s*\d{4,8}\b|\b\d{4,8}\s+is\s+(?:the|your)\s+(?:OTP|code)\b'
                          r'|ওটিপি\s*[:：]?\s*[\d০-৯]{4,8}|গোপন\s*(?:কোড|নম্বর)', re.I)
OTP_RE = OTP_BODY_RE  # alias — same regex
LOW_BAL_RE = re.compile(r'your\s+balance\s+is\s+(very\s+low|finished)', re.I)

# রিয়েল এস্টেট প্রমো — বাদ দেওয়া হবে
REAL_ESTATE_FILTER_RE = re.compile(
    r'জমি\s*কিনুন|ফ্ল্যাট\s*কিনুন|flat\s*(?:sale|sell|buy|for\s*sale)'
    r'|house\s*(?:sale|sell|for\s*sale|wanted)|land\s*(?:sale|sell|for\s*sale|wanted)'
    r'|plot\s*(?:sale|sell|for\s*sale)|apartment\s*for\s*sale'
    r'|জমি\s*বিক্রয়|ফ্ল্যাট\s*বিক্রয়|বাড়ি\s*বিক্রয়|জমি\s*পাওয়া\s*যাচ্ছে'
    r'|real\s*estate|property\s*(?:sale|sell|for\s*sale)'
    r'|বিক্রয়\s*হবে.*(?:জমি|ফ্ল্যাট|বাড়ি|প্লট)'
    r'|রেডি\s*প্লট|ready\s*plot|প্লট\s*বিক্রয়|plot\s*বিক্রয়'
    r'|ফ্ল্যাট\s*বিক্রি|flat\s*বিক্রি|apartment\s*বিক্রয়'
    r'|(?:আবাসিক|বাণিজ্যিক)\s*(?:প্লট|ফ্ল্যাট|জমি).*(?:বিক্রয়|বিক্রি|পাওয়া)'
    r'|duplex|(?:কাঠা|শতাংশ|decimal).*(?:জমি|প্লট).*(?:বিক্রয়|বিক্রি)'
    , re.I)

# ── একটি SMS বিশ্লেষণ ──────────────────────────────────────
# Body থেকে bank/MFS service detect করার জন্য keywords
BODY_BANK_RE = re.compile(
    r'\bislami\s+bank\b|\bIBBL\b|\bBRAC\s+Bank\b|\bTrust\s+Bank\b|'
    r'\bSonali\s+Bank\b|\bDutch.Bangla\b|\bDBBL\b|\bmutual\s+trust\b|'
    r'\bCity\s+Bank\b|\bEastern\s+Bank\b|\bEBL\b',
    re.I
)
BODY_MFS_RE = re.compile(
    r'\bbKash\b|\bNagad\b|\bRocket\b|\bUpay\b|\bMcash\b|\bBinimoy\b', re.I
)

def _bank_from_text(bl):
    """Body বা sender lowercase থেকে bank/mfs service শনাক্ত করে।
    detect_service_from_body() ও get_acct() উভয়ই এটা ব্যবহার করে।
    Returns: ("bank"/"mfs"/"") or service name string when name=True."""
    # MFS
    if re.search(r'\bbkash\b', bl):                          return "mfs"
    if re.search(r'\bnagad\b', bl):                          return "mfs"
    if re.search(r'\bbinimoy\b', bl):                        return "mfs"
    if re.search(r'\brocket\b', bl) and "balance" in bl:     return "mfs"
    # Bank keywords — same list used in both functions
    if "islami bank" in bl or re.search(r'\bibbl\b', bl):   return "bank"
    if "standard islami" in bl or "standardbnk" in bl:       return "bank"
    if "standard bank" in bl:                                 return "bank"
    if "brac bank" in bl or "bracbank" in bl:                 return "bank"
    if "trust bank" in bl or "trustbank" in bl:               return "bank"
    if "sonali bank" in bl:                                    return "bank"
    if "dutch-bangla" in bl or re.search(r'\bdbbl\b', bl):   return "bank"
    if "mutual trust" in bl or re.search(r'\bmtb\b', bl) or "mtblbddh" in bl: return "bank"
    if "city bank" in bl:                                      return "bank"
    if "southeast bank" in bl or "sebl" in bl:                return "bank"
    if "ab bank" in bl:                                        return "bank"
    if "prime bank" in bl:                                     return "bank"
    if "ucb" in bl and "bank" in bl:                           return "bank"
    if "eastern bank" in bl or re.search(r'\bebl\b', bl):    return "bank"
    if "stanchart" in bl or "stancharbd" in bl:               return "bank"
    # Islamic banks
    if "al-arafah" in bl or "al arafah" in bl:               return "bank"
    if "shahjalal" in bl or re.search(r'\bsjibl\b', bl):     return "bank"
    if "exim bank" in bl or "eximbank" in bl:                 return "bank"
    if "first security" in bl or re.search(r'\bfsibl\b', bl): return "bank"
    if "global islami" in bl:                                  return "bank"
    if "union bank" in bl:                                     return "bank"
    # Private banks
    if "mercantile bank" in bl:                               return "bank"
    if "one bank" in bl or re.search(r'\bonebank\b', bl):     return "bank"
    if "nrb bank" in bl or "nrb commercial" in bl:            return "bank"
    if "midland bank" in bl:                                   return "bank"
    if "community bank" in bl:                                 return "bank"
    if "padma bank" in bl:                                     return "bank"
    if "jamuna bank" in bl:                                    return "bank"
    if "modhumoti bank" in bl:                                 return "bank"
    return ""

def detect_service_from_body(body):
    """Body থেকে ব্যাংক/MFS শনাক্ত করে।
    Note: "received from Sonali Bank Account" is bKash SMS — body mention != sender bank"""
    bl = body.lower()
    # bKash/Nagad transfer FROM bank messages — এগুলো MFS transaction
    if re.search(r'received\s+a?\s*deposit\s+of\s+tk.*from\s+\w+\s+bank\s+account', bl): return "mfs"
    # bKash TrxID format: DD[A-Z0-9]{6,} — body-তে "bKash" word না থাকলেও detect করে
    # (personal BD number থেকে bKash notification relay হলে)
    if re.search(r'\bTrxID\s+DD[A-Z0-9]{6,}', body, re.I): return "mfs"
    return _bank_from_text(bl)

def analyze_one(sender, body, date, direction=""):
    st  = stype(sender, body)
    amt = get_amount(body)
    t   = tier(amt)
    cats = []

    # Foreign bank (AlRajhi MY, Maybank etc.) — MYR/SGD/AED amount বের করো
    if st == "foreign_bank":
        f_amt   = get_foreign_amount(body)
        f_ccy   = get_foreign_currency(body)
        if f_amt > 0:
            cats.append("travel_txn")
        elif TRAVEL_RE.search(body):
            cats.append("travel")
        else:
            cats.append("travel_txn")
        acct = get_acct(sender, body, f_amt)
        if not acct.get("service"):
            acct["service"] = "AlRajhi Bank (Malaysia)" if is_foreign_bank(sender) else "Foreign Bank"
        acct["amount"]   = f_amt
        acct["currency"] = f_ccy
        return {"sender":str(sender),"body":body,"date":str(date) if date else "",
                "month":(date.strftime('%Y-%m') if isinstance(date,datetime) else str(date)[:7]),
                "sender_type":"foreign_bank","categories":cats,
                "amount":f_amt,"tier":"","acct":acct,"direction":direction,
                "currency":f_ccy}

    # Personal নম্বর হলেও body-তে bank/MFS থাকলে সেই type অনুযায়ী চলবে
    body_st = detect_service_from_body(body)
    if st == "personal" and body_st in ("bank","mfs"):
        st = body_st

    # MFS/Bank — OTP ও low balance সবার আগে বাদ (tier থাকলেও)
    if st in ("mfs","bank"):
        if OTP_RE.search(body) or LOW_BAL_RE.search(body): return None

    if t: cats.append(f"txn_{t}")

    # MFS/Bank — tier নেই → sidebar summary
    if st in ("mfs","bank") and not t:
        acct_tmp = get_acct(sender, body, amt)
        is_balance_sms = bool(re.search(r'\bbalance\b', body, re.I))
        if acct_tmp.get("balance",0) > 0 or amt > 0 or (is_balance_sms and acct_tmp.get("service")):
            cats.append("bank_info")   # sidebar only
        # bank shortcode থেকে courier/travel/service SMS আসতে পারে (e.g. 16247 = bKash courier relay)
        elif COURIER_RE.search(body):  cats.append("courier")
        elif TRAVEL_RE.search(body):   cats.append("travel")
        elif SHOPPING_RE.search(body): cats.append("shopping")

    if st == "personal":
        if ID_RE.search(body):         cats.append("identity")
        elif LOVE_RE.search(body):     cats.append("love_family")
        elif TRAVEL_RE.search(body):   cats.append("travel")
        elif COURIER_RE.search(body):  cats.append("courier")
        elif CHALAN_RE.search(body):   cats.append("chalan")
        elif EDU_RE.search(body):      cats.append("education")
        elif BO_RE.search(body):       cats.append("bo_account")
        elif SHOPPING_RE.search(body): cats.append("shopping")
        elif len(body.strip()) > 20:   cats.append("personal")
        if not cats: return None

    elif st in ("service","other"):
        if TRAVEL_RE.search(body):        cats.append("travel")
        elif COURIER_RE.search(body):     cats.append("courier")
        elif CHALAN_RE.search(body):      cats.append("chalan")
        elif EDU_RE.search(body):         cats.append("education")
        elif BO_RE.search(body):          cats.append("bo_account")
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
def _norm_acct_id(x):
    """Masking char normalize: x/X → *, runs → single *****।"""
    x = re.sub(r'[xX]', '*', x)
    x = re.sub(r'\*{2,}', '*****', x)
    return x

def extract_account_id(body):
    """মেসেজ থেকে account/card identifier ও ধরন বের করে।
    Returns (display_id, acct_type, canon_key) বা (None, None, None)।
    acct_type: 'card' (ক্রেডিট/ডেবিট কার্ড) | 'account' (সাধারণ একাউন্ট)।
    canon_key = type + last-4 digit — একই কার্ড/একাউন্টের ভিন্ন masking এক গোনা হয়।"""
    b = body
    did = typ = None
    # ── CARD আগে (card মেসেজে A/C-ও থাকতে পারে) ──
    m = re.search(r'(?:credit\s+card|debit\s+card|\bcard)\s*[:#]?\s*(\d{3,6}[\*xX]{1,}\d{3,6})', b, re.I)
    if m: did, typ = _norm_acct_id(m.group(1)), 'card'
    if not did:
        m = re.search(r'card\s+ending\s+(\d{4,6})', b, re.I)
        if m: did, typ = _norm_acct_id(m.group(1)), 'card'
    if not did:
        m = re.search(r'\b(\d{6}[\*xX]{4,8}\d{4})\b', b)   # 16-digit card BIN
        if m: did, typ = _norm_acct_id(m.group(1)), 'card'
    # ── ACCOUNT ──
    if not did:
        m = re.search(r'(?:A/C|A/c|account|acc\.?)\s*(?:no\.?|#|id)?\s*[:.]?\s*'
                      r'(\d{0,6}[\*xX]{2,}\d{3,6}|\*{1,}\d{4,6})', b, re.I)
        if m: did, typ = _norm_acct_id(m.group(1)), 'account'
    if not did:
        return None, None, None
    last4 = re.sub(r'\D', '', did)[-4:]
    canon = f"{typ}:{last4}"
    return did, typ, canon

def get_acct(sender, body, amt):
    sl = str(sender).lower()
    s  = str(sender).strip()
    bl = body.lower()
    info = {"service":"","trx_id":"","direction":"","other_number":"","balance":0,"amount":amt}

    # Short-code bank senders
    if s == "16216": info["service"] = "Dutch-Bangla Bank (DBBL)"
    elif s == "16219": info["service"] = "Mutual Trust Bank (MTB)"
    elif re.match(r'^mtb\.?\s*$', sl): info["service"] = "Mutual Trust Bank (MTB)"
    elif re.match(r'^\[?\s*!\s*\]?$', s): info["service"] = "Dutch-Bangla Bank (DBBL)"

    # Sender keyword detect
    if not info["service"]:
        if   "bkash"    in sl: info["service"]="bKash"
        elif "nagad"    in sl: info["service"]="Nagad"
        elif "rocket"   in sl: info["service"]="Rocket"
        elif "binimoy"  in sl: info["service"]="Binimoy"
        # bKash TrxID format (DD[A-Z0-9]+) — sender personal কিন্তু body bKash
        elif re.search(r'\bTrxID\s+DD[A-Z0-9]{6,}', body, re.I): info["service"]="bKash"
        elif "dutch-bangla" in sl or "dutchbangla" in sl or "dbbl" in sl \
             or "nexuspay" in sl or "nexus pay" in sl: info["service"]="Dutch-Bangla Bank (DBBL)"
        elif "upay"     in sl: info["service"]="Upay"
        elif "ibbl"     in sl or "islami" in sl: info["service"]="Islami Bank"
        elif "trustbank" in sl or "trust bank" in sl: info["service"]="Trust Bank"
        elif "trustcard" in sl: info["service"]="Trust Card"
        elif "standardbnk" in sl or "standard bank" in sl or "standard islami" in sl:
            info["service"]="Standard Bank"
        elif "sonali"   in sl: info["service"]="Sonali Bank"
        elif re.search(r'\bab\s*bank\b|\babbank\b', sl): info["service"]="AB Bank"
        elif "brac"     in sl: info["service"]="BRAC Bank"
        elif "metlife"  in sl: info["service"]="MetLife"
        elif "stanchart" in sl or "stancharbd" in sl: info["service"]="Standard Chartered Bank"

    # Body fallback — যেকোনো নম্বর থেকে আসলেও bank/MFS detect করবে
    # _bank_from_text() helper ব্যবহার করে code duplication এড়ানো হচ্ছে
    if not info["service"]:
        _bt = _bank_from_text(bl)
        if _bt == "mfs":
            # MFS service name
            if re.search(r'\bbkash\b', bl):      info["service"] = "bKash"
            elif re.search(r'\bnagad\b', bl):    info["service"] = "Nagad"
            elif re.search(r'\bbinimoy\b', bl):  info["service"] = "Binimoy"
            elif re.search(r'\brocket\b', bl):   info["service"] = "Rocket"
            else:                                info["service"] = "MFS"
        elif _bt == "bank":
            # Bank service name — keyword-to-name map
            _BANK_MAP = [
                (r'\bibbl\b|islami bank',              "Islami Bank"),
                (r'standard islami|standardbnk|standard bank', "Standard Bank"),
                (r'brac bank|bracbank',                "BRAC Bank"),
                (r'mutual trust|mtblbddh|\bmtb\b',     "Mutual Trust Bank (MTB)"),
                (r'trust bank|trustbank',              "Trust Bank"),
                (r'sonali bank',                       "Sonali Bank"),
                (r'dutch-bangla|\bdbbl\b|nexuspay',    "Dutch-Bangla Bank (DBBL)"),
                (r'southeast bank|sebl',               "Southeast Bank"),
                (r'prime bank',                        "Prime Bank"),
                (r'eastern bank|\bebl\b',              "Eastern Bank"),
                (r'city bank',                         "City Bank"),
                (r'stanchart|stancharbd',              "Standard Chartered Bank"),
                (r'al-arafah|al arafah',               "Al-Arafah Islami Bank"),
                (r'shahjalal|\bsjibl\b',               "Shahjalal Islami Bank"),
                (r'exim bank|eximbank',                "EXIM Bank"),
                (r'first security|\bfsibl\b',          "First Security Islami Bank"),
                (r'global islami',                     "Global Islami Bank"),
                (r'union bank',                        "Union Bank"),
                (r'mercantile bank',                   "Mercantile Bank"),
                (r'one bank|\bonebank\b',              "ONE Bank"),
                (r'nrb bank|nrb commercial',           "NRB Bank"),
                (r'midland bank',                      "Midland Bank"),
                (r'community bank',                    "Community Bank Bangladesh"),
                (r'padma bank',                        "Padma Bank"),
                (r'jamuna bank',                       "Jamuna Bank"),
                (r'modhumoti bank',                    "Modhumoti Bank"),
            ]
            for pat, name in _BANK_MAP:
                if re.search(pat, bl):
                    info["service"] = name
                    break
            if not info["service"]:
                info["service"] = s
        else:
            info["service"] = s

    trx = re.search(r'(?:[Tt]rx|[Tt]xn)\s*[Ii][Dd]?[:\s]+([A-Z0-9]{8,})', body)
    if trx: info["trx_id"]=trx.group(1)
    # বিল-পেমেন্ট: "Biller: DPDC (PREPAID) A/C: 15713091 Contact: 017…" — Contact নম্বরটা
    # গ্রাহকের নিজের, কাউন্টারপার্টি নয়। তাই অপর পক্ষ = বিলার।
    _bm = re.search(r'Biller\s*:\s*([A-Za-z][A-Za-z0-9 .&\-]*?)\s*(?:\(([A-Za-z ]+)\))?\s*(?:A/C|Account|Contact|Amount|,|$)', body)
    if _bm:
        info["biller"] = (_bm.group(1).strip() + (f" ({_bm.group(2).strip()})" if _bm.group(2) else "")).strip()
        _ba = re.search(r'Biller\s*:.*?A/C\s*:\s*(\d{4,})', body)
        if _ba: info["biller_acct"] = _ba.group(1)
    # মার্চেন্ট পেমেন্ট: "Payment of Tk 1,410.00 to Shohoj Limited-1 is successful" /
    # Nagad "Payment to 'MY GP' is Successful"
    _mm = re.search(r"payment\s+(?:of\s+tk\s*[\d,.]+\s+)?to\s+'?(.+?)'?\s+(?:is\s+)?(?:successful|being\s+reserved)", body, re.I)
    if _mm and not _bm:
        _mn = _mm.group(1).strip()
        _mn = re.sub(r'-(?:RM|MYGP)[\w\-]*$', '', _mn)
        _mn = re.sub(r'-\d+$', '', _mn).strip(" -'")
        if _mn and not re.match(r'^01\d{9}$', _mn): info["merchant"] = _mn
    nums = re.findall(r'\b(01[3-9]\d{8})\b', body)
    if _bm:   # বিলের Contact নম্বর বাদ
        _cm = re.search(r'Contact\s*:\s*(01[3-9]\d{8})', body)
        nums = [n for n in nums if not (_cm and n == _cm.group(1))]
    if nums: info["other_number"]=nums[0]

    # Balance — সব bank format cover করে
    bal = re.search(r'C/B\s+Tk([\d,]+\.?\d*)', body)
    if not bal: bal = re.search(r'Avail?\s+Bal[:\s]+BDT\s+([\d,]+\.?\d*)', body, re.I)
    if not bal: bal = re.search(r'Avl\s+Bal[:\s]+([\d,]+\.?\d*)', body)
    if not bal: bal = re.search(r'available\s+balance.*?(?:is\s+)?BDT\s+([\d,]+\.?\d*)', body, re.I)
    if not bal: bal = re.search(r'A/C\s+Bal\s+BDT\s+([\d,.]+)', body)
    # EBL format: "Balance is BDT 77726062.25"
    if not bal: bal = re.search(r'[Bb]alance\s+is\s+BDT\s+([\d,]+\.?\d*)', body)
    # EBL format: "Balance: BDT 166240"
    if not bal: bal = re.search(r'[Bb]alance[:\s]+BDT\s+([\d,]+\.?\d*)', body)
    # General: "Balance Tk/TK X" or "Bal: X"
    if not bal: bal = re.search(r'[Bb]al(?:ance)?[:\s\.]+[Tt]?[Kk]?\.?\s*([\d,]+\.?\d*)', body)
    # Jamuna Bank: "Yearend Balance...was BDT 802.88 as on"
    if not bal: bal = re.search(r'[Bb]alance.*?\bwas\s+BDT\s+([\d,]+\.?\d*)', body, re.I)
    # Mercantile Bank: "balance on 31.12.2025 is 0.00"
    if not bal: bal = re.search(r'balance\s+on\s+[\d./]+\s+is\s+([\d,]+\.?\d*)', body, re.I)
    if bal:
        try: info["balance"]=float(bal.group(1).replace(',',''))
        except (ValueError, AttributeError): pass

    # Account number — masked format (EBL: "AC 116***957", BRAC: "A/C#10656**0001")
    acc_m = re.search(r'AC\s+([\d\*]+(?:\d|\*))', body, re.I)
    if not acc_m: acc_m = re.search(r'A/C\s*#?\s*([\d\*]+(?:\d|\*))', body, re.I)
    # SBI format: "A/c. ID: XXXXX89781XX00" or "SBI account XXXXX89781XX00"
    if not acc_m: acc_m = re.search(r'A/c\.?\s*ID[:\s]+([A-Z\d]+)', body, re.I)
    if not acc_m: acc_m = re.search(r'SBI\s+account\s+([A-Z\d]+)', body, re.I)
    if acc_m:
        # trailing asterisks বাদ দিই — "857**00008501**" → "857**00008501"
        info["account"] = re.sub(r'\*+$', '', acc_m.group(1))

    # নতুন: robust account/card identifier + ধরন (card vs account) শনাক্ত
    _aid, _atype, _acanon = extract_account_id(body)
    if _aid:
        info["acct_id"]   = _aid
        info["acct_type"] = _atype       # 'card' | 'account'
        info["acct_canon"] = _acanon     # একই account/card merge করার key
        if not info.get("account"):
            info["account"] = _aid

    bl2 = body.lower()

    # Priority 0: MetLife / Insurance premium — customer pays = Sent
    if re.search(r'acknowledge\s+receipt.*premium|গৃহীত\s+পলিসি|প্রিমিয়াম.*গৃহীত', bl2):
        info["direction"] = "Sent"
    # IBBL CellFin: "Received X Tk To: [our acc] From Acc: [external]" = Received
    elif re.search(r'cellfin\s+received|islami\s+bank\s+cellfin', bl2):
        info["direction"] = "Received"
    # IBBL USD/foreign transfer to FRSD/IBWHO = foreign payment = Sent
    elif re.search(r'transferred\s+to\s+frsd|transferred\s+to\s+ibwho|usd\s+endorsement', bl2):
        info["direction"] = "Sent"
    # Priority 1: "Credit to your account" / "Debit from your account"
    elif re.search(r'credit(?:ed)?\s+to\s+your', bl2):
        info["direction"] = "Received"
    elif re.search(r'debit(?:ed)?\s+from\s+your', bl2):
        info["direction"] = "Sent"
    # Priority 2: standalone CREDIT / DEBIT (Trust Bank format: "Txn TK XXX CREDIT AC No")
    elif re.search(r'\bcredit\b', bl2):
        info["direction"] = "Received"
    elif re.search(r'\bdebit\b', bl2):
        info["direction"] = "Sent"
    # Priority 3: explicit credited/debited/withdrawn
    elif re.search(r'\bcredited\b', bl2):
        info["direction"] = "Received"
    elif re.search(r'\bdebited\b|\bwithdrawn\b|\bwithdrawal\b', bl2):
        info["direction"] = "Sent"
    # Priority 4: transferred TO/FROM — account perspective
    elif re.search(r'transfer(?:red)?\s+to\s+(?:nagad|bkash|rocket|upay|mfs|mcash)', bl2):
        info["direction"] = "Sent"
    elif re.search(r'transfer(?:red)?\s+to\s+(?:your\s+)?(?:a/c|account)', bl2):
        # "transferred to your A/C" OR "transferred to A/C [number]" = Received
        info["direction"] = "Received"
    elif re.search(r'transfer(?:red)?\s+to\s+a/c\s*\*', bl2):
        # "transferred to A/C 857**" (masked number) = Received
        info["direction"] = "Received"
    elif re.search(r'transfer(?:red)?\s+to\s+(?:\w[\w\s,]*)\s+a/c\s', bl2):
        # "transferred to Bhadra subBr A/C 857**" = Received (target's branch A/C)
        info["direction"] = "Received"
    elif re.search(r'transfer(?:red)?\s+to\b', bl2):
        info["direction"] = "Sent"    # to somewhere else = Sent
    elif re.search(r'transfer(?:red)?\s+from\b', bl2):
        # "transferred from [Branch/A/C] A/C [target]" = target withdrew = Sent
        info["direction"] = "Sent"
    elif re.search(r'deposited\s+(?:from|to\s+your)', bl2):
        info["direction"] = "Received"
    # Priority 5: MFS & general keywords (including Bangla)
    elif re.search(r'\breceived\b|\bcash\s*in\b|\bdeposit(?:ed)?\b|\bsalary\b'
                   r'|পেয়েছেন|ক্যাশ\s*ইন|জমা\s*হয়েছে'
                   r'|interest\s+peyechhen|interest\s+পেয়েছেন', bl2):
        info["direction"] = "Received"
    elif re.search(r'\bcash[\s\-]*out\b|\bcashed\s+out\b|\bpurchase\b|\bpos\s+purchase\b'
                   r'|\bpayment\s+(?:of\s+)?tk.*to\b|\bpayment\s+of\b|\bpaid\b'
                   r'|\bpayment\s+to\b'                       # Nagad: "Payment to 'X' is Successful. Amount: Tk"
                   r'|\bsend\s+money\b|\bsent\b'
                   r'|\bpay\s*bill\b|\brequest\s+of\s+tk.*processed'
                   r'|ক্যাশ\s*আউট|পাঠিয়েছেন|উত্তোলন\s*করেছেন', bl2):
        info["direction"] = "Sent"
    return info

# ── Data load ──────────────────────────────────────────────
def load_xlsx(path):
    # read_only=False — কিছু GP ফাইলে read_only=True তে data row দেখা যায় না
    try:
        wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    except Exception:
        wb = openpyxl.load_workbook(path, read_only=True)
    # ── সব শিট পড়ি (শুধু active শিট নয়) — কিছু এক্সপোর্টে পুরো ডেটা এক শিটে, অংশবিশেষ আরেক শিটে থাকে ──
    def _parse_sheet(rows):
        if not rows: return []
        hdr = [str(c).strip() if c else "" for c in rows[0]]

        def ci(ns):
            for n in ns:
                for i,h in enumerate(hdr):
                    if n.lower() in h.lower(): return i
            return None

        ti = ci(['starttime','start','date','time'])
        ai = ci(['a party','aparty','sender_a'])
        bi = ci(['b party','bparty','sender_b','sender'])
        mi = ci(['message content','message','body','msg'])
        di = ci(['direction'])

        # ── Format detection ─────────────────────────────────────
        # Format A (016-style): AParty=sender_name, BParty=target_number (same number repeats)
        # Format B (017/880-style): AParty=target_number (repeats), BParty=sender_name
        #
        # Detection: SMSMT rows-এ যে column-এ একটাই value বারবার আসে সেটা target।
        # target যদি BParty-তে → Format A;  target যদি AParty-তে → Format B
        mt_a_vals = []
        mt_b_vals = []
        for r in rows[1:]:
            if ai is None or bi is None: break
            if di is not None and di<len(r) and str(r[di]).strip().upper() in ("SMSMT","MTSMS"):
                _body = str(r[mi]).strip() if mi is not None and mi<len(r) else ""
                if _body and _body != "None":
                    if r[ai]: mt_a_vals.append(str(r[ai]).strip())
                    if r[bi]: mt_b_vals.append(str(r[bi]).strip())

        # সবচেয়ে বেশি repeat হওয়া value-র frequency ratio
        def top_freq_ratio(vals):
            if not vals: return 0.0
            c = Counter(vals)
            top_count = c.most_common(1)[0][1]
            return top_count / len(vals)

        a_ratio = top_freq_ratio(mt_a_vals)  # AParty-তে একটা value কতটুকু dominant
        b_ratio = top_freq_ratio(mt_b_vals)  # BParty-তে একটা value কতটুকু dominant

        # যে column-এ dominant repeat বেশি → সেটাই target column
        # Format A: BParty dominant (BParty=target) → sender = AParty
        # Format B: AParty dominant (AParty=target) → sender = BParty
        use_format_a = (b_ratio > a_ratio)   # BParty-তে target repeat বেশি

        out = []
        for r in rows[1:]:
            def g(i): return str(r[i]).strip() if i is not None and i<len(r) and r[i] and str(r[i]).strip() not in ('','None') else ""
            a_val     = g(ai)
            b_val     = g(bi)
            body      = g(mi)
            raw_date  = r[ti] if ti is not None and ti<len(r) else ""
            date      = to_dt(raw_date) if raw_date else ""
            direction = g(di).upper() if di is not None else ""
            if direction in ("SMSMT","MTSMS"): direction = "SMSMT"
            elif direction in ("SMSMO","MOSMS"): direction = "SMSMO"

            if use_format_a:
                # Format A: AParty=sender, BParty=target
                sender = a_val or b_val
                aparty = b_val
            else:
                # Format B: AParty=target, BParty=sender
                if direction == "SMSMO":
                    sender = a_val   # target নিজে outgoing পাঠাচ্ছে
                    aparty = a_val
                else:
                    sender = b_val or a_val
                    aparty = a_val

            if sender and body:
                # Unicode NFC normalize করি — xlsx থেকে আসা বাংলা text-এ
                # "য়" U+09DF (single char) vs U+09AF+U+09BC (two chars) মিস match এড়াতে
                import unicodedata as _ud
                body   = _ud.normalize('NFC', body)
                sender = _ud.normalize('NFC', sender)
                out.append({"sender": sender, "body": body, "date": date,
                            "aparty": aparty, "direction": direction})

        return out

    out = []
    _prior = Counter()          # আগের শিটগুলোর (date, sender, body, direction) গণনা
    for ws in wb.worksheets:
        try:
            rows = list(ws.iter_rows(values_only=True))
        except Exception:
            continue
        if not rows: continue
        _h0 = " ".join(str(c).lower() for c in rows[0] if c)
        if not re.search(r'message|body|msg|content|sms', _h0): continue
        _this = _parse_sheet(rows)
        _used = Counter(); _cnt = Counter()
        for o in _this:
            k = (str(o["date"])[:16], o["sender"], o["body"], o["direction"])
            _cnt[k] += 1
            # আগের শিটে হুবহু একই মেসেজ যতবার আছে, ততবার পর্যন্ত বাদ (ডুপ্লিকেট শিট)
            if _used[k] < _prior[k]:
                _used[k] += 1; continue
            out.append(o)
        _prior.update(_cnt)
    if not out: return []
    # ── Multipart SMS concatenation ───────────────────────────────────────────
    # UCS-2 (বাংলা) SMS per part: max 67 chars। Same sender+date → consecutive parts
    MULTIPART_LEN = 67
    merged = []
    i = 0
    while i < len(out):
        cur = out[i]
        body_len = len(cur["body"])
        # এই row multipart-এর শুরু হতে পারে যদি body ঠিক MULTIPART_LEN char হয়
        if body_len == MULTIPART_LEN:
            combined = cur["body"]
            j = i + 1
            while j < len(out):
                nxt = out[j]
                # Same sender, same direction, same date (minute-level)
                same_sender = nxt["sender"] == cur["sender"]
                same_dir    = nxt["direction"] == cur["direction"]
                same_min    = str(nxt["date"])[:16] == str(cur["date"])[:16]
                if same_sender and same_dir and same_min:
                    combined += nxt["body"]
                    j += 1
                    if len(nxt["body"]) < MULTIPART_LEN:
                        break  # last part (shorter than max)
                else:
                    break
            merged.append({**cur, "body": combined})
            i = j
        else:
            merged.append(cur)
            i += 1
    return merged

def load_csv_f(path):
    """CSV parser — standard SMS export ও MOne SMSC format সাপোর্ট করে।
    MOne format: OrgAddr=target, DestAddr=sender, body column নেই — empty return।"""
    enc = 'utf-8'   # default
    for _enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            with open(path, encoding=_enc, errors='strict') as f:
                f.read(512)
            enc = _enc
            break
        except (OSError, UnicodeDecodeError):
            continue

    with open(path, encoding=enc, errors='replace') as f:
        try:
            dialect = csv.Sniffer().sniff(f.read(4096))
            f.seek(0)
            reader = csv.DictReader(f, dialect=dialect)
        except csv.Error:
            f.seek(0)
            reader = csv.DictReader(f)
        rows = list(reader)

    if not rows: return []
    keys = list(rows[0].keys())

    # MOne SMSC format — OrgAddr column আছে কিন্তু body নেই
    if any('OrgAddr' in k or 'orgaddr' in k.lower() for k in keys):
        return []   # body ছাড়া analysis সম্ভব নয়

    def _fk(names):
        for n in names:
            for k in keys:
                if n.lower() in k.lower(): return k
        return None

    sk   = _fk(['bparty','b party','sender_b','sender'])
    ak   = _fk(['aparty','a party','sender_a'])
    bk   = _fk(['message content','message','body','msg','content'])
    dk   = _fk(['start','date','time'])
    dirk = _fk(['direction'])

    for r in rows:
        def g(k): return str(r.get(k,'')).strip() if k else ''
        s    = g(sk) or g(ak)
        body = g(bk)
        date = g(dk)
        direction = g(dirk).upper() if dirk else ""
        if direction in ("SMSMT","MTSMS"): direction="SMSMT"
        elif direction in ("SMSMO","MOSMS"): direction="SMSMO"
        if s and body:
            out.append({"sender":s,"body":body,"date":date,"direction":direction})
    return out

# ── HTML ───────────────────────────────────────────────────
CAT_META = {
    "txn_high":   {"label":"High Transaction (100K+)",        "color":"#991b1b","icon":"🔴"},
    "txn_med":    {"label":"Medium Transaction (50K–100K)",   "color":"#dc2626","icon":"🟠"},
    "txn_small":  {"label":"Small Transaction (10K–50K)",     "color":"#ea580c","icon":"🟡"},
    "identity":   {"label":"Identity / Profile Info",         "color":"#0369a1","icon":"👤"},
    "love_family":{"label":"Personal / Family",               "color":"#db2777","icon":"❤️"},
    "travel":     {"label":"Travel / Visa / Umrah",           "color":"#0284c7","icon":"✈️"},
    "travel_txn": {"label":"Foreign Card Transaction",        "color":"#0f766e","icon":"💳"},
    "courier":    {"label":"Courier / Delivery",              "color":"#6d28d9","icon":"📦"},
    "chalan":     {"label":"Govt. Challan / Treasury / Court","color":"#047857","icon":"🏛️"},
    "education":  {"label":"Education / Admission / Payment", "color":"#0891b2","icon":"🎓"},
    "bo_account": {"label":"BO Account / Stock Market",       "color":"#7c3aed","icon":"📈"},
    "shopping":   {"label":"Shopping / Billing",              "color":"#059669","icon":"🛒"},
    "service":    {"label":"Service Message",                 "color":"#475569","icon":"🔔"},
    "personal":   {"label":"Personal Conversation",           "color":"#64748b","icon":"💬"},
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
        "last_balance":0.0,"last_balance_date":"","max_txn":0.0,"numbers":set(),
        "accounts": defaultdict(lambda:{"balance":0.0,"balance_date":"","txn_count":0,
                                        "type":"account","display":"","in":0.0,"out":0.0})
    })
    SVC_COL = {
        "bKash":"#e11d48","Nagad":"#ea580c","Rocket":"#7c3aed","Upay":"#16a34a",
        "Islami Bank":"#1d4ed8","Trust Bank":"#0369a1","Trust Card":"#7c3aed",
        "Sonali Bank":"#047857","BRAC Bank":"#b45309","MetLife":"#0284c7",
        "Dutch-Bangla Bank (DBBL)":"#6d28d9",
        "Mutual Trust Bank (MTB)":"#0891b2","Mutual Trust Bank":"#0891b2",
        "Southeast Bank":"#7c2d12","Prime Bank":"#1e40af","Eastern Bank":"#065f46",
        "City Bank":"#9f1239","Standard Chartered Bank":"#1e40af",
    }
    # Investment/FDR/internal cycling transaction types — Received/Sent হিসেবে count করব না
    # কারণ এগুলো একই টাকা ঘুরে আসে (Govt Securities buy+sell, FDR open+close, internal transfer)
    INTERNAL_TXN_TYPES = {
        'government securities', 'new deposit', 'account close out balance',
        'fund transfer request', 'ebl account transfer',
        'profit liquidation', 'ic interest liquidation', 'tax on profit',
        'excise duty', 'witholding source tax', 'withholding tax',
        'credit card si payment', 'account close', 'fdr', 'tdr',
    }

    for s in sms_list:
        ai = s.get("acct",{})
        if not ai or not ai.get("service"): continue
        svc = ai["service"]; d = summ[svc]
        d["service"] = svc
        d["seen"] = True   # এই bank-এর অন্তত একটি SMS পাওয়া গেছে — sidebar-এ দেখাও

        amt = ai.get("amount",0) or 0
        _dir = (ai.get("direction","") or "")
        body_lower = s.get("body","").lower()
        # কার্ড বিল স্টেটমেন্ট (Total Due/Min Due) — লেনদেন নয়, শুধু তথ্য → IN/OUT-এ গোনা হবে না
        _is_statement = bool(re.search(r'card\s+bill.*(?:total\s+due|min\s+due)', body_lower))
        if amt > 0:
            d["txn_count"] += 1
            # txn type body থেকে detect — investment/internal হলে আলাদা bucket
            is_internal = any(t in body_lower for t in INTERNAL_TXN_TYPES)
            if is_internal or _is_statement:
                d["total_cycling"] = d.get("total_cycling", 0.0) + amt
            elif _dir == "Received":
                d["total_in"]  += amt
            elif _dir == "Sent":
                d["total_out"] += amt
            else:
                # direction অনিশ্চিত — ভুল দিকে না গুনে আলাদা bucket-এ রাখি
                d["total_cycling"] = d.get("total_cycling", 0.0) + amt
            if amt > d["max_txn"]: d["max_txn"] = amt

        # per-account/card tracking — canonical key দিয়ে (card vs account আলাদা)
        acc_canon = ai.get("acct_canon","") or ai.get("account","")
        acc_disp  = ai.get("acct_id","") or ai.get("account","")
        acc_type  = ai.get("acct_type","") or "account"
        sms_date = str(s.get("date",""))
        bal = ai.get("balance",0) or 0
        if acc_canon:
            acc_entry = d["accounts"][acc_canon]
            acc_entry["txn_count"] += 1
            acc_entry["type"] = acc_type
            # সবচেয়ে বিস্তারিত display form রাখি (বেশি সংখ্যা দেখা যায় যেটায়)
            _digits = len(re.sub(r'\D','',acc_disp))
            if _digits >= len(re.sub(r'\D','',acc_entry.get("display",""))):
                acc_entry["display"] = acc_disp
            if amt > 0 and not _is_statement:
                if _dir == "Received":
                    acc_entry["in"] = acc_entry.get("in",0.0) + amt
                elif _dir == "Sent":
                    acc_entry["out"] = acc_entry.get("out",0.0) + amt
                # blank direction → account-level in/out-এও গোনা হবে না
            if bal > 0 and sms_date > acc_entry["balance_date"]:
                acc_entry["balance"]      = bal
                acc_entry["balance_date"] = sms_date
        if bal > 0:
            # service-level last balance — সর্বশেষ date-এর balance
            if sms_date > d["last_balance_date"]:
                d["last_balance"] = bal
                d["last_balance_date"] = sms_date

        if ai.get("other_number"):
            d["numbers"].add(ai["other_number"])

    result = {}
    for svc, d in summ.items():
        # multi-account: account-wise last balance-এর সমষ্টি দেখাও
        accs = dict(d["accounts"])
        acc_total_balance = sum(v["balance"] for v in accs.values() if v["balance"] > 0)
        # ১টির বেশি account থাকলে sum ব্যবহার করো, না হলে single last balance
        display_balance = acc_total_balance if len(accs) > 1 and acc_total_balance > 0 \
                          else d["last_balance"]
        result[svc] = {
            **{k:v for k,v in d.items() if k not in ("accounts",)},
            "numbers": list(d["numbers"])[:4],
            "color": SVC_COL.get(svc, "#2563eb"),
            "accounts": {k: {"balance": v["balance"], "txn_count": v["txn_count"],
                             "type": v.get("type","account"),
                             "display": v.get("display",k),
                             "in": v.get("in",0.0), "out": v.get("out",0.0)}
                         for k,v in accs.items() if v["txn_count"] > 0},
            "display_balance": display_balance,
            "multi_account": len(accs) > 1,
            "total_cycling": d.get("total_cycling", 0.0),
        }

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


# ════════════════════════════════════════════════════════════════════════════
#  INTELLIGENCE METRICS — Timeline, Counterparty, Behavioral flags
# ════════════════════════════════════════════════════════════════════════════
def _im_parse_dt(s):
    """তারিখ string → datetime (একাধিক ফরম্যাট সহ্য করে)।"""
    s = str(s).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try: return datetime.strptime(s[:19].strip() if len(s) >= 19 else s, fmt)
        except: pass
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})[ T]?(\d{2})?:?(\d{2})?', s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4) or 0), int(m.group(5) or 0))
        except: return None
    return None

def _im_eff_dir(s):
    """লেনদেনের কার্যকর দিক: IN (credit) না OUT (debit)।
    credit/debit chart-এর মতো একই লজিক — সামঞ্জস্য বজায় রাখে।"""
    ai = s.get("acct", {}) or {}
    _st = s.get("sender_type", ""); _ad = ai.get("direction", ""); _dir = s.get("direction", "")
    if _st in ("mfs", "bank") and _ad: return "IN" if _ad == "Received" else "OUT"
    if _dir == "SMSMT": return "IN"
    if _dir == "SMSMO": return "OUT"
    if _ad == "Received": return "IN"
    if _ad == "Sent": return "OUT"
    return ""

def _im_acct_lbl(a):
    """একাউন্ট/কার্ড লেবেল — যেমন "A/C 1111*****50314" বা "Card 5457*****4686"।"""
    aid = (a or {}).get("acct_id", "")
    if not aid: return ""
    return ("Card " if (a or {}).get("acct_type") == "card" else "A/C ") + aid

_IM_KIND = [
    (r'deposit\s+from\s+ibanking',              'iBanking deposit'),
    (r'cash\s*deposit',                           'Cash deposit'),
    (r'fund\s*transfer',                          'Fund transfer (iBanking)'),
    (r'\bnpsb\b',                                'NPSB transfer'),
    (r'\bbeftn\b',                               'BEFTN transfer'),
    (r'\brtgs\b',                                'RTGS transfer'),
    (r'cheque',                                   'Cheque'),
    (r'po\s*en(?:c)?hashed|pay\s*order',          'Pay order encashed'),
    (r'\batm\b',                                 'ATM withdrawal'),
    (r'cash[\s\-]*out|cashed\s+out|ক্যাশ\s*আউট',   'Cash-out (agent)'),
    (r'cash\s*in',                                'Cash-in (agent)'),
    (r'send\s+money|\bsent\b|পাঠিয়েছেন',          'Send money'),
    (r'payment\s+bdt.*?to\s+card',                'Card bill payment'),
    (r'\bpayment\b|\bpaid\b|bill',                'Payment'),
    (r'\breceived\b|পেয়েছেন',                    'Received'),
    (r'credited|debited|transfer',                'Bank transfer'),
]

def _im_party(s, side):
    """লেনদেনের অপর পক্ষ/উৎস/গন্তব্য — মোবাইল নম্বর, ব্যাংক, মার্চেন্ট বা লেনদেনের ধরন।"""
    a = s.get("acct", {}) or {}
    b = s.get("body", "") or ""
    kind = next((lbl for p, lbl in _IM_KIND if re.search(p, b, re.I)), "")
    who = a.get("other_number", "") or ""
    if not who:
        m = re.search(r'from\s+ibanking.*?from\s+([A-Za-z][A-Za-z .&]+?\bBank)', b, re.I)
        if m: who = m.group(1).strip()
    if not who:
        m = re.search(r'\bfrom\s+(ABDIRECT|[A-Z]{4,}DIRECT)', b)
        if m: who = m.group(1) + " (online banking)"
    if not who and side == "OUT":   # "To Card …" গন্তব্য — শুধু OUT পাশে প্রযোজ্য
        m = re.search(r'\bto\s+card\s+(\d+\*+\d+)', b, re.I)
        if m: who = "Card " + _norm_acct_id(m.group(1))
    if not who and side == "OUT":
        m = re.search(r"\bto\s+'?([A-Za-z][\w&.\- ]{2,40}?)'?\s+(?:is\s+)?successful", b, re.I)
        if m: who = m.group(1).strip()
    if who and kind: return f"{who} · {kind}"
    return who or kind or "—"

def compute_intel_metrics(display_list, own=""):
    """Timeline, counterparty ranking ও behavioral flag হিসাব করে।"""
    from collections import defaultdict
    M = {}
    txns = [s for s in display_list if (s.get("amount", 0) or 0) > 0]

    # ── ক: Timeline / সময়-প্যাটার্ন ──
    monthly = defaultdict(lambda: {"in": 0.0, "out": 0.0, "cnt": 0})
    hours = Counter(); night = []; dts = []
    for s in txns:
        amt = s["amount"]; eff = _im_eff_dir(s); mon = s.get("month", "")
        if mon:
            if eff == "IN":  monthly[mon]["in"]  += amt
            elif eff == "OUT": monthly[mon]["out"] += amt
            monthly[mon]["cnt"] += 1
        dt = _im_parse_dt(s.get("date", ""))
        if dt:
            dts.append(dt); hours[dt.hour] += 1
            if 0 <= dt.hour <= 5: night.append((dt, amt, eff, s))
    M["monthly"]   = dict(sorted(monthly.items()))
    M["night_cnt"] = len(night)
    M["night_total"] = sum(x[1] for x in night)
    M["night_items"] = sorted(night, key=lambda x: x[0])
    M["first_dt"]  = min(dts) if dts else None
    M["last_dt"]   = max(dts) if dts else None
    M["active_days"] = ((max(dts) - min(dts)).days + 1) if dts else 0
    buckets = {"Night (12AM–6AM)": 0, "Morning (6AM–12PM)": 0,
               "Afternoon (12PM–6PM)": 0, "Evening (6PM–12AM)": 0}
    for h, c in hours.items():
        if   0 <= h < 6:  buckets["Night (12AM–6AM)"] += c
        elif h < 12:      buckets["Morning (6AM–12PM)"] += c
        elif h < 18:      buckets["Afternoon (12PM–6PM)"] += c
        else:             buckets["Evening (6PM–12AM)"] += c
    M["buckets"] = buckets

    # ── খ: Counterparty / Network ──
    recv = defaultdict(lambda: {"amt": 0.0, "cnt": 0})
    sent = defaultdict(lambda: {"amt": 0.0, "cnt": 0})
    freq = Counter()
    cp_svc = defaultdict(Counter)   # counterparty → কোন সার্ভিস/ব্যাংকের মাধ্যমে (bKash/Nagad/Rocket/ব্যাংক)
    cp_svc_in  = defaultdict(lambda: defaultdict(float))   # cp → {service: IN amount}
    cp_svc_out = defaultdict(lambda: defaultdict(float))   # cp → {service: OUT amount}
    for s in txns:
        _a = s.get("acct", {}) or {}
        cp = str(_a.get("other_number", "") or "")
        # কাউন্টারপার্টি নেটওয়ার্কে শুধু মোবাইল নম্বর — বিল (DPDC), মার্চেন্ট/অপারেটর (GP, Banglalink ইত্যাদি)
        # এবং সাবজেক্টের নিজের নম্বর বাদ
        if not re.match(r'^01\d{9}$', cp): continue
        if own and cp == own: continue
        if _a.get("biller") or _a.get("merchant"): continue
        _svc = ((s.get("acct", {}) or {}).get("service", "") or str(s.get("sender", "")).strip() or "Unknown")
        cp_svc[cp][_svc] += 1
        eff = _im_eff_dir(s); amt = s["amount"]; freq[cp] += 1
        if eff == "IN":
            recv[cp]["amt"] += amt; recv[cp]["cnt"] += 1; cp_svc_in[cp][_svc] += amt
        elif eff == "OUT":
            sent[cp]["amt"] += amt; sent[cp]["cnt"] += 1; cp_svc_out[cp][_svc] += amt
    M["top_recv"]  = sorted(recv.items(), key=lambda x: -x[1]["amt"])[:10]
    M["top_sent"]  = sorted(sent.items(), key=lambda x: -x[1]["amt"])[:10]
    M["high_freq"] = [(n, c) for n, c in freq.most_common(10) if c >= 3]
    # প্রতি নম্বরের সার্ভিস লেবেল — বেশি ব্যবহৃতটা আগে (যেমন "bKash", "bKash, Nagad", "AB Bank")
    M["cp_svc"] = {cp: ", ".join(k for k, _ in c.most_common()) for cp, c in cp_svc.items()}
    M["cp_svc_in"]  = {cp: dict(d) for cp, d in cp_svc_in.items()}
    M["cp_svc_out"] = {cp: dict(d) for cp, d in cp_svc_out.items()}

    # ── গ: Behavioral flags ──
    # structuring: threshold-এর ঠিক নিচে বারবার (9k–10k, 19k–20k, 49k–50k, 99k–100k)
    bands = [(9000, 10000), (19000, 20000), (49000, 50000), (99000, 100000)]
    M["structuring"] = sorted(
        [s for s in txns if any(lo <= s["amount"] < hi for lo, hi in bands)],
        key=lambda s: -s["amount"])
    # rapid in-out: IN তারপর ≤60min-এ OUT, পরিমাণ ±25%
    # প্রতিটি কেসে উৎস (কোথা থেকে), কোন মাধ্যমে/একাউন্টে ঢুকল, কোন মাধ্যম/একাউন্ট থেকে বের হলো,
    # এবং গন্তব্য (কোথায় গেল) — সব রাখা হয়।
    seq = sorted([(_im_parse_dt(s.get("date", "")), s) for s in txns
                  if _im_parse_dt(s.get("date", ""))], key=lambda x: x[0])
    rapid = []
    for i in range(len(seq) - 1):
        d1, s1 = seq[i]
        if _im_eff_dir(s1) != "IN" or s1["amount"] <= 0: continue
        for j in range(i + 1, min(i + 6, len(seq))):
            d2, s2 = seq[j]
            if (d2 - d1).total_seconds() > 3600: break
            if _im_eff_dir(s2) == "OUT" and abs(s2["amount"] - s1["amount"]) / s1["amount"] <= 0.25:
                a1 = s1.get("acct", {}) or {}; a2 = s2.get("acct", {}) or {}
                rec = {
                    "in_dt": d1, "in_amt": s1["amount"],
                    "in_from": _im_party(s1, "IN"),
                    "in_svc": a1.get("service", "") or str(s1.get("sender", "")),
                    "in_acct": _im_acct_lbl(a1),
                    "out_dt": d2, "out_amt": s2["amount"],
                    "out_svc": a2.get("service", "") or str(s2.get("sender", "")),
                    "out_acct": _im_acct_lbl(a2),
                    "out_to": _im_party(s2, "OUT"),
                    "cp": a2.get("other_number", ""),
                }
                # একই টাকা নিজের দুই একাউন্টের মধ্যে (সমান অঙ্ক, ভিন্ন একাউন্ট/সার্ভিস,
                # কোনো বাইরের নম্বর নেই) → own-account transfer হওয়ার সম্ভাবনা
                _same_amt = abs(s1["amount"] - s2["amount"]) < 1
                _diff_acct = (rec["in_svc"], rec["in_acct"]) != (rec["out_svc"], rec["out_acct"])
                _no_ext = not a1.get("other_number") and not a2.get("other_number")
                _xref = bool(re.search(re.escape(str(rec["out_svc"]).split()[0]), s1.get("body", ""), re.I)) \
                        if rec["out_svc"] else False
                if _same_amt and _diff_acct and _no_ext and ((d2 - d1).total_seconds() <= 1800 or _xref):
                    rec["note"] = "own"
                    # নিজের একাউন্টের মধ্যে — গন্তব্য = যে একাউন্টে ঢুকেছে
                    _own_in = f'{rec["in_svc"]} {rec["in_acct"]}'.strip()
                    rec["out_to"] = f'Own {_own_in} (same money)'
                else:
                    rec["note"] = "pass"
                rapid.append(rec); break
    M["rapid"] = rapid
    # unusual volume: median-এর 2x-এর বেশি মোট ভলিউম যেসব counterparty-র
    allcp = defaultdict(float)
    for cp, d in list(recv.items()) + list(sent.items()): allcp[cp] += d["amt"]
    if allcp:
        vals = sorted(allcp.values()); med = vals[len(vals) // 2]
        M["high_vol"] = [(cp, v, freq.get(cp, 0)) for cp, v in
                         sorted(allcp.items(), key=lambda x: -x[1])
                         if v > max(med * 2, 50000)][:8]
    else:
        M["high_vol"] = []
    return M

def build_intel_html(M):
    """Intelligence metrics → HTML সেকশন (dark theme, existing style অনুসরণ করে)।"""
    MONTH_SHORT = {"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"May","06":"Jun",
                   "07":"Jul","08":"Aug","09":"Sep","10":"Oct","11":"Nov","12":"Dec"}
    def mlabel(m):
        try: y, mo = m.split("-"); return f"{MONTH_SHORT.get(mo, mo)} {y[2:]}"
        except: return m
    H = ['<div class="im-wrap">']

    # ── ক: Timeline ──
    H.append('<div class="im-sec"><div class="im-head">📅 Activity Timeline & Time Pattern</div>')
    if M["first_dt"] and M["last_dt"]:
        H.append('<div class="im-period">'
                 f'<div class="im-pd"><span>First Transaction</span><b>{M["first_dt"].strftime("%d %b %Y, %I:%M %p")}</b></div>'
                 f'<div class="im-pd"><span>Last Transaction</span><b>{M["last_dt"].strftime("%d %b %Y, %I:%M %p")}</b></div>'
                 f'<div class="im-pd"><span>Active Span</span><b>{M["active_days"]} days</b></div>'
                 '</div>')
    # monthly credit/debit bars
    mon = M["monthly"]
    if mon:
        mx = max((max(d["in"], d["out"]) for d in mon.values()), default=1) or 1
        H.append('<div class="im-mtitle">Monthly Credit (IN) vs Debit (OUT)</div><div class="im-months">')
        for m, d in mon.items():
            ih = d["in"] / mx * 100; oh = d["out"] / mx * 100
            H.append(
                '<div class="im-mcol">'
                f'<div class="im-bars"><div class="im-bin" style="height:{ih:.0f}%" title="IN ৳{d["in"]:,.0f}"></div>'
                f'<div class="im-bout" style="height:{oh:.0f}%" title="OUT ৳{d["out"]:,.0f}"></div></div>'
                f'<div class="im-mlbl">{mlabel(m)}</div><div class="im-mcnt">{d["cnt"]}</div>'
                '</div>')
        H.append('</div><div class="im-leg"><span class="im-dot-in"></span>Credit (IN)'
                 '<span class="im-dot-out"></span>Debit (OUT)<span class="im-leg-n">number = txn count</span></div>')
        H.append('<table class="im-tbl im-mtbl"><tr><th>Month</th><th>Credit IN (৳)</th><th>Debit OUT (৳)</th>'
                 '<th>Net (৳)</th><th>Txns</th></tr>')
        for m, d in mon.items():
            _n = d["in"] - d["out"]
            H.append(f'<tr><td>{mlabel(m)}</td><td style="color:#16a34a;font-weight:700">৳{d["in"]:,.0f}</td>'
                     f'<td style="color:#dc2626;font-weight:700">৳{d["out"]:,.0f}</td>'
                     f'<td style="font-weight:700">{"+" if _n >= 0 else "−"}৳{abs(_n):,.0f}</td><td>{d["cnt"]}</td></tr>')
        H.append('</table>')
    # day-part
    b = M["buckets"]; tb = sum(b.values()) or 1
    H.append('<div class="im-mtitle">Time-of-Day Activity</div><div class="im-buckets">')
    for lbl, c in b.items():
        hot = 'im-bk-night' if lbl.startswith('Night') and c > 0 else ''
        H.append(f'<div class="im-bk {hot}"><div class="im-bk-n">{c}</div>'
                 f'<div class="im-bk-l">{lbl}</div><div class="im-bk-p">{c/tb*100:.0f}%</div></div>')
    H.append('</div>')
    if M["night_cnt"]:
        H.append(f'<div class="im-flag im-flag-amber">🌙 {M["night_cnt"]} late-night transaction(s) '
                 f'(12AM–6AM) totaling ৳{M["night_total"]:,.0f} — worth review.</div>')
    H.append('</div>')  # end timeline sec

    # ── খ: Counterparty ──
    _CPS = M.get("cp_svc", {}) or {}
    def _svc_badge(n):
        sv = _CPS.get(n, "")
        if not sv: return '<span class="im-svc im-svc-x">—</span>'
        out = []
        for one in sv.split(", "):
            ol = one.lower()
            cls = ("bk" if "bkash" in ol else "ng" if "nagad" in ol else
                   "rk" if ("rocket" in ol or "dbbl" in ol or "dutch" in ol) else
                   "up" if "upay" in ol else "bn")
            out.append(f'<span class="im-svc im-svc-{cls}">{hl.escape(one)}</span>')
        return ' '.join(out)
    def _svc_cls(one):
        ol = one.lower()
        return ("bk" if "bkash" in ol else "ng" if "nagad" in ol else
                "rk" if ("rocket" in ol or "dbbl" in ol or "dutch" in ol) else
                "up" if "upay" in ol else "bn")
    def _svc_amt_badges(n, amap):
        """একাধিক মাধ্যম থাকলে প্রতিটির অঙ্ক — যেমন "bKash ৳80,000 · Nagad ৳44,430"।"""
        d = (amap or {}).get(n, {})
        d = {k: v for k, v in d.items() if v > 0}
        if len(d) <= 1:
            return _svc_badge(n) if not d else \
                   f'<span class="im-svc im-svc-{_svc_cls(next(iter(d)))}">{hl.escape(next(iter(d)))}</span>'
        return '<br>'.join(
            f'<span class="im-svc im-svc-{_svc_cls(k)}">{hl.escape(k)}</span>'
            f'<span class="im-svc-amt">৳{v:,.0f}</span>'
            for k, v in sorted(d.items(), key=lambda x: -x[1]))
    def _svc_plain(svc):
        return f'<span class="im-svc im-svc-{_svc_cls(svc)}">{hl.escape(svc)}</span>' if svc else '—'
    H.append('<div class="im-sec"><div class="im-head">🔗 Top Counterparties (Network)</div>')
    def cp_table(title, rows, color, amap):
        if not rows:
            return (f'<div class="im-ctitle">{title}</div>'
                    '<div class="im-noflag">No transfers to/from individual mobile numbers found in the SMS '
                    '(bill & merchant payments are excluded). If money left an account without SMS, see '
                    '⚖️ Balance Reconciliation above.</div>')
        t = [f'<div class="im-ctitle">{title}</div><table class="im-tbl">'
             '<tr><th>#</th><th>Number</th><th>Via (Service) — breakdown</th><th>Total (৳)</th><th>Txns</th></tr>']
        for i, (n, d) in enumerate(rows, 1):
            t.append(f'<tr><td>{i}</td><td class="im-num">{n}</td><td>{_svc_amt_badges(n, amap)}</td>'
                     f'<td style="color:{color};font-weight:700">৳{d["amt"]:,.0f}</td><td>{d["cnt"]}</td></tr>')
        t.append('</table>')
        return ''.join(t)
    H.append('<div class="im-cwrap">')
    H.append(cp_table("💰 Top Received-From", M["top_recv"], "#16a34a", M.get("cp_svc_in", {})))
    H.append(cp_table("💸 Top Sent-To", M["top_sent"], "#dc2626", M.get("cp_svc_out", {})))
    H.append('</div>')
    if M["high_freq"]:
        H.append('<div class="im-ctitle">🔁 High-Frequency Contacts (≥3 txns)</div><div class="im-chips">')
        for n, c in M["high_freq"]:
            H.append(f'<div class="im-chip"><b>{n}</b>{_svc_badge(n)}<span>{c}×</span></div>')
        H.append('</div>')
    H.append('</div>')  # end counterparty sec

    # ── গ: Behavioral flags ──
    has_flags = M["structuring"] or M["rapid"] or M["high_vol"]
    H.append('<div class="im-sec"><div class="im-head">🚩 Behavioral Flags</div>')
    if not has_flags:
        H.append('<div class="im-noflag">✓ No structuring, rapid in-out, or unusual-volume patterns detected.</div>')
    # structuring
    if M["structuring"]:
        H.append(f'<div class="im-flag im-flag-red">⚠ <b>Structuring suspicion:</b> {len(M["structuring"])} '
                 'transaction(s) just below round thresholds (e.g. 9,000–10,000 / 49,000–50,000) — '
                 'a pattern associated with reporting-limit avoidance.</div>')
        H.append('<table class="im-tbl im-tbl-flag"><tr><th>Date</th><th>Amount (৳)</th><th>Direction</th></tr>')
        for s in M["structuring"][:12]:
            eff = _im_eff_dir(s); dlbl = 'IN' if eff == 'IN' else ('OUT' if eff == 'OUT' else '—')
            H.append(f'<tr><td>{str(s.get("date",""))[:16]}</td>'
                     f'<td style="font-weight:700">৳{s["amount"]:,.0f}</td><td>{dlbl}</td></tr>')
        H.append('</table>')
    # rapid in-out
    if M["rapid"]:
        _n_own = sum(1 for r in M["rapid"] if r.get("note") == "own")
        _n_pass = len(M["rapid"]) - _n_own
        H.append(f'<div class="im-flag im-flag-amber">⚡ <b>Rapid in-out:</b> {len(M["rapid"])} case(s) where '
                 'money was received then sent out within 60 minutes at a similar amount. '
                 f'<b>{_n_pass}</b> look like pass-through / layering; <b>{_n_own}</b> look like the same money '
                 'moving between the subject\'s own accounts (shown on both sides) — verify before concluding.</div>')
        H.append('<div class="im-scroll"><table class="im-tbl im-tbl-flag im-rapid">'
                 '<tr><th>IN time</th><th>IN ৳</th><th>Source (from)</th><th>Received into</th>'
                 '<th>OUT time</th><th>OUT ৳</th><th>Paid out from</th><th>Destination (to)</th><th>Note</th></tr>')
        for r in M["rapid"][:12]:
            gap = int((r["out_dt"] - r["in_dt"]).total_seconds() // 60)
            _into = _svc_plain(r["in_svc"]) + (f'<div class="im-acct">{hl.escape(r["in_acct"])}</div>' if r["in_acct"] else '')
            _outf = _svc_plain(r["out_svc"]) + (f'<div class="im-acct">{hl.escape(r["out_acct"])}</div>' if r["out_acct"] else '')
            _note = ('<span class="im-note-own">🔁 Own-account transfer (likely)</span>' if r["note"] == "own"
                     else '<span class="im-note-pass">⚡ Pass-through</span>')
            H.append(f'<tr><td>{r["in_dt"].strftime("%d %b %y %H:%M")}</td>'
                     f'<td style="color:#16a34a;font-weight:700">৳{r["in_amt"]:,.0f}</td>'
                     f'<td>{hl.escape(r["in_from"])}</td><td>{_into}</td>'
                     f'<td>{r["out_dt"].strftime("%H:%M")} <span class="im-gap">(+{gap}m)</span></td>'
                     f'<td style="color:#dc2626;font-weight:700">৳{r["out_amt"]:,.0f}</td>'
                     f'<td>{_outf}</td><td>{hl.escape(r["out_to"])}</td><td>{_note}</td></tr>')
        H.append('</table></div>')
    # high volume
    if M["high_vol"]:
        H.append('<div class="im-ctitle">📊 Unusually High-Volume Counterparties</div>'
                 '<table class="im-tbl"><tr><th>#</th><th>Number</th><th>Via (Service)</th><th>Total Volume (৳)</th><th>Txns</th></tr>')
        for i, (cp, v, c) in enumerate(M["high_vol"], 1):
            _comb = {}
            for _m in (M.get("cp_svc_in", {}), M.get("cp_svc_out", {})):
                for k, v2 in (_m.get(cp, {}) or {}).items(): _comb[k] = _comb.get(k, 0) + v2
            H.append(f'<tr><td>{i}</td><td class="im-num">{cp}</td><td>{_svc_amt_badges(cp, {cp: _comb})}</td>'
                     f'<td style="font-weight:700">৳{v:,.0f}</td><td>{c}</td></tr>')
        H.append('</table>')
    H.append('</div>')  # end flags sec

    H.append('</div>')  # end im-wrap
    return ''.join(H)

IM_CSS = """
.im-wrap{margin:18px 0}
.im-bottom{max-width:1400px;margin:28px auto 40px;padding:0 20px;border-top:3px solid #e2e8f0;padding-top:18px}
.im-bottom-title{font-size:17px;font-weight:800;color:#0f172a;margin-bottom:6px}
.im-bottom-hint{font-size:11.5px;color:#475569;background:#f1f5f9;border:1px dashed #cbd5e1;border-radius:8px;padding:6px 10px;margin:4px 0 6px}
.deleted{display:none!important}
.dz-target{position:relative}
.dz-del{background:#fff;border:1px solid #fecaca;color:#dc2626;border-radius:6px;font-size:11px;line-height:1;padding:3px 5px;cursor:pointer;opacity:.55;transition:opacity .15s}
.dz-del:hover{opacity:1;background:#fef2f2}
.im-head > .dz-del,.im-ctitle > .dz-del{float:right;margin-left:8px}
.pc-item > .dz-del,.im-flag > .dz-del,.im-noflag > .dz-del,.im-chip > .dz-del,.im-pd > .dz-del,.im-bk > .dz-del{position:absolute;top:4px;right:4px}
.dz-td,.dz-th{width:28px;text-align:center;padding:2px!important;background:transparent!important}
.im-note{font-size:11.5px;color:#64748b;margin:0 0 10px;line-height:1.5}
.pc-cat{margin-bottom:12px}
.pc-n{color:#94a3b8;font-weight:600}
.pc-item{border:1px solid #e8eaed;border-left:3px solid #6366f1;border-radius:8px;padding:8px 34px 8px 10px;margin:6px 0;background:#fafbff}
.pc-top{display:flex;flex-wrap:wrap;gap:6px 10px;align-items:baseline}
.pc-lbl{font-weight:800;color:#1e293b;font-size:12.5px}
.pc-meta{font-size:11px;color:#64748b;font-family:monospace}
.pc-det{font-size:11.5px;color:#334155;margin-top:3px}
.pc-src{font-size:11px;color:#475569;background:#f1f5f9;border-radius:6px;padding:5px 7px;margin-top:5px;white-space:pre-wrap;word-break:break-word}
.pc-snd{font-weight:700;color:#334155}
.rc-tbl td{vertical-align:top}
.rc-body{font-size:10.5px;color:#475569;max-width:420px;white-space:pre-wrap;word-break:break-word}
.rc-ok{color:#047857;font-weight:700;white-space:nowrap}
.rc-bad{color:#dc2626;font-weight:700;white-space:nowrap}
.rc-warn{color:#b45309;font-weight:700;white-space:nowrap}
.im-mtbl{margin-top:10px}
@media print{.dz-del,.dz-td,.dz-th,.im-bottom-hint{display:none!important}.deleted{display:none!important}.pc-item{padding-right:10px}}
.im-svc{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:5px;margin:1px 2px 1px 0;white-space:nowrap;border:1px solid}
.im-svc-bk{background:#fdf2f8;color:#be185d;border-color:#f9a8d4}
.im-svc-ng{background:#fff7ed;color:#c2410c;border-color:#fdba74}
.im-svc-rk{background:#f5f3ff;color:#6d28d9;border-color:#c4b5fd}
.im-svc-up{background:#ecfeff;color:#0e7490;border-color:#67e8f9}
.im-svc-bn{background:#eff6ff;color:#1d4ed8;border-color:#93c5fd}
.im-svc-x{background:#f8fafc;color:#94a3b8;border-color:#e2e8f0}
.im-svc-amt{font-family:monospace;font-size:11px;font-weight:700;color:#334155;margin-left:3px}
.im-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
.im-rapid td{vertical-align:top;font-size:11.5px}
.im-acct{font-family:monospace;font-size:10.5px;color:#475569;margin-top:2px}
.im-gap{color:#94a3b8;font-size:10.5px}
.im-note-own{background:#ecfdf5;color:#047857;border:1px solid #6ee7b7;border-radius:5px;padding:1px 6px;font-size:10.5px;font-weight:700;white-space:nowrap}
.im-note-pass{background:#fff7ed;color:#c2410c;border:1px solid #fdba74;border-radius:5px;padding:1px 6px;font-size:10.5px;font-weight:700;white-space:nowrap}
.im-sec{background:#fff;border:1px solid #e8eaed;border-radius:12px;padding:16px 18px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.im-head{font-size:15px;font-weight:800;color:#1e293b;margin-bottom:14px;padding-bottom:10px;border-bottom:2px solid #f1f5f9}
.im-period{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px}
.im-pd{background:#f8fafc;border:1px solid #e8eaed;border-radius:9px;padding:10px 12px;text-align:center}
.im-pd span{display:block;font-size:11px;color:#64748b;margin-bottom:3px}
.im-pd b{font-size:13.5px;color:#0f172a}
.im-mtitle{font-size:12.5px;font-weight:700;color:#475569;margin:14px 0 8px}
.im-months{display:flex;gap:6px;align-items:flex-end;overflow-x:auto;padding:6px 0 2px;min-height:130px}
.im-mcol{display:flex;flex-direction:column;align-items:center;gap:3px;min-width:44px}
.im-bars{display:flex;gap:2px;align-items:flex-end;height:100px}
.im-bin{width:10px;background:linear-gradient(180deg,#22c55e,#16a34a);border-radius:3px 3px 0 0;min-height:2px}
.im-bout{width:10px;background:linear-gradient(180deg,#ef4444,#dc2626);border-radius:3px 3px 0 0;min-height:2px}
.im-mlbl{font-size:10px;color:#475569;white-space:nowrap}
.im-mcnt{font-size:10px;color:#94a3b8;font-weight:700}
.im-leg{display:flex;align-items:center;gap:6px;font-size:11px;color:#64748b;margin-top:8px;flex-wrap:wrap}
.im-dot-in,.im-dot-out{width:10px;height:10px;border-radius:3px;display:inline-block}
.im-dot-in{background:#16a34a}.im-dot-out{background:#dc2626;margin-left:10px}
.im-leg-n{margin-left:auto;font-style:italic}
.im-buckets{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.im-bk{background:#f8fafc;border:1px solid #e8eaed;border-radius:9px;padding:11px;text-align:center}
.im-bk-night{background:rgba(220,38,38,.07);border-color:rgba(220,38,38,.3)}
.im-bk-n{font-size:20px;font-weight:800;color:#1e293b}
.im-bk-l{font-size:10px;color:#64748b;margin:3px 0}
.im-bk-p{font-size:11px;color:#94a3b8;font-weight:700}
.im-cwrap{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.im-ctitle{font-size:12.5px;font-weight:700;color:#475569;margin:12px 0 6px}
.im-tbl{width:100%;border-collapse:collapse;font-size:12px}
.im-tbl th{background:#f1f5f9;color:#475569;text-align:left;padding:6px 8px;font-size:10.5px;text-transform:uppercase;letter-spacing:.3px}
.im-tbl td{padding:6px 8px;border-bottom:1px solid #f1f5f9;color:#334155}
.im-num{font-family:monospace;font-size:11.5px}
.im-chips{display:flex;flex-wrap:wrap;gap:8px}
.im-chip{background:#eff6ff;border:1px solid #bfdbfe;border-radius:20px;padding:5px 12px;font-size:12px;color:#1e40af}
.im-chip b{font-family:monospace}.im-chip span{color:#3b82f6;margin-left:5px;font-weight:700}
.im-flag{border-radius:9px;padding:10px 13px;font-size:12.5px;line-height:1.5;margin:10px 0}
.im-flag-red{background:rgba(220,38,38,.08);border:1px solid rgba(220,38,38,.35);color:#991b1b}
.im-flag-amber{background:rgba(217,119,6,.09);border:1px solid rgba(217,119,6,.35);color:#92400e}
.im-tbl-flag{margin-top:4px}
.im-noflag{background:rgba(22,163,74,.08);border:1px solid rgba(22,163,74,.3);border-radius:9px;padding:11px 13px;font-size:12.5px;color:#15803d}
@media (max-width:640px){.im-period,.im-buckets{grid-template-columns:1fr 1fr}.im-cwrap{grid-template-columns:1fr}}
@media print{.im-sec{break-inside:avoid}}
"""


def generate_html(sms_list, stats, out_path, tn=""):
    now = datetime.now().strftime("%d %B %Y, %I:%M %p")

    # effective direction (IN/OUT) — body-detected credited/debited সবচেয়ে নির্ভরযোগ্য
    def _eff_dir_of(s):
        ai = s.get("acct", {}) or {}
        _st = s.get("sender_type", ""); _ad = ai.get("direction", ""); _dir = s.get("direction", "")
        # mfs/bank: শুধু body-detected direction নির্ভরযোগ্য। SMSMT/SMSMO সব notification-এই
        # আসে, তাই তা দিয়ে IN/OUT ঠিক করা যাবে না — direction খালি হলে গোনা হবে না।
        if _st in ("mfs","bank"):
            if _ad == "Received": return "IN"
            if _ad == "Sent":     return "OUT"
            return ""
        if _dir == "SMSMT": return "IN"
        if _dir == "SMSMO": return "OUT"
        if _ad == "Received": return "IN"
        if _ad == "Sent": return "OUT"
        return ""

    # ── Credit/Debit টোটাল: সব BDT লেনদেন থেকে (10K-এর নিচেরগুলোসহ) ──────────────
    # চার্ট ও সাইডবার সব transaction নিয়ে হিসাব করে; শুধু card-এ দেখানোর ক্ষেত্রে threshold।
    # পাশাপাশি 10K+ card-worthy txn গুনি — display floor নির্ধারণে কাজে লাগে।
    _credit_total = 0.0; _debit_total = 0.0
    _credit_cnt = 0; _debit_cnt = 0
    _above10k = 0
    for s in sms_list:
        amt = s.get("amount", 0) or 0
        if amt <= 0: continue
        if "travel_txn" in s.get("categories", []): continue   # বিদেশি মুদ্রা → BDT চার্টে নয়
        eff = _eff_dir_of(s)
        if eff == "IN":  _credit_total += amt; _credit_cnt += 1
        elif eff == "OUT": _debit_total += amt; _debit_cnt += 1
        if amt >= TIER_SMALL and s.get("sender_type") in ("mfs","bank"):
            _above10k += 1
    _cd_grand = _credit_total + _debit_total
    _cr_pct = (_credit_total / _cd_grand * 100) if _cd_grand else 0
    _db_pct = (_debit_total / _cd_grand * 100) if _cd_grand else 0
    _net_flow = _credit_total - _debit_total

    # ── Dynamic display floor (message/card দেখানোর threshold) ──
    # 10K+ লেনদেন 5টির বেশি থাকলে শুধু 10K+ card; নইলে 5K+ card-ও দেখাই।
    _disp_floor = TIER_SMALL if _above10k > 5 else 5000

    # display_list: card-worthy txn (>= _disp_floor) + non-txn categories (personal, travel…)
    # foreign_bank SMS → সর্বোচ্চ 20টি card, বাকিগুলো summary block-এ
    def _show_in_cards(s):
        cats = s.get("categories", [])
        if "travel_txn" in cats: return False
        if "bank_info" in cats:                       # 10K-এর নিচের mfs/bank txn
            return (s.get("amount", 0) or 0) >= _disp_floor
        if cats == [f"txn_{s.get('tier','')}"] and s.get("tier","") == "":
            return False
        return True
    _non_ftxn = [s for s in sms_list if _show_in_cards(s)]
    _ftxn_all  = [s for s in sms_list if "travel_txn" in s.get("categories",[])]
    _ftxn_show = _ftxn_all[:20]
    _ftxn_hidden = _ftxn_all[20:]

    # Foreign txn summary stats
    _ftxn_by_ccy = {}
    for s in _ftxn_all:
        ccy = s.get("currency","MYR") or "MYR"
        amt = s.get("amount",0) or 0
        if ccy not in _ftxn_by_ccy:
            _ftxn_by_ccy[ccy] = {"count":0,"total":0.0,"max":0.0}
        _ftxn_by_ccy[ccy]["count"] += 1
        _ftxn_by_ccy[ccy]["total"] += amt
        if amt > _ftxn_by_ccy[ccy]["max"]:
            _ftxn_by_ccy[ccy]["max"] = amt

    display_list = _non_ftxn + _ftxn_show

    # Intelligence metrics: timeline, counterparty, behavioral flags।
    # পূর্ণ BDT লেনদেন সেটে চালাই (ছোট bank_info txn সহ, বিদেশি মুদ্রা বাদে) —
    # structuring/night/rapid বিশ্লেষণে সব লেনদেন লাগে, শুধু কার্ড-দেখানো গুলো নয়।
    _intel_src = [s for s in sms_list
                  if (s.get("amount", 0) or 0) > 0
                  and "travel_txn" not in s.get("categories", [])]
    _own_num = (stats or {}).get("own_number", "") or (("0" + re.sub(r'^\+?880', '', re.sub(r'\D', '', tn))[-10:]) if tn else "")
    _intel = compute_intel_metrics(_intel_src, own=_own_num)
    _intel_html = build_intel_html(_intel)
    _profile_html = build_profile_html((stats or {}).get("profile") or {})
    _recon_html = build_recon_html((stats or {}).get("recon") or [])

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
              "txn_small":"Small Transaction (10K–50K)","travel_txn":"💳 Foreign Card Transaction",
              "identity":"Identity / Profile Info",
              "love_family":"Personal / Family","travel":"Travel / Visa / Umrah",
              "courier":"Courier / Delivery","chalan":"Govt. Challan / Treasury / Court",
              "education":"Education / Admission / Payment",
              "bo_account":"BO Account / Stock Market",
              "shopping":"Shopping / Billing",
              "service":"Service Messages","personal":"Personal Conversation"}
    def esc(t): return hl.escape(str(t))
    def cat_tags(cats):
        o=""
        for c in cats:
            if c == "bank_info": continue   # internal marker — tag হিসেবে দেখাব না
            m=CAT_META.get(c,{}); col=m.get("color","#555")
            lbl=CAT_EN.get(c,m.get("label",c))
            o+=(f'<span class="ctag" style="background:{col}12;border:1px solid {col}35;color:{col}">'
                f'{m.get("icon","🔍")} {lbl}</span>')
        return o
    def hl_body(body, amt, stype_="", ccy=""):
        r=esc(body)
        if stype_ == "foreign_bank" and amt > 0 and ccy:
            for v in [f'{ccy}{amt:,.2f}', f'{ccy} {amt:,.2f}', f'{ccy}{amt:.2f}']:
                try: r=re.sub(re.escape(v),f'<mark>{v}</mark>',r,count=1,flags=re.I)
                except (ValueError, TypeError): pass
            for v in [f'{amt:,.2f}', f'{amt:.2f}']:
                try: r=re.sub(re.escape(v),f'<mark>{v}</mark>',r,count=1,flags=re.I)
                except (ValueError, TypeError): pass
        elif amt>=_disp_floor:
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
        _s_type = s.get("sender_type","")
        _ccy    = s.get("currency","")
        if _s_type == "foreign_bank" and s["amount"] > 0:
            amt_tag = f'<span class="amt-tag" style="background:#0f766e">💳 {_ccy} {s["amount"]:,.2f}</span>'
        else:
            amt_tag=(f'<span class="amt-tag" style="background:{bdr if t else "#64748b"}">💰 Tk {s["amount"]:,.2f}</span>'
                     if s["amount"]>=_disp_floor else "")
        tier_badge=(f'<span class="rbadge" style="background:{bdr}">{TIER_EN_LBL.get(t,"")}</span>' if t else "")
        # MFS/Bank SMS: body-detected direction সবচেয়ে নির্ভরযোগ্য (SMSMT সব notification-এই আসে)
        _acct_dir = ai.get("direction","") if ai else ""
        if _s_type in ("mfs","bank") and _acct_dir:
            _eff_dir = "IN" if _acct_dir=="Received" else "OUT"
        elif direction=="SMSMT":  _eff_dir = "IN"
        elif direction=="SMSMO":  _eff_dir = "OUT"
        elif _acct_dir=="Received": _eff_dir = "IN"
        elif _acct_dir=="Sent":     _eff_dir = "OUT"
        else: _eff_dir = ""
        dir_badge=""
        if _eff_dir=="IN":  dir_badge='<span class="dir-in">● IN</span>'
        elif _eff_dir=="OUT": dir_badge='<span class="dir-out">● OUT</span>'
        acct_mini=""
        if _s_type == "foreign_bank":
            pts=[]
            if ai and ai.get("service"): pts.append(f'<b>{ai["service"]}</b>')
            if _ccy and s["amount"] > 0:  pts.append(f'<b style="color:#0f766e">{_ccy} {s["amount"]:,.2f}</b>')
            if ai and ai.get("trx_id"):  pts.append(f'TrxID <code>{ai["trx_id"]}</code>')
            if pts: acct_mini=f'<div class="acct-mini">{" &nbsp;·&nbsp; ".join(pts)}</div>'
        elif ai and ai.get("service"):
            pts=[]
            if ai.get("direction"):    pts.append(f'<b>{ai["direction"]}</b>')
            if ai.get("other_number"): pts.append(f'📱 {ai["other_number"]}')
            if ai.get("trx_id"):       pts.append(f'TrxID <code>{ai["trx_id"]}</code>')
            if ai.get("balance"):      pts.append(f'Balance: Tk {ai["balance"]:,.2f}')
            if pts: acct_mini=f'<div class="acct-mini">{" &nbsp;·&nbsp; ".join(pts)}</div>'
        merged_tag=(f'<span class="merged-tag">🔗 {s["parts"]} parts</span>' if s.get("parts",1)>1 else "")
        cards+=f'''
        <div class="scard" id="card-{i}" data-tier="{t}" data-month="{s["month"]}" data-cats="{" ".join(s["categories"])}" data-dir="{direction}">
          <div class="sc-head">
            <div class="sc-left">
              <span class="snum">#{i}</span>
              <div>
                <div class="snm">{dir_badge} 📱 {esc(s["sender"])}</div>
                <div class="sdate">🕐 {esc(s["date"])} {merged_tag}</div>
              </div>
            </div>
            <div class="sc-right">{tier_badge}{amt_tag}
              <button class="del-btn no-print" onclick="delCard('card-{i}')" title="Delete this card">🗑</button>
            </div>
          </div>
          {acct_mini}
          <div class="sbody" style="background:{bg}">{hl_body(s["body"],s["amount"],_s_type,_ccy)}</div>
          <div class="sc-foot">{cat_tags(s["categories"])}</div>
        </div>'''
    if not cards: cards='<div class="no-result"><div style="font-size:48px">✅</div><div>No SMS records found</div></div>'

    # ── Foreign Transaction Summary block ──────────────────────────────────
    ftxn_summary_html = ""
    if _ftxn_all:
        rows_html = ""
        for ccy, d in _ftxn_by_ccy.items():
            rows_html += (f'<div class="fts-row">'
                          f'<span class="fts-ccy">{ccy}</span>'
                          f'<span class="fts-val">{d["count"]} transactions</span>'
                          f'<span class="fts-val" style="color:#0f766e;font-weight:700">'
                          f'Total: {ccy} {d["total"]:,.2f}</span>'
                          f'<span class="fts-val" style="color:#475569">'
                          f'Max: {ccy} {d["max"]:,.2f}</span>'
                          f'</div>')
        hidden_note = (f'<div class="fts-note">Showing {len(_ftxn_show)} of {len(_ftxn_all)} transactions as cards. '
                       f'{len(_ftxn_hidden)} hidden (summary only).</div>'
                       if _ftxn_hidden else
                       f'<div class="fts-note">All {len(_ftxn_all)} transactions shown as cards.</div>')
        ftxn_summary_html = f'''
        <div class="ftxn-summary" id="ftxn-summary">
          <div class="fts-header">💳 Foreign Card Transaction Summary</div>
          {rows_html}
          {hidden_note}
        </div>'''
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
        if d["txn_count"] == 0 and d.get("display_balance",0) == 0 and d["last_balance"] == 0 and not d.get("seen"): continue
        display_bal = d.get("display_balance", d["last_balance"])
        is_multi    = d.get("multi_account", False)
        accs        = d.get("accounts", {})

        # একাধিক account/card থাকলে — প্রতিটি আলাদা (card vs account), txn ও in/out সহ।
        # একটিমাত্র হলেও যদি কার্ড হয়, breakdown দেখাই (কার্ড আলাদা করে বোঝা জরুরি)।
        _show_breakdown = (len(accs) > 1) or any(a.get("type") == "card" for a in accs.values())
        if _show_breakdown and accs:
            acc_rows = ""
            for _canon, av in sorted(accs.items(), key=lambda x: -x[1]["txn_count"]):
                _is_card = av.get("type") == "card"
                _icon = "💳" if _is_card else "🏦"
                _tlbl = "Card" if _is_card else "A/C"
                _disp = esc(av.get("display", _canon))
                _bits = f'<span class="acc-cnt">{av["txn_count"]} txn</span>'
                if av.get("in", 0) > 0:  _bits += f'<span class="acc-in">IN ৳{av["in"]:,.0f}</span>'
                if av.get("out", 0) > 0: _bits += f'<span class="acc-out">OUT ৳{av["out"]:,.0f}</span>'
                _balp = f'<span class="acc-bal">Bal ৳{av["balance"]:,.0f}</span>' if av.get("balance", 0) > 0 else ""
                acc_rows += (f'<div class="acc-row2">'
                             f'<div class="acc-id"><span class="acc-ic {"card" if _is_card else "reg"}">{_icon} {_tlbl}</span> '
                             f'<code>{_disp}</code></div>'
                             f'<div class="acc-meta">{_bits}{_balp}</div>'
                             f'</div>')
            bal_html = (f'<div class="ag-accts">'
                        f'<div class="ag-accts-hdr">🗂️ Accounts / Cards ({len(accs)})</div>'
                        f'{acc_rows}</div>')
        else:
            bal_html = (f'<div class="ag-bal">💰 Last Balance: Tk {display_bal:,.2f}</div>'
                        if display_bal else "")

        bank_sb+=f'''<div class="acct-card" style="border-left:3px solid {col}">
          <div class="acct-svc" style="color:{col}">💳 {esc(svc)}</div>
          <div class="acct-grid4">
            <div class="ag-c"><div class="ag-v">{d["txn_count"]}</div><div class="ag-l">Transactions</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#16a34a;font-size:11px">৳{d["total_in"]:,.0f}</div><div class="ag-l">Received</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#dc2626;font-size:11px">৳{d["total_out"]:,.0f}</div><div class="ag-l">Sent</div></div>
            <div class="ag-c"><div class="ag-v" style="color:#d97706;font-size:10px">৳{d["max_txn"]:,.0f}</div><div class="ag-l">Highest</div></div>
          </div>
          {f'<div class="ag-cycling">🔄 Investment/Internal: ৳{d["total_cycling"]:,.0f} <span class="cyc-note">(Govt Sec/FDR/Transfer — excluded from IN/OUT)</span></div>' if d.get("total_cycling",0) > 0 else ""}
          {bal_html}
          {nums}</div>'''
    if not bank_sb: bank_sb='<div style="color:#9ca3af;font-size:11px;padding:4px 0">No MFS/Bank data</div>'
    def fb(lbl,js,col=None,data=""):
        st=f' style="border-color:{col}"' if col else ""
        return f'<button class="fbtn"{st} onclick="{js}" {data}>{lbl}</button>'
    fbtns =f'<button class="fbtn active" data-all="1" onclick="fAll(this)">All ({len(display_list)})</button>'
    fbtns+=fb(f'🔴 High ({tier_cnt.get("high",0)})',"fTier(this,'high')","#991b1b","data-tier='high'")
    fbtns+=fb(f'🟠 Medium ({tier_cnt.get("med",0)})',"fTier(this,'med')","#dc2626","data-tier='med'")
    fbtns+=fb(f'🟡 Small ({tier_cnt.get("small",0)})',"fTier(this,'small')","#ea580c","data-tier='small'")
    if smsmt_count: pass
    if smsmo_count: pass
    for ck in ["identity","love_family","travel","travel_txn","courier","chalan","education","bo_account","shopping","personal"]:
        c=cat_cnt.get(ck,0)
        if not c: continue
        meta=CAT_META.get(ck,{}); lbl=CAT_EN.get(ck,meta.get("label",ck))
        fbtns+=fb(f'{meta.get("icon","")} {lbl} ({c})',f"fCat(this,'{ck}')",meta.get("color"),f"data-cat='{ck}'")
    mbtns="".join(
        f'<button class="mbtn" data-month="{m}" onclick="fMonth(this,\'{m}\')">'
        f'{MONTH_EN.get(m.split("-")[1],m)} {m.split("-")[0]} ({mnt_cnt[m]})</button>'
        for m in months)
    tnb=f'<div class="tnbadge">📱 {esc(tn)}</div>' if tn else ""

    _dz_js = '''
// ── নিচের Intelligence অংশ: প্রতিটি সেকশন/সারি/আইটেমে ডিলিট বাটন ──────────────
(function(){
  const root=document.querySelector('.im-bottom'); if(!root) return;
  function mkBtn(){const b=document.createElement('button');b.type='button';b.className='dz-del no-print';
    b.title='Delete';b.textContent='🗑';b.setAttribute('onclick','delEl(this)');return b;}
  root.querySelectorAll('.im-sec, .pc-cat').forEach(el=>{
    el.classList.add('dz-target');
    const h=el.querySelector(':scope > .im-head, :scope > .im-ctitle');
    if(h && !h.querySelector(':scope > .dz-del')) h.appendChild(mkBtn());
  });
  root.querySelectorAll('.pc-item, .im-flag, .im-noflag, .im-chip, .im-pd, .im-bk').forEach(el=>{
    el.classList.add('dz-target');
    if(!el.querySelector(':scope > .dz-del')) el.appendChild(mkBtn());
  });
  root.querySelectorAll('table.im-tbl').forEach(t=>{
    t.querySelectorAll('tr').forEach(tr=>{
      if(tr.querySelector(':scope > .dz-td, :scope > .dz-th')) return;
      if(tr.querySelector(':scope > th')){const th=document.createElement('th');th.className='dz-th no-print';tr.appendChild(th);}
      else{tr.classList.add('dz-target');const td=document.createElement('td');td.className='dz-td no-print';td.appendChild(mkBtn());tr.appendChild(td);}
    });
  });
})();
function delEl(btn){
  const el=btn.closest('.dz-target'); if(!el) return;
  el.classList.add('deleted'); _delStack.push(el);
  document.getElementById('del-cnt').textContent=_delStack.length;
  _undoBar.classList.add('show'); clearTimeout(_undoTimer);
  _undoTimer=setTimeout(()=>_undoBar.classList.remove('show'),4000);
}
'''
    _word_js = '''
// ── Word (docx.js) ──────────────────────────────────────────────────────────
function downloadWord(){
  const btn=document.getElementById('word-btn');
  btn.textContent='⏳ Generating...'; btn.disabled=true;
  if(!window.docx){
    const s=document.createElement('script');
    s.src='https://cdn.jsdelivr.net/npm/docx@8/build/index.js';
    s.onload=()=>_buildWord(btn);
    s.onerror=()=>{ btn.textContent='❌ No Internet'; btn.disabled=false; };
    document.head.appendChild(s);
  } else { _buildWord(btn); }
}

// ── Word: নিচের Intelligence অংশ (ডিলিট করা অংশ বাদ) ────────────────────────
function _bottomBlocks(){
  const {Table,TableRow,TableCell,Paragraph,TextRun,WidthType,ShadingType}=docx;
  const out=[]; const root=document.querySelector('.im-bottom'); if(!root) return out;
  const vis=el=>!el.closest('.deleted') && !el.classList.contains('no-print');
  const txt=el=>(el.innerText||el.textContent||'').replace(/🗑/g,'').replace(/[ \\t]+\\n/g,'\\n').replace(/\\n{2,}/g,'\\n').trim();
  const P=(t,o)=>new Paragraph({spacing:{after:60},children:t.split('\\n').map((l,i)=>new TextRun(Object.assign({text:l.trim(),size:18,font:'Arial',break:i?1:0},o||{})))});
  out.push(new Paragraph({spacing:{before:320,after:120},children:[new TextRun({text:'INTELLIGENCE ANALYSIS',bold:true,size:30,color:'1A1A2E',font:'Arial'})]}));
  function tbl(t){
    const rows=[...t.querySelectorAll('tr')].filter(vis).map(tr=>{
      const cells=[...tr.children].filter(c=>!c.classList.contains('no-print'));
      if(!cells.length) return null;
      const isH=cells[0].tagName==='TH';
      return new TableRow({children:cells.map(c=>new TableCell({
        shading:isH?{type:ShadingType.CLEAR,color:'auto',fill:'E2E8F0'}:undefined,
        children:[new Paragraph({children:txt(c).split('\\n').map((l,i)=>new TextRun({text:l.trim(),size:16,bold:isH,font:'Arial',break:i?1:0}))})]}))});
    }).filter(Boolean);
    if(rows.length>1) out.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows}));
    out.push(new Paragraph({spacing:{after:80}}));
  }
  const LEAF=['pc-item','im-flag','im-noflag','im-chip','im-pd','im-bk','im-note'];
  const SKIP=['im-months','im-leg','im-bars','im-bottom-hint'];
  function walk(el){
    if(!vis(el)) return;
    if(el.tagName==='TABLE'){tbl(el);return;}
    if(['SCRIPT','STYLE','BUTTON'].includes(el.tagName)) return;
    if(SKIP.some(c=>el.classList.contains(c))) return;
    if(el.classList.contains('im-head')){out.push(P(txt(el),{bold:true,size:24,color:'1A1A2E'}));return;}
    if(el.classList.contains('im-ctitle')||el.classList.contains('im-mtitle')){out.push(P(txt(el),{bold:true,size:20,color:'334155'}));return;}
    const blocks=[...el.children].filter(c=>!['BR','B','SPAN','CODE','I','STRONG','EM','A','SUP','SUB'].includes(c.tagName));
    if(!blocks.length || LEAF.some(c=>el.classList.contains(c))){const t=txt(el); if(t) out.push(P(t)); return;}
    blocks.forEach(walk);
  }
  [...root.children].filter(c=>!c.classList.contains('im-bottom-title')).forEach(walk);
  return out;
}

function _buildWord(btn){
  const {Document,Packer,Table,TableRow,TableCell,Paragraph,TextRun,
         WidthType,BorderStyle,AlignmentType,VerticalAlign,ShadingType}=docx;
  const cards=[...document.querySelectorAll('.scard')].filter(c=>
    !c.classList.contains('deleted') && c.style.display!=='none'
  );
  const tn_val=document.querySelector('.tnbadge')?document.querySelector('.tnbadge').textContent.replace('📱','').trim():'';
  const now_str=new Date().toLocaleString('en-BD',{dateStyle:'long',timeStyle:'short'});

  function _rgb(h){const n=parseInt(h,16);return {r:(n>>16)&255,g:(n>>8)&255,b:n&255};}
  function shd(fill){return {type:ShadingType.CLEAR,fill:'#'+fill};}
  const brd=(color)=>({
    top:{style:BorderStyle.SINGLE,size:4,color:color||'CCCCCC'},
    bottom:{style:BorderStyle.SINGLE,size:4,color:color||'CCCCCC'},
    left:{style:BorderStyle.SINGLE,size:4,color:color||'CCCCCC'},
    right:{style:BorderStyle.SINGLE,size:4,color:color||'CCCCCC'},
  });

  function cell(text,opts){
    opts=opts||{};
    const fill=opts.fill||'FFFFFF', bold=opts.bold||false;
    const color=opts.color||'111111', sz=(opts.size||10)*2;
    const align=opts.align||AlignmentType.LEFT, width=opts.width||2;
    return new TableCell({
      shading:shd(fill),
      width:{size:Math.round(width*568),type:WidthType.DXA},
      verticalAlign:VerticalAlign.TOP,
      borders:brd('CCCCCC'),
      children:[new Paragraph({
        alignment:align,
        children:[new TextRun({text:String(text||''),bold:bold,size:sz,color:color,font:'Arial'})],
      })],
    });
  }
  function hdrCell(text,width){
    return cell(text,{fill:'D6E4F0',bold:true,color:'1A1A2E',size:9,align:AlignmentType.CENTER,width:width});
  }
  function secHead(num,title){
    return new Paragraph({
      spacing:{before:240,after:80},
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'C0392B',space:1}},
      children:[
        new TextRun({text:num+'. ',bold:true,size:28,color:'C0392B',font:'Arial'}),
        new TextRun({text:title,bold:true,size:28,color:'1A1A2E',font:'Arial'}),
      ],
    });
  }

  // Bank summary
  const bankRows=[];
  document.querySelectorAll('.acct-card').forEach(ac=>{
    const svc=(ac.querySelector('.acct-svc')||{}).textContent||'';
    const nums=[...ac.querySelectorAll('.ag-v')].map(v=>v.textContent.trim());
    const bal=(ac.querySelector('.ag-bal')||{}).textContent||'';
    if(svc.trim()) bankRows.push({
      svc:svc.replace('💳','').trim(),
      txn:nums[0]||'0',recv:nums[1]||'0',sent:nums[2]||'0',max:nums[3]||'0',
      bal:bal.replace('Last Balance:','').trim()
    });
  });
  const bHdrs=['Bank / MFS Service','Txn Count','Received (IN)','Sent (OUT)','Max Txn','Last Balance'];
  const bW=[3.5,1.8,2.8,2.8,2.8,2.8];
  const bRows=[new TableRow({children:bHdrs.map((h,i)=>hdrCell(h,bW[i]))})];
  bankRows.forEach((r,i)=>{
    const bg=i%2===0?'FFFFFF':'F5F5F5';
    bRows.push(new TableRow({children:[
      cell(r.svc,{fill:bg,bold:true,color:'1A1A2E',size:9,width:bW[0]}),
      cell(r.txn,{fill:bg,color:'444444',size:9,align:AlignmentType.CENTER,width:bW[1]}),
      cell(r.recv,{fill:bg,color:'1A7A4A',size:9,align:AlignmentType.RIGHT,width:bW[2]}),
      cell(r.sent,{fill:bg,color:'CC2200',size:9,align:AlignmentType.RIGHT,width:bW[3]}),
      cell(r.max,{fill:bg,color:'7A5A00',size:9,align:AlignmentType.RIGHT,width:bW[4]}),
      cell(r.bal,{fill:bg,color:'1A4A8A',size:9,align:AlignmentType.RIGHT,width:bW[5]}),
    ]}));
  });

  // SMS rows
  const sHdrs=['SL','Sender','Date / Time','Type','Message Body','Remarks'];
  const sW=[0.8,2.5,2.5,1.2,8.5,2.5];
  const sRows=[new TableRow({children:sHdrs.map((h,i)=>hdrCell(h,sW[i]))})];
  cards.forEach((c,i)=>{
    const din=c.querySelector('.dir-in'), dout=c.querySelector('.dir-out');
    const mtype=din?'IN':dout?'OUT':'--';
    const sender=(c.querySelector('.snm')||{}).textContent||'';
    const dt=(c.querySelector('.sdate')||{}).textContent||'';
    const body=(c.querySelector('.sbody')||{}).textContent||'';
    const cats=[...c.querySelectorAll('.ctag')].map(t=>t.textContent.trim()).join(', ');
    const acct=((c.querySelector('.acct-mini')||{}).textContent||'').trim();
    const rem=[cats,acct].filter(Boolean).join(' | ');
    const t=c.dataset.tier;
    const rbg=t==='high'?'FEF2F2':t==='med'?'FFF5F5':t==='small'?'FFFBF0':(i%2===0?'FFFFFF':'F7F7F7');
    const tbg=mtype==='IN'?'E8F5EE':mtype==='OUT'?'FFF0F0':rbg;
    const tcol=mtype==='IN'?'1A7A4A':mtype==='OUT'?'CC2200':'666666';
    sRows.push(new TableRow({children:[
      cell(String(i+1),{fill:rbg,bold:true,color:'555555',size:8,align:AlignmentType.CENTER,width:sW[0]}),
      cell((sender||'').replace('📱','').replace('●','').replace('•','').trim(),{fill:rbg,bold:true,color:'1A4A8A',size:8,width:sW[1]}),
      cell(dt.replace('🕐','').trim().slice(0,16),{fill:rbg,color:'444444',size:7,width:sW[2]}),
      cell(mtype,{fill:tbg,bold:true,color:tcol,size:9,align:AlignmentType.CENTER,width:sW[3]}),
      cell(body,{fill:rbg,color:'111111',size:8,width:sW[4]}),
      cell(rem,{fill:rbg,color:'333333',size:8,width:sW[5]}),
    ]}));
  });

  // Info rows
  const iData=[['Target',tn_val||'All'],['SMS Count',String(cards.length)],['Generated',now_str]];
  const iRows=iData.map((r,i)=>new TableRow({children:[
    cell(r[0],{fill:i%2===0?'EEF2FF':'FFFFFF',bold:true,size:10,color:'1A1A2E',width:5}),
    cell(r[1],{fill:i%2===0?'EEF2FF':'FFFFFF',bold:true,size:10,color:'C0392B',width:6}),
  ]}));

  // Assemble
  const hdrPara=new Paragraph({
    border:{bottom:{style:BorderStyle.SINGLE,size:6,color:'C0392B',space:1}},
    spacing:{after:40},
    children:[
      new TextRun({text:'SMS INTELLIGENCE REPORT',bold:true,size:18,color:'1A1A2E',font:'Arial'}),
      new TextRun({text:tn_val?'   |   Target: '+tn_val:'',size:18,color:'555555',font:'Arial'}),
      new TextRun({text:'   |   CONFIDENTIAL',bold:true,size:18,color:'C0392B',font:'Arial'}),
    ],
  });
  const ftrPara=new Paragraph({
    border:{top:{style:BorderStyle.SINGLE,size:6,color:'C0392B',space:1}},
    children:[new TextRun({text:'CDR Forensic Analysis | Bangladesh Law Enforcement | CONFIDENTIAL',size:16,color:'888888',font:'Arial'})],
  });

  const doc=new Document({
    creator:'SMS Intelligence',title:'SMS Intelligence Report',
    sections:[{
      properties:{page:{size:{width:12240,height:15840},margin:{top:1080,bottom:1080,left:850,right:850}}},
      headers:{default:new docx.Header({children:[hdrPara]})},
      footers:{default:new docx.Footer({children:[ftrPara]})},
      children:[
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:400,after:120},children:[new TextRun({text:'SMS INTELLIGENCE REPORT',bold:true,size:52,color:'1A1A2E',font:'Arial'})]}),
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:320},children:[new TextRun({text:'Forensic SMS Analysis — CDR Intelligence',italics:true,size:24,color:'555555',font:'Arial'})]}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:iRows}),
        new Paragraph({spacing:{after:120}}),
        secHead('1','MFS / Bank Transaction Summary'),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:bRows}),
        new Paragraph({spacing:{after:120}}),
        secHead('2','Important SMS Messages — All ('+cards.length+' records)'),
        new Paragraph({spacing:{after:60},children:[new TextRun({text:'Deleted cards excluded. Generated: '+now_str,italics:true,size:18,color:'888888',font:'Arial'})]}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:sRows}),
        new Paragraph({spacing:{after:60}}),
        ..._bottomBlocks(),
      ],
    }],
  });

  Packer.toBlob(doc).then(blob=>{
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download='SMS_Intelligence_Report_'+(tn_val||'report')+'.docx';
    a.click();
    btn.textContent='📄 Download Word'; btn.disabled=false;
  }).catch(e=>{ btn.textContent='❌ Error'; btn.disabled=false; console.error(e); });
}


'''

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
.cd-summary{{margin-top:14px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:12px;padding:14px 16px}}
.cd-head{{font-size:13px;font-weight:700;color:#e2e8f0;margin-bottom:12px;letter-spacing:.3px}}
.cd-stats{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px}}
.cd-stat{{text-align:center;padding:10px;border-radius:10px;border:1px solid rgba(255,255,255,.08)}}
.cd-in{{background:rgba(22,163,74,.14);border-color:rgba(22,163,74,.4)}}
.cd-out{{background:rgba(220,38,38,.14);border-color:rgba(220,38,38,.4)}}
.cd-net{{background:rgba(255,255,255,.06)}}
.cd-amt{{font-size:20px;font-weight:800;line-height:1.1}}
.cd-in .cd-amt{{color:#86efac}}.cd-out .cd-amt{{color:#fca5a5}}
.cd-lbl{{font-size:10.5px;color:#94a3b8;margin-top:4px}}
.cd-bar{{display:flex;height:26px;border-radius:7px;overflow:hidden;border:1px solid rgba(255,255,255,.12)}}
.cd-bar-in{{background:linear-gradient(90deg,#16a34a,#22c55e);color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;min-width:2px}}
.cd-bar-out{{background:linear-gradient(90deg,#dc2626,#ef4444);color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;min-width:2px}}
@media (max-width:640px){{.cd-stats{{grid-template-columns:1fr}}}}
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
.fbar-hint{{font-size:10px;color:#9ca3af;margin-left:4px;white-space:nowrap}}
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
.ag-bal-multi{{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;padding:6px 8px;margin-top:6px}}
.ag-bal-hdr{{font-size:10px;font-weight:700;color:#15803d;margin-bottom:4px}}
.acc-row{{display:flex;justify-content:space-between;align-items:center;padding:2px 0;border-bottom:1px solid #dcfce7}}
.acc-row:last-child{{border-bottom:none}}
.acc-no{{font-family:'JetBrains Mono',monospace;font-size:10px;color:#374151}}
.acc-bal{{font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700;color:#15803d}}
.ag-accts{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:6px 8px;margin-top:6px}}
.ag-accts-hdr{{font-size:10px;font-weight:700;color:#334155;margin-bottom:5px}}
.acc-row2{{padding:4px 0;border-bottom:1px solid #eef2f7}}
.acc-row2:last-child{{border-bottom:none}}
.acc-id{{display:flex;align-items:center;gap:5px;margin-bottom:2px}}
.acc-id code{{font-family:'JetBrains Mono',monospace;font-size:10px;color:#1e293b;font-weight:700}}
.acc-ic{{font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;white-space:nowrap}}
.acc-ic.card{{background:#fef3c7;color:#b45309;border:1px solid #fcd34d}}
.acc-ic.reg{{background:#dbeafe;color:#1d4ed8;border:1px solid #93c5fd}}
.acc-meta{{display:flex;flex-wrap:wrap;gap:6px;padding-left:2px}}
.acc-cnt{{font-size:9.5px;color:#64748b;font-weight:600}}
.acc-in{{font-size:9.5px;color:#16a34a;font-family:'JetBrains Mono',monospace}}
.acc-out{{font-size:9.5px;color:#dc2626;font-family:'JetBrains Mono',monospace}}
.ag-cycling{{background:#fefce8;border:1px solid #fde047;border-radius:5px;padding:4px 7px;margin-top:5px;font-size:10px;color:#854d0e}}
.cyc-note{{color:#a16207;font-style:italic}}
.ftxn-summary{{background:#f0fdfa;border:1px solid #99f6e4;border-left:4px solid #0f766e;border-radius:10px;padding:14px 18px;margin-bottom:12px}}
.fts-header{{font-size:13px;font-weight:700;color:#0f766e;margin-bottom:10px}}
.fts-row{{display:flex;gap:16px;align-items:center;padding:6px 0;border-bottom:1px solid #ccfbf1;flex-wrap:wrap}}
.fts-row:last-of-type{{border-bottom:none}}
.fts-ccy{{font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;color:#134e4a;min-width:40px}}
.fts-val{{font-size:12px;font-family:'JetBrains Mono',monospace;color:#374151}}
.fts-note{{font-size:10px;color:#6b7280;margin-top:8px;font-style:italic}}
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
.del-btn{{background:none;border:1px solid #fca5a5;color:#dc2626;font-size:14px;
  padding:3px 8px;border-radius:6px;cursor:pointer;margin-left:6px;
  transition:all .2s;line-height:1;flex-shrink:0}}
.del-btn:hover{{background:#fef2f2;border-color:#dc2626;transform:scale(1.1)}}
.scard.deleted{{display:none!important}}
.undo-bar{{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);
  background:#1a1a2e;color:#fff;padding:10px 20px;border-radius:24px;
  font-size:13px;font-weight:600;display:flex;align-items:center;gap:12px;
  box-shadow:0 4px 20px rgba(0,0,0,.3);z-index:9999;
  opacity:0;transition:opacity .3s;pointer-events:none}}
.undo-bar.show{{opacity:1;pointer-events:all}}
.undo-btn{{background:#c0392b;color:#fff;border:none;border-radius:12px;
  padding:4px 14px;cursor:pointer;font-size:12px;font-weight:700;font-family:inherit}}
.del-count{{background:rgba(255,255,255,.15);padding:2px 8px;border-radius:10px;font-size:12px}}
.footer{{background:#1a1a2e;padding:16px 28px;
  display:flex;justify-content:space-between;font-size:10px;color:rgba(255,255,255,.35)}}
@media print{{.topbar{{position:static}}.summary-sheet{{position:static}}
  .sidebar{{position:static;height:auto}}.layout{{grid-template-columns:1fr}}
  .fbar{{display:none}}.scard{{break-inside:avoid}}
  .del-btn,.undo-bar,.no-print{{display:none!important}}
  .scard.deleted{{display:none!important}}}}
{IM_CSS}
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
    <button class="dl-top" onclick="downloadWord()" id="word-btn">📄 Download Word</button>
    <button class="dl-top" onclick="window.print()" style="margin-left:0">🖨 Print / Save PDF</button>
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
  <div class="cd-summary">
    <div class="cd-head">💰 Total Credit vs Debit Summary</div>
    <div class="cd-stats">
      <div class="cd-stat cd-in">
        <div class="cd-amt">৳{_credit_total:,.0f}</div>
        <div class="cd-lbl">Total Credit (IN) · {_credit_cnt} txn</div>
      </div>
      <div class="cd-stat cd-out">
        <div class="cd-amt">৳{_debit_total:,.0f}</div>
        <div class="cd-lbl">Total Debit (OUT) · {_debit_cnt} txn</div>
      </div>
      <div class="cd-stat cd-net">
        <div class="cd-amt" style="color:{'#86efac' if _net_flow>=0 else '#fca5a5'}">৳{_net_flow:,.0f}</div>
        <div class="cd-lbl">Net Flow {'(Surplus)' if _net_flow>=0 else '(Deficit)'}</div>
      </div>
    </div>
    <div class="cd-bar">
      <div class="cd-bar-in" style="width:{_cr_pct:.1f}%" title="Credit {_cr_pct:.1f}%">{(f'{_cr_pct:.0f}% IN' if _cr_pct>=12 else '')}</div>
      <div class="cd-bar-out" style="width:{_db_pct:.1f}%" title="Debit {_db_pct:.1f}%">{(f'{_db_pct:.0f}% OUT' if _db_pct>=12 else '')}</div>
    </div>
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
    <span class="fbar-hint">✦ multiple select</span>
    {fbtns}
    <div class="fsep"></div>
    <span class="flbl">Month:</span>
    <button class="mbtn active" data-all="1" onclick="fMonthAll(this)">All</button>
    {mbtns}
    <input class="sbox" type="text" placeholder="🔍 Search messages..." oninput="fSearch(this.value)">
  </div>
  <div class="content">
    <div class="rh"><div class="rinfo">Showing: <span id="vc">{len(display_list)}</span> SMS records</div></div>
    <div class="scards" id="cards">{ftxn_summary_html}{cards}</div>
  </div>
</div>
</div>

<!-- Intelligence analysis — মেসেজের পরে, একদম নিচে -->
<div class="im-bottom">
<div class="im-bottom-title">🧠 Intelligence Analysis — Profile · Reconciliation · Timeline · Counterparties · Flags</div>
<div class="im-bottom-hint no-print">🗑 Any section, row or item below can be deleted (↩ Undo available). Then use <b>📄 Download Word</b> or <b>🖨 Print / Save PDF</b> — deleted items are left out.</div>
<div class="im-wrap">{_profile_html}{_recon_html}</div>
{_intel_html}
</div>

<!-- Undo bar -->
<div id="undo-bar" class="undo-bar no-print">
  <span>🗑 <span id="del-cnt" class="del-count">0</span> item(s) deleted</span>
  <button class="undo-btn" onclick="undoLast()">↩ Undo</button>
  <button class="undo-btn" style="background:#475569" onclick="clearAll()">✕</button>
</div>

<div class="footer">
  <span>🔒 Confidential — For Law Enforcement Use Only</span>
  <span>{"Target: " + esc(tn) + " | " if tn else ""}Generated: {now} | SMS Intelligence Dashboard</span>
</div>
<script>
const allC=document.querySelectorAll('.scard');

// ── Multi-select filter engine ──────────────────────────────────────────────
// selTiers, selCats, selMonths: Set of selected values. Empty = "All"
let selTiers=new Set(), selCats=new Set(), selMonths=new Set(), sQ='';

function apply(){{
  let v=0;
  allC.forEach(c=>{{
    if(c.classList.contains('deleted'))return;
    const tier  = c.dataset.tier  || '';
    const cats  = c.dataset.cats  || '';
    const month = c.dataset.month || '';
    const txt   = c.innerText.toLowerCase();

    // Filter logic: tier OR cat = union (not intersection)
    const hasFilter = selTiers.size>0 || selCats.size>0;
    const tierMatch = selTiers.size>0 && selTiers.has(tier);
    const catMatch  = selCats.size>0  && [...selCats].some(x=>cats.includes(x));
    const okMain    = !hasFilter || tierMatch || catMatch;
    const okMonth = selMonths.size===0 || selMonths.has(month);
    const okSrch  = !sQ || txt.includes(sQ.toLowerCase());

    const ok = okMain && okMonth && okSrch;
    c.style.display = ok ? '' : 'none';
    if(ok) v++;
  }});
  document.getElementById('vc').textContent = v;
  _sync();
}}

function _sync(){{
  // Tier / Category / "All" filter buttons
  const noFilter = selTiers.size===0 && selCats.size===0;
  document.querySelectorAll('.fbtn').forEach(b=>{{
    if('all' in b.dataset)       b.classList.toggle('active', noFilter);
    else if('tier' in b.dataset) b.classList.toggle('active', selTiers.has(b.dataset.tier));
    else if('cat'  in b.dataset) b.classList.toggle('active', selCats.has(b.dataset.cat));
  }});
  // Month buttons
  const noMonth = selMonths.size===0;
  document.querySelectorAll('.mbtn').forEach(b=>{{
    if('all'   in b.dataset) b.classList.toggle('active', noMonth);
    else if('month' in b.dataset) b.classList.toggle('active', selMonths.has(b.dataset.month));
  }});
}}

// All: clear all selections → show everything
function fAll(b){{ selTiers.clear(); selCats.clear(); apply(); }}

// Tier toggle: click once = select, click again = deselect
function fTier(b,t){{
  if(selTiers.has(t)) selTiers.delete(t); else selTiers.add(t);
  apply();
}}

// Category toggle
function fCat(b,c){{
  if(selCats.has(c)) selCats.delete(c); else selCats.add(c);
  apply();
}}

// Month All: clear month filter
function fMonthAll(b){{ selMonths.clear(); apply(); }}

// Month toggle
function fMonth(b,m){{
  if(selMonths.has(m)) selMonths.delete(m); else selMonths.add(m);
  apply();
}}

function fSearch(v){{sQ=v;apply();}}
function downloadHTML(){{const blob=new Blob([document.documentElement.outerHTML],{{type:'text/html'}});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='SMS_Intelligence_Report_{esc(tn) if tn else "report"}.html';a.click();}}

{_word_js}
// ── Delete card with Undo ──────────────────────────────────────────────────
let _delStack=[], _undoTimer=null;
const _undoBar=document.getElementById('undo-bar');

function delCard(id){{
  const el=document.getElementById(id);
  if(!el) return;
  el.classList.add('deleted');
  _delStack.push(id);
  document.getElementById('del-cnt').textContent=_delStack.length;
  _undoBar.classList.add('show');
  clearTimeout(_undoTimer);
  _undoTimer=setTimeout(()=>_undoBar.classList.remove('show'),4000);
  apply();
}}
function undoLast(){{
  const id=_delStack.pop();
  if(typeof id==='string'){{ document.getElementById(id)?.classList.remove('deleted'); }}
  else if(id){{ id.classList.remove('deleted'); }}
  document.getElementById('del-cnt').textContent=_delStack.length;
  if(!_delStack.length) _undoBar.classList.remove('show');
  apply();
}}
function clearAll(){{ _delStack=[]; _undoBar.classList.remove('show'); }}
{_dz_js}
</script>
</body></html>'''
    with open(out_path,'w',encoding='utf-8') as f: f.write(html)

# ═══════════════════════════════════════════════════════════════════════════
#  SUBJECT PROFILE CLUES — চাকরি, বাসস্থান, যানবাহন, ভ্রমণ, চিকিৎসা, নাম, খরচের স্থান
#  (ফিল্টার হওয়া মেসেজসহ সব মেসেজ থেকে বের করা হয়)
# ═══════════════════════════════════════════════════════════════════════════
_PC_CATS = [
    ("employment", "💼 Employment & Income"),
    ("residence",  "🏠 Residence, Utility & Property"),
    ("vehicle",    "🚗 Vehicle"),
    ("travel",     "✈️ Travel"),
    ("health",     "🏥 Health / Medical Institutions"),
    ("names",      "👤 Names Addressed in SMS"),
    ("spending",   "🛍️ Card Spending Locations"),
    ("finance",    "🏦 Other Financial Footprint"),
    ("business",   "📑 Business / Govt Procurement"),
]
_PC_VEH_RE = re.compile(
    r'\b((?:DHAKA|CHATTA|CHATTOGRAM|CHITTAGONG|CTG|KHULNA|RAJSHAHI|SYLHET|BARISHAL|BARISAL|RANGPUR|'
    r'MYMENSINGH|GAZIPUR|NARAYANGANJ|CUMILLA|COMILLA|BOGURA|BOGRA|JASHORE|JESSORE|NARSINGDI|SAVAR)'
    r'(?:\s+METRO)?[\s\-]+[A-Z]{1,4}[\s\-]+\d{2}[\s\-]+\d{4})\b')
_PC_NAME_STOP = {"customer","sir","madam","valued","consumer","user","cardholder","parents","member",
                 "subscriber","client","guardian","student","patient","all","concern","friend","applicant"}
_PC_HEALTH_RE = re.compile(r'hospital|diagnostic|clinic|medical|labaid|popular|square\s*hosp|ibn\s*sina|evercare|\bbsh\b|healthcare', re.I)
_PC_HEALTH_ENGAGE = re.compile(r'registration|uhid|report|thank\s*you\s*for\s*choosing|feedback|appointment|patient|serial|health\s*card|admi(?:t|ssion)|discharge|purchase|bill|lab\b|test', re.I)

def _pc_dt(d):
    try: return to_dt(d)
    except Exception: return None

def extract_profile_clues(items):
    """items: merged SMS list ({sender, body, date, ...}) — ফিল্টার করার আগের।
    Returns: {cat: [ {label, detail, count, first, last, amount, body, sender}, ... ]}"""
    agg = {}   # (cat, key) -> dict
    def add(cat, key, label, detail="", it=None, amount=0.0):
        k = (cat, key)
        d = agg.get(k)
        dt = _pc_dt(it.get("date","")) if it else None
        if not d:
            d = agg[k] = {"cat": cat, "label": label, "detail": detail, "count": 0, "first": dt, "last": dt,
                          "amount": 0.0, "amounts": [], "body": "", "sender": "", "_bodies": set()}
        d["count"] += 1
        if cat in ("health", "spending") and len(label) > len(d["label"]): d["label"] = label
        if amount: d["amount"] += amount; d["amounts"].append(amount)
        if dt:
            if not d["first"] or dt < d["first"]: d["first"] = dt
            if not d["last"] or dt >= d["last"]:
                d["last"] = dt
                if it: d["body"] = it.get("body",""); d["sender"] = str(it.get("sender",""))
        elif it and not d["body"]:
            d["body"] = it.get("body",""); d["sender"] = str(it.get("sender",""))
        if detail and detail not in d["detail"]:
            d["detail"] = (d["detail"] + " · " + detail).strip(" ·") if d["detail"] else detail

    for it in items:
        b = str(it.get("body","") or ""); snd = str(it.get("sender","") or "")
        if not b.strip(): continue
        bl = b.lower()

        # ── Employment / Income ──
        m = re.search(r'EFT\s+Advice.*?Salary\s+for\s+the\s+month\s+of\s+([A-Za-z]+\s*\d{4}).*?Tk\.?\s*([\d,]+)', b, re.I | re.S)
        if m:
            amt = float(m.group(2).replace(",", ""))
            add("employment", "salary_eft", "Monthly salary received via Govt. EFT (iBAS++)",
                f"{m.group(1).strip()}: ৳{amt:,.0f}", it, amt)
        m = re.search(r'salary\s+bill\s+for\s+the\s+month.*?(?:forwarded|sent)\s+(?:by\s+DDO\s+)?to\s+(.+?)\s*-\s*iBas', b, re.I | re.S)
        if m:
            off = m.group(1).strip()
            mm = re.search(r'Ministry\s+of\s+([A-Za-z &]+)|CAFO\s*,?\s*([A-Za-z &]+)', off)
            if mm:
                minis = (mm.group(1) or mm.group(2)).strip()
                minis = re.sub(r'^(?:Ministry\s+of|Office\s+of\s+the)\s+', '', minis, flags=re.I).strip()
                add("employment", "employer_" + minis.lower(), f"Employer (Govt. salary bill via iBAS++): Ministry of {minis}",
                    "", it)
        if re.search(r'health\s*card.*sarkari\s*karmachari', b, re.I):
            add("employment", "govt_health_card", "Govt. employee health card (Sarkari Karmachari Hospital)", "", it)
        m = re.search(r'\bsalary\b.*?(?:credited|BDT|Tk)\s*([\d,]+\.?\d*)', b, re.I)
        if m and not re.search(r'EFT\s+Advice|salary\s+bill|loan|offer|account\s+open', b, re.I) and re.search(r'credited|deposit', b, re.I):
            try:
                amt = float(m.group(1).replace(",", ""))
                add("employment", "salary_credit_" + snd.lower(), f"Salary credited ({snd})", f"৳{amt:,.0f}", it, amt)
            except ValueError: pass

        # ── Business / Procurement ──
        m = re.search(r'e-?GP.*?Tender\s*ID\s*[:#]?\s*(\d+)', b, re.I | re.S)
        if m:
            add("business", "egp_" + m.group(1), f"e-GP tender ID {m.group(1)}", "", it)

        # ── Vehicle ──
        for v in set(re.sub(r'\s+', ' ', x).replace(" -", "-").replace("- ", "-") for x in _PC_VEH_RE.findall(b.upper())):
            ev = ("Tax-Token" if "tax-token" in bl else "Fitness" if "fitness" in bl else
                  "Trustee board cert." if "trustee" in bl else "BRTA notice")
            dm = re.search(r'(?:expire[sd]?\s+on|is)\s+(\d{1,2}-[A-Z]{3}-\d{2,4})', b, re.I)
            add("vehicle", "veh_" + v, f"Vehicle {v}", f"{ev}{(' ' + dm.group(1)) if dm else ''}", it)

        # ── Residence / Utility / Property ──
        m = re.search(r'Biller\s*:\s*([A-Za-z][A-Za-z0-9 .&\-]*?)\s*(?:\(([A-Za-z ]+)\))?\s*A/C\s*:\s*(\d{4,})', b)
        if m:
            nm = m.group(1).strip() + (f" ({m.group(2)})" if m.group(2) else "")
            am = re.search(r'Amount\s*:\s*Tk\s*([\d,]+\.?\d*)', b, re.I)
            add("residence", "util_" + m.group(3), f"Utility account — {nm} A/C {m.group(3)}",
                "", it, float(am.group(1).replace(",", "")) if am else 0.0)
        m = re.search(r'(DPDC|DESCO|NESCO|WZPDCL|BPDB|BREB|Palli\s*Bidyut)?\s*Prepaid\s+Token.*?Customer\s*No\s*[:.]?\s*(\d{4,})', b, re.I | re.S)
        if m:
            add("residence", "util_" + m.group(2), f"Utility account — {(m.group(1) or 'Prepaid meter').upper()} Customer No {m.group(2)}", "(prepaid meter token)", it)
        m = re.search(r'consumer\s*(?:number|no\.?)\s*[:#]?\s*(\d{4,}).*?meter\s*(?:number|no\.?)\s*[:#]?\s*(\d{4,})', b, re.I | re.S)
        if m:
            add("residence", "meter_" + m.group(2), f"Electricity consumer {m.group(1)}, meter {m.group(2)}", "", it)
        m = re.search(r'\b(DNCC|DSCC|CCC|RCC|KCC|SCC|BCC|GCC|NCC|CUCC|MCC)\s+Holding\s+Tax\s*(?:Id|ID)?\s*(\d+)?(?:.*?Bill\s*No\.?\s*(\d+))?', b, re.I | re.S)
        if m:
            am = re.search(r'Tk\.?\s*([\d,]+\.?\d*)\s*paid', b, re.I)
            add("residence", "holding_" + m.group(1).upper() + (m.group(2) or ""),
                f"{m.group(1).upper()} Holding Tax paid (property)" + (f" — Holding ID {m.group(2)}" if m.group(2) else ""),
                f"Bill {m.group(3)}" if m.group(3) else "", it, float(am.group(1).replace(",", "")) if am else 0.0)
        # কুরিয়ার বুকিং: "Mango booked from Vomra, Satkhira. to Shyamoli,Dhaka. CN: …"
        m = re.search(r'booked\s+from\s+(.+?)\.?\s+to\s+([A-Za-z][A-Za-z ,\-]+?)\.\s*CN', b, re.I)
        if m:
            _dst = re.sub(r'\s*,\s*', ', ', m.group(2).strip())
            _src = re.sub(r'\s*,\s*', ', ', m.group(1).strip().rstrip('.'))
            add("residence", "cdst_" + _dst.lower(), f"Courier deliveries addressed to: {_dst}", f"from {_src}", it)
        m = re.search(r'local\s+post\s+office\s*\((\d{4})\)', b, re.I)
        if m:
            add("residence", "post_" + m.group(1), f"Parcel delivered to local post office — Post code {m.group(1)}", "", it)
        m = re.search(r'ready\s+at\s+(.+?)(?:,\s*Pls|\.\s|,\s*Contact|$)', b, re.I)
        if m and re.search(r'courier|parcel|product|fruits|delivery|sundarban|pathao|redx|steadfast|sa\s*paribahan', b + " " + snd, re.I):
            add("residence", "pickup_" + m.group(1).strip().lower()[:40], f"Courier pickup point: {m.group(1).strip()}", f"via {snd}", it)

        # ── Travel ──
        m = re.search(r'PNR[\-:\s]*([A-Z0-9]{5,7})\b(.*?)(?:Ticket\s+amount\s*([\d,.]+))?', b, re.S)
        if m and re.search(r'[A-Z]{3}-[A-Z]{3}', b):
            segs = re.findall(r'\b([A-Z0-9]{2}\s?\d{2,4})\s+([A-Z]{3})-([A-Z]{3})\s+on\s+(\d{1,2}\s+[A-Za-z]{3})(?:\s+at\s+([\d:]+\s*[AP]M))?', b)
            route = "; ".join(f"{f} {a}→{c} {d}{(' ' + t) if t else ''}" for f, a, c, d, t in segs) or "flight"
            fare = re.search(r'Ticket\s+amount\s*([\d,]+\.?\d*)', b, re.I)
            add("travel", "pnr_" + m.group(1), f"Flight PNR {m.group(1)} ({snd}) — {route}",
                "", it, float(fare.group(1).replace(",", "")) if fare else 0.0)
        m = re.search(r'COACH\s*:\s*(\S+)\s+([A-Z][A-Z\-]+)', b)
        if m and re.search(r'FARE', b):
            fare = re.search(r'FARE\s*:\s*(\d+)', b); seat = re.search(r'TS\s*:\s*([\w,\-]+)', b)
            op = re.search(r'\]\s*(.+?)(?:\s+TK|$)', b)
            add("travel", "bus_" + m.group(2) + m.group(1), f"Bus ticket {m.group(2)} (coach {m.group(1)})",
                (f"seat {seat.group(1)}" if seat else "") + (f" · {op.group(1).strip()[:60]}" if op else ""),
                it, float(fare.group(1)) if fare else 0.0)

        # ── Health ──
        if (_PC_HEALTH_RE.search(snd) or _PC_HEALTH_RE.search(b)) and _PC_HEALTH_ENGAGE.search(b) \
                and not re.search(r'Purchase\s+BDT', b, re.I) \
                and not re.search(r'শুভেচ্ছা|অভিনন্দন|greetings?\s+from.*?(?:eid|new\s+year)|offer|discount|ছাড়|package', b, re.I):
            nm = re.search(r'([A-Z][A-Za-z.&]*(?:\s+[A-Z][A-Za-z.&]*)*\s+(?:Hospital|Hospitals|Diagnostic(?:\s+Cent(?:re|er))?|Clinic|Medical(?:\s+College)?)'
                           r'(?:\s*(?:Ltd\.?|PLC))?(?:\s*\([A-Za-z ]+\))?)', b)
            if nm or _PC_HEALTH_RE.search(snd):
                inst = (nm.group(1).strip() if nm else snd).strip()
                ids = re.findall(r'(UHID\s*(?:is|no\.?|:)?\s*[A-Z]{0,4}\d{5,}|Registration\s*ID\s*(?:is|:)?\s*\d{5,}|LAB\s*/\s*RIS\s*number\s*is\s*\d+)', b, re.I)
                _hk = re.sub(r'\s*\(.*?\)|\b(?:ltd|plc|limited|centre|center|lab|labs)\b\.?', '', inst.lower())
                _hk = " ".join(re.sub(r'[^a-z ]', ' ', _hk).split()[:2]) or inst.lower()[:20]
                add("health", "h_" + _hk, inst, " · ".join(i.strip() for i in ids), it)

        # ── Names addressed ──
        for nm in re.findall(r'\bDear\s+(?:Mr|Mrs|Ms|Miss|Md|Dr|Engr)\.?\s+([A-Z][A-Za-z.]*(?:\s{1,2}[A-Z][A-Za-z.]*){0,3})', b):
            n2 = re.sub(r'\s+', ' ', nm).strip(" .,")
            if n2.split()[0].lower() in _PC_NAME_STOP or len(n2) < 3: continue
            add("names", "n_" + n2.lower(), n2.title() if n2.isupper() else n2, f"by {snd}", it)

        for nm in re.findall(r'\bDear\s+([\u0980-\u09FF]{2,}(?:\s[\u0980-\u09FF]{2,}){0,3})\s*,', b):
            n2 = nm.strip()
            if n2 in ("গ্রাহক", "সম্মানিত", "প্রিয়", "অভিভাবক", "শিক্ষার্থী"): continue
            add("names", "n_" + n2, n2, f"by {snd}", it)

        # ── Card spending locations ──
        m = re.search(r'Purchase\s+BDT\s*([\d,]+\.?\d*)\s+from\s+([^,]+),\s*([A-Z][A-Z ]+?),\s*([^.]+?)\.\s', b)
        if m:
            mer = m.group(2).strip().rstrip(" &.")
            _city = m.group(3).strip().title(); _ar = m.group(4).strip().strip(", ")
            _area = ", ".join(x for x in ([_ar] if _ar and _ar.upper() != _city.upper() else []) + [_city] if x)
            add("spending", "sp_" + re.sub(r'[^a-z0-9]', '', mer.lower()), mer, _area, it,
                float(m.group(1).replace(",", "")))

        # ── Other financial footprint ──
        m = re.search(r'\(([A-Z]{3,10})\)\s*এ\s+আপনার\s+([\d,]+)\s*/-\s*সঞ্চয়\s+জমা', b)
        if m:
            add("finance", "coop_" + m.group(1), f"Cooperative savings deposits — {m.group(1)}", "", it,
                float(m.group(2).replace(",", "")))
        m = re.search(r'received\s+(?:Digital\s+)?Loan\s+(?:of\s+)?Tk\s*([\d,]+\.?\d*)\s+from\s+([A-Za-z ]+?Bank)', b, re.I)
        if m:
            add("finance", "loan_" + m.group(2).lower(), f"Loan received — {m.group(2).strip()}", "", it,
                float(m.group(1).replace(",", "")))
        m = re.search(r'Loan\s+Repayment\s+of\s+Tk\s*([\d,]+\.?\d*)', b, re.I)
        if m:
            add("finance", "loanrep", "Loan repayment", "", it, float(m.group(1).replace(",", "")))

    out = {}
    for (cat, key), d in agg.items():
        d.pop("_bodies", None)
        out.setdefault(cat, []).append(d)
    for cat in out:
        out[cat].sort(key=lambda d: (-(d["amount"] or 0), -d["count"]))
    return out


# ═══════════════════════════════════════════════════════════════════════════
#  BALANCE RECONCILIATION — SMS-এ দেখা IN/OUT বনাম ব্যালেন্সের আসল পরিবর্তন
# ═══════════════════════════════════════════════════════════════════════════
def compute_reconciliation(results, tol=200.0):
    """প্রতিটি একাউন্ট/ওয়ালেটে ব্যালেন্স-চেকপয়েন্টের মধ্যে SMS-এ না-থাকা লেনদেন বের করে।
    ক্রেডিট কার্ড বাদ (কার্ডের 'Balance' আসলে available limit)।"""
    from collections import defaultdict as _dd
    groups = _dd(list)
    for k, s in enumerate(results):
        a = s.get("acct", {}) or {}
        svc = a.get("service", "")
        if not svc or s.get("sender_type") not in ("mfs", "bank"): continue
        if a.get("acct_type") == "card": continue
        if re.search(r'card\s+bill|total\s+due', s.get("body", ""), re.I): continue
        acc = a.get("acct_id", "") if a.get("acct_type") == "account" else ""
        groups[(svc, acc)].append((to_dt(s.get("date", "")), k, s))
    out = []
    for (svc, acc), evs in groups.items():
        evs = [e for e in evs if e[0]]
        evs.sort(key=lambda e: (e[0], e[1]))
        def _io(s):
            a = s.get("acct", {}) or {}; amt = s.get("amount", 0) or 0
            d = a.get("direction", "")
            return (amt if d == "Received" else 0.0), (amt if d == "Sent" else 0.0)
        def _bl(s): return (s.get("acct", {}) or {}).get("balance", 0) or 0
        # ── একই মিনিটের মেসেজ: (১) একই ব্যালেন্স লেখা একাধিক মেসেজ = একটি ব্যাচ (ব্যাংক
        #    সম্মিলিত চূড়ান্ত ব্যালেন্স দেয়) → যোগ করে একটি চেকপয়েন্ট; (২) বাকিগুলো ব্যালেন্স-শৃঙ্খল
        #    মিলিয়ে ক্রমে সাজানো।
        units = []   # (dt, in, out, bal, body)
        _i = 0; _run = None
        while _i < len(evs):
            _j = _i
            while _j < len(evs) and evs[_j][0] == evs[_i][0]: _j += 1
            grp = evs[_i:_j]; batch = {}; singles = []
            for e in grp:
                b = _bl(e[2]); ii, oo = _io(e[2])
                if b > 0:
                    k = round(b, 2)
                    if k in batch:
                        batch[k][1] += ii; batch[k][2] += oo; batch[k][4] += " ‖ " + e[2].get("body", "")
                    else:
                        batch[k] = [e[0], ii, oo, b, e[2].get("body", "")]
                else:
                    singles.append([e[0], ii, oo, 0.0, e[2].get("body", "")])
            pool = list(batch.values()) + singles
            while pool:
                pick = None
                if _run is not None:
                    for u in pool:
                        if u[3] > 0 and abs(u[3] - (_run + u[1] - u[2])) <= tol: pick = u; break
                if pick is None: pick = pool[0]
                pool.remove(pick); units.append(pick)
                _run = pick[3] if pick[3] > 0 else ((_run + pick[1] - pick[2]) if _run is not None else None)
            _i = _j
        running = None; pending = 0.0
        first_bal = last_bal = None; first_dt = last_dt = prev_dt = None
        sms_in = sms_out = 0.0; gaps = []; n_ck = 0
        for dt, ii, oo, bal, body in units:
            delta = ii - oo
            if bal > 0:
                n_ck += 1
                if running is None:
                    running = bal; first_bal = bal; first_dt = dt; pending = 0.0
                else:
                    sms_in += ii; sms_out += oo
                    gap = bal - (running + pending + delta)
                    if abs(gap) > tol:
                        gaps.append({"from": prev_dt, "to": dt, "gap": gap, "body": body})
                    running = bal; pending = 0.0
                last_bal = bal; last_dt = dt; prev_dt = dt
            elif running is not None:
                pending += delta; sms_in += ii; sms_out += oo
        if n_ck < 2: continue
        # একই মিনিটের ক্রম-বিভ্রাটে তৈরি +/− জোড়া বাদ
        clean = []; i = 0
        while i < len(gaps):
            g = gaps[i]
            if i + 1 < len(gaps):
                h = gaps[i + 1]
                if abs(g["gap"] + h["gap"]) <= tol and h["to"] and g["to"] and abs((h["to"] - g["to"]).total_seconds()) <= 180:
                    i += 2; continue
            clean.append(g); i += 1
        exp_close = first_bal + sms_in - sms_out
        net = last_bal - exp_close
        u_out = sum(-g["gap"] for g in clean if g["gap"] < 0)
        u_in = sum(g["gap"] for g in clean if g["gap"] > 0)
        out.append({"service": svc, "acct": acc, "first_dt": first_dt, "last_dt": last_dt,
                    "open": first_bal, "close": last_bal, "sms_in": sms_in, "sms_out": sms_out,
                    "exp_close": exp_close, "net": net, "u_out": u_out, "u_in": u_in,
                    "gaps": sorted(clean, key=lambda g: g["gap"])[:8], "checkpoints": n_ck})
    out.sort(key=lambda r: -(r["u_out"] + r["u_in"]))
    return out

def _fmt_d(dt, f="%d %b %Y"):
    try: return dt.strftime(f) if dt else ""
    except Exception: return ""

def build_profile_html(P):
    """Subject Profile Clues → HTML (প্রতিটি আইটেম আলাদাভাবে ডিলিটযোগ্য)।"""
    if not P: return ""
    LIM = {"spending": 12, "names": 15, "health": 12}
    H = ['<div class="im-sec pc-sec"><div class="im-head">🧩 Subject Profile Clues</div>',
         '<div class="im-note">Identity, employment, residence, vehicle, travel and spending clues found across '
         '<b>all</b> SMS — including ones filtered out of the message list above (e.g. salary bills, meter tokens, BRTA notices). '
         'Each clue shows how often it appeared, the date range, and the latest source SMS in full.</div>']
    any_item = False
    for cat, title in _PC_CATS:
        items = P.get(cat) or []
        if not items: continue
        any_item = True
        H.append(f'<div class="pc-cat"><div class="im-ctitle">{title} <span class="pc-n">({len(items)})</span></div>')
        for d in items[:LIM.get(cat, 30)]:
            rng = _fmt_d(d.get("first"))
            if d.get("last") and d.get("first") and _fmt_d(d["last"]) != rng: rng += " → " + _fmt_d(d["last"])
            meta = [f'{d["count"]}×']
            if rng: meta.append(rng)
            if d.get("amount"): meta.append(f'৳{d["amount"]:,.0f}')
            H.append('<div class="pc-item">'
                     f'<div class="pc-top"><span class="pc-lbl">{hl.escape(d["label"])}</span>'
                     f'<span class="pc-meta">{" · ".join(meta)}</span></div>'
                     + (f'<div class="pc-det">{hl.escape(d["detail"])}</div>' if d.get("detail") else '')
                     + (f'<div class="pc-src"><span class="pc-snd">{hl.escape(d.get("sender",""))}:</span> '
                        f'{hl.escape(d.get("body",""))}</div>' if d.get("body") else '')
                     + '</div>')
        if len(items) > LIM.get(cat, 30):
            H.append(f'<div class="im-note">+ {len(items) - LIM.get(cat, 30)} more not shown</div>')
        H.append('</div>')
    if not any_item:
        H.append('<div class="im-noflag">No profile clues found in this SMS set.</div>')
    H.append('</div>')
    return "".join(H)

def build_recon_html(R):
    """Balance Reconciliation → HTML।"""
    if not R: return ""
    H = ['<div class="im-sec rc-sec"><div class="im-head">⚖️ Balance Reconciliation — SMS vs Actual Balance</div>',
         '<div class="im-note">For each wallet / bank account the balance printed in the SMS is chained from one message to the next. '
         'If the balance moves by more than the SMS transactions explain, money moved <b>without an SMS in this export</b> '
         '(e.g. cash-out / send-money / transfers whose SMS were not captured). Credit cards are excluded '
         '(their "balance" is the available limit). Small differences can be fees/charges.</div>']
    flags = []
    for r in R:
        thr = max(1000.0, 0.005 * (r["sms_in"] + r["sms_out"]))
        nm = f'{r["service"]} {r["acct"]}'.strip()
        per = f'{_fmt_d(r["first_dt"])} – {_fmt_d(r["last_dt"])}'
        if r["net"] < -thr:
            flags.append(f'<div class="im-flag im-flag-red">💸 <b>{hl.escape(nm)}:</b> about <b>৳{-r["net"]:,.0f}</b> left this account '
                         f'between {per} with <b>no SMS record</b> — outgoing transactions (cash-out / send money / transfer) '
                         'are missing from this SMS export.</div>')
        elif r["net"] > thr:
            flags.append(f'<div class="im-flag im-flag-amber">📥 <b>{hl.escape(nm)}:</b> about <b>৳{r["net"]:,.0f}</b> arrived in this account '
                         f'between {per} with <b>no SMS record</b> — an incoming credit SMS is missing.</div>')
    H.extend(flags)
    H.append('<div class="im-scroll"><table class="im-tbl rc-tbl"><tr><th>Account / Wallet</th><th>Period</th>'
             '<th>Opening bal.</th><th>SMS IN</th><th>SMS OUT</th><th>Expected closing</th><th>Actual closing</th>'
             '<th>Unexplained</th><th>Status</th></tr>')
    for r in R:
        thr = max(1000.0, 0.005 * (r["sms_in"] + r["sms_out"]))
        if abs(r["net"]) <= thr: st = '<span class="rc-ok">✓ Consistent</span>'
        elif r["net"] < 0:      st = '<span class="rc-bad">⚠ Untracked OUT</span>'
        else:                   st = '<span class="rc-warn">⚠ Untracked IN</span>'
        _nc = "#dc2626" if r["net"] < -thr else ("#b45309" if r["net"] > thr else "#475569")
        H.append(f'<tr><td><b>{hl.escape(r["service"])}</b>'
                 + (f'<div class="im-acct">{hl.escape(r["acct"])}</div>' if r["acct"] else '')
                 + f'</td><td>{_fmt_d(r["first_dt"], "%d %b %y")} → {_fmt_d(r["last_dt"], "%d %b %y")}'
                 f'<div class="im-acct">{r["checkpoints"]} balance points</div></td>'
                 f'<td>৳{r["open"]:,.0f}</td><td style="color:#16a34a">৳{r["sms_in"]:,.0f}</td>'
                 f'<td style="color:#dc2626">৳{r["sms_out"]:,.0f}</td><td>৳{r["exp_close"]:,.0f}</td>'
                 f'<td><b>৳{r["close"]:,.0f}</b></td>'
                 f'<td style="color:{_nc};font-weight:800">{"+" if r["net"] > 0 else ""}৳{r["net"]:,.0f}</td><td>{st}</td></tr>')
    H.append('</table></div>')
    # largest unexplained movements
    rows = []
    for r in R:
        for g in r["gaps"]:
            rows.append((abs(g["gap"]), r, g))
    rows.sort(key=lambda x: -x[0])
    if rows:
        H.append('<div class="im-ctitle">🔍 Largest Unexplained Balance Movements</div>'
                 '<div class="im-scroll"><table class="im-tbl rc-tbl"><tr><th>Account</th><th>Between</th>'
                 '<th>Unexplained</th><th>Detected at (next SMS with balance)</th></tr>')
        for _, r, g in rows[:12]:
            nm = f'{r["service"]} {r["acct"]}'.strip()
            col = "#dc2626" if g["gap"] < 0 else "#b45309"
            H.append(f'<tr><td>{hl.escape(nm)}</td>'
                     f'<td>{_fmt_d(g["from"], "%d %b %y %H:%M")} → {_fmt_d(g["to"], "%d %b %y %H:%M")}</td>'
                     f'<td style="color:{col};font-weight:800">{"−" if g["gap"] < 0 else "+"}৳{abs(g["gap"]):,.0f}</td>'
                     f'<td class="rc-body">{hl.escape(g["body"])}</td></tr>')
        H.append('</table></div>')
    H.append('</div>')
    return "".join(H)

def _dedupe_same_amount(results):
    """একই লেনদেনের দুটি মেসেজ (যেমন "payment reserved" + "payment successful",
    বা "received Loan" + "received Digital Loan") দুবার গোনা এড়ায়।

    নিয়ম (user-specified): একই service + একই amount + একই direction হলে, এবং
    অল্প সময়ের ব্যবধানে (≤180s) এলে, আর ব্যালেন্সে টাকা একবারই কর্তন হলে
    (দুই মেসেজের balance সমান, বা একটিতে balance নেই) → ০১টি লেনদেন হিসেবে গোনা।
    ব্যালেন্স দুবার কমলে (amount অনুযায়ী) → আলাদা লেনদেন, দুটোই রাখি।
    """
    if len(results) < 2: return results, 0
    # index সহ sort — time, তারপর original order
    idx = sorted(range(len(results)),
                 key=lambda i: (str(results[i].get("date","")), i))
    drop = set()
    # (service, amount, direction) অনুযায়ী consecutive candidate খুঁজি
    def _key(s):
        ai = s.get("acct",{}) or {}
        return (ai.get("service",""), round(s.get("amount",0) or 0, 2), ai.get("direction",""))
    def _bal(s): return (s.get("acct",{}) or {}).get("balance",0) or 0
    def _is_pending(s):
        b = s.get("body","").lower()
        return bool(re.search(r'being\s+reserved|is\s+reserved|reserved\s+for|pending|initiat', b))
    def _is_success(s):
        b = s.get("body","").lower()
        return bool(re.search(r'successful|completed|is\s+complete|সফল', b))

    for a in range(len(idx)):
        i = idx[a]
        if i in drop: continue
        si = results[i]
        if (si.get("amount",0) or 0) <= 0: continue
        ki = _key(si); ti = to_dt(si.get("date",""))
        for b in range(a+1, len(idx)):
            j = idx[b]
            if j in drop: continue
            sj = results[j]
            if _key(sj) != ki: continue
            tj = to_dt(sj.get("date",""))
            if ti and tj and abs((tj-ti).total_seconds()) > 180: break  # time উইন্ডো পার
            bi, bj = _bal(si), _bal(sj)
            # ব্যালেন্স দুবার কমেছে (amount অনুযায়ী) → আলাদা লেনদেন, রাখি
            if bi > 0 and bj > 0 and abs(abs(bi - bj) - ki[1]) < 1:
                continue
            # ব্যালেন্স সমান, বা একটিতে balance নেই → একই লেনদেনের দুই মেসেজ → merge
            same_bal = (bi > 0 and bj > 0 and abs(bi - bj) < 1)
            one_missing = (bi == 0) ^ (bj == 0)
            if same_bal or one_missing:
                # successful/final মেসেজটা রাখি, pending/reserved টা বাদ দিই
                if _is_pending(si) and not _is_pending(sj):
                    drop.add(i)
                else:
                    drop.add(j)
                if i in drop: break   # si বাদ পড়লে আর তুলনা করার দরকার নেই
    if not drop: return results, 0
    kept = [s for k,s in enumerate(results) if k not in drop]
    return kept, len(drop)

def _filter_and_analyze(raw, target_number="", input_path=""):
    """raw SMS list → merge → filter → analyze → (results, stats, tn)"""
    if target_number:
        tn_clean = re.sub(r'[^\d]','', target_number.strip())
        tn_variants = {tn_clean}
        if tn_clean.startswith('01'): tn_variants.add('880'+tn_clean)
        if tn_clean.startswith('880') and len(tn_clean) >= 13: tn_variants.add(tn_clean[3:])
        def is_related(item):
            sender_digits = re.sub(r'[^\d]','', str(item.get("sender","")))
            aparty_digits = re.sub(r'[^\d]','', str(item.get("aparty","")))
            return (any(v in sender_digits for v in tn_variants)
                    or any(v in aparty_digits for v in tn_variants)
                    or any(v in item.get("body","") for v in tn_variants))
        raw = [r for r in raw if is_related(r)]

    merged_ = merge(raw)
    filtered_cnt = 0; results = []; seen_keyed = set(); seen_body_only = set()
    seen_full_body = set()  # global body dedup — club/org mass SMS retry এড়াতে
    for item in merged_:
        if not item["body"].strip(): continue
        if is_filtered(item["sender"], item["body"]): filtered_cnt+=1; continue
        body_key  = item["body"].strip()
        _dt       = to_dt(item["date"])
        dt_min    = _dt.strftime('%Y%m%d%H%M') if _dt != datetime.min else str(item["date"])[:12]

        _st = stype(item["sender"], body_key)
        if _st in ("foreign_bank", "service", "other"):
            if body_key in seen_body_only: filtered_cnt+=1; continue
            seen_body_only.add(body_key)
        else:
            dt_hour   = _dt.strftime('%Y%m%d%H') if _dt != datetime.min else str(item["date"])[:11]
            if _st == "personal":
                # personal: sender+body+hour — ১ ঘন্টার মধ্যে same sender same body = ১টাই
                dk = (item["sender"], body_key, dt_hour)
            else:
                # bank/mfs: sender+body+minute (repeated txn each minute is valid)
                dk = (item["sender"], body_key, dt_min)
            if dk in seen_keyed: filtered_cnt+=1; continue
            seen_keyed.add(dk)
            # personal SMS global body dedup — club/org mass retry
            if _st == "personal" and len(body_key) > 20:
                if body_key in seen_full_body: filtered_cnt+=1; continue
                seen_full_body.add(body_key)
        res = analyze_one(item["sender"], item["body"], item["date"], item.get("direction",""))
        if res:
            res["parts"] = item.get("parts",1)
            results.append(res)

    # একই পেমেন্টের "reserved + successful" (বা দুই notification) দুবার গোনা এড়াই
    results, _dupcnt = _dedupe_same_amount(results)
    # সাবজেক্টের নিজের নম্বর — AParty কলামে সবচেয়ে বেশি আসা নম্বর (বা target number)
    _own = ""
    if target_number:
        _d = re.sub(r'\D', '', target_number)
        if len(_d) >= 10: _own = "0" + _d[-10:]
    if not _own:
        _ac = Counter()
        for _it in merged_:
            _d = re.sub(r'\D', '', str(_it.get("aparty", "") or ""))
            if len(_d) >= 10 and _d[-10:].startswith("1"): _ac["0" + _d[-10:]] += 1
        if _ac: _own = _ac.most_common(1)[0][0]
    stats = {"total":len(raw),"after_merge":len(merged_),"filtered":filtered_cnt,"deduped":_dupcnt,
             "own_number": _own,
             "profile": extract_profile_clues(merged_),
             "recon": compute_reconciliation(results)}
    tn = re.sub(r'[^\d]','', target_number) if target_number else ""
    if not tn and input_path:
        base = os.path.splitext(os.path.basename(input_path))[0]
        m2 = re.search(r'(\d{10,})', base)
        tn = m2.group(1) if m2 else ""
    if tn.startswith('880') and len(tn)>11: tn='0'+tn[3:]
    return results, stats, tn


def generate_docx(results, stats, tn=""):
    """results → Word (.docx) bytes (python-docx).
    - Clean light-coloured table headers (no dark background)
    - If tn set: Section 3 = all SMS, Section 4 = target number SMS separately
    - Message body preserved as-is; all other text in English.
    """
    from docx import Document as WDoc
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    CAT_LABEL = {
        "txn_high":"High Transaction (100K+)","txn_med":"Medium Transaction (50K-100K)",
        "txn_small":"Small Transaction (10K-50K)","identity":"Identity / Profile",
        "love_family":"Personal / Family","travel":"Travel / Visa / Umrah",
        "courier":"Courier / Delivery","chalan":"Govt. Challan / Treasury",
        "education":"Education / Admission","bo_account":"BO Account / Stock Market",
        "shopping":"Shopping / Billing","personal":"Personal Conversation",
    }

    def _rgb(h): return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

    def _shd(cell, fill):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
        shd.set(qn('w:fill'),fill); tcPr.append(shd)

    def _brd(cell, color="CCCCCC"):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        tcB=OxmlElement('w:tcBorders')
        for side in ('top','left','bottom','right'):
            el=OxmlElement(f'w:{side}')
            el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4')
            el.set(qn('w:space'),'0'); el.set(qn('w:color'),color)
            tcB.append(el)
        tcPr.append(tcB)

    def _run(para, text, bold=False, italic=False, size=10, color="111111"):
        r=para.add_run(str(text or ""))
        r.bold=bold; r.italic=italic
        r.font.name="Arial"; r.font.size=Pt(size)
        r.font.color.rgb=_rgb(color); return r

    def _cell(tbl_row, col_idx, text, width_cm, bold=False, italic=False,
              size=10, color="111111", fill="FFFFFF",
              align=WD_ALIGN_PARAGRAPH.LEFT, brd_col="CCCCCC"):
        c=tbl_row.cells[col_idx]; c.width=Cm(width_cm)
        _shd(c,fill); _brd(c,brd_col)
        c.vertical_alignment=WD_ALIGN_VERTICAL.TOP
        p=c.paragraphs[0]; p.clear(); p.alignment=align
        _run(p,text,bold=bold,italic=italic,size=size,color=color)

    def _sec_heading(doc, num, title):
        """Clean section heading — no background, just bold text with red accent."""
        p=doc.add_paragraph()
        p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(5)
        _run(p,f"{num}. ",bold=True,size=14,color="C0392B")
        _run(p,title,bold=True,size=14,color="1A1A2E")
        # thin red underline
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single')
        bot.set(qn('w:sz'),'4'); bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),'C0392B')
        pBdr.append(bot); pPr.append(pBdr)

    def _tbl_hdr_cell(tbl_row, col_idx, text, width_cm):
        """Clean light-blue header cell — no dark background."""
        _cell(tbl_row, col_idx, text, width_cm,
              bold=True, size=9, color="1A1A2E",
              fill="D6E4F0",   # light steel blue
              align=WD_ALIGN_PARAGRAPH.CENTER, brd_col="7FB3D3")

    def _msg_type(r):
        ai=r.get("acct",{})
        sdir=r.get("direction","")
        stype=r.get("sender_type","")
        # MFS/Bank SMS: body থেকে detected direction সবচেয়ে নির্ভরযোগ্য
        # কারণ MFS notification সব সময় SMSMT হয় — operator নিজে পাঠায়
        if stype in ("mfs","bank") and ai.get("direction"):
            return "IN" if ai["direction"]=="Received" else "OUT"
        # Personal SMS বা CDR direction স্পষ্ট
        if sdir=="SMSMT" and not ai.get("direction"):  return "IN"
        if sdir=="SMSMO" and not ai.get("direction"):  return "OUT"
        if ai.get("direction")=="Received": return "IN"
        if ai.get("direction")=="Sent":     return "OUT"
        if sdir=="SMSMT": return "IN"
        if sdir=="SMSMO": return "OUT"
        return "—"

    def _remarks(r):
        ai=r.get("acct",{}); rems=[]
        if ai.get("service"):  rems.append(ai["service"])
        if r["amount"]>=10000: rems.append(f"Tk {r['amount']:,.0f}")
        if ai.get("trx_id"):   rems.append(f"TrxID: {ai['trx_id']}")
        if ai.get("balance"):  rems.append(f"Bal: {ai['balance']:,.0f}")
        cats=[c for c in r.get("categories",[]) if c!="bank_info"]
        cat_s=", ".join(CAT_LABEL.get(c,c) for c in cats)
        if cat_s: rems.append(cat_s)
        return " | ".join(rems)

    def _write_sms_table(doc, msg_list):
        """Write SMS messages table with given list."""
        if not msg_list:
            p=doc.add_paragraph(); _run(p,"No records found.",italic=True,size=9,color="888888")
            return
        mH=["SL","B Party (Sender)","Date / Time","Type (In/Out)","Full Message (Body)","Remarks"]
        mW=[0.8, 2.5, 2.5, 1.4, 8.0, 2.8]   # total ~18cm
        mt=doc.add_table(rows=1+len(msg_list),cols=6)
        mt.alignment=WD_TABLE_ALIGNMENT.LEFT
        for j,(h,w) in enumerate(zip(mH,mW)):
            _tbl_hdr_cell(mt.rows[0],j,h,w)
        for idx,r in enumerate(msg_list):
            mtype=_msg_type(r)
            t2=r.get("tier","")
            rbg={"high":"FEF2F2","med":"FFF5F5","small":"FFFBF0"}.get(t2,"FFFFFF" if idx%2==0 else "F7F7F7")
            tbg="E8F5EE" if mtype=="IN" else ("FFF0F0" if mtype=="OUT" else rbg)
            tcol="1A7A4A" if mtype=="IN" else ("CC2200" if mtype=="OUT" else "666666")
            row=mt.rows[idx+1]
            data=[
                (str(idx+1),    "555555", WD_ALIGN_PARAGRAPH.CENTER, rbg,  True,  8),
                (r["sender"],   "1A4A8A", WD_ALIGN_PARAGRAPH.LEFT,   rbg,  True,  8),
                (str(r.get("date",""))[:16], "444444", WD_ALIGN_PARAGRAPH.LEFT, rbg, False, 7.5),
                (mtype,         tcol,     WD_ALIGN_PARAGRAPH.CENTER, tbg,  True,  9),
                (r["body"],     "111111", WD_ALIGN_PARAGRAPH.LEFT,   rbg,  False, 8.5),
                (_remarks(r),   "333333", WD_ALIGN_PARAGRAPH.LEFT,   rbg,  False, 8),
            ]
            for j,(v,c,a,bg,bd,fs) in enumerate(data):
                _cell(row,j,v,mW[j],bold=bd,size=fs,color=c,fill=bg,align=a)

    # ── Build doc ──────────────────────────────────────────────
    doc=WDoc()
    sec=doc.sections[0]
    sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.left_margin=sec.right_margin=Cm(1.5)
    sec.top_margin=sec.bottom_margin=Cm(1.8)

    # Header
    hp=sec.header.paragraphs[0]; hp.clear()
    _run(hp,"SMS INTELLIGENCE REPORT",bold=True,size=9,color="1A1A2E")
    _run(hp,f"   |   Target: {tn}" if tn else "",size=9,color="555555")
    _run(hp,"   |   CONFIDENTIAL",bold=True,size=9,color="C0392B")
    hp.paragraph_format.space_after=Pt(2)
    pPr=hp._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
    bot2=OxmlElement('w:bottom'); bot2.set(qn('w:val'),'single')
    bot2.set(qn('w:sz'),'6'); bot2.set(qn('w:space'),'1'); bot2.set(qn('w:color'),'C0392B')
    pBdr.append(bot2); pPr.append(pBdr)

    # Footer
    fp=sec.footer.paragraphs[0]; fp.clear()
    _run(fp,"CDR Forensic Analysis  |  Bangladesh Law Enforcement  |  CONFIDENTIAL",size=8,color="888888")
    fp.paragraph_format.space_before=Pt(2)
    pPr2=fp._p.get_or_add_pPr(); pBdr2=OxmlElement('w:pBdr')
    top3=OxmlElement('w:top'); top3.set(qn('w:val'),'single')
    top3.set(qn('w:sz'),'6'); top3.set(qn('w:space'),'1'); top3.set(qn('w:color'),'C0392B')
    pBdr2.append(top3); pPr2.append(pBdr2)

    from datetime import datetime as _dt
    now_str=_dt.now().strftime("%d %B %Y, %I:%M %p")
    disp=[r for r in results if "bank_info" not in r.get("categories",[])]

    # ── Cover ──────────────────────────────────────────────────
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(20); p.paragraph_format.space_after=Pt(6)
    _run(p,"SMS INTELLIGENCE REPORT",bold=True,size=26,color="1A1A2E")
    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after=Pt(16)
    _run(p2,"Forensic SMS Analysis Dashboard — CDR Intelligence",italic=True,size=12,color="555555")

    info_rows=[
        ("Target Number",       tn or "All Numbers"),
        ("Total SMS in File",   str(stats["total"])),
        ("After Dedup / Merge", str(stats["after_merge"])),
        ("Operator Filtered",   str(stats["filtered"])),
        ("Identified / Shown",  str(len(disp))),
        ("Report Generated",    now_str),
    ]
    t0=doc.add_table(rows=len(info_rows),cols=2); t0.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,(lbl,val) in enumerate(info_rows):
        bg="EEF2FF" if i%2==0 else "FFFFFF"
        _cell(t0.rows[i],0,lbl,5,bold=True,size=10,color="1A1A2E",fill=bg)
        _cell(t0.rows[i],1,val,6,bold=True,size=10,color="C0392B",fill=bg)
    doc.add_paragraph().paragraph_format.space_after=Pt(4)

    # ── Section 1: Bank/MFS Summary ────────────────────────────
    _sec_heading(doc,"1","MFS / Bank Transaction Summary")
    bank_summ=build_bank_summary(results)
    bH=["Bank / MFS Service","Txn Count","Total Received (IN)","Total Sent (OUT)","Max Transaction","Last Balance"]
    bW=[4.0,2.0,3.0,3.0,3.0,3.0]
    bt=doc.add_table(rows=1+len(bank_summ)+1,cols=6); bt.alignment=WD_TABLE_ALIGNMENT.LEFT
    for j,(h,w) in enumerate(zip(bH,bW)): _tbl_hdr_cell(bt.rows[0],j,h,w)
    fmt=lambda n: f"Tk {n:,.0f}" if n>0 else "—"
    ti=to=0
    for idx,(svc,d) in enumerate(bank_summ.items()):
        if d["txn_count"]==0 and d["last_balance"]==0: continue
        bg="FFFFFF" if idx%2==0 else "F5F5F5"
        row=bt.rows[idx+1]
        vals=[svc,str(d["txn_count"]),fmt(d["total_in"]),fmt(d["total_out"]),fmt(d["max_txn"]),fmt(d["last_balance"])]
        clrs=["1A1A2E","444444","1A7A4A","CC2200","7A5A00","1A4A8A"]
        alns=[WD_ALIGN_PARAGRAPH.LEFT,WD_ALIGN_PARAGRAPH.CENTER,
              WD_ALIGN_PARAGRAPH.RIGHT,WD_ALIGN_PARAGRAPH.RIGHT,
              WD_ALIGN_PARAGRAPH.RIGHT,WD_ALIGN_PARAGRAPH.RIGHT]
        for j in range(6): _cell(row,j,vals[j],bW[j],bold=(j==0),size=9,color=clrs[j],fill=bg,align=alns[j])
        ti+=d["total_in"]; to+=d["total_out"]
    # Total row — light green/red background instead of dark
    tr=bt.rows[-1]
    for j,(v,c,a,bg) in enumerate(zip(
        ["TOTAL (ALL SERVICES)","—",fmt(ti),fmt(to),"—","—"],
        ["1A1A2E","666666","1A7A4A","CC2200","666666","666666"],
        [WD_ALIGN_PARAGRAPH.LEFT,WD_ALIGN_PARAGRAPH.CENTER,
         WD_ALIGN_PARAGRAPH.RIGHT,WD_ALIGN_PARAGRAPH.RIGHT,
         WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_PARAGRAPH.CENTER],
        ["EEF2FF","EEF2FF","E8F5EE","FFF0F0","EEF2FF","EEF2FF"]
    )):
        _cell(tr,j,v,bW[j],bold=True,size=9,color=c,fill=bg,align=a,brd_col="AAAAAA")
    doc.add_paragraph().paragraph_format.space_after=Pt(2)

    # ── Section 2: Tier Breakdown ───────────────────────────────
    _sec_heading(doc,"2","Transaction Tier Breakdown")
    tiers=[("High (100K+)","txn_high","991B1B","FEF2F2"),
           ("Medium (50K-100K)","txn_med","DC2626","FFF5F5"),
           ("Small (10K-50K)","txn_small","EA580C","FFFBF0")]
    tH=["Tier","Total Count","IN (Received)","OUT (Sent)","Direction Unknown"]
    tW=[4.5,2.0,3.0,3.0,3.5]
    tt=doc.add_table(rows=1+len(tiers),cols=5); tt.alignment=WD_TABLE_ALIGNMENT.LEFT
    for j,(h,w) in enumerate(zip(tH,tW)): _tbl_hdr_cell(tt.rows[0],j,h,w)
    for idx,(lbl,key,col,bg) in enumerate(tiers):
        msgs=[r for r in results if key in r.get("categories",[])]
        ins =sum(1 for r in msgs if _msg_type(r)=="IN")
        outs=sum(1 for r in msgs if _msg_type(r)=="OUT")
        unk =max(len(msgs)-ins-outs,0)
        row=tt.rows[idx+1]
        for j,(v,c,a) in enumerate(zip(
            [lbl,str(len(msgs)),str(ins),str(outs),str(unk)],
            [col,col,"1A7A4A","CC2200","888888"],
            [WD_ALIGN_PARAGRAPH.LEFT]+[WD_ALIGN_PARAGRAPH.CENTER]*4
        )):
            _cell(row,j,v,tW[j],bold=(j==0),size=9,color=c,fill=bg,align=a)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)

    # ── Section 3: All SMS Messages ─────────────────────────────
    _sec_heading(doc,"3","Important SMS Messages — All")
    _write_sms_table(doc, disp)

    # ── Section 4: Target Number SMS (only if tn provided) ──────
    if tn:
        doc.add_paragraph().paragraph_format.space_after=Pt(6)
        _sec_heading(doc,"4",f"SMS Related to Target Number: {tn}")
        tn_clean=re.sub(r'[^\d]','',tn)
        tn_variants={tn_clean}
        if tn_clean.startswith('01'): tn_variants.add('880'+tn_clean)
        if tn_clean.startswith('880') and len(tn_clean)>=13: tn_variants.add(tn_clean[3:])
        def _is_tn(r):
            sd=re.sub(r'[^\d]','',str(r.get("sender","")))
            bd=r.get("body","")
            ai=r.get("acct",{})
            on=re.sub(r'[^\d]','',str(ai.get("other_number","")))
            return (any(v in sd for v in tn_variants)
                    or any(v in on for v in tn_variants)
                    or any(v in bd for v in tn_variants))
        tn_msgs=[r for r in disp if _is_tn(r)]
        note=doc.add_paragraph()
        note.paragraph_format.space_after=Pt(4)
        _run(note,f"Showing {len(tn_msgs)} SMS records related to {tn}  |  ",size=9,color="555555",italic=True)
        _run(note,"This section includes SMS where the target number appears as sender, recipient, or in message body.",
             size=9,color="888888",italic=True)
        _write_sms_table(doc,tn_msgs)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()


def process_file(input_path, html_out, docx_out=None, target_number=""):
    ext = os.path.splitext(input_path)[1].lower()
    raw = load_xlsx(input_path) if ext!='.csv' else load_csv_f(input_path)
    results, stats, tn = _filter_and_analyze(raw, target_number, input_path)
    generate_html(results, stats, html_out, tn)
    if docx_out:
        docx_bytes = generate_docx(results, stats, tn)
        with open(docx_out,'wb') as f: f.write(docx_bytes)
    tc = Counter(s["tier"] for s in results)
    display_count = sum(1 for s in results if "bank_info" not in s.get("categories",[]))
    return {"target_number":tn,"total_sms":len(raw),"after_merge":stats["after_merge"],
            "filtered":stats["filtered"],"total_shown":display_count,
            "high":tc.get("high",0),"med":tc.get("med",0),"small":tc.get("small",0),
            "_results":results,"_stats":stats,"_tn":tn}

def main():
    if len(sys.argv)<2: print("python3 sms_analyzer.py input.xlsx [output.html]"); sys.exit(1)
    inp=sys.argv[1]; out=sys.argv[2] if len(sys.argv)>2 else re.sub(r'\.(xlsx?|csv)$','.html',inp)
    r=process_file(inp,out)
    print(f"✅ Total: {r['total_shown']} | 🔴High:{r['high']} | 🟠Medium:{r['med']} | 🟡Small:{r['small']}")
    print(f"📄 {os.path.abspath(out)}")

if __name__=="__main__": main()
