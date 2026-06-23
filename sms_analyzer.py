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
    r'মিস করেছেন|কল মিস|missed call|miss call|কলটি মিস|কলটি গ্রহণ করতে পারবেন',
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

def is_filtered(sender, body):
    s  = str(sender).strip()
    sl = s.lower()
    st = stype(s, body)
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
    nums = re.findall(r'\b(01[3-9]\d{8})\b', body)
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
    elif re.search(r'\bcash\s*out\b|\bpurchase\b|\bpos\s+purchase\b'
                   r'|\bpayment\s+(?:of\s+)?tk.*to\b|\bpayment\s+of\b|\bpaid\b'
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
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
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
        if di is not None and di<len(r) and str(r[di]).strip().upper()=="SMSMT":
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
        "accounts": defaultdict(lambda:{"balance":0.0,"balance_date":"","txn_count":0})
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
        if amt > 0:
            d["txn_count"] += 1
            # txn type body থেকে detect — investment/internal হলে আলাদা bucket
            body_lower = s.get("body","").lower()
            is_internal = any(t in body_lower for t in INTERNAL_TXN_TYPES)
            if is_internal:
                d["total_cycling"] = d.get("total_cycling", 0.0) + amt
            elif "Received" in (ai.get("direction","") or ""):
                d["total_in"]  += amt
            else:
                d["total_out"] += amt
            if amt > d["max_txn"]: d["max_txn"] = amt

        # per-account balance tracking
        acc_no = ai.get("account","")
        sms_date = str(s.get("date",""))
        bal = ai.get("balance",0) or 0
        if bal > 0:
            # account-level last balance (তারিখ অনুযায়ী)
            if acc_no:
                acc_entry = d["accounts"][acc_no]
                acc_entry["txn_count"] += 1
                if sms_date > acc_entry["balance_date"]:
                    acc_entry["balance"]      = bal
                    acc_entry["balance_date"] = sms_date
            # service-level last balance — সব account-এর সর্বশেষ balance এর সমষ্টি নয়,
            # বরং সর্বশেষ date-এর balance (single-account case-এ সঠিক থাকে)
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
            "accounts": {k: {"balance": v["balance"], "txn_count": v["txn_count"]}
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


def generate_html(sms_list, stats, out_path, tn=""):
    now = datetime.now().strftime("%d %B %Y, %I:%M %p")

    # bank_info — sidebar summary-র জন্য, card-এ দেখাবে না
    # display_list-এ শুধু 10K+ transaction এবং অন্য categories (personal, travel ইত্যাদি)
    # foreign_bank SMS → সর্বোচ্চ 20টি card, বাকিগুলো summary block-এ
    _non_ftxn = [s for s in sms_list
                 if "bank_info" not in s.get("categories",[])
                 and "travel_txn" not in s.get("categories",[])
                 and not (s.get("categories") == [f"txn_{s.get('tier','')}"]
                          and s.get("tier","") == "")]
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
        elif amt>=TIER_SMALL:
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
            amt_tag=(f'<span class="amt-tag" style="background:{bdr}">💰 Tk {s["amount"]:,.2f}</span>'
                     if s["amount"]>=TIER_SMALL else "")
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

        # Single account: last balance দেখাও
        # Multi account: প্রতিটির balance আলাদা দেখাও
        if is_multi and accs:
            acc_rows = ""
            for acc_no, av in sorted(accs.items()):
                if av["balance"] > 0:
                    acc_rows += (f'<div class="acc-row">'
                                 f'<span class="acc-no">AC {esc(acc_no)}</span>'
                                 f'<span class="acc-bal">BDT {av["balance"]:,.2f}</span>'
                                 f'</div>')
            bal_html = (f'<div class="ag-bal-multi">'
                        f'<div class="ag-bal-hdr">💰 Account Balances</div>'
                        f'{acc_rows}'
                        f'</div>') if acc_rows else ""
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

<!-- Undo bar -->
<div id="undo-bar" class="undo-bar no-print">
  <span>🗑 <span id="del-cnt" class="del-count">0</span> card(s) deleted</span>
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
  if(id){{ document.getElementById(id)?.classList.remove('deleted'); }}
  document.getElementById('del-cnt').textContent=_delStack.length;
  if(!_delStack.length) _undoBar.classList.remove('show');
  apply();
}}
function clearAll(){{ _delStack=[]; _undoBar.classList.remove('show'); }}
</script>
</body></html>'''
    with open(out_path,'w',encoding='utf-8') as f: f.write(html)

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

    stats = {"total":len(raw),"after_merge":len(merged_),"filtered":filtered_cnt}
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
