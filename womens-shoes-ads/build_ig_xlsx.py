"""data/ig_ads.jsonl -> pinkwear-womens-shoes-instagram.xlsx (compares with the Phase-1 Facebook workbook if given)
usage: python build_ig_xlsx.py [path/to/pinkwear-womens-shoes-competitors-FULL.xlsx]"""
import json, re, sys, datetime as dt
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TODAY = dt.date(2026, 10, 3)
ads = [json.loads(l) for l in open("data/ig_ads.jsonl", encoding="utf-8")]
FB = sys.argv[1] if len(sys.argv) > 1 else None
fb_pages = {}
if FB:
    ws = openpyxl.load_workbook(FB, read_only=True)["Competitors"]
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[1]: fb_pages[str(r[1]).strip().lower()] = r[3]

SHOE = re.compile(r"shoes?\b|\bheels?\b|sandals?\b|slippers?\b|\bflats?\b|\bpumps?\b|loafers?\b|\bboots?\b|nagra|kolhapuri|khussa|footwear|jutar?|জুতা|জুতো|স্যান্ডেল|স্যান্ডাল|হিল|স্লিপার|চপ্পল|নাগরা|বুট|লোফার|ফ্ল্যাট|\bশু\b|কোলাপুরি", re.I)
WOMEN = re.compile(r"ladies|lady|women|woman|female|bridal|ballerina|wedge|heel|মহিলা|মেয়েদের|লেডি|আপু|হিল|ব্রাইডাল|ওয়েজ|ladys", re.I)
EXCL = [("বাচ্চাদের", r"kids|baby|boys?\b|children|শিশু|বাচ্চা|কিডস"), ("ছেলেদের", r"\bmen'?s\b|\bgents?\b|\bmale\b|পুরুষ|ছেলেদের|জেন্টস|punjabi|পাঞ্জাবি|formal shoe for men"),
        ("স্নিকার্স/স্পোর্টস", r"sneaker|running|sports? shoe|স্নিকার|স্পোর্টস"), ("ইনসোল/এক্সেসরিজ", r"insole|ইনসোল|rack|র.?্যাক|polish|পালিশ"),
        ("পাইকারি/বিদেশি", r"wholesale|manufacturer|supplier|factory price")]
BD = re.compile(r"[ঀ-৿]|৳|\btk\b|taka|bdt|\bbd\b|bangladesh|dhaka|cod\b|cash on delivery|inbox|messenger", re.I)
PROD = [("হিল (সব)", r"heel|হিল"), ("ব্লক হিল", r"block|box heel|ব্লক"), ("ওয়েজ", r"wedge|ওয়েজ"), ("স্যান্ডেল", r"sandal|স্যান্ডেল|স্যান্ডাল"),
        ("ফ্ল্যাট/পাম্প/ব্যালে", r"\bflats?\b|ballerina|pump|ফ্ল্যাট|ব্যালে|পাম্প"), ("স্লিপার", r"slipper|slide|flip|স্লিপার|চপ্পল"), ("বুট", r"boot|বুট"),
        ("লোফার", r"loafer|লোফার"), ("নাগরা/খুসসা/কোলাপুরি", r"nagra|khussa|kolhapuri|নাগরা|কোলাপুরি"), ("পার্টি/ব্রাইডাল", r"party|bridal|wedding|পার্টি|বিয়ে|ব্রাইডাল"),
        ("পিভিসি/জেলি", r"pvc|jelly|glassy|পিভিসি"), ("লেদার", r"leather|লেদার|চামড়া"), ("কমফোর্ট/ডক্টর", r"comfort|doctor|orthopedic|আরাম|ডক্টর")]
OFFER = [("COD / অগ্রিম নয়", r"cash on delivery|\bcod\b|ক্যাশ অন|অগ্রিম|advance"), ("ডিসকাউন্ট/অফার", r"discount|offer|% ?off|sale|ছাড়|অফার|সেল"),
         ("ফ্রি/কম ডেলিভারি", r"free delivery|ফ্রি ডেলিভারি|delivery free"), ("এক্সচেঞ্জ/রিটার্ন", r"exchange|return|এক্সচেঞ্জ|রিটার্ন"),
         ("ওয়ারেন্টি/গ্যারান্টি", r"warranty|guarantee|ওয়ারেন্টি|গ্যারান্টি"), ("রেডি স্টক", r"ready stock|রেডি স্টক"), ("প্রি-অর্ডার/চায়না", r"pre-?order|china|প্রি-?অর্ডার|চায়না"),
         ("দাম লেখা", r"৳|tk\.?\s*\d|taka|টাকা|bdt"), ("কম্বো", r"combo|কম্বো|2 pair|২ জোড়া"), ("সিজনাল", r"puja|eid|winter|পূজা|ঈদ|শীত")]
D = lambda s: dt.datetime.strptime(s, "%b %d, %Y").date() if s else None
days = lambda a: (TODAY - D(a["start"])).days if a.get("start") else ""
T = lambda a: f"{a['page']} {a['copy']}"
def num(s): return int(s.translate(str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")).replace(",", ""))
def prices(t): return [num(x) for x in re.findall(r"[৳]\s*([0-9০-৯,]{3,6})|(?:tk\.?|taka|টাকা|bdt)\s*([0-9০-৯,]{3,6})", t, re.I) for x in x if x] if False else \
    [num(g) for m in re.finditer(r"৳\s*([0-9০-৯,]{3,6})|(?:tk\.?|bdt)\s*([0-9০-৯,]{3,6})|([0-9০-৯,]{3,6})\s*(?:টাকা|taka|tk)", t, re.I) for g in m.groups() if g]

rel, excl = [], []
for a in ads:
    t = T(a)
    if not SHOE.search(a["copy"]): excl.append((a, "জুতা সংক্রান্ত নয়")); continue
    if re.search(r"সালোয়ার|কামিজ|শাড়ি|ফ্যাব্রিক|fabric|ডায়াবেটিস|ব্লাড সার্কুলেশন|diabetes|ভ্রমণ|\bO-MEN\b", t, re.I) and len(SHOE.findall(a["copy"])) < 2: excl.append((a, "অন্য পণ্য/সেবা")); continue
    bad = [n for n, rx in EXCL if re.search(rx, t, re.I)]
    if bad and not re.search(r"ladies|women|woman|মহিলা|লেডি|heel|হিল|মেয়েদের", a["copy"], re.I): excl.append((a, bad[0])); continue
    if re.search(r"baby|kids|শিশু|বাচ্চা|কিডস", a["copy"], re.I) and not re.search(r"ladies|women|woman|মহিলা|লেডি", a["copy"], re.I): excl.append((a, "বাচ্চাদের")); continue
    if not BD.search(t) and a["page"].strip().lower() not in fb_pages: excl.append((a, "বাংলাদেশি নয় (সম্ভবত বিদেশি)")); continue
    if not WOMEN.search(t) and re.search(EXCL[1][1], t, re.I): excl.append((a, "ছেলেদের")); continue
    rel.append(a)

pages = defaultdict(list)
for a in rel: pages[a["page"]].append(a)
pr = lambda L: sorted({p for a in L for p in prices(T(a)) if 300 <= p <= 6000})
ILL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
clean = lambda v: ILL.sub("", v) if isinstance(v, str) else v
H = PatternFill("solid", fgColor="1F3A5F")
wb = openpyxl.Workbook(); wb.remove(wb.active)
def sheet(name, head, data, widths):
    ws = wb.create_sheet(name); ws.append(head)
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = H; c.alignment = Alignment(wrap_text=True)
    for d in data: ws.append([clean(x) for x in d])
    for i, w in enumerate(widths): ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    for r in ws.iter_rows(min_row=2):
        for c in r: c.alignment = Alignment(wrap_text=True, vertical="top")

prod_cnt = {n: sum(1 for a in rel if re.search(rx, T(a), re.I)) for n, rx in PROD}
off_cnt = {n: sum(1 for a in rel if re.search(rx, T(a), re.I)) for n, rx in OFFER}
allp = [p for a in rel for p in prices(T(a)) if 300 <= p <= 6000]
med = sorted(allp)[len(allp) // 2] if allp else None
big = sorted(pages.items(), key=lambda kv: -len(kv[1]))
both = [p for p in pages if p.strip().lower() in fb_pages]
age = [days(a) for a in rel if days(a) != ""]
pc = lambda n: f"{n} ({n * 100 // max(1, len(rel))}%)"
top = lambda d, k=6: " · ".join(f"{n} {pc(c)}" for n, c in sorted(d.items(), key=lambda x: -x[1])[:k])

S = [
 ["Pinkwear — মহিলাদের জুতা: Instagram রিসার্চ (Meta Ad Library, Instagram প্ল্যাটফর্ম ফিল্টার) | তারিখ: ৩ অক্টোবর ২০২৬ · শুধু বাংলাদেশ · শুধু চালু অ্যাড"],
 [""], ["১. পরিধি"],
 [f"Meta Ad Library-তে প্ল্যাটফর্ম = Instagram ফিল্টার দিয়ে {len({k for a in ads for k in a.get('keywords', [])})}টি কিওয়ার্ডে (বাংলা, ইংরেজি, বাংলিশ) সার্চ। মোট {len(ads)}টি ইউনিক অ্যাড পাওয়া গেছে; এর মধ্যে {len(rel)}টি সত্যিই বাংলাদেশি মহিলাদের জুতার অ্যাড, এসেছে {len(pages)}টি পেজ থেকে। বাকি {len(excl)}টি (ধর্মীয়/বই/বিদেশি/ছেলেদের/বাচ্চাদের ইত্যাদি) 'Excluded' শিটে কারণসহ আছে।"],
 [""], ["২. এক নজরে"],
 [f"ধরনের হিসাব — {top(prod_cnt, 8)}"],
 [f"অফার/আস্থা — {top(off_cnt, 8)}"],
 [f"দাম — অ্যাডে লেখা দাম {len(allp)}টি; মাঝামাঝি (median) ৳{med}. রেঞ্জ ৳{min(allp) if allp else '-'}–{max(allp) if allp else '-'}। Pinkwear-এর ফোকাস ৳১০০০–২৫০০-এর মধ্যে পড়ে {sum(1 for p in allp if 1000 <= p <= 2500)}টি দাম।"],
 [f"কতদিন চলছে — ৭ দিনের কম {sum(1 for d in age if d < 7)} · ৭–৩০ দিন {sum(1 for d in age if 7 <= d <= 30)} · ৩১–৯০ দিন {sum(1 for d in age if 31 <= d <= 90)} · ৯০ দিনের বেশি {sum(1 for d in age if d > 90)}। ৯০+ দিনের অ্যাডগুলোই সবচেয়ে প্রমাণিত।"],
 [f"পেজ বিভাজন — ১টি অ্যাড: {sum(1 for L in pages.values() if len(L)==1)} পেজ · ২–৪টি: {sum(1 for L in pages.values() if 2<=len(L)<=4)} পেজ · ৫+টি: {sum(1 for L in pages.values() if len(L)>=5)} পেজ (আসল প্রতিযোগী)।"],
 [""], ["৩. বড় প্রতিযোগী (Instagram-এ সবচেয়ে বেশি অ্যাড)"],
 *[[f"{n} — {len(L)}টি অ্যাড · দাম: {('৳' + '–'.join(map(str, (pr(L)[0], pr(L)[-1])))) if pr(L) else 'লেখা নেই'} · Facebook-এও আছে: {'হ্যাঁ' if n.strip().lower() in fb_pages else 'না/অজানা'}"] for n, L in big[:12]],
 [""], ["৪. Facebook বনাম Instagram"],
 [f"{len(both)}টি পেজ Facebook রিসার্চেও (১১৭ পেজ) ছিল, মানে ওরা Facebook + Instagram দুই জায়গায় চালায়। বাকি {len(pages)-len(both)}টি পেজ নতুন — এগুলো Facebook রিসার্চে আসেনি (তালিকা Accounts শিটে 'In FB research?' কলামে)।"],
 [""], ["৫. Pinkwear-এর জন্য সুপারিশ"],
 ["• Instagram-এ প্রতিযোগিতা কম (প্রাসঙ্গিক অ্যাড মাত্র সীমিত সংখ্যক) — Reels + Stories-এ পরিষ্কার দাম, ক্যাশ অন ডেলিভারি, রেডি স্টক ও সাইজ ৩৪–৪৩ দিয়ে আলাদা হোন।"],
 ["• Facebook-এর মতোই আস্থার লাইন (COD, এক্সচেঞ্জ, ওয়ারেন্টি) প্রথম লাইনে রাখুন; Instagram-এ ছোট, ভিজ্যুয়াল-প্রধান কপি।"],
 ["• ট্রাই-অন/হাঁটার Reels, ফ্যাক্টরি প্রসেস, আনবক্সিং — বেশিরভাগ প্রতিযোগী শুধু প্রোডাক্ট-ছবি দেয়।"],
 [""], ["৬. সীমাবদ্ধতা"],
 ["• Ad Library শুধু পেইড অ্যাড দেখায়; Instagram-এর অর্গানিক পোস্ট/Reels, ফলোয়ার, লাইক/ভিউ এখানে নেই (লগইন লাগে — বাইপাস করা হয়নি)।"],
 ["• Ad Library খরচ/রিচ দেয় না। সফলতা অনুমান হয়েছে অ্যাড কতদিন চলছে ও একাধিক ভার্সন দেখে।"],
 ["• প্ল্যাটফর্ম ফিল্টার Instagram — তবে একই অ্যাড Facebook-এও চলতে পারে। প্রতি সার্চে Facebook ~৩০–৬০টির বেশি অ্যাড দেখায় না, তাই বড় অ্যাডভার্টাইজারের সব অ্যাড হয়তো আসেনি।"],
 ["• প্রাসঙ্গিকতা অটো-ফিল্টারে বাছা (শব্দ ধরে) — কিছু ভুল বাদ/অন্তর্ভুক্ত থাকতে পারে; Excluded শিট দেখে মিলিয়ে নিন।"]]
ws = wb.create_sheet("সারসংক্ষেপ (বাংলা)")
for s in S: ws.append([clean(x) for x in s])
ws.column_dimensions["A"].width = 160
for r in ws.iter_rows():
    for c in r: c.alignment = Alignment(wrap_text=True, vertical="top")
for i, s in enumerate(S, 1):
    if s[0][:2] in ("১.", "২.", "৩.", "৪.", "৫.", "৬.") or i == 1: ws.cell(i, 1).font = Font(bold=True, size=12)

sheet("Accounts", ["#", "Page", "Instagram handle", "Ads found", "Oldest ad (days)", "Prices seen (৳)", "Main products", "Top offers", "In FB research?", "Sample ad link", "Sample hook"],
      [[i, n, next((a["instagram_handle"] for a in L if a["instagram_handle"]), ""), len(L), max([days(a) for a in L if days(a) != ""] or [""]),
        f"{pr(L)[0]}–{pr(L)[-1]}" if pr(L) else "", ", ".join(k for k, rx in PROD if any(re.search(rx, T(a), re.I) for a in L))[:80],
        ", ".join(k for k, rx in OFFER if any(re.search(rx, T(a), re.I) for a in L))[:80], "Yes" if n.strip().lower() in fb_pages else "No",
        max(L, key=lambda a: days(a) if days(a) != "" else -1)["ad_url"], L[0]["copy"][:120]] for i, (n, L) in enumerate(big, 1)], [5, 28, 18, 8, 10, 14, 34, 34, 10, 48, 60])
sheet("Ads", ["Page", "Ad link (permanent)", "Library ID", "Started", "Days running", "Multi-version", "Products", "Offers", "Prices (৳)", "Keywords", "Ad copy"],
      [[a["page"], a["ad_url"], a["ad_id"], a["start"], days(a), "Yes" if a.get("multi_version") else "", ", ".join(k for k, rx in PROD if re.search(rx, T(a), re.I)),
        ", ".join(k for k, rx in OFFER if re.search(rx, T(a), re.I)), ", ".join(map(str, sorted(set(prices(T(a)))))), "; ".join(a.get("keywords", []))[:120], a["copy"]]
       for a in sorted(rel, key=lambda a: (-(days(a) if days(a) != "" else 0)))], [26, 46, 18, 13, 9, 9, 26, 28, 12, 28, 80])
sheet("Trends", ["Metric", "Ads", "Share"], [[f"Product: {k}", v, f"{v * 100 // max(1, len(rel))}%"] for k, v in sorted(prod_cnt.items(), key=lambda x: -x[1])] +
      [[f"Offer: {k}", v, f"{v * 100 // max(1, len(rel))}%"] for k, v in sorted(off_cnt.items(), key=lambda x: -x[1])], [40, 10, 10])
kw_c = Counter(k for a in rel for k in a.get("keywords", []))
sheet("Keywords", ["Keyword", "Relevant ads found"], [[k, c] for k, c in kw_c.most_common()], [40, 18])
sheet("Excluded", ["Page", "Ad link", "Reason", "Copy (start)"], [[a["page"], a["ad_url"], why, a["copy"][:150]] for a, why in excl], [28, 46, 26, 80])
wb.save("pinkwear-womens-shoes-instagram.xlsx"); print("saved:", len(ads), "ads;", len(rel), "relevant;", len(pages), "pages;", len(both), "also in FB")
