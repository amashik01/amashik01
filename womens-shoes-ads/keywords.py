"""Generates 200+ Bangla / English / Banglish keywords for women's footwear (BD)."""
import itertools

# (english, banglish, bangla) product terms. Only footwear a BD factory can make.
PRODUCTS = [
    ("ladies shoes", "ladies jutar", "মহিলাদের জুতা"),
    ("women shoes", "mohilader juta", "মেয়েদের জুতা"),
    ("ladies sandal", "ladies sandel", "লেডিস স্যান্ডেল"),
    ("ladies slipper", "ladies slipar", "লেডিস স্লিপার"),
    ("high heel", "hil juta", "হিল জুতা"),
    ("block heel", "block hil", "ব্লক হিল"),
    ("pencil heel", "pencil hil", "পেন্সিল হিল"),
    ("platform heel", "platform hil", "প্ল্যাটফর্ম হিল"),
    ("wedge heel", "wedge hil", "ওয়েজ হিল"),
    ("flat sandal", "flat sandel", "ফ্ল্যাট স্যান্ডেল"),
    ("ballerina flat", "ballerina juta", "বেলেরিনা জুতা"),
    ("pump shoes", "pump shoe", "পাম্প শু"),
    ("loafer", "ladies loafer", "লেডিস লোফার"),
    ("kolhapuri", "kolhapuri sandel", "কোলাপুরি"),
    ("khussa", "khussa juta", "খুসসা"),
    ("nagra", "nagra juta", "নাগরা জুতা"),
    ("bridal shoes", "biyer juta", "বিয়ের জুতা"),
    ("party shoes", "party juta", "পার্টি জুতা"),
    ("wedding heel", "biye hil", "বিয়ের হিল"),
    ("office shoes ladies", "office juta", "অফিস জুতা"),
    ("casual sandal", "casual sandel", "ক্যাজুয়াল স্যান্ডেল"),
    ("comfort sandal", "aramdayok sandel", "আরামদায়ক স্যান্ডেল"),
    ("medicated slipper", "medicated slipar", "মেডিকেটেড স্লিপার"),
    ("home slipper", "ghorer slipar", "ঘরের স্লিপার"),
    ("platform slipper", "platform slipar", "প্ল্যাটফর্ম স্লিপার"),
    ("slide sandal", "slide sandel", "স্লাইড স্যান্ডেল"),
    ("jelly shoes", "jelly juta", "জেলি জুতা"),
    ("leather sandal", "chamrar sandel", "চামড়ার স্যান্ডেল"),
    ("leather shoes ladies", "chamrar juta", "চামড়ার জুতা"),
    ("ankle strap heel", "ankle strap hil", "অ্যাংকেল স্ট্র্যাপ হিল"),
    ("embroidered shoes", "embroidery juta", "এমব্রয়ডারি জুতা"),
    ("pearl sandal", "pearl sandel", "পার্ল স্যান্ডেল"),
    ("stone work heel", "stone work hil", "স্টোন ওয়ার্ক হিল"),
    ("boots ladies", "ladies boot", "লেডিস বুট"),
    ("ladies chappal", "ladies chappal", "লেডিস চপ্পল"),
    ("fancy sandal", "fancy sandel", "ফ্যান্সি স্যান্ডেল"),
    ("eid collection shoes", "eid juta collection", "ঈদ কালেকশন জুতা"),
    ("girls shoes", "meyeder juta", "গার্লস শু"),
    ("handmade shoes", "handmade juta", "হ্যান্ডমেড জুতা"),
    ("peep toe", "peep toe hil", "পিপ টো"),
]
MODIFIERS = [
    "price in bangladesh", "bd", "new collection", "cash on delivery",
    "order now", "offer", "discount", "online shop bd",
]
BN_MODIFIERS = ["দাম", "অর্ডার করুন", "ক্যাশ অন ডেলিভারি", "নতুন কালেকশন", "অফার"]

def build():
    kws = []
    for en, bl, bn in PRODUCTS:
        kws += [en, bl, bn]
    for en, bl, bn in PRODUCTS[:20]:
        for m in MODIFIERS[:4]:
            kws.append(f"{en} {m}")
        kws.append(f"{bl} dam")
        for m in BN_MODIFIERS[:3]:
            kws.append(f"{bn} {m}")
    for m in ["shoe bd", "juta bd", "jutar dokan", "জুতার দোকান", "জুতা কালেকশন",
              "ladies footwear bangladesh", "footwear brand bd", "shoe shop dhaka",
              "juta shop dhaka", "women footwear factory price", "wholesale ladies shoes bd",
              "jutar pikar", "ladies juta cod", "বিয়ের জুতা কালেকশন", "মেয়েদের স্যান্ডেল"]:
        kws.append(m)
    seen, out = set(), []
    for k in kws:
        if k.lower() not in seen:
            seen.add(k.lower()); out.append(k)
    return out

if __name__ == "__main__":
    k = build()
    print(len(k))
