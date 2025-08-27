# -*- coding: utf-8 -*-
import os
import sys
import json
import threading
import winshell
from pathlib import Path
from typing import Dict, List, Tuple, Generator
from win32com.client import Dispatch
import ctypes
from ctypes import wintypes
import win32con
import win32api
import win32gui
import time

# For system tray
import pystray
from PIL import Image, ImageDraw

# For keyboard hook
from pynput import keyboard

# --------------------------
# Khipro Mapping groups (exactly as provided)
# --------------------------

SHOR: Dict[str, str] = {
    "o": "অ", "oo": "ঽ",
    "fuf": "‌ু", "fuuf": "‌ূ", "fqf": "‌ৃ",
    "fa": "া", "a": "আ",
    "fi": "ি", "i": "ই",
    "fii": "ী", "ii": "ঈ",
    "fu": "ু", "u": "উ",
    "fuu": "ূ", "uu": "ঊ",
    "fq": "ৃ", "q": "ঋ",
    "fe": "ে", "e": "এ",
    "foi": "ৈ", "oi": "ঐ",
    "fw": "ো", "w": "ও",
    "fou": "ৌ", "ou": "ঔ",
    "fae": "্যা", "ae": "অ্যা",
    "wa": "ওয়া", "fwa": "োয়া",
    "wae": "ওয়্যা",
    "we": "ওয়ে", "fwe": "োয়ে",

    "ngo": "ঙ", "nga": "ঙা", "ngi": "ঙি", "ngii": "ঙী", "ngu": "ঙু",
    "nguff": "ঙ", "nguu": "ঙূ", "nguuff": "ঙ", "ngq": "ঙৃ", "nge": "ঙে",
    "ngoi": "ঙৈ", "ngw": "ঙো", "ngou": "ঙৌ", "ngae": "ঙ্যা",
}

BYANJON: Dict[str, str] = {
    "k": "ক", "kh": "খ", "g": "গ", "gh": "ঘ",
    "c": "च", "ch": "ছ", "j": "জ", "jh": "ঝ", "nff": "ঞ",
    "tf": "ট", "tff": "ঠ", "tfh": "ঠ", "df": "ড", "dff": "ঢ", "dfh": "ঢ", "nf": "ণ",
    "t": "ত", "th": "থ", "d": "দ", "dh": "ধ", "n": "ন",
    "p": "প", "ph": "ফ", "b": "ব", "v": "ভ", "m": "ম",
    "z": "য", "l": "ল", "sh": "শ", "sf": "ষ", "s": "স", "h": "হ",
    "y": "য়", "rf": "ড়", "rff": "ঢ়",
    ",,": "়",
}

JUKTOBORNO: Dict[str, str] = {
    "rz": "র‍্য",
    "kk": "ক্ক", "ktf": "ক্ট", "ktfr": "ক্ট্র", "kt": "ক্ত", "ktr": "ক্ত্র", "kb": "ক্ব", "km": "ক্ম", "kz": "ক্য", "kr": "ক্র", "kl": "ক্ল",
    "kf": "ক্ষ", "ksf": "ক্ষ", "kkh": "ক্ষ", "kfnf": "ক্ষ্ণ", "kfn": "ক্ষ্ণ", "ksfnf": "ক্ষ্ণ", "ksfn": "ক্ষ্ণ", "kkhn": "ক্ষ্ণ", "kkhnf": "ক্ষ্ণ",
    "kfb": "ক্ষ্ব", "ksfb": "ক্ষ্ব", "kkhb": "ক্ষ্ব", "kfm": "ক্ষ্ম", "kkhm": "ক্ষ্ম", "ksfm": "ক্ষ্ম", "kfz": "ক্ষ্য", "ksfz": "ক্ষ্য", "kkhz": "ক্ষ্য",
    "ks": "ক্স",
    "khz": "খ্য", "khr": "খ্র",
    "ggg": "গ্গ", "gnf": "গ্‌ণ", "gdh": "গ্ধ", "gdhz": "গ্ধ্য", "gdhr": "গ্ধ্র", "gn": "গ্ন", "gnz": "গ্ন্য", "gb": "গ্ব", "gm": "গ্ম", "gz": "গ্য", "gr": "গ্র", "grz": "গ্র্য", "gl": "গ্ল",
    "ghn": "ঘ্ন", "ghr": "ঘ্র",
    "ngk": "ঙ্ক", "ngkt": "ঙ্‌ক্ত", "ngkz": "ঙ্ক্য", "ngkr": "ঙ্ক্র", "ngkkh": "ঙ্ক্ষ", "ngksf": "ঙ্ক্ষ", "ngkh": "ঙ্খ", "ngg": "ঙ্গ", "nggz": "ঙ্গ্য", "nggh": "ঙ্ঘ", "ngghz": "ঙ্ঘ্য", "ngghr": "ঙ্ঘ্র", "ngm": "ঙ্ম",
    "cc": "চ্চ", "cch": "চ্ছ", "cchb": "চ্ছ্ব", "cchr": "চ্ছ্র", "cnff": "চ্ঞ", "cb": "চ্ব", "cz": "চ্য",
    "jj": "জ্জ", "jjb": "জ্জ্ব", "jjh": "জ্ঝ", "jnff": "জ্ঞ", "gg": "জ্ঞ", "jb": "জ্ব", "jz": "জ্য", "jr": "জ্র",
    "nc": "ঞ্চ", "nffc": "ঞ্চ", "nj": "ঞ্জ", "nffj": "ঞ্জ", "njh": "ঞ্ঝ", "nffjh": "ঞ্ঝ", "nch": "ঞ্ছ", "nffch": "ঞ্ছ",
    "ttf": "ট্ট", "tftf": "ট্ট", "tfb": "ট্ব", "tfm": "ट्म", "tfz": "ট্য", "tfr": "ট্র",
    "ddf": "ড্ড", "dfdf": "ড্ড", "dfb": "ড্ব", "dfz": "ড্য", "dfr": "ড্র", "rfg": "ড়্‌গ",
    "dffz": "ঢ্য", "dfhz": "ঢ্য", "dffr": "ঢ্র", "dfhr": "ढ्र",
    "nftf": "ণ্ট", "nftff": "ণ্ঠ", "nftfh": "ণ্ঠ", "nftffz": "ণ্ঠ্য", "nftfhz": "ণ্ঠ্য", "nfdf": "ণ্ড", "nfdfz": "ণ্ড্য", "nfdfr": "ণ্ড্র", "nfdff": "ণ্ঢ", "nfdfh": "ণ্ঢ", "nfnf": "ণ্ণ", "nfn": "ণ্ণ", "nfb": "ণ্ব", "nfm": "ণ্ম", "nfz": "ণ্য",
    "tt": "ত্ত", "ttb": "ত্ত্ব", "ttz": "ত্ত্য", "tth": "ত্থ", "tn": "ত্ন", "tb": "ত্ব", "tm": "ত্ম", "tmz": "ত্ম্য", "tz": "ত্য", "tr": "ত্র", "trz": "ত্র্য",
    "thb": "থ্ব", "thz": "থ্য", "thr": "থ্র",
    "dg": "দ্‌গ", "dgh": "দ্‌ঘ", "dd": "দ্দ", "ddb": "দ্দ্ব", "ddh": "দ্ধ", "db": "দ্ব", "dv": "দ্ভ", "dvr": "দ্ভ্র", "dm": "দ্ম", "dz": "দ্য", "dr": "দ্র", "drz": "দ্র্য",
    "dhn": "ধ্ন", "dhb": "ধ্ব", "dhm": "ধ্ম", "dhz": "ধ্য", "dhr": "ধ্র",
    "ntf": "ন্ট", "ntfr": "ন্ট্র", "ntff": "ন্ঠ", "ntfh": "ন্ঠ", "ndf": "ন্ড", "ndfr": "ন্ড্র", "nt": "ন্ত", "ntb": "ন্ত্ব", "ntr": "ন্ত্র", "ntrz": "ন্ত্র্য", "nth": "ন্থ", "nthr": "ন্থ্র", "nd": "ন্দ", "ndb": "ন্দ্ব", "ndz": "ন্দ্য",
    "ndr": "ন্দ্র", "ndh": "ন্ধ", "ndhz": "ন্ধ্য", "ndhr": "ন্ধ্র", "nn": "ন্ন", "nb": "ন্ব", "nm": "ন্ম", "nz": "ন্য", "ns": "ন্স",
    "ptf": "প্ট", "pt": "প্ত", "pn": "প্ন", "pp": "প্প", "pz": "প্য", "pr": "প্র", "pl": "প্ল", "ps": "প্স",
    "phr": "ফ্র", "phl": "ফ্ল",
    "bj": "ব্জ", "bd": "ব্দ", "bdh": "ব্ধ", "bb": "ব্ব", "bz": "ব্য", "br": "ব্র", "bl": "ব্ল", "vb": "ভ্ব", "vz": "ভ্য", "vr": "ভ্র", "vl": "ভ্ল",
    "mn": "ম্ন", "mp": "ম্প", "mpr": "ম্প্র", "mph": "ম্ফ", "mb": "ম্ব", "mbr": "म्ब्र", "mv": "ম্ভ", "mvr": "ম্ভ্র", "mm": "ম্ম", "mz": "ম্য", "mr": "ম্র", "ml": "ম্ল",
    "zz": "য্য",
    "lk": "ল্ক", "lkz": "ল্ক্য", "lg": "ল্গ", "ltf": "ল্ট", "ldf": "ল্ড", "lp": "ল্প", "lph": "ল্ফ", "lb": "ল্ব", "lv": "ল্‌ভ", "lm": "ল্ম", "lz": "ল্য", "ll": "ল্ল",
    "shc": "শ্চ", "shch": "শ্ছ", "shn": "শ্ন", "shb": "শ্ব", "shm": "শ্ম", "shz": "শ্য", "shr": "শ্র", "shl": "শ্ল",
    "sfk": "ষ্ক", "sfkr": "ষ্ক্র", "sftf": "ষ্ট", "sftfz": "ষ্ট্য", "sftfr": "ষ্ট্র", "sftff": "ষ্ঠ", "sftfh": "ষ্ঠ", "sftffz": "ষ্ঠ্য", "sftfhz": "ষ্ঠ্য", "sfnf": "ষ্ণ", "sfn": "ষ্ণ",
    "sfp": "ষ্প", "sfpr": "ষ্প্র", "sfph": "ষ্ফ", "sfb": "ষ্ব", "sfm": "ষ্ম", "sfz": "ষ্য",
    "sk": "স্ক", "skr": "স্ক্র", "skh": "স্খ", "stf": "স্ট", "stfr": "স্ট্র", "st": "স্ত", "stb": "স্ত্ব", "stz": "স্ত্য", "str": "স্ত্র", "sth": "স্থ", "sthz": "স্থ্য", "sn": "স্ন",
    "sp": "স্প", "spr": "স্প্র", "spl": "স্প্ল", "sph": "স্ফ", "sb": "স্ব", "sm": "স্ম", "sz": "স্য", "sr": "স্র", "sl": "স্ল",
    "hn": "হ্ন", "hnf": "হ্ণ", "hb": "হ্ব", "hm": "হ্ম", "hz": "হ্য", "hr": "হ্র", "hl": "হ্ল",

    # oshomvob juktoborno
    "ksh": "কশ", "nsh": "নश", "psh": "পশ", "ld": "লদ", "gd": "গদ", "ngkk": "ঙ্কক", "ngks": "ঙ্কস", "cn": "চন", "cnf": "চণ", "jn": "জন", "jnf": "জণ", "tft": "টত", "dfd": "ডদ",
    "nft": "ণত", "nfd": "ণদ", "lt": "লত", "sft": "ষত", "nfth": "ণথ", "nfdh": "ণধ", "sfth": "ষথ",
    "ktff": "কঠ", "ktfh": "কঠ", "ptff": "পঠ", "ptfh": "पঠ", "ltff": "लঠ", "ltfh": "লঠ", "stff": "सঠ", "stfh": "सঠ", "dfdff": "ডঢ", "dfdfh": "ডঢ", "ndff": "नढ", "ndfh": "नढ",
    "ktfrf": "क्टড়", "ktfrff": "क्टঢ়", "kth": "कथ", "ktrf": "क्तড়", "ktrff": "क्तঢ়", "krf": "कড়", "krff": "कঢ়", "khrf": "खড়", "khrff": "खঢ়", "gggh": "ज्ञঘ", "gdff": "গঢ", "gdfh": "গঢ", "gdhrf": "গ্ধড়",
    "gdhrff": "গ্ধঢ়", "grf": "গড়", "grff": "গঢ়", "ghrf": "ঘড়", "ghrff": "ঘঢ়", "ngkth": "ঙ্কথ", "ngkrf": "ঙ্কড়", "ngkrff": "ঙ্কঢ়", "ngghrf": "ঙ্ঘড়", "ngghrff": "ঙ্ঘঢ়", "cchrf": "চ্ছড়", "cchrff": "চ্ছঢ়",
    "tfrf": "টড়", "tfrff": "টঢ়", "dfrf": "ডড়", "dfrff": "ডঢ়", "rfgh": "ড়্‌ঘ", "dffrf": "ঢড়", "dfhrf": "ঢড়", "dffrff": "ঢঢ়", "dfhrff": "ঢঢ়", "nfdfrf": "ণ্ডড়", "nfdfrff": "ণ্ডঢ়", "trf": "তড়", "trff": "তঢ়", "thrf": "থড়", "thrff": "थढ়",
    "dvrf": "দ্ভড়", "dvrff": "দ্ভঢ়", "drf": "দড়", "drff": "দঢ়", "dhrf": "ধড়", "dhrff": "ধঢ়", "ntfrf": "ন্টড়", "ntfrff": "ন্টঢ়", "ndfrf": "ন্ডড়", "ndfrff": "ন্ডঢ়", "ntrf": "ন্তড়", "ntrff": "ন্তঢ়", "nthrf": "ন্থড়",
    "nthrff": "ন্থঢ়", "ndrf": "ন্দড়", "ndrff": "ন্দঢ়", "ndhrf": "ন্ধড়", "ndhrff": "ন্ধঢ়", "pth": "পথ", "pph": "পফ", "prf": "পড়", "prff": "পঢ়", "phrf": "ফড়", "phrff": "ফঢ়", "bjh": "বঝ", "brf": "বড়", "brff": "বঢ়",
    "vrf": "ভড়", "vrff": "ভঢ়", "mprf": "ম্পড়", "mprff": "ম্পঢ়", "mbrf": "ম্বড়", "mbrff": "ম্বঢ়", "mvrf": "ম্ভড়", "mvrff": "ম্ভঢ়", "mrf": "মড়", "mrff": "মঢ়", "lkh": "লখ", "lgh": "লঘ", "shrf": "শড়", "shrff": "শঢ়", "sfkh": "ষখ",
    "sfkrf": "ষ্কড়", "sfkrff": "ষ্কঢ়", "sftfrf": "ষ্টড়", "sftfrff": "ষ্টঢ়", "sfprf": "ষ্পড়", "sfprff": "ষ্পঢ়", "skrf": "স্কড়", "skrff": "স্কঢ়", "stfrf": "স্টড়", "stfrff": "স্টঢ়", "strf": "স্তড়", "strff": "স্তঢ়", "sprf": "স্পড়", "sprff": "স্পঢ়",
    "srf": "সড়", "srff": "সঢ়", "hrf": "হড়", "hrff": "হঢ়", "ldh": "লধ", "ngksh": "ঙ্কশ", "tfth": "টথ", "dfdh": "ডধ", "lth": "লথ",
}

REPH: Dict[str, str] = {
    "rr": "র্",
    "r": "র",
}

PHOLA: Dict[str, str] = {
    "r": "র",
    "z": "য",
}

KAR: Dict[str, str] = {
    "o": "", "of": "অ",
    "a": "া", "af": "আ",
    "i": "ि", "if": "ই",
    "ii": "ী", "iif": "ঈ",
    "u": "ু", "uf": "উ",
    "uu": "ূ", "uuf": "ঊ",
    "q": "ৃ", "qf": "ঋ",
    "e": "ে", "ef": "এ",
    "oi": "ৈ", "oif": "ই",
    "w": "ো", "wf": "ও",
    "ou": "ৌ", "ouf": "উ",
    "ae": "্যা", "aef": "অ্যা",
    "uff": "‌ু", "uuff": "‌ূ", "qff": "‌ৃ",
    "we": "োয়ে", "wef": "ওয়ে",
    "waf": "ওয়া", "wa": "োয়া",
    "wae": "ওয়্যা",
}

ONGKO: Dict[str, str] = {
    ".1": ".১", ".2": ".২", ".3": ".৩", ".4": ".৪", ".5": ".৫", ".6": ".৬", ".7": ".৭", ".8": ".৮", ".9": ".৯", ".0": ".০",
    "1": "১", "2": "২", "3": "৩", "4": "৪", "5": "५", "6": "৬", "7": "৭", "8": "৮", "9": "৯", "0": "০",
}

DIACRITIC: Dict[str, str] = {
    "qq": "্", "xx": "্‌", "t/": "ৎ", "x": "ঃ", "ng": "ং", "ngf": "ং", "/": "ঁ", "//": "/", "`": "‌", "``": "‍",
}

BIRAM: Dict[str, str] = {
    ".": "।", "...": "...", "..": ".", "$": "৳", "$f": "₹", ",,,": ",,", ".f": "॥", ".ff": "৺", "+f": "×", "-f": "÷",
}

PRITHAYOK: Dict[str, str] = {
    ";": "", ";;": ";",
}

AE: Dict[str, str] = {
    "ae": "‍্যা",
}

# --------------------------
# State machine configuration
# --------------------------

INIT = "init"
SHOR_STATE = "shor-state"
REPH_STATE = "reph-state"
BYANJON_STATE = "byanjon-state"

GROUP_MAPS: Dict[str, Dict[str, str]] = {
    "shor": SHOR,
    "byanjon": BYANJON,
    "juktoborno": JUKTOBORNO,
    "reph": REPH,
    "phola": PHOLA,
    "kar": KAR,
    "ongko": ONGKO,
    "diacritic": DIACRITIC,
    "biram": BIRAM,
    "prithayok": PRITHAYOK,
    "ae": AE,
}

# Group order per state (priority used when same-length matches)
STATE_GROUP_ORDER: Dict[str, List[str]] = {
    INIT: ["diacritic", "shor", "prithayok", "ongko", "biram", "reph", "juktoborno", "byanjon"],
    SHOR_STATE: ["diacritic", "shor", "biram", "prithayok", "ongko", "reph", "juktoborno", "byanjon"],
    REPH_STATE: ["prithayok", "ae", "juktoborno", "byanjon", "kar"],
    BYANJON_STATE: ["diacritic", "prithayok", "ongko", "biram", "kar", "juktoborno", "phola", "byanjon"],
}

# Precompute max key length per group for greedy matching
MAXLEN_PER_GROUP: Dict[str, int] = {g: (max((len(k) for k in m.keys()), default=0)) for g, m in GROUP_MAPS.items()}


def _find_longest(state: str, text: str, i: int) -> Tuple[str, str, str]:
    """Return (group, key, value) for the longest match allowed in current state. If none, return ("", "", "")."""
    allowed = STATE_GROUP_ORDER[state]
    # Determine the max lookahead we need
    maxlen = 0
    for g in allowed:
        maxlen = max(maxlen, MAXLEN_PER_GROUP[g])
    end = min(len(text), i + maxlen)
    best_group = ""
    best_key = ""
    best_val = ""
    best_len = 0

    # Try lengths from longest to shortest to implement greedy matching
    for L in range(end - i, 0, -1):
        chunk = text[i:i + L]
        # Check groups by priority
        for g in allowed:
            m = GROUP_MAPS[g]
            if chunk in m:
                # First match at this length wins due to priority order
                return (g, chunk, m[chunk])
    return ("", "", "")


def _apply_transition(state: str, group: str) -> str:
    """Return the next state after consuming a token of 'group' in 'state'."""
    if state == INIT:
        if group == "diacritic":
            return SHOR_STATE
        if group == "shor":
            return SHOR_STATE
        if group in ("prithayok",):
            return INIT
        if group in ("ongko", "biram"):
            return INIT
        if group == "reph":
            return REPH_STATE
        if group in ("juktoborno", "byanjon"):
            return BYANJON_STATE
        return state

    if state == SHOR_STATE:
        if group in ("diacritic", "shor"):
            return SHOR_STATE
        if group in ("biram", "prithayok", "ongko"):
            return INIT
        if group == "reph":
            return REPH_STATE
        if group in ("juktoborno", "byanjon"):
            return BYANJON_STATE
        return state

    if state == REPH_STATE:
        if group == "prithayok":
            return INIT
        if group == "ae":
            return SHOR_STATE
        if group in ("juktoborno", "byanjon"):
            return BYANJON_STATE
        if group == "kar":
            return SHOR_STATE
        return state

    if state == BYANJON_STATE:
        if group in ("diacritic", "kar"):
            return SHOR_STATE
        if group in ("prithayok", "ongko", "biram"):
            return INIT
        # juktoborno, phola, byanjon keep BYANJON_STATE
        return BYANJON_STATE

    return state


def convert(text: str) -> str:
    """Convert an ASCII input string to Bengali output using the bn-khipro state machine."""
    i = 0
    n = len(text)
    state = INIT
    out: List[str] = []

    while i < n:
        group, key, val = _find_longest(state, text, i)
        if not group:
            # No mapping: pass through this char and reset to INIT
            out.append(text[i])
            i += 1
            state = INIT
            continue

        # Special handling: PHOLA in BYANJON_STATE inserts virama before mapped char
        if state == BYANJON_STATE and group == "phola":
            out.append("্")
            out.append(val)
        else:
            out.append(val)

        i += len(key)
        state = _apply_transition(state, group)

    return "".join(out)


# --------------------------
# IME Implementation
# --------------------------

class KhiproIME:
    def __init__(self):
        self.bengali_mode = False
        self.buffer = ""
        self.listener = None
        self.tray_icon = None
        self.setup_tray_icon()
        
    def setup_tray_icon(self):
        """Create system tray icon with menu"""
        # Create an image for the icon
        image = Image.new('RGB', (64, 64), 'black')
        dc = ImageDraw.Draw(image)
        dc.rectangle((0, 0, 64, 64), fill='green' if self.bengali_mode else 'red')
        dc.text((10, 10), "BN" if self.bengali_mode else "EN", fill='white')
        
        # Create menu
        menu = pystray.Menu(
            pystray.MenuItem('Toggle Bengali/English (F12)', self.toggle_mode),
            pystray.MenuItem('Start on Boot', self.toggle_startup),
            pystray.MenuItem('Exit', self.exit_app)
        )
        
        self.tray_icon = pystray.Icon("khipro_ime", image, "Khipro IME", menu)
        
    def toggle_mode(self):
        """Toggle between Bengali and English mode"""
        self.bengali_mode = not self.bengali_mode
        self.update_tray_icon()
        
    def update_tray_icon(self):
        """Update the tray icon based on current mode"""
        image = Image.new('RGB', (64, 64), 'black')
        dc = ImageDraw.Draw(image)
        dc.rectangle((0, 0, 64, 64), fill='green' if self.bengali_mode else 'red')
        dc.text((10, 10), "BN" if self.bengali_mode else "EN", fill='white')
        
        if self.tray_icon:
            self.tray_icon.icon = image
            
    def toggle_startup(self):
        """Toggle whether to start on boot"""
        try:
            startup_dir = winshell.startup()
            shortcut_path = os.path.join(startup_dir, "KhiproIME.lnk")
            
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                print("Removed from startup")
            else:
                # Create shortcut
                target = sys.executable
                wDir = os.path.dirname(sys.executable)
                icon = sys.executable
                
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target
                shortcut.Arguments = f'"{sys.argv[0]}"'
                shortcut.WorkingDirectory = wDir
                shortcut.IconLocation = icon
                shortcut.save()
                print("Added to startup")
        except Exception as e:
            print(f"Error toggling startup: {e}")
            
    def exit_app(self):
        """Exit the application"""
        if self.listener:
            self.listener.stop()
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)
        
    def on_press(self, key):
        """Handle key press events"""
        try:
            if key == keyboard.Key.f12:
                self.toggle_mode()
                return False  # Don't propagate F12
                
            if not self.bengali_mode:
                return True  # Pass through if not in Bengali mode
                
            # Handle special keys
            if key == keyboard.Key.space:
                self.flush_buffer()
                return True
            elif key == keyboard.Key.backspace:
                if self.buffer:
                    self.buffer = self.buffer[:-1]
                return True
            elif key == keyboard.Key.enter:
                self.flush_buffer()
                return True
                
            # Handle character keys
            if hasattr(key, 'char') and key.char:
                self.buffer += key.char
                converted = convert(self.buffer)
                
                # Simulate backspacing and typing the converted text
                self.simulate_backspace(len(self.buffer))
                self.simulate_type(converted)
                
                # If the conversion consumed the whole buffer, clear it
                if converted:
                    self.buffer = ""
                    
                return False  # Suppress the original key
                
        except Exception as e:
            print(f"Error in key handler: {e}")
            
        return True
        
    def flush_buffer(self):
        """Flush the current buffer"""
        if self.buffer:
            converted = convert(self.buffer)
            self.simulate_backspace(len(self.buffer))
            self.simulate_type(converted)
            self.buffer = ""
            
    def simulate_backspace(self, count):
        """Simulate pressing backspace count times"""
        for _ in range(count):
            win32api.keybd_event(win32con.VK_BACK, 0, 0, 0)
            win32api.keybd_event(win32con.VK_BACK, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.01)
            
    def simulate_type(self, text):
        """Simulate typing the given text"""
        for char in text:
            # We need to use the unicode version of keybd_event
            win32api.keybd_event(0, 0, 0, 0)  # dummy call
            # This is a simplified approach - might need more complex handling for all characters
            for key_code in self.char_to_keycode(char):
                win32api.keybd_event(key_code, 0, 0, 0)
                win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.01)
            
    def char_to_keycode(self, char):
        """Convert a character to virtual key codes (simplified)"""
        # This is a very simplified version - would need proper mapping
        # For Bengali characters, we might need to use unicode input methods
        try:
            # Try to use the char directly
            vk = win32api.VkKeyScan(char)
            return [vk & 0xff]  # Return the virtual key code
        except:
            return []
            
    def start(self):
        """Start the IME"""
        # Start the tray icon in a separate thread
        tray_thread = threading.Thread(target=self.tray_icon.run)
        tray_thread.daemon = True
        tray_thread.start()
        
        # Start the keyboard listener
        with keyboard.Listener(on_press=self.on_press) as listener:
            self.listener = listener
            listener.join()


# --------------------------
# Main application
# --------------------------

if __name__ == "__main__":
    ime = KhiproIME()
    ime.start()
