#!/usr/bin/env python3
"""
MigrationHunter — Config Loader
بارگذاری تنظیمات داینامیک از config.json
هیچ نام یا اطلاعات شخصی hardcoded نیست
"""
import os, json

BASE = os.path.dirname(os.path.abspath(__file__))

_config_cache = None

def load_config():
    """بارگذاری config.json — cached"""
    global _config_cache
    if _config_cache is not None:
        return _config_cache
    fp = os.path.join(BASE, "config.json")
    if not os.path.exists(fp):
        _config_cache = {"applicants": []}
        return _config_cache
    with open(fp, "r", encoding="utf-8") as f:
        _config_cache = json.load(f)
    return _config_cache

def get_applicants():
    """لیست متقاضیان"""
    return load_config().get("applicants", [])

def get_applicant_by_id(app_id):
    """پیدا کردن متقاضی با id"""
    app_id_lower = app_id.lower()
    for a in get_applicants():
        if a["id"] == app_id_lower:
            return a
    return None

def get_applicant_label(app_id):
    """برچسب نمایشی — مثلاً '👨 توحید' یا '👩 ندا'"""
    a = get_applicant_by_id(app_id)
    if a:
        return f"{a.get('emoji', '?')} {a.get('name_fa', app_id)}"
    return app_id

def get_applicant_name(app_id):
    """نام فارسی"""
    a = get_applicant_by_id(app_id)
    return a.get("name_fa", app_id) if a else app_id

def get_applicant_emoji(app_id):
    """ایموجی"""
    a = get_applicant_by_id(app_id)
    return a.get("emoji", "?") if a else "?"

def get_applicant_keywords(app_id):
    """کلمات کلیدی برای تشخیص ایمیل"""
    a = get_applicant_by_id(app_id)
    return a.get("keywords", []) if a else []

def get_all_applicant_labels():
    """دیکشنری id → label برای نمایش"""
    result = {}
    for a in get_applicants():
        result[a["id"].upper()] = f"{a.get('emoji', '?')} {a.get('name_fa', a['id'])}"
    return result

def get_applicant_colors():
    """دیکشنری id → رنگ برای Excel"""
    colors = {}
    for i, a in enumerate(get_applicants()):
        colors[a["id"].upper()] = "8E44AD" if i == 0 else "2E86C1"
    return colors

def detect_applicant_from_text(text):
    """تشخیص متقاضی از متن ایمیل — داینامیک از config"""
    text_lower = text.lower()
    scores = {}
    for a in get_applicants():
        score = sum(1 for k in a.get("keywords", []) if k.lower() in text_lower)
        scores[a["id"].upper()] = score
    if not scores:
        return "UNKNOWN"
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "UNKNOWN"
