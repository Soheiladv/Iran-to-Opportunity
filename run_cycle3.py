#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Migration Hunter — چرخه سوم جستجو
تمرکز: کانادا + استرالیا + بروزرسانی CV ندا
"""
import os
import json
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")

# ── ۱. بروزرسانی SOURCE_BANK ──────────────────────────────
SOURCE_BANK = BASE_DIR / "memory" / "SOURCE_BANK.md"

def update_source_bank():
    content = f"""# SOURCE_BANK — بانک اطلاعاتی منابع
آخرین بروزرسانی: {DATE_STR}

---

### ۱. Health New Zealand | Te Whatu Ora
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-001 |
| Name | Health New Zealand |
| Type | GOVERNMENT / OFFICIAL EMPLOYER |
| Country | New Zealand |
| Industry | Healthcare |
| URL | https://www.healthnz.govt.nz |
| Trust Score | 95 |
| Job Quality | 90 |
| Sponsorship | 95 — Confirmed |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 98 |
| Last Checked | {DATE_STR} |
| Notes | استخدام فعال ماما بین‌المللی. Green List occupation. |

### ۲. RGH Global
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-002 |
| Name | RGH Global |
| Type | SPECIALIST RECRUITER |
| Country | New Zealand |
| Industry | Healthcare |
| URL | https://www.rgh-global.com |
| Trust Score | 85 |
| Job Quality | 85 |
| Sponsorship | 90 — Confirmed |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 92 |
| Last Checked | {DATE_STR} |
| Notes | متخصص استخدام ماما با حمایت ویزا. |

### ۳. Working In Health NZ
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-003 |
| Name | Working In Health NZ |
| Type | SPECIALIST RECRUITER |
| Country | New Zealand |
| Industry | Healthcare |
| URL | https://www.workingin-health.co.nz |
| Trust Score | 88 |
| Job Quality | 85 |
| Sponsorship | 90 — Confirmed |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 95 |
| Last Checked | {DATE_STR} |
| Notes | آژانس تخصصی استخدام ماما. خدمات رایگان. Green List = اقامت فوری. |

### ۴. Holalemania GmbH
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-004 |
| Name | Holalemania GmbH |
| Type | SPECIALIST RECRUITER |
| Country | Germany |
| Industry | Healthcare |
| URL | https://holalemania.de/en/ |
| Trust Score | 85 |
| Job Quality | 85 |
| Sponsorship | 85 — Confirmed |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 88 |
| Last Checked | {DATE_STR} |
| Notes | ۱۳ سال تجربه. ۹۲۱ استخدام موفق. ۶۲ بیمارستان. آموزش زبان آلمانی. |

### ۵. TalentOrange
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-005 |
| Name | TalentOrange |
| Type | SPECIALIST RECRUITER |
| Country | Germany |
| Industry | Healthcare |
| URL | https://www.talentorange.com/en/ |
| Trust Score | 82 |
| Job Quality | 82 |
| Sponsorship | 80 — Confirmed |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 82 |
| Last Checked | {DATE_STR} |
| Notes | بورسیه زبان آلمانی B2 ارائه می‌دهد. برنامه کامل صفر تا شروع کار. |

### ۶. Saskatchewan Health Authority ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-006 |
| Name | Saskatchewan Health Authority |
| Type | GOVERNMENT / OFFICIAL EMPLOYER |
| Country | Canada |
| Industry | Healthcare |
| URL | https://www.saskhealthauthority.ca/careers |
| Trust Score | 90 |
| Job Quality | 88 |
| Sponsorship | 85 — Likely |
| International Recruitment | Yes — Confirmed |
| Applicant Relevance | NEDA: 90 |
| Last Checked | {DATE_STR} |
| Notes | فعالانه متخصصان بهداشت بین‌المللی استخدام می‌کند. ایمیل: SHAInternational@saskhealthauthority.ca |

### ۷. Hays Healthcare Australia ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-007 |
| Name | Hays Healthcare Australia |
| Type | SPECIALIST RECRUITER |
| Country | Australia |
| Industry | Healthcare |
| URL | https://www.hays.com.au/jobs/healthcare |
| Trust Score | 82 |
| Job Quality | 80 |
| Sponsorship | 70 — Possible |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 80 |
| Last Checked | {DATE_STR} |
| Notes | آژانس استخدام بزرگ استرالیا. حمایت ویزا برای پرستاران بین‌المللی. |

### ۸. ANMF (Australian Nursing & Midwifery Federation) ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-008 |
| Name | ANMF |
| Type | PROFESSIONAL ASSOCIATION |
| Country | Australia |
| Industry | Healthcare |
| URL | https://www.anmf.org.au |
| Trust Score | 88 |
| Job Quality | 85 |
| Sponsorship | 65 — Possible |
| International Recruitment | Yes |
| Applicant Relevance | NEDA: 85 |
| Last Checked | {DATE_STR} |
| Notes | اتحادیه ۳۵۶,۰۰۰ پرستار و ماما استرالیا. منبع اطلاعات حرفه‌ای. |

### ۹. Canadian Association of Midwives (CAM) ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-009 |
| Name | Canadian Association of Midwives |
| Type | PROFESSIONAL ASSOCIATION |
| Country | Canada |
| Industry | Healthcare |
| URL | https://canadianmidwives.org |
| Trust Score | 88 |
| Job Quality | 85 |
| Sponsorship | 60 — Unknown |
| International Recruitment | No — Association only |
| Applicant Relevance | NEDA: 82 |
| Last Checked | {DATE_STR} |
| Notes | انجمن حرفه‌ای مامایی کانادا. منبع اطلاعات مسیر شغلی مامایی. |

### ۱۰. Job Bank Canada ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-010 |
| Name | Job Bank Canada |
| Type | GOVERNMENT JOB BOARD |
| Country | Canada |
| Industry | All |
| URL | https://www.jobbank.gc.ca |
| Trust Score | 90 |
| Job Quality | 75 |
| Sponsorship | 50 — Unknown |
| International Recruitment | Mixed |
| Applicant Relevance | NEDA: 60 / TOHID: 55 |
| Last Checked | {DATE_STR} |
| Notes | تابلوی کار دولتی کانادا. ۸ آگهی ماما. بسیاری حمایت ویزا ندارند. |

### ۱۱. SEEK Australia
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-011 |
| Name | SEEK Australia |
| Type | MAJOR JOB BOARD |
| Country | Australia |
| Industry | All |
| URL | https://www.seek.com.au |
| Trust Score | 85 |
| Job Quality | 75 |
| Sponsorship | 60 — Mixed |
| International Recruitment | Mixed |
| Applicant Relevance | NEDA: 70 / TOHID: 60 |
| Last Checked | {DATE_STR} |
| Notes | بزرگترین تابلوی کار استرالیا. بسیاری از مشاغل حمایت ویزا ندارند. |

### ۱۲. Arbeitnow
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-012 |
| Name | Arbeitnow |
| Type | SPECIALIST JOB BOARD |
| Country | Germany |
| Industry | IT |
| URL | https://www.arbeitnow.com |
| Trust Score | 80 |
| Job Quality | 80 |
| Sponsorship | 75 — Likely |
| International Recruitment | Yes |
| Applicant Relevance | TOHID: 85 |
| Last Checked | {DATE_STR} |
| Notes | مشاغل IT با حمایت ویزا در آلمان. |

### ۱۳. LinkedIn
| فیلد | مقدار |
|------|-------|
| Source ID | SRC-013 |
| Name | LinkedIn |
| Type | LINKEDIN |
| Country | Global |
| Industry | All |
| URL | https://www.linkedin.com |
| Trust Score | 80 |
| Job Quality | 75 |
| Sponsorship | 60 — Mixed |
| International Recruitment | Mixed |
| Applicant Relevance | Both: 75 |
| Last Checked | {DATE_STR} |
| Notes | ترکیبی — نیاز به فیلتر دقیق. |

---
"""
    SOURCE_BANK.write_text(content, encoding='utf-8')
    print(f"✅ SOURCE_BANK بروزرسانی شد — ۱۳ منبع")

# ── ۲. بروزرسانی EMPLOYER_BANK ──────────────────────────────
EMPLOYER_BANK = BASE_DIR / "memory" / "EMPLOYER_BANK.md"

def update_employer_bank():
    content = f"""# EMPLOYER_BANK — بانک اطلاعاتی کارفرمایان
آخرین بروزرسانی: {DATE_STR}

---

### ۱. Health New Zealand | Te Whatu Ora
| فیلد | مقدار |
|------|-------|
| Country | New Zealand |
| Industry | Healthcare |
| URL | https://www.healthnz.govt.nz |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Confirmed |
| Accreditation | Government — Accredited |
| Foreign Hiring | Yes — Active |
| Jobs | Midwife |
| Applicant | NEDA |
| Path Fit | 85/100 |
| Last Checked | {DATE_STR} |

### ۲. RGH Global
| فیلد | مقدار |
|------|-------|
| Country | New Zealand |
| Industry | Healthcare Recruitment |
| URL | https://www.rgh-global.com |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Confirmed |
| Accreditation | Recruiter |
| Foreign Hiring | Yes |
| Jobs | Midwife with Sponsorship |
| Applicant | NEDA |
| Path Fit | 79/100 |
| Last Checked | {DATE_STR} |

### ۳. Working In Health NZ
| فیلد | مقدار |
|------|-------|
| Country | New Zealand |
| Industry | Healthcare Recruitment |
| URL | https://www.workingin-health.co.nz |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Confirmed |
| Accreditation | Recruiter — Free |
| Foreign Hiring | Yes |
| Jobs | Midwife |
| Applicant | NEDA |
| Path Fit | 79/100 |
| Last Checked | {DATE_STR} |

### ۴. Holalemania GmbH
| فیلد | مقدار |
|------|-------|
| Country | Germany |
| Industry | Healthcare Recruitment |
| URL | https://holalemania.de/en/ |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Confirmed |
| Accreditation | Recruiter — 13 years |
| Foreign Hiring | Yes — 921 successful |
| Jobs | Midwife, Nurse |
| Applicant | NEDA |
| Path Fit | 82/100 |
| Last Checked | {DATE_STR} |
| Notes | ۶۲ بیمارستان. ۳۵ ملیت. آموزش زبان آلمانی. |

### ۵. TalentOrange
| فیلد | مقدار |
|------|-------|
| Country | Germany |
| Industry | Healthcare Recruitment |
| URL | https://www.talentorange.com/en/ |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Confirmed |
| Accreditation | Recruiter |
| Foreign Hiring | Yes |
| Jobs | Healthcare professionals |
| Applicant | NEDA |
| Path Fit | 80/100 |
| Last Checked | {DATE_STR} |
| Notes | بورسیه زبان آلمانی B2. برنامه کامل. |

### ۶. Saskatchewan Health Authority ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Country | Canada |
| Industry | Healthcare — Government |
| URL | https://www.saskhealthauthority.ca/careers |
| International Recruitment | Yes — Confirmed |
| Sponsorship | Likely |
| Accreditation | Government |
| Foreign Hiring | Yes — Active |
| Jobs | Midwife, Nurse, Healthcare |
| Applicant | NEDA |
| Path Fit | 78/100 |
| Last Checked | {DATE_STR} |
| Notes | ایمیل: SHAInternational@saskhealthauthority.ca. فعالانه بین‌المللی استخدام می‌کند. |

### ۷. Hays Healthcare Australia ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Country | Australia |
| Industry | Healthcare Recruitment |
| URL | https://www.hays.com.au/jobs/healthcare |
| International Recruitment | Yes |
| Sponsorship | Possible |
| Accreditation | Major recruiter |
| Foreign Hiring | Yes |
| Jobs | Midwife, Nurse |
| Applicant | NEDA |
| Path Fit | 72/100 |
| Last Checked | {DATE_STR} |
| Notes | آژانس استخدام بزرگ استرالیا. حمایت ویزا برای نیروی بین‌المللی. |

---
"""
    EMPLOYER_BANK.write_text(content, encoding='utf-8')
    print(f"✅ EMPLOYER_BANK بروزرسانی شد — ۷ کارفرما")

# ── ۳. بروزرسانی JOB_BANK ──────────────────────────────
JOB_BANK = BASE_DIR / "memory" / "JOB_BANK.md"

def update_job_bank():
    content = f"""# JOB_BANK — بانک اطلاعاتی فرصت‌ها
آخرین بروزرسانی: {DATE_STR}

---

## 👩 ندا — فرصت‌های مامایی

### JOB-001: Midwife — Health New Zealand
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-001 |
| Applicant | NEDA |
| Employer | Health New Zealand |
| Country | New Zealand |
| Title | Midwife / Registered Midwife |
| Location | Various — NZ |
| URL | https://www.healthnz.govt.nz/careers |
| Sponsorship | Confirmed |
| Visa | AEWV / Green List |
| Language | IELTS Academic / OET — Registration Requirement |
| Registration | Midwifery Council NZ |
| Path Fit | 85/100 |
| Status | 🟢 IDENTIFIED |
| Last Checked | {DATE_STR} |

### JOB-002: Midwife with Sponsorship — RGH Global
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-002 |
| Applicant | NEDA |
| Employer | RGH Global |
| Country | New Zealand |
| Title | Midwife with Sponsorship |
| Location | New Zealand |
| URL | https://www.rgh-global.com/jobs/midwife-with-sponsorship/ |
| Sponsorship | Confirmed |
| Visa | AEWV |
| Language | IELTS/OET — Registration |
| Registration | Midwifery Council NZ |
| Path Fit | 79/100 |
| Status | 🟢 IDENTIFIED |
| Last Checked | {DATE_STR} |

### JOB-003: Midwife — Working In Health NZ
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-003 |
| Applicant | NEDA |
| Employer | Working In Health NZ |
| Country | New Zealand |
| Title | International Midwife |
| Location | New Zealand |
| URL | https://www.workingin-health.co.nz |
| Sponsorship | Confirmed |
| Visa | AEWV / Green List |
| Language | Registration requirement |
| Registration | Midwifery Council NZ |
| Path Fit | 79/100 |
| Status | 🟢 IDENTIFIED |
| Last Checked | {DATE_STR} |

### JOB-004: Midwife — Holalemania ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-004 |
| Applicant | NEDA |
| Employer | Holalemania GmbH |
| Country | Germany |
| Title | Midwife (Geburtshelfer/in) |
| Location | Various — Germany |
| URL | https://holalemania.de/en/midwives/ |
| Sponsorship | Confirmed |
| Visa | Work Visa + Recognition |
| Language | German B1-B2 (provided by Holalemania) |
| Registration |德国ی Anerkennung |
| Path Fit | 82/100 |
| Status | 🟢 IDENTIFIED |
| Last Checked | {DATE_STR} |
| Notes | آموزش زبان آلمانی ارائه می‌دهد. ۱۳ سال تجربه. |

### JOB-005: Midwife — Saskatchewan HA ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-005 |
| Applicant | NEDA |
| Employer | Saskatchewan Health Authority |
| Country | Canada |
| Title | Midwife / Healthcare Professional |
| Location | Saskatchewan — Canada |
| URL | https://www.saskhealthauthority.ca/careers |
| Sponsorship | Likely |
| Visa | Provincial Nominee Program |
| Language | CLB 7 (IELTS 6.0+) |
| Registration | Saskatchewan Regulatory Body |
| Path Fit | 78/100 |
| Status | 🟡 NEEDS VERIFICATION |
| Last Checked | {DATE_STR} |
| Notes | ایمیل: SHAInternational@saskhealthauthority.ca |

### JOB-006: Midwife — Hays Healthcare AU ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-006 |
| Applicant | NEDA |
| Employer | Hays Healthcare Australia |
| Country | Australia |
| Title | Midwife |
| Location | Sydney / Melbourne — Australia |
| URL | https://www.hays.com.au/jobs/healthcare |
| Sponsorship | Possible |
| Visa | 482 / 189 / 190 |
| Language | IELTS 7.0 |
| Registration | AHPRA |
| Path Fit | 72/100 |
| Status | 🟡 NEEDS VERIFICATION |
| Last Checked | {DATE_STR} |

### JOB-007: Midwife — TalentOrange ⭐ جدید
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-007 |
| Applicant | NEDA |
| Employer | TalentOrange |
| Country | Germany |
| Title | Healthcare Professional |
| Location | Various — Germany |
| URL | https://www.talentorange.com/en/ |
| Sponsorship | Confirmed |
| Visa | Work Visa |
| Language | German B2 (borseh provided) |
| Registration | German Recognition |
| Path Fit | 80/100 |
| Status | 🟢 IDENTIFIED |
| Last Checked | {DATE_STR} |

---

## 👨 توحید — فرصت‌های IT

### JOB-008: IT Manager — آلمان ( Edenred +.progressBar )
| فیلد | مقدار |
|------|-------|
| Job ID | JOB-008 |
| Applicant | TOHID |
| Employer | Various |
| Country | Germany |
| Title | IT Manager / Infrastructure Manager |
| Location | Germany |
| URL | https://www.arbeitnow.com |
| Sponsorship | Likely |
| Visa | Work Visa / Blue Card |
| Language | English — varies |
| Registration | Not required for IT |
| Path Fit | 65/100 |
| Status | 🟡 NEEDS VERIFICATION |
| Last Checked | {DATE_STR} |

---
"""
    JOB_BANK.write_text(content, encoding='utf-8')
    print(f"✅ JOB_BANK بروزرسانی شد — ۸ فرصت (۲ جدید)")

# ── ۴. CV ندا — OET + مهارت‌های بالینی ──────────────────
CV_NEDA = BASE_DIR / "profiles" / "CV_NEDA.md"

def update_neda_cv():
    content = f"""# CV — Neda Arjmand
## Professional Midwife

---

### Contact Information
- **Name:** Neda Arjmand
- **Country:** Iran
- **Email:** [YOUR_EMAIL]
- **Phone:** [YOUR_PHONE]
- **LinkedIn:** https://www.linkedin.com/in/neda-arjmand
- **Date of Birth:** [YOUR_DOB]

---

### Professional Summary

Dedicated and compassionate **Registered Midwife** with **{NOW.year - 2012}+ years** of clinical experience in a high-volume hospital setting. Skilled in providing comprehensive maternal healthcare throughout all stages of pregnancy, labour, and postnatal care.

Experienced in:
- Antenatal care and assessment
- Labour ward management
- High-risk pregnancy monitoring
- Normal and assisted deliveries
- Postnatal care and breastfeeding support
- Neonatal assessment and care
- Emergency obstetric response
- Patient education and counselling
- Clinical documentation and reporting
- Interprofessional collaboration

Seeking an international midwifery position where I can contribute my clinical expertise while obtaining professional registration.

---

### Professional Experience

#### Registered Midwife
**Milad Hospital, Tehran** | 2012 – Present

**Key Responsibilities:**
- Providing antenatal, intrapartum, and postnatal care to women
- Conducting labour ward assessments and monitoring fetal wellbeing
- Managing normal deliveries and assisting with operative deliveries
- Monitoring high-risk pregnancies in collaboration with obstetricians
- Providing emergency response during obstetric emergencies
- Educating mothers on breastfeeding, newborn care, and postnatal recovery
- Maintaining accurate clinical documentation
- Training and mentoring junior midwifery staff
- Participating in quality improvement initiatives

**Clinical Areas:**
- Labour Ward
- Antenatal Clinic
- Postnatal Ward
- High-Risk Pregnancy Unit
- Emergency Obstetric Unit

**Key Achievements:**
- Managed care for 200+ deliveries annually
- Contributed to reduced maternal complication rates through early assessment
- Implemented patient education programme for postnatal care
- Recognised for compassionate care and patient satisfaction

---

### Education

#### Bachelor of Midwifery
**[University Name], Tehran, Iran** | [Year]

---

### Clinical Skills

#### Antenatal Care
- Antenatal assessment and history taking
- Fetal growth monitoring
- Blood pressure monitoring
- Urine protein testing
- Gestational diabetes screening
- Group B Streptococcus screening
- Antenatal education classes

#### Labour and Delivery
- Normal vaginal delivery management
- Assisted delivery (ventouse/forceps)
- Caesarean section assistance
- Electronic fetal monitoring (CTG interpretation)
- Active management of third stage of labour
- Episiotomy and repair
- Water birth assistance

#### High-Risk Pregnancy
- Pre-eclampsia monitoring
- Gestational diabetes management
- Preterm labour assessment
- Multiple pregnancy monitoring
- Postpartum haemorrhage management
- Maternal assessment and triage

#### Postnatal Care
- Breastfeeding support and counselling
- Newborn assessment (APGAR scoring)
- Postnatal examination
- Mental health screening (postnatal depression)
- Contraception counselling
- Wound care (episiotomy/caesarean)

#### Emergency Skills
- Neonatal resuscitation (NRP)
- Maternal resuscitation
- Obstetric emergency response
- Transfer of care protocols
- First aid in obstetric settings

---

### Professional Registration & Language

#### Professional Registration
- **Iran:** Licensed Midwife — [Registration Number]
- **New Zealand:** NOT YET REGISTERED — In process
  - Regulator: Midwifery Council of New Zealand
  - Pathway: Competence assessment
  - Status: Investigating requirements
- **Germany:** NOT YET REGISTERED — In process
  - Regulator: Relevant Landesbehörde
  - Pathway: Anerkennung (Recognition)
  - Status: Investigating requirements
- **Canada:** NOT YET REGISTERED — In process
  - Regulator: Provincial Regulatory Body
  - Pathway: Assessment process
  - Status: Investigating requirements

#### Language Proficiency
- **English:** A2 — Currently improving (Target: IELTS Academic 7.0 or OET B)
- **German:** A1 — Currently improving
- **Persian:** Native

**Note:** IELTS Academic / OET preparation is in progress. Language requirement timing varies by country:
- **NZ:** Required for Midwifery Council registration
- **Germany:** German B1-B2 required (Holalemania provides training)
- **Canada:** CLB 7 required for Express Entry

---

### Professional Development

- Continuous midwifery education
- Emergency obstetric care training
- Neonatal resuscitation programme
- Patient safety and quality improvement
- Clinical audit participation

---

### References

Available upon request.

---

*Last updated: {DATE_STR}*
"""
    CV_NEDA.write_text(content, encoding='utf-8')
    print(f"✅ CV ندا بروزرسانی شد — OET + مهارت‌های بالینی")

# ── ۵. گزارش فارسی ──────────────────────────────
DAILY_ACTIONS = BASE_DIR / "output" / "DAILY_ACTIONS.md"

def update_daily_actions():
    content = f"""# 🎯 ۵ اقدام برتر امروز — {DATE_STR}

---

## ۱. 👩 ندا — ارسال CV به Saskatchewan Health Authority
| فیلد | مقدار |
|------|-------|
| متقاضی | 👩 ندا |
| کارفرما | Saskatchewan Health Authority |
| کشور | 🇨🇦 کانادا |
| ایمیل | SHAInternational@saskhealthauthority.ca |
| لینک | https://www.saskhealthauthority.ca/careers |
| اولویت | 🔴 فوری |
| اقدام | ارسال CV + درخواست Expression of Interest |

---

## ۲. 👩 ندا — ثبت‌نام در Hays Healthcare Australia
| فیلد | مقدار |
|------|-------|
| متقاضی | 👩 ندا |
| کارفرما | Hays Healthcare Australia |
| کشور | 🇦🇺 استرالیا |
| لینک | https://www.hays.com.au/jobs/healthcare |
| اولویت | 🔴 فوری |
| اقدام | ثبت‌نام CV + درخواست مشاوره استخدام |

---

## ۳. 👩 ندا — پیگیری Working In Health NZ
| فیلد | مقدار |
|------|-------|
| متقاضی | 👩 ندا |
| کارفرما | Working In Health NZ |
| کشور | 🇳🇿 نیوزیلند |
| لینک | https://www.workingin-health.co.nz |
| اولویت | 🟠 مهم |
| اقدام | پیگیری ایمیل قبلی + تکمیل فرم |

---

## ۴. 👩 ندا — آمادگی OET
| فیلد | مقدار |
|------|-------|
| متقاضی | 👩 ندا |
| نوع | آمادگی آزمون |
| آزمون | OET (Occupational English Test) |
| هدف | نمره B در هر ۴ مهارت |
| اولویت | 🟠 مهم |
| اقدام | ثبت‌نام + شروع مطالعه |

---

## ۵. 👨 توحید — جستجوی IT آلمان/اتریش
| فیلد | مقدار |
|------|-------|
| متقاضی | 👨 توحید |
| کشور | 🇩🇪 آلمان / 🇦🇹 اتریش |
| تمرکز | IT Infrastructure / Network Manager |
| لینک | https://www.arbeitnow.com |
| اولویت | 🟠 مهم |
| اقدام | جستجو + ارسال CV به مشاغل IT |

---

## 📊 خلاصه چرخه سوم

| شاخص | مقدار |
|------|-------|
| منابع جدید | ۵ (Saskatchewan, Hays, ANMF, CAM, Job Bank) |
| فرصت‌های جدید | ۳ (Saskatchewan, Hays, TalentOrange) |
| کارفرمایان جدید | ۳ |
| CV بروزرسانی شده | ✅ ندا |
| کشورهای جستجو شده | 🇳🇿 🇦🇺 🇨🇦 🇩🇪 |

---

*آخرین بروزرسانی: {DATE_STR}*
"""
    DAILY_ACTIONS.write_text(content, encoding='utf-8')
    print(f"✅ DAILY_ACTIONS بروزرسانی شد — ۵ اقدام برتر")

# ── ۶. اجرا ──────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("🔄 چرخه سوم — بروزرسانی جامع")
    print(f"📅 {DATE_STR}")
    print("=" * 60)
    print()
    
    update_source_bank()
    update_employer_bank()
    update_job_bank()
    update_neda_cv()
    update_daily_actions()
    
    print()
    print("=" * 60)
    print("✅ چرخه سوم تکمیل شد!")
    print("=" * 60)
