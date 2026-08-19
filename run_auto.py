# -*- coding: utf-8 -*-
"""
Migration Hunter — Auto Run Script v2.0
نسخه خودکار با جستجوی وب + AI API + ارسال ایمیل

نحوه اجرا:
    python run_auto.py

یا:
    python run_auto.py --country nz
    python run_auto.py --country de
    python run_auto.py --full
    python run_auto.py --ai          # با AI API
    python run_auto.py --email       # با ارسال ایمیل
    python run_auto.py --ai --email  # هر دو
"""

import os
import sys
import json
import re
import hashlib
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

# Import from run.py
sys.path.insert(0, str(Path(__file__).parent))
from run import (
    MemoryBank, JobCollector, JobAnalyzer,
    ReportGenerator, ExcelGenerator, MigrationHunter
)

# ==========================================
# CONFIGURATION
# ==========================================

BASE_DIR = Path(__file__).parent
ENV_FILE = BASE_DIR / ".env"
CONFIG_FILE = BASE_DIR / "config.json"

# ==========================================
# AI API INTEGRATION
# ==========================================

class AIApiClient:
    """اتصال به AI API برای تحلیل هوشمند"""
    
    def __init__(self):
        self.provider = None
        self.api_key = None
        self._load_config()
    
    def _load_config(self):
        """بارگذاری تنظیمات از .env یا config.json"""
        # Try .env first
        if ENV_FILE.exists():
            content = ENV_FILE.read_text(encoding='utf-8')
            for line in content.split('\n'):
                if '=' in line and not line.startswith('#'):
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key == 'AI_PROVIDER':
                        self.provider = value
                    elif key == 'AI_API_KEY':
                        self.api_key = value
        
        # Try config.json
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.provider = config.get('ai_provider', self.provider)
                    self.api_key = config.get('ai_api_key', self.api_key)
            except:
                pass
        
        if self.provider and self.api_key:
            print(f"✓ AI API configured: {self.provider}")
        else:
            print("⚠️ AI API not configured. Using rule-based analysis.")
            print("  Create .env with:")
            print("    AI_PROVIDER=openai")
            print("    AI_API_KEY=your_api_key")
    
    def analyze_job(self, job, profile):
        """تحلیل هوشمند شغل با AI"""
        if not self.provider or not self.api_key:
            return None
        
        prompt = f"""
Analyze this job opportunity for an applicant from Iran.

Job:
- Title: {job.get('title', 'Unknown')}
- Employer: {job.get('employer', 'Unknown')}
- Country: {job.get('country', 'Unknown')}
- Sponsorship: {job.get('sponsorship', 'Unknown')}
- Language: {job.get('language', 'Unknown')}

Applicant Profile:
- Name: {profile.get('name', 'Unknown')}
- Profession: {profile.get('profession', 'Unknown')}
- Age: {profile.get('age', 'Unknown')}
- English: {profile.get('english', 'Unknown')}
- German: {profile.get('german', 'Unknown')}

Provide:
1. Path Fit Score (0-100)
2. Sponsorship assessment
3. Language gap analysis
4. Registration requirements
5. Risk factors
6. Recommendation (APPLY/CONTACT/RESEARCH/SKIP)

Respond in JSON format.
"""
        
        try:
            if self.provider == 'openai':
                return self._call_openai(prompt)
            elif self.provider == 'gemini':
                return self._call_gemini(prompt)
        except Exception as e:
            print(f"  ⚠️ AI API error: {e}")
            return None
    
    def generate_email(self, job, profile, job_type='application'):
        """تولید ایمیل با AI"""
        if not self.provider or not self.api_key:
            return self._fallback_email(job, profile)
        
        prompt = f"""
Write a professional {job_type} email for this job opportunity.

Job Details:
- Title: {job.get('title', 'Unknown')}
- Employer: {job.get('employer', 'Unknown')}
- Country: {job.get('country', 'Unknown')}
- Email: {job.get('email', 'Unknown')}

Applicant:
- Name: {profile.get('name', 'Unknown')}
- Profession: {profile.get('profession', 'Unknown')}
- Experience: {profile.get('experience', 'Unknown')}
- Current Location: Iran

Write a professional, concise email in English.
Do NOT use "Dear Sir/Madam".
Address the hiring manager directly.
Mention willingness to relocate.
Do NOT exaggerate.
"""
        
        try:
            if self.provider == 'openai':
                return self._call_openai(prompt)
            elif self.provider == 'gemini':
                return self._call_gemini(prompt)
        except Exception as e:
            print(f"  ⚠️ AI API error: {e}")
            return self._fallback_email(job, profile)
    
    def generate_cover_letter(self, job, profile):
        """تولید کاور لیتر با AI"""
        if not self.provider or not self.api_key:
            return self._fallback_cover_letter(job, profile)
        
        prompt = f"""
Write a professional cover letter for this job opportunity.

Job Details:
- Title: {job.get('title', 'Unknown')}
- Employer: {job.get('employer', 'Unknown')}
- Country: {job.get('country', 'Unknown')}

Applicant:
- Name: {profile.get('name', 'Unknown')}
- Profession: {profile.get('profession', 'Unknown')}
- Experience: {profile.get('experience', 'Unknown')}
- Skills: {profile.get('skills', 'Unknown')}
- Current Location: Iran

Write a professional, concise cover letter (max 300 words).
Focus on relevant experience and willingness to relocate.
Do NOT exaggerate.
"""
        
        try:
            if self.provider == 'openai':
                return self._call_openai(prompt)
            elif self.provider == 'gemini':
                return self._call_gemini(prompt)
        except Exception as e:
            print(f"  ⚠️ AI API error: {e}")
            return self._fallback_cover_letter(job, profile)
    
    def _call_openai(self, prompt):
        """فراخوانی OpenAI API"""
        try:
            import requests
            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }
            data = {
                'model': 'gpt-4',
                'messages': [{'role': 'user', 'content': prompt}],
                'temperature': 0.7,
                'max_tokens': 1000
            }
            response = requests.post(
                'https://api.openai.com/v1/chat/completions',
                headers=headers,
                json=data,
                timeout=30
            )
            if response.status_code == 200:
                return response.json()['choices'][0]['message']['content']
        except Exception as e:
            print(f"  ⚠️ OpenAI error: {e}")
        return None
    
    def _call_gemini(self, prompt):
        """فراخوانی Gemini API"""
        try:
            import requests
            url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key={self.api_key}'
            data = {
                'contents': [{'parts': [{'text': prompt}]}]
            }
            response = requests.post(url, json=data, timeout=30)
            if response.status_code == 200:
                return response.json()['candidates'][0]['content']['parts'][0]['text']
        except Exception as e:
            print(f"  ⚠️ Gemini error: {e}")
        return None
    
    def _fallback_email(self, job, profile):
        """ایمیل پیش‌فرض بدون AI"""
        return f"""Subject: {job.get('title', 'Position')} — Application

Dear Hiring Manager,

I am writing to express my interest in the {job.get('title', 'position')} role at {job.get('employer', 'your organization')}.

I am a qualified {profile.get('profession', 'professional')} with experience in {profile.get('profession', 'my field')}. I am seeking opportunities with employers who sponsor international candidates.

I would welcome the opportunity to discuss how my skills and experience could contribute to your team.

Please find my CV attached for your review.

Best regards,
{profile.get('name', 'Applicant')}
"""
    
    def _fallback_cover_letter(self, job, profile):
        """کاور لیتر پیش‌فرض بدون AI"""
        return f"""Dear Hiring Manager,

I am writing to apply for the {job.get('title', 'position')} role at {job.get('employer', 'your organization')}.

As a qualified {profile.get('profession', 'professional')} with experience in {profile.get('profession', 'my field')}, I believe my skills would be a valuable addition to your team.

I am particularly drawn to this opportunity because of your organization's commitment to international recruitment. I am eager to relocate and contribute to your team's success.

I have attached my CV for your review and would welcome the opportunity to discuss my application further.

Best regards,
{profile.get('name', 'Applicant')}
"""


# ==========================================
# EMAIL SENDER
# ==========================================

class EmailSender:
    """ارسال ایمیل با تأیید کاربر"""
    
    def __init__(self):
        self.sent_emails = []
    
    def send_with_confirmation(self, email_data, recipient_name=""):
        """ارسال ایمیل با تأیید کاربر"""
        print("\n" + "=" * 60)
        print("📧 ایمیل آماده ارسال")
        print("=" * 60)
        print(f"\nگیرنده: {recipient_name}")
        print(f"موضوع: {email_data.get('subject', 'N/A')}")
        print(f"\nمتن ایمیل:")
        print("-" * 40)
        print(email_data.get('body', ''))
        print("-" * 40)
        
        print("\n⚠️ گزینه‌ها:")
        print("  1. ارسال ایمیل (نیاز به پیکربندی SMTP)")
        print("  2. ذخیره به عنوان فایل")
        print("  3. کپی به clipboard")
        print("  4. لغو")
        
        choice = input("\nانتخاب شما (1-4): ").strip()
        
        if choice == '1':
            return self._send_smtp(email_data)
        elif choice == '2':
            return self._save_to_file(email_data)
        elif choice == '3':
            return self._copy_to_clipboard(email_data)
        else:
            print("❌ لغو شد")
            return False
    
    def _send_smtp(self, email_data):
        """ارسال از طریق SMTP"""
        print("\n⚠️ ارسال SMTP نیاز به پیکربندی دارد.")
        print("این قابلیت در نسخه آینده اضافه خواهد شد.")
        print("فعلاً ایمیل ذخیره شد.")
        return self._save_to_file(email_data)
    
    def _save_to_file(self, email_data):
        """ذخیره ایمیل به عنوان فایل"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"email_{timestamp}.txt"
        filepath = BASE_DIR / "output" / filename
        
        content = f"""To: {email_data.get('to', 'N/A')}
Subject: {email_data.get('subject', 'N/A')}

{email_data.get('body', '')}
"""
        
        filepath.write_text(content, encoding='utf-8')
        print(f"✅ ایمیل ذخیره شد: {filepath}")
        self.sent_emails.append(email_data)
        return True
    
    def _copy_to_clipboard(self, email_data):
        """کپی ایمیل به clipboard"""
        try:
            import subprocess
            content = f"To: {email_data.get('to', 'N/A')}\nSubject: {email_data.get('subject', 'N/A')}\n\n{email_data.get('body', '')}"
            
            # Windows
            if sys.platform == 'win32':
                subprocess.run(['clip'], input=content.encode('utf-16le'), check=True)
            
            print("✅ ایمیل به clipboard کپی شد")
            return True
        except Exception as e:
            print(f"⚠️ خطا در کپی: {e}")
            return self._save_to_file(email_data)


# ==========================================
# WEB SCRAPER
# ==========================================

class WebScraper:
    """خواندن صفحات وب"""

    def __init__(self):
        self.session = None
        self._init_session()

    def _init_session(self):
        """مقداردهی اولیه session"""
        try:
            import requests
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            print("✓ requests available")
        except ImportError:
            print("⚠️ requests not installed. Run: pip install requests")

    def fetch_page(self, url, timeout=10):
        """خواندن صفحه وب"""
        if not self.session:
            return None

        try:
            response = self.session.get(url, timeout=timeout)
            response.raise_for_status()
            return response.text
        except Exception as e:
            print(f"  ⚠️ Error fetching {url}: {e}")
            return None

    def parse_jobs_from_html(self, html, source):
        """پارس کردن مشاغل از HTML"""
        jobs = []

        # Simple regex patterns for job titles
        patterns = [
            r'<h[23][^>]*>(.*?)</h[23]>',
            r'<a[^>]*class="[^"]*job[^"]*"[^>]*>(.*?)</a>',
            r'<div[^>]*class="[^"]*title[^"]*"[^>]*>(.*?)</div>',
        ]

        for pattern in patterns:
            matches = re.findall(pattern, html, re.IGNORECASE | re.DOTALL)
            for match in matches:
                clean = re.sub(r'<[^>]+>', '', match).strip()
                if clean and len(clean) > 5:
                    jobs.append({
                        "title": clean,
                        "source": source,
                        "collected_at": datetime.now().isoformat()
                    })

        return jobs[:10]  # Limit to 10


# ==========================================
# JOB SOURCES
# ==========================================

class JobSources:
    """منابع شغلی"""

    SOURCES = {
        "nz": [
            {
                "name": "Health New Zealand",
                "url": "https://www.healthnz.govt.nz/careers/international",
                "type": "government",
                "trust_score": 95
            },
            {
                "name": "Working In Health NZ",
                "url": "https://www.workingin-health.co.nz/midwifery-jobs/",
                "type": "recruiter",
                "trust_score": 88
            },
            {
                "name": "RGH Global",
                "url": "https://www.rgh-global.com/jobs/midwife-with-sponsorship/",
                "type": "recruiter",
                "trust_score": 85
            },
        ],
        "de": [
            {
                "name": "Holalemania",
                "url": "https://holalemania.de/en/",
                "type": "recruiter",
                "trust_score": 85
            },
            {
                "name": "TalentOrange",
                "url": "https://www.talentorange.com/en/",
                "type": "recruiter",
                "trust_score": 82
            },
            {
                "name": "Arbeitnow",
                "url": "https://www.arbeitnow.com",
                "type": "job_board",
                "trust_score": 80
            },
        ],
        "au": [
            {
                "name": "SEEK Australia",
                "url": "https://www.seek.com.au",
                "type": "job_board",
                "trust_score": 85
            },
            {
                "name": "Hays Healthcare",
                "url": "https://www.hays.com.au/healthcare",
                "type": "recruiter",
                "trust_score": 80
            },
        ],
        "ca": [
            {
                "name": "Job Bank Canada",
                "url": "https://www.jobbank.gc.ca",
                "type": "government",
                "trust_score": 90
            },
        ],
    }

    @classmethod
    def get_sources(cls, country=None):
        """دریافت منابع"""
        if country:
            return cls.SOURCES.get(country, [])
        return cls.SOURCES


# ==========================================
# ENHANCED COLLECTOR
# ==========================================

class EnhancedCollector(JobCollector):
    """جمع‌آوری پیشرفته با وب اسکرپینگ"""

    def __init__(self):
        super().__init__()
        self.scraper = WebScraper()

    def collect_from_web(self, countries=None):
        """جمع‌آوری از وب"""
        print("\n🌐 جمع‌آوری از وب...")

        if not countries:
            countries = ["nz", "de"]

        for country in countries:
            sources = JobSources.get_sources(country)
            for source in sources:
                print(f"\n  📡 {source['name']} ({country})...")

                html = self.scraper.fetch_page(source["url"])
                if html:
                    web_jobs = self.scraper.parse_jobs_from_html(html, source["name"])
                    print(f"    ✓ {len(web_jobs)} مشاغل وب یافت شد")

                    # Convert to our format
                    for job in web_jobs:
                        self.jobs.append({
                            "id": f"WEB-{hashlib.md5(job['title'].encode()).hexdigest()[:8].upper()}",
                            "employer": source["name"],
                            "country": country,
                            "title": job["title"],
                            "url": source["url"],
                            "sponsorship": "UNKNOWN",
                            "applicant": "NEDA" if country in ["nz", "de"] else "TOHID",
                            "path_fit_score": 60,  # Will be calculated
                            "status": "NEW",
                            "source": "web",
                            "collected_at": datetime.now().isoformat()
                        })

        return self.jobs


# ==========================================
# MAIN ENHANCED RUNNER
# ==========================================

class MigrationHunterAuto(MigrationHunter):
    """نسخه خودکار"""

    def __init__(self, use_ai=False, send_email=False):
        super().__init__()
        self.collector = EnhancedCollector()
        self.ai_client = AIApiClient() if use_ai else None
        self.email_sender = EmailSender() if send_email else None
        self.use_ai = use_ai
        self.send_email = send_email

    def run(self, countries=None, applicant=None):
        """اجرای خودکار"""
        print("=" * 60)
        print("🚀 MIGRATION HUNTER — AUTO RUN v2.0")
        print(f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"🤖 AI API: {'Enabled' if self.use_ai else 'Disabled'}")
        print(f"📧 Email: {'Enabled' if self.send_email else 'Disabled'}")
        print("=" * 60)

        # Step 1: Load memory
        self.load_memory()

        # Step 2: Load profiles
        profiles = self.load_profiles()

        # Step 3: Collect from web
        if not countries:
            countries = ["nz", "de"]
        self.collector.collect_from_web(countries)

        # Step 4: Collect from known sources
        self.collector.collect_all()

        # Step 5: Analyze jobs (with AI if enabled)
        if self.use_ai and self.ai_client:
            analyzed_jobs = self._analyze_with_ai(self.collector.jobs, profiles)
        else:
            analyzed_jobs = self.analyzer.analyze_jobs(self.collector.jobs, profiles)

        # Step 6: Generate emails (with AI if enabled)
        if self.send_email:
            self._generate_emails(analyzed_jobs, profiles)

        # Step 7: Update memory
        for job in analyzed_jobs:
            self.memory_banks["job"].add_item({
                "name": job.get("id", ""),
                "employer": job.get("employer", ""),
                "country": job.get("country", ""),
                "applicant": job.get("applicant", ""),
                "title": job.get("title", ""),
                "score": job.get("path_fit_score", 0),
                "status": job.get("status", ""),
                "collected_at": job.get("collected_at", "")
            })

        # Step 8: Generate reports
        self.reporter.generate_daily_actions(analyzed_jobs)
        self.reporter.generate_top_jobs(analyzed_jobs, "NEDA")
        self.reporter.generate_top_jobs(analyzed_jobs, "TOHID")

        # Step 9: Generate Excel
        self.excel_gen.generate(analyzed_jobs)

        # Step 10: Save memory
        self.save_memory()

        # Summary
        print("\n" + "=" * 60)
        print("✅ اجرای خودکار تکمیل شد!")
        print("=" * 60)
        print(f"\n📊 خلاصه:")
        print(f"  - کل فرصت‌ها: {len(analyzed_jobs)}")
        print(f"  - NEDA: {len([j for j in analyzed_jobs if j.get('applicant') == 'NEDA'])}")
        print(f"  - TOHID: {len([j for j in analyzed_jobs if j.get('applicant') == 'TOHID'])}")
        if self.send_email:
            print(f"  - ایمیل‌های ارسالی: {len(self.email_sender.sent_emails)}")

        return analyzed_jobs

    def _analyze_with_ai(self, jobs, profiles):
        """تحلیل مشاغل با AI"""
        print("\n🤖 تحلیل هوشمند با AI...")
        
        analyzed = []
        for job in jobs:
            applicant = job.get("applicant", "")
            profile = profiles.get(applicant, {})
            
            # Try AI analysis
            ai_result = self.ai_client.analyze_job(job, profile)
            
            if ai_result:
                try:
                    # Parse AI result (assuming JSON)
                    if isinstance(ai_result, str):
                        ai_data = json.loads(ai_result)
                    else:
                        ai_data = ai_result
                    
                    job["path_fit_score"] = ai_data.get("path_fit_score", 70)
                    job["ai_analysis"] = ai_data
                    print(f"  ✓ {job['employer']}: {job['path_fit_score']}/100 (AI)")
                except:
                    # Fallback to rule-based
                    score = self.analyzer.calculate_path_fit(job, profile)
                    job["path_fit_score"] = score
                    print(f"  ✓ {job['employer']}: {score}/100 (rule-based)")
            else:
                # Fallback to rule-based
                score = self.analyzer.calculate_path_fit(job, profile)
                job["path_fit_score"] = score
                print(f"  ✓ {job['employer']}: {score}/100 (rule-based)")
            
            analyzed.append(job)
        
        return sorted(analyzed, key=lambda x: x.get("path_fit_score", 0), reverse=True)

    def _generate_emails(self, jobs, profiles):
        """تولید ایمیل‌ها"""
        print("\n📧 تولید ایمیل‌ها...")
        
        emails_generated = 0
        
        for job in jobs[:5]:  # Top 5 jobs
            applicant = job.get("applicant", "")
            profile = profiles.get(applicant, {})
            
            # Generate email
            if self.ai_client:
                email_body = self.ai_client.generate_email(job, profile)
            else:
                email_body = self.ai_client._fallback_email(job, profile) if self.ai_client else "Email generation requires AI API"
            
            email_data = {
                "to": job.get("email", "recruiter@" + job.get("employer", "company.com").replace(" ", "").lower() + ".com"),
                "subject": f"{job.get('title', 'Position')} — Application from International Candidate",
                "body": email_body,
                "job": job,
                "applicant": applicant
            }
            
            print(f"\n  📧 Email for {job['employer']}:")
            print(f"     Subject: {email_data['subject']}")
            
            # Save email
            self._save_email(email_data)
            emails_generated += 1
        
        print(f"\n  ✅ {emails_generated} ایمیل تولید شد")
        return emails_generated

    def _save_email(self, email_data):
        """ذخیره ایمیل"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"email_{timestamp}.txt"
        filepath = BASE_DIR / "output" / filename
        
        content = f"""To: {email_data.get('to', 'N/A')}
Subject: {email_data.get('subject', 'N/A')}

{email_data.get('body', '')}
"""
        
        filepath.write_text(content, encoding='utf-8')
        print(f"     💾 ذخیره شد: {filepath}")

    def run_interactive_email(self):
        """حالت تعاملی برای ارسال ایمیل"""
        if not self.email_sender:
            print("❌ Email sending not enabled. Use --email flag.")
            return
        
        print("\n📧 حالت تعاملی ارسال ایمیل")
        print("=" * 60)
        
        # Load latest emails
        emails_dir = BASE_DIR / "output"
        email_files = list(emails_dir.glob("email_*.txt"))
        
        if not email_files:
            print("❌ هیچ ایمیلی ذخیره نشده است.")
            return
        
        print(f"\n📄 {len(email_files)} ایمیل ذخیره شده:")
        for i, f in enumerate(email_files[:5], 1):
            print(f"  {i}. {f.name}")
        
        choice = input("\nشماره ایمیل برای ارسال (یا Enter برای لغو): ").strip()
        
        if choice and choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(email_files):
                email_file = email_files[idx]
                content = email_file.read_text(encoding='utf-8')
                
                print(f"\n📧 محتوای ایمیل:")
                print("-" * 40)
                print(content)
                print("-" * 40)
                
                confirm = input("\nآیا می‌خواهید این ایمیل را ارسال کنید؟ (y/N): ").strip().lower()
                
                if confirm == 'y':
                    print("✅ ایمیل آماده ارسال است.")
                    print("⚠️ لطفاً ایمیل را از طریق مرورگر یا اپلیکیشن ایمیل ارسال کنید.")
                    print(f"📁 فایل: {email_file}")
                else:
                    print("❌ لغو شد.")


# ==========================================
# ENTRY POINT
# ==========================================

def main():
    """نقطه ورود"""
    import argparse

    parser = argparse.ArgumentParser(description='Migration Hunter Auto v2.0 — Job Hunting Engine with AI')
    parser.add_argument('--country', type=str, help='Country to search (nz, de, au, ca)')
    parser.add_argument('--applicant', type=str, help='Applicant (neda, tohid)')
    parser.add_argument('--full', action='store_true', help='Full run')
    parser.add_argument('--ai', action='store_true', help='Enable AI API analysis')
    parser.add_argument('--email', action='store_true', help='Enable email generation')
    parser.add_argument('--interactive-email', action='store_true', help='Interactive email sending')

    args = parser.parse_args()

    hunter = MigrationHunterAuto(use_ai=args.ai, send_email=args.email)

    if args.interactive_email:
        hunter.run_interactive_email()
    elif args.country:
        print(f"🔍 جستجوی {args.country}...")
        hunter.run(countries=[args.country])
    elif args.applicant:
        print(f"🔍 جستجوی {args.applicant}...")
        hunter.run(applicant=args.applicant)
    else:
        hunter.run()


if __name__ == "__main__":
    main()
