# -*- coding: utf-8 -*-
"""
LinkedIn Crawler — کراول پروفایل و آگهی‌های شغلی
برای پروژه Iran-to-Opportunity

نحوه اجرا:
    python linkedin_crawler.py --profile YOUR_LINKEDIN_URL
    python linkedin_crawler.py --jobs "midwife" --country "New Zealand"
    python linkedin_crawler.py --export profile.json
"""

import os
import sys
import json
import re
from datetime import datetime
from pathlib import Path

# ==========================================
# CONFIGURATION
# ==========================================

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "linkedin_data"
OUTPUT_DIR.mkdir(exist_ok=True)

# ==========================================
# LINKEDIN PROFILE CRAWLER
# ==========================================

class LinkedInProfileCrawler:
    """کراول پروفایل LinkedIn"""
    
    def __init__(self, use_selenium=False):
        self.use_selenium = use_selenium
        self.session = None
        self._init_session()
    
    def _init_session(self):
        """مقداردهی اولیه session"""
        try:
            import requests
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
            })
            print("✓ requests available")
        except ImportError:
            print("⚠️ requests not installed. Run: pip install requests")
    
    def crawl_profile(self, linkedin_url):
        """کراول پروفایل LinkedIn"""
        print(f"\n🔍 کراول پروفایل: {linkedin_url}")
        
        # Extract username from URL
        username = self._extract_username(linkedin_url)
        if not username:
            print("❌ نام کاربری قابل استخراج نیست")
            return None
        
        print(f"  📧 نام کاربری: {username}")
        
        # Try to fetch public profile
        profile_data = self._fetch_public_profile(username)
        
        if profile_data:
            print(f"  ✅ پروفایل یافت شد")
            return profile_data
        else:
            print(f"  ⚠️ پروفایل قابل دسترسی نیست (نیاز به login)")
            return self._create_template_profile(username)
    
    def _extract_username(self, url):
        """استخراج نام کاربری از URL"""
        patterns = [
            r'linkedin\.com/in/([^/?]+)',
            r'linkedin\.com/profile/view\?id=([^&]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        
        return None
    
    def _fetch_public_profile(self, username):
        """دریافت پروفایل عمومی"""
        if not self.session:
            return None
        
        url = f"https://www.linkedin.com/in/{username}"
        
        try:
            response = self.session.get(url, timeout=10)
            
            if response.status_code == 200:
                # Parse basic info from HTML
                return self._parse_profile_html(response.text, username)
            else:
                print(f"  ⚠️ Status: {response.status_code}")
                return None
        except Exception as e:
            print(f"  ⚠️ Error: {e}")
            return None
    
    def _parse_profile_html(self, html, username):
        """پارس کردن HTML پروفایل"""
        profile = {
            "username": username,
            "url": f"https://www.linkedin.com/in/{username}",
            "crawled_at": datetime.now().isoformat(),
            "source": "linkedin_public"
        }
        
        # Try to extract basic info
        # Note: LinkedIn blocks most scraping, so this is limited
        
        # Extract title/headline
        title_match = re.search(r'<title>(.*?)</title>', html)
        if title_match:
            profile["headline"] = title_match.group(1).split(" - ")[0].strip()
        
        # Extract meta description
        desc_match = re.search(r'<meta[^>]*name="description"[^>]*content="([^"]*)"', html)
        if desc_match:
            profile["description"] = desc_match.group(1)
        
        return profile
    
    def _create_template_profile(self, username):
        """ایجاد پروفایل الگو"""
        return {
            "username": username,
            "url": f"https://www.linkedin.com/in/{username}",
            "crawled_at": datetime.now().isoformat(),
            "source": "template",
            "note": "LinkedIn blocks direct scraping. Use LinkedIn API or Selenium for full data.",
            "template": {
                "headline": "[YOUR HEADLINE]",
                "location": "[YOUR LOCATION]",
                "connections": "[NUMBER]",
                "experience": [
                    {
                        "title": "[JOB TITLE]",
                        "company": "[COMPANY]",
                        "duration": "[DURATION]"
                    }
                ],
                "education": [
                    {
                        "school": "[UNIVERSITY]",
                        "degree": "[DEGREE]",
                        "field": "[FIELD]"
                    }
                ],
                "skills": ["[SKILL 1]", "[SKILL 2]", "[SKILL 3]"]
            }
        }


# ==========================================
# LINKEDIN JOBS CRAWLER
# ==========================================

class LinkedInJobsCrawler:
    """کراول آگهی‌های شغلی LinkedIn"""
    
    def __init__(self):
        self.session = None
        self._init_session()
    
    def _init_session(self):
        """مقداردهی اولیه session"""
        try:
            import requests
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            })
        except ImportError:
            print("⚠️ requests not installed")
    
    def search_jobs(self, keywords, location="", limit=10):
        """جستجوی مشاغل"""
        print(f"\n🔍 جستجوی مشاغل: {keywords} در {location}")
        
        jobs = []
        
        # LinkedIn Jobs URL
        params = {
            'keywords': keywords,
            'location': location,
            'f_TPR': 'r604800',  # Last week
        }
        
        # Note: LinkedIn blocks direct scraping
        # This is a template for when API access is available
        
        print(f"  ⚠️ LinkedIn مستقیماً کراول را بلاک می‌کند")
        print(f"  📝 از LinkedIn API یا Selenium استفاده کنید")
        
        return jobs


# ==========================================
# DATA EXPORTER
# ==========================================

class LinkedInDataExporter:
    """خروجی گرفتن از داده‌های LinkedIn"""
    
    def __init__(self):
        self.output_dir = OUTPUT_DIR
    
    def export_to_json(self, data, filename):
        """خروجی به JSON"""
        filepath = self.output_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"  ✅ ذخیره شد: {filepath}")
        return filepath
    
    def export_to_csv(self, data, filename):
        """خروجی به CSV"""
        import csv
        
        filepath = self.output_dir / filename
        
        if isinstance(data, list) and len(data) > 0:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=data[0].keys())
                writer.writeheader()
                writer.writerows(data)
            
            print(f"  ✅ ذخیره شد: {filepath}")
            return filepath
        
        return None
    
    def export_profile_for_migration(self, profile_data, applicant_name):
        """خروجی پروفایل برای Migration Hunter"""
        
        # Create profile markdown
        markdown = f"""# {applicant_name} — LinkedIn Profile

## Basic Information

| Field | Value |
|-------|-------|
| **Name** | {applicant_name} |
| **LinkedIn URL** | {profile_data.get('url', 'N/A')} |
| **Headline** | {profile_data.get('headline', 'N/A')} |
| **Crawled At** | {profile_data.get('crawled_at', 'N/A')} |

## Template Data

.fill this manually based on your LinkedIn profile:

| Field | Your Value |
|-------|------------|
| **Location** | [YOUR LOCATION] |
| **Current Company** | [YOUR COMPANY] |
| **Current Title** | [YOUR TITLE] |
| **Experience Years** | [YEARS] |
| **Education** | [YOUR DEGREE] |
| **Skills** | [YOUR SKILLS] |

---
"""
        
        filepath = self.output_dir / f"{applicant_name.lower().replace(' ', '_')}_linkedin.md"
        filepath.write_text(markdown, encoding='utf-8')
        
        print(f"  ✅ پروفایل ذخیره شد: {filepath}")
        return filepath


# ==========================================
# MAIN
# ==========================================

def main():
    """نقطه ورود"""
    import argparse
    
    parser = argparse.ArgumentParser(description='LinkedIn Crawler for Job Hunting')
    parser.add_argument('--profile', type=str, help='LinkedIn profile URL')
    parser.add_argument('--jobs', type=str, help='Job search keywords')
    parser.add_argument('--country', type=str, help='Country for job search')
    parser.add_argument('--export', type=str, help='Export filename')
    parser.add_argument('--applicant', type=str, help='Applicant name (for profile export)')
    
    args = parser.parse_args()
    
    if args.profile:
        # Crawl profile
        crawler = LinkedInProfileCrawler()
        profile = crawler.crawl_profile(args.profile)
        
        if profile:
            exporter = LinkedInDataExporter()
            exporter.export_to_json(profile, f"profile_{profile.get('username', 'unknown')}.json")
            
            if args.applicant:
                exporter.export_profile_for_migration(profile, args.applicant)
    
    elif args.jobs:
        # Search jobs
        crawler = LinkedInJobsCrawler()
        jobs = crawler.search_jobs(args.jobs, args.country or "")
        
        if jobs:
            exporter = LinkedInDataExporter()
            exporter.export_to_json(jobs, f"jobs_{args.jobs.replace(' ', '_')}.json")
    
    else:
        print("📋 LinkedIn Crawler — Usage:")
        print("  python linkedin_crawler.py --profile https://www.linkedin.com/in/username")
        print("  python linkedin_crawler.py --jobs 'midwife' --country 'New Zealand'")
        print("  python linkedin_crawler.py --profile URL --applicant 'John Doe'")


if __name__ == "__main__":
    main()
