import streamlit as st
import tempfile
import os
from io import BytesIO
import pandas as pd
import re
import PyPDF2
from typing import List, Dict
import datetime

# NEW: for DOCX support
from docx import Document

# --------------------------
# Utilities: PDF/DOCX -> text
# --------------------------
def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """Extract all text from PDF file"""
    try:
        reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
    except Exception:
        return ""
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)

def extract_text_from_docx_bytes(docx_bytes: bytes) -> str:
    """Extract text from DOCX (paragraphs + tables)"""
    try:
        doc = Document(BytesIO(docx_bytes))
    except Exception:
        return ""
    parts = []
    # paragraphs
    for p in doc.paragraphs:
        if p.text:
            parts.append(p.text)
    # tables
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text and cell.text.strip()]
            if row_text:
                parts.append(" | ".join(row_text))
    return "\n".join(parts)

def extract_text_from_upload(file_name: str, file_bytes: bytes) -> str:
    """Dispatch extraction based on extension"""
    name = file_name.lower()
    if name.endswith(".pdf"):
        return extract_text_from_pdf_bytes(file_bytes)
    elif name.endswith(".docx"):  # .doc not supported reliably without external deps
        return extract_text_from_docx_bytes(file_bytes)
    else:
        return ""

# --------------------------
# Section extraction
# --------------------------
def extract_relevant_sections(text: str) -> str:
    """
    Extract only Projects, Internships, and Experience sections from resume
    Returns concatenated text from these sections only
    """
    text_lower = text.lower()

    section_keywords = [
        r'\bprojects?\b',
        r'\bexperience\b',
        r'\bwork\s+experience\b',
        r'\bprofessional\s+experience\b',
        r'\binternships?\b',
        r'\btraining\b',
        r'\bindustrial\s+training\b',
        r'\bpractical\s+experience\b',
        r'\bwork\s+history\b',
        r'\bemployment\b',
        r'\bprofessional\s+background\b'
    ]

    exclusion_keywords = [
        r'\beducation\b',
        r'\bskills?\b',
        r'\btechnical\s+skills?\b',
        r'\bcertifications?\b',
        r'\bawards?\b',
        r'\bachievements?\b',
        r'\bhobbies\b',
        r'\binterests?\b',
        r'\blanguages?\b',
        r'\breferences?\b',
        r'\bdeclaration\b',
        r'\bpersonal\s+details?\b',
        r'\bcontact\b'
    ]

    sections = []
    for keyword_pattern in section_keywords:
        for match in re.finditer(keyword_pattern, text_lower):
            sections.append(('relevant', match.start(), match.group()))
    for keyword_pattern in exclusion_keywords:
        for match in re.finditer(keyword_pattern, text_lower):
            sections.append(('exclude', match.start(), match.group()))

    sections.sort(key=lambda x: x[1])

    relevant_text_parts = []
    for i, (section_type, start_pos, keyword) in enumerate(sections):
        if section_type == 'relevant':
            end_pos = len(text)
            for j in range(i + 1, len(sections)):
                _, next_start, _ = sections[j]
                end_pos = next_start
                break
            relevant_text_parts.append(text[start_pos:end_pos])

    if not relevant_text_parts:
        return ""
    return "\n".join(relevant_text_parts)

# --------------------------
# Information extraction
# --------------------------
def extract_name_from_filename(filename: str) -> str:
    """
    Extract and format name from filename.
    Converts formats like:
    - Piyush_Raj_Lenka.pdf -> Piyush Raj Lenka
    - Piyush-Raj-Lenka.docx -> Piyush Raj Lenka
    - PiyushRajLenka.pdf -> Piyushrajlenka (keeps as is if no separators)
    """
    # Remove file extension
    name_without_ext = os.path.splitext(filename)[0]
    
    # Replace underscores and hyphens with spaces
    name_with_spaces = name_without_ext.replace('_', ' ').replace('-', ' ')
    
    # Remove any extra whitespace and capitalize each word properly
    name_parts = name_with_spaces.split()
    formatted_name = ' '.join(word.capitalize() for word in name_parts if word)
    
    return formatted_name if formatted_name else filename

def extract_name(text: str) -> str:
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if not lines:
        return ""
    first_line = re.sub(r'^(resume|curriculum vitae|cv)[\s:]*', '', lines[0], flags=re.I)
    words = first_line.split()
    if 2 <= len(words) <= 4 and all(word.replace('.', '').isalpha() for word in words):
        return first_line
    if len(lines) > 1:
        second_line = lines[1]
        words = second_line.split()
        if 2 <= len(words) <= 4 and all(word.replace('.', '').isalpha() for word in words):
            return second_line
    return first_line[:50]

def extract_email(text: str) -> str:
    """
    Ultra-robust email extraction that handles all common resume formats.
    Works with PDFs, Word docs, and various text extraction issues.
    """
    if not text:
        return ""
    
    # Step 1: Clean and prepare text
    # Replace newlines with spaces to handle multi-line emails
    text_cleaned = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    
    # Also keep original text for some patterns
    text_original = text
    
    all_emails = []
    
    # Pattern 1: Standard email format (most reliable)
    pattern1 = r'\b[a-zA-Z0-9][a-zA-Z0-9._-]*[a-zA-Z0-9]@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,}\b'
    emails1 = re.findall(pattern1, text_cleaned, re.IGNORECASE)
    all_emails.extend(emails1)
    
    # Pattern 2: Email with potential spaces (PDF extraction issue)
    # Example: "user @domain.com" or "user@ domain.com"
    pattern2 = r'\b([a-zA-Z0-9][a-zA-Z0-9._-]*)\s*@\s*([a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,})\b'
    emails2 = re.findall(pattern2, text_cleaned, re.IGNORECASE)
    for username, domain in emails2:
        email = f"{username.strip()}@{domain.strip()}"
        all_emails.append(email)
    
    # Pattern 3: Email after common labels (case-insensitive)
    label_patterns = [
        r'(?:email|e-mail|mail|e mail|email id|email address)\s*[:\-]?\s*([a-zA-Z0-9][a-zA-Z0-9._-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,})',
        r'(?:contact|write to|reach me)\s*[:\-]?\s*([a-zA-Z0-9][a-zA-Z0-9._-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,})',
    ]
    for pattern in label_patterns:
        emails3 = re.findall(pattern, text_cleaned, re.IGNORECASE)
        all_emails.extend(emails3)
    
    # Pattern 4: Email in parentheses, brackets, or quotes
    pattern4 = r'[\(\[\{<"\']([a-zA-Z0-9][a-zA-Z0-9._-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,})[\)\]\}>"\']'
    emails4 = re.findall(pattern4, text_cleaned, re.IGNORECASE)
    all_emails.extend(emails4)
    
    # Pattern 5: More flexible pattern for difficult cases
    # This handles cases where email might have extra characters around it
    pattern5 = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    emails5 = re.findall(pattern5, text_cleaned, re.IGNORECASE)
    all_emails.extend(emails5)
    
    # Pattern 6: Handle emails split across lines in original text
    # Look for @ symbol and grab context around it
    lines = text_original.split('\n')
    for i, line in enumerate(lines):
        if '@' in line:
            # Combine current line with previous and next if they exist
            context = ''
            if i > 0:
                context += lines[i-1] + ' '
            context += line
            if i < len(lines) - 1:
                context += ' ' + lines[i+1]
            
            # Now search in this context
            emails6 = re.findall(pattern1, context, re.IGNORECASE)
            all_emails.extend(emails6)
    
    # Pattern 7: Handle emails with [at] or (at) instead of @
    pattern7 = r'\b([a-zA-Z0-9][a-zA-Z0-9._-]*)\s*[\[\(]?\s*at\s*[\]\)]?\s*([a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,})\b'
    emails7 = re.findall(pattern7, text_cleaned, re.IGNORECASE)
    for username, domain in emails7:
        email = f"{username.strip()}@{domain.strip()}"
        all_emails.append(email)
    
    # Pattern 8: Handle emails with [dot] or (dot) instead of .
    pattern8 = r'\b([a-zA-Z0-9][a-zA-Z0-9._-]*@[a-zA-Z0-9][a-zA-Z0-9-]*)\s*[\[\(]?\s*dot\s*[\]\)]?\s*([a-zA-Z]{2,})\b'
    emails8 = re.findall(pattern8, text_cleaned, re.IGNORECASE)
    for base, tld in emails8:
        email = f"{base.strip()}.{tld.strip()}"
        all_emails.append(email)
    
    # Clean and validate all found emails
    validated_emails = []
    for email in all_emails:
        # Remove any whitespace
        email = ''.join(email.split())
        
        # Convert to lowercase for consistency
        email = email.lower()
        
        # Remove any trailing/leading special characters
        email = email.strip('.,;:!?\'"()[]{}\\/<>|')
        
        # Validate email format
        if not email or len(email) < 6:  # Minimum: a@b.co
            continue
            
        if '@' not in email:
            continue
        
        # Split into username and domain
        parts = email.split('@')
        if len(parts) != 2:
            continue
            
        username, domain = parts
        
        # Validate username (at least 1 char, only valid characters)
        if not username or not re.match(r'^[a-zA-Z0-9._-]+$', username):
            continue
        
        # Validate domain (must have at least one dot and valid TLD)
        if '.' not in domain:
            continue
            
        domain_parts = domain.split('.')
        if len(domain_parts) < 2:
            continue
            
        # Check TLD is at least 2 characters
        tld = domain_parts[-1]
        if len(tld) < 2 or not tld.isalpha():
            continue
        
        # Check domain name is valid
        domain_name = '.'.join(domain_parts[:-1])
        if not domain_name or not re.match(r'^[a-zA-Z0-9.-]+$', domain_name):
            continue
        
        # Avoid common false positives
        false_positives = ['example.com', 'domain.com', 'email.com', 'test.com', 'sample.com']
        if email in false_positives:
            continue
        
        validated_emails.append(email)
    
    # Remove duplicates while preserving order
    seen = set()
    unique_emails = []
    for email in validated_emails:
        if email not in seen:
            seen.add(email)
            unique_emails.append(email)
    
    # Return the first valid email found
    return unique_emails[0] if unique_emails else ""

def extract_phone(text: str) -> str:
    """
    Enhanced phone number extraction with multiple patterns.
    Handles Indian and international phone formats with various separators.
    """
    if not text:
        return ""
    
    # Remove extra whitespace and newlines for better matching
    text_normalized = ' '.join(text.split())
    
    # Comprehensive phone patterns
    phone_patterns = [
        # International format with country code
        r'\+\d{1,3}[\s.-]?\d{3,5}[\s.-]?\d{3,5}[\s.-]?\d{3,4}',
        r'\+\d{1,3}[\s.-]?\(?\d{3,5}\)?[\s.-]?\d{3,5}[\s.-]?\d{3,4}',
        
        # Indian format with +91
        r'\+91[\s.-]?\d{10}',
        r'\+91[\s.-]?\d{5}[\s.-]?\d{5}',
        r'\+91[\s.-]?\d{4}[\s.-]?\d{3}[\s.-]?\d{3}',
        r'\+91[\s.-]?\(?\d{3,5}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}',
        
        # Indian format without +
        r'\b91[\s.-]?\d{10}\b',
        r'\b0091[\s.-]?\d{10}\b',
        
        # 10-digit phone numbers (Indian mobile)
        r'\b[6-9]\d{9}\b',
        r'\b[6-9]\d{4}[\s.-]\d{5}\b',
        
        # With parentheses
        r'\(?\d{3,5}\)?[\s.-]?\d{3,5}[\s.-]?\d{3,4}',
        
        # After common labels
        r'(?:phone|mobile|cell|contact|tel|telephone)[\s:]*[\+]?[\d\s\(\)\.-]{10,}',
        
        # Separated by spaces or dashes
        r'\b\d{3}[\s.-]\d{3}[\s.-]\d{4}\b',
        r'\b\d{4}[\s.-]\d{3}[\s.-]\d{3}\b',
        r'\b\d{5}[\s.-]\d{5}\b',
        
        # Just digits (10 or more)
        r'\b\d{10,12}\b',
    ]
    
    all_phones = []
    
    for pattern in phone_patterns:
        matches = re.findall(pattern, text_normalized, re.IGNORECASE)
        for match in matches:
            # Clean the phone number
            phone = re.sub(r'[^\d+]', '', match)
            
            # Validate phone number
            # Must have at least 10 digits
            digits_only = re.sub(r'\D', '', phone)
            
            if len(digits_only) >= 10:
                # Skip if it looks like a date, year, or pin code
                if not re.match(r'^(19|20)\d{2}', digits_only) and \
                   not re.match(r'^\d{6}$', digits_only) and \
                   not re.match(r'^[01]\d{9}$', digits_only):
                    all_phones.append(match.strip())
    
    # Remove duplicates while preserving order
    seen = set()
    unique_phones = []
    for phone in all_phones:
        phone_normalized = re.sub(r'\D', '', phone)
        if phone_normalized not in seen:
            seen.add(phone_normalized)
            unique_phones.append(phone)
    
    # Return the first valid phone found
    return unique_phones[0] if unique_phones else ""

def extract_education(text: str) -> str:
    """
    Enhanced education extraction with comprehensive degree detection.
    Captures full degree information including specialization and institution.
    """
    if not text:
        return ""
    
    text_lower = text.lower()
    
    # Step 1: Find education section
    education_keywords = [
        r'\beducation\b',
        r'\bacademic\s+background\b',
        r'\bacademic\s+qualification\b',
        r'\bqualifications?\b',
        r'\beducational\s+background\b',
        r'\beducational\s+qualification\b',
        r'\bacademics\b'
    ]
    
    education_start = -1
    for keyword in education_keywords:
        match = re.search(keyword, text_lower)
        if match:
            education_start = match.start()
            break
    
    # If education section found, extract from there
    if education_start != -1:
        # Find next major section to limit search
        next_section_keywords = [
            r'\bexperience\b',
            r'\bwork\s+experience\b',
            r'\bprofessional\s+experience\b',
            r'\bemployment\b',
            r'\bskills?\b',
            r'\btechnical\s+skills?\b',
            r'\bprojects?\b',
            r'\bcertifications?\b',
            r'\bawards?\b',
            r'\bachievements?\b'
        ]
        
        education_end = len(text)
        for keyword in next_section_keywords:
            match = re.search(keyword, text_lower[education_start + 50:])
            if match:
                education_end = education_start + 50 + match.start()
                break
        
        education_text = text[education_start:education_end]
    else:
        # No explicit section, search entire document
        education_text = text
    
    # Step 2: Comprehensive degree patterns
    degree_patterns = [
        # Full degree with specialization
        r'((?:Bachelor|Master|B\.?Tech|M\.?Tech|B\.?E\.?|M\.?E\.?|B\.?Sc|M\.?Sc|BCA|MCA|MBA|PhD|Ph\.?D\.?|Doctorate)[^,\n]{0,100}(?:in|of|from)[^,\n]{0,100})',
        
        # Degree with year and institution
        r'((?:Bachelor|Master|B\.?Tech|M\.?Tech|B\.?E\.?|M\.?E\.?|B\.?Sc|M\.?Sc|BCA|MCA|MBA|PhD|Ph\.?D\.?)[^,\n]{0,150}(?:19|20)\d{2})',
        
        # Degree with CGPA/Percentage
        r'((?:Bachelor|Master|B\.?Tech|M\.?Tech|B\.?E\.?|M\.?E\.?|B\.?Sc|M\.?Sc|BCA|MCA|MBA)[^,\n]{0,100}(?:CGPA|GPA|Percentage|%|\d+\.\d+))',
        
        # Standard degree names
        r'(Bachelor\s+of\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
        r'(Master\s+of\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
        r'(B\.?Tech\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(M\.?Tech\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(B\.?E\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(M\.?E\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(B\.?Sc\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(M\.?Sc\.?\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(BCA\s+[^\n]{0,50})',
        r'(MCA\s+[^\n]{0,50})',
        r'(MBA\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(PhD\s+(?:in\s+)?[A-Za-z\s&]+)',
        r'(Diploma\s+in\s+[A-Za-z\s&]+)',
        
        # Short forms
        r'\b(B\.?Tech|M\.?Tech|B\.?E|M\.?E|BCA|MCA|MBA|B\.?Sc|M\.?Sc|PhD)\b',
    ]
    
    all_degrees = []
    for pattern in degree_patterns:
        matches = re.findall(pattern, education_text, re.IGNORECASE)
        all_degrees.extend(matches)
    
    if all_degrees:
        # Clean and format the best match (longest and most complete)
        best_degree = max(all_degrees, key=len)
        best_degree = best_degree.strip()
        
        # Clean up extra whitespace
        best_degree = re.sub(r'\s+', ' ', best_degree)
        
        # Remove trailing incomplete text
        best_degree = re.sub(r'\s+[a-z]+$', '', best_degree)
        
        # Capitalize properly
        # Keep acronyms uppercase, capitalize other words
        words = best_degree.split()
        formatted_words = []
        for word in words:
            # If word is likely an acronym (short and uppercase)
            if len(word) <= 6 and word.upper() == word and '.' in word:
                formatted_words.append(word.upper())
            elif word.upper() in ['B.TECH', 'M.TECH', 'B.E', 'M.E', 'BCA', 'MCA', 'MBA', 'BSC', 'MSC', 'PHD']:
                formatted_words.append(word.upper())
            elif word.lower() in ['in', 'of', 'and', 'from', 'the', 'with']:
                formatted_words.append(word.lower())
            else:
                formatted_words.append(word.capitalize())
        
        return ' '.join(formatted_words)[:150]  # Limit length
    
    return ""

def extract_location(text: str) -> str:
    """
    Enhanced location extraction from resume.
    Checks multiple patterns and contexts to find the most likely location.
    """
    if not text:
        return ""
    
    # Strategy 1: Look for explicit location labels in first 500 chars
    top_text = text[:500]
    
    location_label_patterns = [
        r'(?:Location|Address|City|Residence|Based in|Current Location|Permanent Address|Present Address)\s*[:\-]?\s*([A-Z][a-zA-Z\s,\-]+(?:\d{5,6})?)',
        r'(?:Location|Address|City)\s*[:\-]?\s*([^\n]{5,50})',
    ]
    
    for pattern in location_label_patterns:
        matches = re.findall(pattern, top_text, re.IGNORECASE)
        if matches:
            location = matches[0].strip()
            # Clean up
            location = re.sub(r'\s+', ' ', location)
            # Remove phone/email if accidentally captured
            if '@' in location or re.search(r'\d{10}', location):
                continue
            # Remove trailing junk
            location = re.sub(r'[,\s]+$', '', location)
            if len(location.split()) <= 6 and len(location) <= 60:
                return location
    
    # Strategy 2: Common location patterns (City, State or City-PIN)
    location_patterns = [
        # City, State format
        r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?),\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b',
        # City - PIN format (Indian)
        r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s*[-–]\s*(\d{5,6})\b',
        # City, State, Country
        r'\b([A-Z][a-z]+),\s+([A-Z][a-z]+),\s+([A-Z][a-z]+)\b',
    ]
    
    # Search in top 1000 characters
    search_text = text[:1000]
    
    for pattern in location_patterns:
        matches = re.findall(pattern, search_text)
        if matches:
            if isinstance(matches[0], tuple):
                location = ', '.join([part.strip() for part in matches[0] if part.strip()])
            else:
                location = matches[0].strip()
            
            # Validate it's not part of education/company name
            if not re.search(r'\b(University|College|Institute|School|Company|Ltd|Inc|Corp)\b', location, re.IGNORECASE):
                return location
    
    # Strategy 3: Look in first 15 lines for a line that looks like a location
    lines = [l.strip() for l in text.split('\n')[:20] if l.strip()]
    
    for i, line in enumerate(lines):
        # Skip first 2 lines (likely name/title)
        if i < 2:
            continue
        
        # Check if line looks like a location
        # Has comma, no @, no excessive digits, reasonable length
        if (',' in line and 
            '@' not in line and 
            not re.search(r'\d{10}', line) and  # Not phone
            len(line.split()) >= 2 and 
            len(line.split()) <= 6 and
            len(line) <= 60):
            
            # Should start with capital letter
            if line[0].isupper():
                # Clean up
                location = re.sub(r'\s+', ' ', line)
                location = re.sub(r'[,\s]+$', '', location)
                
                # Validate not a sentence
                if not re.search(r'\b(is|am|are|was|were|have|has|had|the|a|an)\b', location.lower()):
                    return location
    
    # Strategy 4: Look for Indian cities in the text
    indian_cities = [
        'Mumbai', 'Delhi', 'Bangalore', 'Bengaluru', 'Hyderabad', 'Chennai', 'Kolkata',
        'Pune', 'Ahmedabad', 'Surat', 'Jaipur', 'Lucknow', 'Kanpur', 'Nagpur',
        'Indore', 'Thane', 'Bhopal', 'Visakhapatnam', 'Pimpri', 'Patna', 'Vadodara',
        'Ghaziabad', 'Ludhiana', 'Agra', 'Nashik', 'Faridabad', 'Meerut', 'Rajkot',
        'Varanasi', 'Srinagar', 'Aurangabad', 'Dhanbad', 'Amritsar', 'Navi Mumbai',
        'Allahabad', 'Ranchi', 'Howrah', 'Coimbatore', 'Jabalpur', 'Gwalior', 'Noida',
        'Panvel', 'Thane', 'Greater Noida', 'Gurgaon', 'Gurugram'
    ]
    
    for city in indian_cities:
        # Look for city in first 1000 chars
        city_pattern = r'\b' + re.escape(city) + r'\b[,\s]+([A-Z][a-z]+)?'
        match = re.search(city_pattern, search_text, re.IGNORECASE)
        if match:
            # Found city, try to get state too
            context_start = max(0, match.start() - 20)
            context_end = min(len(search_text), match.end() + 30)
            context = search_text[context_start:context_end]
            
            # Extract just the city and state
            location_match = re.search(r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?(?:,\s*[A-Z][a-z]+)?)', context, re.IGNORECASE)
            if location_match:
                return location_match.group(1).strip()
    
    return ""

def is_internship_entry(text_block: str) -> bool:
    """
    Determine if a work entry is an internship.
    Returns True if it's an internship, False otherwise.
    """
    text_lower = text_block.lower()
    
    # Keywords that indicate internship
    internship_keywords = [
        r'\bintern\b',
        r'\binternship\b',
        r'\btrainee\b',
        r'\btraining\b',
        r'\bsummer\s+intern\b',
        r'\bwinter\s+intern\b',
        r'\bindustrial\s+training\b',
        r'\bpractical\s+training\b',
        r'\bapprentice\b',
        r'\bapprenticeship\b',
    ]
    
    for keyword in internship_keywords:
        if re.search(keyword, text_lower):
            return True
    
    return False

def extract_total_experience(text: str) -> float:
    """
    Ultra-robust work experience extraction.
    Calculates ONLY full-time work experience, EXCLUDES internships.
    Uses multiple strategies to calculate accurately.
    """
    if not text:
        return 0.0
    
    now = datetime.datetime.now()
    current_year = now.year
    current_month = now.month
    
    # Month mapping
    month_map = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12,
    }
    
    # Strategy 1: Direct extraction of "X years of experience" (excluding internship mentions)
    direct_patterns = [
        r'(\d+\.?\d*)\s*\+?\s*(years?|yrs?)\s*(?:of)?\s*(?:total\s+)?(?:work\s+)?(?:professional\s+)?experience',
        r'(?:total\s+)?(?:work\s+)?experience\s*[:\-]?\s*(\d+\.?\d*)\s*\+?\s*(years?|yrs?)',
        r'(\d+\.?\d*)\s*\+?\s*(years?|yrs?)\s*(?:of\s+)?(?:professional\s+)?(?:work\s+)?experience',
    ]
    
    # Look for these in sections NOT labeled as internship
    text_lower = text.lower()
    
    # Only trust direct statements if they're NOT in an internship context
    for pattern in direct_patterns:
        matches = re.findall(pattern, text_lower)
        if matches:
            max_exp = 0.0
            for match in matches:
                # Get context around the match to check if it's internship
                match_text = str(match)
                match_pos = text_lower.find(match_text)
                if match_pos != -1:
                    context_start = max(0, match_pos - 100)
                    context_end = min(len(text_lower), match_pos + 100)
                    context = text_lower[context_start:context_end]
                    
                    # Skip if in internship context
                    if is_internship_entry(context):
                        continue
                
                if isinstance(match, tuple):
                    num = float(match[0])
                    unit = match[1] if len(match) > 1 else 'years'
                else:
                    num = float(match)
                    unit = 'years'
                
                if 'month' in unit:
                    num /= 12
                max_exp = max(max_exp, num)
            
            if max_exp > 0:
                return round(max_exp, 1)
    
    # Strategy 2: Calculate from date ranges, EXCLUDING internships
    # Find experience section
    exp_keywords = [
        r'\bwork\s+experience\b',
        r'\bprofessional\s+experience\b',
        r'\bemployment\s+history\b',
        r'\bwork\s+history\b',
        r'\bcareer\s+history\b',
        r'\bprofessional\s+background\b',
        r'\bexperience\b'
    ]
    
    exp_start = -1
    for keyword in exp_keywords:
        match = re.search(keyword, text_lower)
        if match:
            exp_start = match.start()
            break
    
    if exp_start == -1:
        # Try with the whole document
        exp_text = text
    else:
        # Find end of experience section
        end_keywords = [
            r'\beducation\b',
            r'\bskills?\b',
            r'\btechnical\s+skills?\b',
            r'\bcertifications?\b',
            r'\bprojects?\b',
            r'\bawards?\b',
            r'\bachievements?\b',
            r'\binternships?\b'  # Separate internship section
        ]
        
        exp_end = len(text)
        for keyword in end_keywords:
            match = re.search(keyword, text_lower[exp_start + 50:])
            if match:
                exp_end = exp_start + 50 + match.start()
                break
        
        exp_text = text[exp_start:exp_end]
    
    # Split experience section into individual job entries
    # Look for patterns that indicate new job entries
    job_separators = [
        r'\n\s*\n',  # Double newline
        r'\n(?=[A-Z][a-zA-Z\s,]+\|)',  # New line starting with company name
        r'\n(?=[A-Z][a-zA-Z\s&]+(?:Ltd|Inc|Corp|Pvt|Private|Limited))',  # Company with suffix
    ]
    
    # For now, split by double newlines or major gaps
    job_entries = re.split(r'\n\s*\n+', exp_text)
    
    # Comprehensive date range patterns
    date_patterns = [
        # Format: YYYY - YYYY
        (r'(\d{4})\s*[-–—to]\s*(\d{4})', 'year_year'),
        
        # Format: YYYY - Present/Current
        (r'(\d{4})\s*[-–—to]\s*(present|current|till\s+date|ongoing|now|till\s+now)', 'year_present'),
        
        # Format: Month YYYY - Month YYYY
        (r'([a-z]{3,9})\s+(\d{4})\s*[-–—to]\s*([a-z]{3,9})\s+(\d{4})', 'month_year_month_year'),
        
        # Format: Month YYYY - Present
        (r'([a-z]{3,9})\s+(\d{4})\s*[-–—to]\s*(present|current|till\s+date|ongoing|now)', 'month_year_present'),
        
        # Format: MM/YYYY - MM/YYYY
        (r'(\d{1,2})/(\d{4})\s*[-–—to]\s*(\d{1,2})/(\d{4})', 'mm_yyyy_mm_yyyy'),
        
        # Format: MM/YYYY - Present
        (r'(\d{1,2})/(\d{4})\s*[-–—to]\s*(present|current|till\s+date|ongoing|now)', 'mm_yyyy_present'),
        
        # Format: Month, YYYY - Month, YYYY
        (r'([a-z]{3,9}),?\s+(\d{4})\s*[-–—to]\s*([a-z]{3,9}),?\s+(\d{4})', 'month_comma_year'),
        
        # Format: YYYY-MM - YYYY-MM
        (r'(\d{4})-(\d{2})\s*[-–—to]\s*(\d{4})-(\d{2})', 'yyyy_mm_yyyy_mm'),
    ]
    
    all_durations = []
    
    for job_entry in job_entries:
        # Check if this entry is an internship - if so, skip it
        if is_internship_entry(job_entry):
            continue
        
        # Process this non-internship entry for date ranges
        for pattern, pattern_type in date_patterns:
            matches = re.findall(pattern, job_entry, re.IGNORECASE)
            
            for match in matches:
                try:
                    start_year = start_month = end_year = end_month = 0
                    
                    if pattern_type == 'year_year':
                        start_year = int(match[0])
                        end_year = int(match[1])
                        start_month = 1
                        end_month = 12
                    
                    elif pattern_type == 'year_present':
                        start_year = int(match[0])
                        end_year = current_year
                        start_month = 1
                        end_month = current_month
                    
                    elif pattern_type == 'mm_yyyy_present':
                        start_month = int(match[0])
                        start_year = int(match[1])
                        end_year = current_year
                        end_month = current_month
                    
                    elif pattern_type == 'month_year_present':
                        start_month = month_map.get(match[0].lower()[:3], 1)
                        start_year = int(match[1])
                        end_year = current_year
                        end_month = current_month
                    
                    elif pattern_type in ['month_year_month_year', 'month_comma_year']:
                        start_month = month_map.get(match[0].lower()[:3], 1)
                        start_year = int(match[1])
                        end_month = month_map.get(match[2].lower()[:3], 12)
                        end_year = int(match[3])
                    
                    elif pattern_type == 'mm_yyyy_mm_yyyy':
                        start_month = int(match[0])
                        start_year = int(match[1])
                        end_month = int(match[2])
                        end_year = int(match[3])
                    
                    elif pattern_type == 'yyyy_mm_yyyy_mm':
                        start_year = int(match[0])
                        start_month = int(match[1])
                        end_year = int(match[2])
                        end_month = int(match[3])
                    
                    # Validate years
                    if start_year < 1970 or start_year > current_year:
                        continue
                    if end_year < start_year or end_year > current_year + 1:
                        continue
                    
                    # Calculate duration
                    if start_year and end_year:
                        years = end_year - start_year
                        months = (end_month - start_month) / 12
                        duration = years + months
                        
                        if duration > 0 and duration <= 50:  # Sanity check
                            all_durations.append(duration)
                
                except (ValueError, KeyError):
                    continue
    
    # Calculate total work experience (excluding internships)
    if all_durations:
        # Sum all durations (in case of multiple jobs)
        total_exp = sum(all_durations)
        
        # But cap at 50 years for sanity
        total_exp = min(total_exp, 50)
        
        return round(total_exp, 1)
    
    return 0.0

# --------------------------
# Skill / Industry matching
# --------------------------
def normalize_skill_list(skills: List[str]) -> List[str]:
    """Deduplicate case-insensitive while preserving order and trim whitespace."""
    seen = set()
    normalized = []
    for s in skills:
        key = s.strip()
        if not key:
            continue
        lower = key.lower()
        if lower not in seen:
            seen.add(lower)
            # Use title-case for display but preserve some acronyms fully uppercase (heuristic)
            if key.isupper() and len(key) <= 5:
                display = key  # keep short all-caps like "IT", "SAAS"
            else:
                display = key.title()
            normalized.append(display)
    return normalized

# A dictionary of regex patterns (list) to match common variations for industries/verticals
INDUSTRY_PATTERNS = {
    'Pharma': [r'\bpharma\b', r'\bpharmaceuticals?\b', r'\bpharmaceutical\b'],
    'Hospitality': [r'\bhospitalit(y|ies)\b', r'\bhotels?\b', r'\bfood\s+and\s+beverage\b', r'\bfnb\b'],
    'Enterprise Software': [r'\benterprise[\s\-]?software\b', r'\benterprise\s+apps?\b', r'\benterprise\s+solutions?\b'],
    'Real Estate': [r'\breal[\s\-]?estate\b', r'\bproperty\s+development\b', r'\bproperty\s+management\b'],
    'Agritech': [r'\bagritech\b', r'\bagri[\s\-]?tech\b', r'\bagriculture\b', r'\bfarming\b'],
    'Sales': [r'\bsales\b', r'\bsales\s+professional\b', r'\bsales\s+executive\b'],
    'Business Development': [r'\bbusiness\s+development\b', r'\bbd\s+manager\b', r'\bbusiness\s+dev\b', r'\bbd\b'],
    'HoReCa': [r'\bhoreca\b', r'\bhoreca\b', r'\bhotel\s+restaurant\s+cafe\b'],
    'Banking': [r'\bbank(ing)?\b', r'\bfinancial\s+services\b'],
    'FMCG': [r'\bfmcg\b', r'\bfast\s+moving\s+consumer\s+goods\b'],
    'TELECOM': [r'\btelecom\b', r'\btelecommunications?\b', r'\btelecoms?\b'],
    'INSURANCE': [r'\binsurance\b', r'\binsurance\s+industry\b'],
    'FINTECH': [r'\bfintech\b', r'\bfinancial\s+technology\b'],
    'IT': [r'\bit\s+sector\b', r'\binformation\s+technology\b', r'\bit\s+services\b', r'\btechnology\s+company\b'],
    'SAAS': [r'\bsaas\b', r'\bsoftware\s+as\s+a\s+service\b'],
    'B2b': [r'\bb2b\b', r'\bb2\b[\s\-]?b\b', r'\bbusiness\-to\-business\b'],
    'Edtech': [r'\bedtech\b', r'\beducation\s+technology\b', r'\beducational\s+technology\b'],
    'BFSI': [r'\bbfsi\b', r'\bbanking\s+finance\s+and\s+insurance\b'],
    'Logistic': [r'\blogistic(s)?\b', r'\bsupply\s+chain\b', r'\blogistics?\b'],
    'ECommerce': [r'\be[\s\-]?commerce\b', r'\becommerce\b', r'\bonline\s+retail\b']
}

def check_skill_present(text: str, skill_display_name: str) -> int:
    """Return 1 if any pattern for the skill matches in text (case-insensitive), else 0."""
    if not text or not skill_display_name:
        return 0
    patterns = INDUSTRY_PATTERNS.get(skill_display_name, [r'\b' + re.escape(skill_display_name.lower()) + r'\b'])
    for pat in patterns:
        if re.search(pat, text, flags=re.I):
            return 1
    return 0

# --------------------------
# Resume processing
# --------------------------
def process_single_resume(file_bytes: bytes, filename: str, skills_to_check: List[str]) -> Dict:
    """Process a single resume and return extracted data. Skills are checked in the entire resume text."""
    full_text = extract_text_from_upload(filename, file_bytes)
    if not full_text.strip():
        return None
    
    # Extract name from filename
    name = extract_name_from_filename(filename)
    
    email = extract_email(full_text)
    phone = extract_phone(full_text)
    education = extract_education(full_text)
    location = extract_location(full_text)
    total_experience = extract_total_experience(full_text)

    data = {
        'Filename': filename,
        'Name': name,
        'Email': email,
        'Phone Number': phone,
        'Education': education,
        'Location': location,
        'Total Years of Work Experience': total_experience  # Updated column name
    }

    # Check each industry/vertical against the WHOLE resume (full_text)
    for skill in skills_to_check:
        data[skill] = check_skill_present(full_text, skill)

    return data

# --------------------------
# Excel generation
# --------------------------
def generate_excel_from_data(all_data: List[Dict]) -> bytes:
    df = pd.DataFrame(all_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Resume_Data")
    return output.getvalue()

# --------------------------
# Streamlit UI
# --------------------------
st.set_page_config(page_title="Resume Industry Extractor", layout="wide")
st.title("📄 Batch Resume → Industry/Vertical Extractor (Work Experience Only)")
st.markdown(
    """
    Upload **multiple resumes (PDF or Word .docx)** and get a **single Excel file** with:
    - **Highly accurate extraction** of Name, Email, Phone, Education, Location
    - **Total Years of Work Experience** (Internships NOT included)
    - Industry/vertical presence indicators (1 if present, 0 if not) - **searched across the entire resume text**
    """
)

# --------------------------
# YOUR UPDATED INDUSTRY / VERTICAL LIST (normalized & deduped)
# --------------------------
RAW_SKILLS = [
    "Pharma",
    "Hospitality",
    "Enterprise software",
    "Real Estate",
    "Agritech",
    "SALES",
    "Business Development",
    "HoReCa",
    "Banking",
    "FMCG",
    "TELECOM",
    "INSURANCE",
    "FINTECH",
    "IT",
    "SAAS",
    "B2b",
    "sales",
    "EdTech",
    "BFSI",
    "Logistic",
    "ECommerce"
]

SKILLS_TO_CHECK = normalize_skill_list(RAW_SKILLS)

col1, col2 = st.columns([1, 2])

with col1:
    uploaded_files = st.file_uploader(
        "Upload Resumes (PDF or DOCX) — up to 100 files",
        type=["pdf", "docx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        # Hard cap at 100 files
        if len(uploaded_files) > 100:
            st.warning(f"⚠️ You uploaded {len(uploaded_files)} files. Processing the first 100.")
            uploaded_files = uploaded_files[:100]
        st.success(f"✅ {len(uploaded_files)} file(s) ready to process")

    st.info(
        "**Enhanced Extraction Features:**\n\n"
        "✨ **Name**: From filename (formatted)\n"
        "✨ **Email**: 8 pattern strategies\n"
        "✨ **Phone**: Indian + International formats\n"
        "✨ **Education**: Full degree with specialization\n"
        "✨ **Location**: 4 detection strategies\n"
        "✨ **Work Experience**: Full-time jobs ONLY (internships excluded)\n"
        "✨ **Industries**: Whole-resume matching\n\n"
        "📎 Supported: **PDF**, **Word (.docx)**"
    )

    process_button = st.button("🚀 Process All Resumes", type="primary")

with col2:
    st.header("How it works")
    st.write("1. **Upload** multiple resumes (PDF or DOCX).")
    st.write("2. Click **Process All Resumes**.")
    st.write("3. **Download** the Excel file with all extracted information.")
    st.write("")
    st.success("✅ **Work Experience Only**: Internships are automatically excluded from the total!")
    st.write("")

    st.write("**Industries / Verticals Checked:**")
    cols = st.columns(3)
    for idx, skill in enumerate(SKILLS_TO_CHECK):
        with cols[idx % 3]:
            st.write(f"  • {skill}")

# Main processing
if process_button:
    if not uploaded_files:
        st.error("⚠️ Please upload at least one resume first.")
    else:
        all_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Processing {idx + 1}/{len(uploaded_files)}: {uploaded_file.name}")
            try:
                raw_bytes = uploaded_file.read()
                data = process_single_resume(raw_bytes, uploaded_file.name, SKILLS_TO_CHECK)
                if data:
                    all_data.append(data)
                else:
                    st.warning(f"⚠️ Could not extract text from: {uploaded_file.name} (unsupported/empty/corrupt)")
            except Exception as e:
                st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.empty()
        progress_bar.empty()

        if not all_data:
            st.error("❌ Could not extract data from any of the uploaded files.")
        else:
            st.success(f"✅ Successfully processed {len(all_data)} out of {len(uploaded_files)} file(s)!")

            st.subheader("📊 Processing Summary")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Total Files Uploaded", len(uploaded_files))
            with col_b:
                st.metric("Successfully Processed", len(all_data))
            with col_c:
                st.metric("Failed", len(uploaded_files) - len(all_data))

            st.subheader("📈 Industry/Vertical Statistics (from whole resume)")
            skill_counts = {}
            for skill in SKILLS_TO_CHECK:
                count = sum(1 for data in all_data if data.get(skill) == 1)
                skill_counts[skill] = count

            num_cols = 4
            skill_items = list(skill_counts.items())
            for i in range(0, len(skill_items), num_cols):
                skill_cols = st.columns(num_cols)
                for j in range(num_cols):
                    if i + j < len(skill_items):
                        skill, count = skill_items[i + j]
                        with skill_cols[j]:
                            percentage = (count / len(all_data) * 100) if all_data else 0
                            st.metric(skill, f"{count}/{len(all_data)}", f"{percentage:.0f}%")

            st.subheader("Complete Data Table")
            df_display = pd.DataFrame(all_data)
            st.dataframe(df_display, use_container_width=True)

            with st.spinner("📝 Generating Excel file..."):
                excel_bytes = generate_excel_from_data(all_data)

            st.download_button(
                label="📥 Download Excel File with All Data",
                data=excel_bytes,
                file_name="batch_resume_industries_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.success("🎉 Processing complete! Click the button above to download your Excel file.")

st.markdown("---")
st.caption("Have a GOOD DAY!!!")    