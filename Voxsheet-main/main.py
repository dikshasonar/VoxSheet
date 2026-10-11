"""
VoxSheet — Voice-Activated Student Attendance and Roster Management System
===========================================================================
Single-file consolidated architecture featuring:
  - Precise Class Summary (When asking for TE, gives TE only; same for FE, SE, BE, or All)
  - Audio Voice Mute/Unmute Toggle (Speaking is optional; all tasks can be done via typing/clicking)
  - Modern Web Dashboard UI (Glassmorphism dark theme, interactive soundwave mic,
    live roster table, real-time stats cards, NLP entity breakdown, and Excel export)
  - Pure Natural Language Understanding (Automatically detects Intent and Entities
    directly from speech without requiring manual mode switching)
  - Strict Data Integrity Guard (Only valid student records saved to students.xlsx;
    conversational text, queries, and system commands are NEVER saved to Excel)
  - Thread-Safe Speech Synthesis (Windows SAPI5 with CoInitialize + pyttsx3 fallback)
  - Built-in Multi-threaded HTTP API Server (Zero external framework dependencies)
  - EntityExtractor: Order-independent Hybrid NLP (spaCy NER + Regex + Token Filtering)
  - IntentClassifier: Natural language intent recognition
  - RosterRepository: Excel database management (students.xlsx) with live audit timestamps
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
import http.server
import json
import os
import re
import socketserver
import sys
import threading
from typing import Any, Callable, Dict, List, Optional, Tuple
import urllib.parse
import webbrowser

import pandas as pd
import speech_recognition as sr


# =============================================================================
# 1. CONFIGURATION & CONSTANTS
# =============================================================================

class Config:
    """Central configuration for application paths, models, and defaults."""
    PORT: int = 5050
    BASE_DIR: str = os.path.dirname(os.path.abspath(__file__))
    EXCEL_FILE: str = os.path.join(BASE_DIR, "students.xlsx")
    PARENT_EXCEL_FILE: str = os.path.join(os.path.dirname(BASE_DIR), "students.xlsx")
    
    DEFAULT_CLASSES: Tuple[str, ...] = ("FE", "SE", "TE", "BE")
    COLUMNS: List[str] = ["Name", "RollNo", "Class", "Status", "Date"]
    
    MAX_DIALOG_RETRIES: int = 2
    MIC_TIMEOUT: int = 5
    MIC_PHRASE_TIME_LIMIT: int = 8
    
    DEFAULT_ROSTER: List[Dict[str, str]] = [
        {"Name": "Diksha", "RollNo": "49", "Class": "TE", "Status": "Absent", "Date": "-"},
        {"Name": "Isha Patil", "RollNo": "48", "Class": "TE", "Status": "Absent", "Date": "-"},
        {"Name": "Ayush Sharma", "RollNo": "34", "Class": "TE", "Status": "Absent", "Date": "-"},
        {"Name": "Shreya Deshmukh", "RollNo": "12", "Class": "TE", "Status": "Absent", "Date": "-"},
        {"Name": "Rohan Sharma", "RollNo": "12", "Class": "SE", "Status": "Absent", "Date": "-"},
        {"Name": "Aarav Mehta", "RollNo": "20", "Class": "SE", "Status": "Absent", "Date": "-"},
        {"Name": "Rahul Patil", "RollNo": "5", "Class": "FE", "Status": "Absent", "Date": "-"},
        {"Name": "Neha Kulkarni", "RollNo": "8", "Class": "FE", "Status": "Absent", "Date": "-"},
        {"Name": "Priya Deshmukh", "RollNo": "15", "Class": "BE", "Status": "Absent", "Date": "-"},
        {"Name": "Siddharth Joshi", "RollNo": "22", "Class": "BE", "Status": "Absent", "Date": "-"},
    ]


class Intent(str, Enum):
    REGISTER = "register"
    MARK_ATTENDANCE = "mark_attendance"
    SUMMARY = "summary"
    QUERY = "query"
    LOOKUP = "lookup"
    LOCATION_REJECTED = "location_rejected"
    UNKNOWN = "unknown"


@dataclass
class ExtractedEntities:
    """Holds parsed entities from natural language speech or text."""
    name: Optional[str] = None
    roll_no: Optional[str] = None
    class_name: Optional[str] = None

    def is_complete(self, required: Tuple[str, ...]) -> bool:
        return all(getattr(self, f, None) is not None for f in required)

    def missing_fields(self, required: Tuple[str, ...]) -> List[str]:
        labels = {"name": "Student Name", "roll_no": "Roll Number", "class_name": "Class"}
        return [labels.get(f, f) for f in required if getattr(self, f, None) is None]

    def update_from(self, other: ExtractedEntities) -> None:
        if not self.name and other.name:
            self.name = other.name
        if not self.roll_no and other.roll_no:
            self.roll_no = other.roll_no
        if not self.class_name and other.class_name:
            self.class_name = other.class_name

    def to_dict(self) -> Dict[str, Optional[str]]:
        return {"name": self.name, "roll_no": self.roll_no, "class": self.class_name}


# =============================================================================
# 2. VOICE SERVICE (TTS + STT - THREAD-SAFE)
# =============================================================================

class VoiceService:
    """Manages audio output (TTS) and audio input (STT) with graceful fallbacks."""

    def __init__(self):
        self.recognizer = sr.Recognizer()

    def speak(self, text: str) -> None:
        """Thread-safe speech synthesis with Windows SAPI5 / pyttsx3."""
        print(f"\n[SYSTEM VOICE]: {text}\n")

        # 1. Primary: Windows SAPI5 with per-thread COM initialization
        if sys.platform == "win32":
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                speaker.Speak(text)
                return
            except Exception:
                pass

        # 2. Secondary: pyttsx3 fallback
        try:
            import pyttsx3
            engine = pyttsx3.init()
            engine.setProperty("rate", 160)
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass

    def listen(self) -> Optional[str]:
        """Captures spoken audio from microphone and returns transcribed text."""
        try:
            with sr.Microphone() as source:
                print("🎤 Listening... please speak clearly into your microphone.")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.4)
                audio = self.recognizer.listen(
                    source,
                    timeout=Config.MIC_TIMEOUT,
                    phrase_time_limit=Config.MIC_PHRASE_TIME_LIMIT
                )

            text = self.recognizer.recognize_google(audio)
            print(f"👉 Transcribed: \"{text}\"")
            return text.strip()
        except (AttributeError, OSError) as e:
            print(f"[Microphone Notice]: Audio input device unavailable ({e}).")
            return None
        except sr.WaitTimeoutError:
            print("[STT Notice]: Listening timed out. No speech detected.")
            return None
        except sr.UnknownValueError:
            print("[STT Notice]: Speech was unclear.")
            return None
        except sr.RequestError as e:
            print(f"[STT Error]: Cloud recognition service issue ({e}).")
            return None
        except Exception as e:
            print(f"[STT Note]: {e}")
            return None


# =============================================================================
# 3. ENTITY EXTRACTION MODULE (HYBRID NLP / spaCy + REGEX)
# =============================================================================

class EntityExtractor:
    """
    Extracts structured fields (Name, Roll No, Class) from speech transcripts
    regardless of word order using regex patterns and spaCy NER.
    """

    def __init__(self):
        self._nlp = None
        self._init_nlp()
        self._init_regexes()

    def _init_nlp(self):
        import spacy
        try:
            self._nlp = spacy.load("en_core_web_sm")
        except Exception:
            try:
                import spacy.cli
                print("Downloading spaCy model 'en_core_web_sm'...")
                spacy.cli.download("en_core_web_sm")
                self._nlp = spacy.load("en_core_web_sm")
            except Exception as e:
                print(f"[NLP Notice]: Running in regex-only fallback mode ({e}).")
                self._nlp = None

    def _init_regexes(self):
        self.roll_patterns = [
            re.compile(r"roll\s*(?:no|number)?\.?\s*(?:is|was|=|:)?\s*[:\-]?\s*(\d+)", re.I),
            re.compile(r"roll\s*[:\-]?\s*(\d+)", re.I),
            re.compile(r"number\s*[:\-]?\s*(\d+)", re.I),
            re.compile(r"\b(\d+)\b"),
        ]

        self.intro_pattern = re.compile(
            r"(?:i am|i'm|my name is|this is|myself|name\s+is)\s+([a-zA-Z]+(?:\s+[a-zA-Z]+)?)"
            r"(?=\s+(?:roll|from|class|year|and|,|$)|$)",
            re.I
        )
        self.mark_pattern = re.compile(
            r"(?:mark|attendance\s+for)\s+([a-zA-Z]+(?:\s+[a-zA-Z]+)?)"
            r"(?=\s+(?:as|present|absent|roll|class|year|,|$)|$)",
            re.I
        )

        # Precise class extraction patterns supporting spelling and speech variants like "T E", "T.E."
        self.class_patterns = [
            ("TE", [
                r"\bt\.?\s*e\.?\b",
                r"\bthird\s*year\b",
                r"\b3rd\s*year\b",
                r"\b3rd\b",
                r"\bthird\b"
            ]),
            ("FE", [
                r"\bf\.?\s*e\.?\b",
                r"\bfirst\s*year\b",
                r"\b1st\s*year\b",
                r"\bfresher\b",
                r"\b1st\b",
                r"\bfirst\b"
            ]),
            ("SE", [
                r"\bs\.?\s*e\.?\b",
                r"\bsecond\s*year\b",
                r"\b2nd\s*year\b",
                r"\b2nd\b",
                r"\bsecond\b"
            ]),
            ("BE", [
                r"\bb\.?\s*e\.?\b",
                r"\bfinal\s*year\b",
                r"\bfourth\s*year\b",
                r"\b4th\s*year\b",
                r"\blast\s*year\b",
                r"\b4th\b",
                r"\bfourth\b"
            ]),
        ]

        # Stopwords & command words
        self.command_words = {
            "change", "intent", "intents", "switch", "mode", "option", "options",
            "select", "menu", "system", "voice", "sheet", "bot", "assistant",
            "help", "quit", "exit", "python", "user", "users", "audio", "mic",
            "summary", "report", "count", "lookup", "find", "check"
        }

        self.class_words = {"fe", "se", "te", "be", "first", "second", "third", "fourth", "final", "year", "fresher"}

        self.stopwords = {
            "my", "name", "is", "i", "am", "i'm", "myself", "this", "roll", "no", "number",
            "class", "year", "and", "from", "the", "in", "of", "for", "attendance", "register",
            "present", "absent", "mark", "student", "here", "hi", "hello", "hey", "please",
            "st", "nd", "rd", "th", "was", "are", "marked", "to", "with", "details", "info",
            "detail", "record", "records"
        } | self.class_words | self.command_words

    def extract_class(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        """Finds specific class (FE, SE, TE, BE, or ALL) from the transcript."""
        text_lower = text.lower()

        # Check for ALL / OVERALL
        if re.search(r"\b(all|overall|entire|complete|whole|every)\b", text_lower):
            return "ALL", "all"

        for canonical, patterns in self.class_patterns:
            for pattern in patterns:
                match = re.search(pattern, text_lower)
                if match:
                    return canonical, match.group(0)

        return None, None

    def extract_roll_no(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        for pattern in self.roll_patterns:
            match = pattern.search(text)
            if match:
                return match.group(1), match.group(0)
        return None, None

    def extract_name(
        self, text: str, matched_class: Optional[str] = None, matched_roll: Optional[str] = None
    ) -> Optional[str]:
        clean_text = text
        if matched_class:
            clean_text = re.sub(r"\b" + re.escape(matched_class) + r"\b", "", clean_text, flags=re.I)
        if matched_roll:
            clean_text = re.sub(re.escape(matched_roll), "", clean_text, flags=re.I)

        # Stage 1: Explicit self-introduction
        match = self.intro_pattern.search(text)
        if match:
            tokens = [w for w in match.group(1).split() if w.lower() not in self.stopwords]
            if tokens and not any(w.lower() in self.command_words for w in tokens):
                return " ".join(tokens).title()

        # Stage 2: Attendance phrases
        match = self.mark_pattern.search(text)
        if match:
            tokens = [w for w in match.group(1).split() if w.lower() not in self.stopwords]
            if tokens and not any(w.lower() in self.command_words for w in tokens):
                return " ".join(tokens).title()

        # Stage 3: spaCy NER
        if self._nlp:
            doc = self._nlp(clean_text)
            for ent in doc.ents:
                if ent.label_ == "PERSON":
                    tokens = [w for w in ent.text.split() if w.lower() not in self.stopwords and len(w) > 1]
                    if tokens and not any(w.lower() in self.command_words for w in tokens):
                        return " ".join(tokens).title()

        # Stage 4: Residual token analysis
        words_only = re.sub(r"[^\w\s]", " ", clean_text)
        words_only = re.sub(r"\d+", " ", words_only)
        candidate_tokens = [
            t for t in words_only.split()
            if t.lower() not in self.stopwords and len(t) > 1 and not t.isdigit()
        ]

        if candidate_tokens:
            name_cand = " ".join(candidate_tokens[:2]).strip().title()
            if not any(w.lower() in self.command_words for w in name_cand.split()):
                return name_cand

        return None

    def extract(self, text: str) -> ExtractedEntities:
        cls_val, cls_raw = self.extract_class(text)
        roll_val, roll_raw = self.extract_roll_no(text)
        name_val = self.extract_name(text, matched_class=cls_raw, matched_roll=roll_raw)
        return ExtractedEntities(name=name_val, roll_no=roll_val, class_name=cls_val)

    @staticmethod
    def has_location(text: str, nlp_model=None) -> bool:
        """
        Detects if user mentions location, address, city, or geographical place.
        """
        if not text:
            return False

        text_lower = text.lower().strip()

        # 1. Direct location keywords
        location_keywords = [
            r"\blocation\b", r"\blocations\b", r"\baddress\b", r"\baddresses\b",
            r"\bcity\b", r"\bcities\b", r"\bhometown\b", r"\bdistrict\b",
            r"\bvillage\b", r"\btown\b", r"\bstate\b", r"\bcountry\b",
            r"\bresidence\b", r"\bresiding\b", r"\bnative\s+place\b",
            r"\barea\b", r"\bpin\s*code\b", r"\bzip\s*code\b"
        ]
        if any(re.search(pat, text_lower) for pat in location_keywords):
            return True

        # 2. Living / staying / coming from phrases
        verb_phrases = [
            r"\b(?:live|living|lives|lived)\s+in\b",
            r"\b(?:stay|staying|stays|stayed)\s+in\b",
            r"\b(?:reside|residing|resides)\s+in\b",
            r"\b(?:located|situated)\s+(?:in|at)\b",
            r"\b(?:come|coming|came)\s+from\b",
            r"\bwhere\s+i\s+live\b",
            r"\bmy\s+place\b"
        ]
        if any(re.search(pat, text_lower) for pat in verb_phrases):
            return True

        # 3. 'from <location>' pattern where target is not class/year/college/school
        from_match = re.search(r"\bfrom\s+([a-zA-Z]+)\b", text_lower)
        if from_match:
            target = from_match.group(1).lower()
            not_location = {"fe", "se", "te", "be", "class", "year", "college", "school", "first", "second", "third", "fourth", "final", "fresher"}
            if target not in not_location:
                return True

        # 4. Known major Indian and global cities / regions
        known_places = {
            "pune", "mumbai", "delhi", "bangalore", "bengaluru", "hyderabad", "chennai",
            "kolkata", "ahmedabad", "surat", "jaipur", "lucknow", "kanpur", "nagpur",
            "indore", "thane", "bhopal", "visakhapatnam", "pimpri", "chinchwad", "patna",
            "vadodara", "ghaziabad", "ludhiana", "agra", "nashik", "nasik", "faridabad",
            "meerut", "rajkot", "kalyan", "dombivli", "vasai", "virar", "varanasi",
            "srinagar", "aurangabad", "sambhajinagar", "dhanbad", "amritsar", "navi mumbai",
            "allahabad", "prayagraj", "ranchi", "howrah", "jabalpur", "gwalior", "vijayawada",
            "jodhpur", "madurai", "raipur", "kota", "guwahati", "chandigarh", "solapur",
            "hubli", "dharwad", "bareilly", "moradabad", "mysore", "mysuru", "gurgaon",
            "gurugram", "aligarh", "jalandhar", "tiruchirappalli", "bhubaneswar", "salem",
            "warangal", "jalgaon", "guntur", "thiruvananthapuram", "bhiwandi", "saharanpur",
            "gorakhpur", "bikaner", "amravati", "noida", "jamshedpur", "bhilai", "cuttack",
            "firozabad", "kochi", "nellore", "bhavnagar", "dehradun", "durgapur", "asansol",
            "rourkela", "nanded", "kolhapur", "ajmer", "akola", "gulbarga", "jamnagar",
            "ujjain", "loni", "siliguri", "jhansi", "ulhasnagar", "jammu", "sangli",
            "mangalore", "erode", "belgaum", "malegaon", "gaya", "jalna", "udaipur",
            "latur", "dhule", "rohtak", "korba", "bhilwara", "ahmednagar", "mathura",
            "satara", "bijapur", "chandrapur", "junagadh", "thrissur", "alwar", "panipat",
            "ichalkaranji", "karnal", "bathinda", "barshi", "yavatmal", "wardha", "gondia",
            "bhandara", "osmanabad", "dharashiv", "beed", "washim", "buldhana", "nandurbar",
            "ratnagiri", "sindhudurg", "raigad", "palghar", "goa", "maharashtra", "gujarat",
            "karnataka", "kerala", "tamil nadu", "punjab", "haryana", "rajasthan"
        }
        words = set(re.findall(r"\b[a-zA-Z]+\b", text_lower))
        if words & known_places:
            return True

        # 5. spaCy Named Entity Recognition for GPE / LOC / FAC
        if nlp_model:
            doc = nlp_model(text)
            excluded_ents = {
                "fe", "se", "te", "be", "class", "year", "first", "second", "third", "fourth",
                "final", "present", "absent", "roll", "student", "attendance", "roster"
            }
            for ent in doc.ents:
                if ent.label_ in ("GPE", "LOC", "FAC"):
                    if ent.text.strip().lower() not in excluded_ents:
                        return True

        return False


# =============================================================================
# 4. INTENT CLASSIFIER MODULE
# =============================================================================

class IntentClassifier:
    """Classifies user queries into discrete executable intents based on natural speech."""

    LOOKUP_KEYWORDS = ("what is", "what's", "tell me the", "which class is", "find student", "lookup", "who is")
    SUMMARY_KEYWORDS = ("summary", "report", "overall", "total present", "total absent", "absent students", "present students")
    QUERY_KEYWORDS = ("how many", "count", "present today", "attendance of", "headcount")
    MARK_KEYWORDS = ("mark", "present", "absent", "came to class", "is here")
    REGISTER_KEYWORDS = ("i am", "i'm", "my name is", "myself", "this is", "name is", "register", "enroll")

    @classmethod
    def classify(cls, text: str) -> Intent:
        if not text:
            return Intent.UNKNOWN

        text_lower = text.lower()

        if any(kw in text_lower for kw in cls.LOOKUP_KEYWORDS):
            return Intent.LOOKUP
        if any(kw in text_lower for kw in cls.SUMMARY_KEYWORDS):
            return Intent.SUMMARY
        if any(kw in text_lower for kw in cls.QUERY_KEYWORDS):
            return Intent.QUERY
        if any(kw in text_lower for kw in cls.MARK_KEYWORDS):
            return Intent.MARK_ATTENDANCE
        if any(kw in text_lower for kw in cls.REGISTER_KEYWORDS):
            return Intent.REGISTER

        tokens = text_lower.split()
        class_clues = {"roll", "class", "from", "fe", "se", "te", "be", "year"}
        if any(tok in class_clues for tok in tokens) or any(char.isdigit() for char in text):
            return Intent.REGISTER

        if len(text_lower.strip()) > 0:
            return Intent.MARK_ATTENDANCE

        return Intent.UNKNOWN


# =============================================================================
# 5. ROSTER REPOSITORY (EXCEL DATABASE LAYER)
# =============================================================================

class RosterRepository:
    """Handles Excel file I/O, student lookups, fuzzy matching, and updates."""

    def __init__(self):
        self.lock = threading.Lock()
        self.df: pd.DataFrame = self._load()

    @staticmethod
    def normalize_roll(val: Any) -> str:
        s = str(val).strip().replace(".0", "")
        return s

    @classmethod
    def _enforce_primary_key(cls, df: pd.DataFrame) -> pd.DataFrame:
        """Ensures that RollNo acts as a unique primary key within each class."""
        if df.empty:
            return df
        # Resolve existing duplicate of 48 for Shreya in TE by restoring her default roll 12 if free
        shreya_dup = (
            (df["Class"].astype(str).str.upper() == "TE") &
            (df["RollNo"].astype(str) == "48") &
            (df["Name"].astype(str).str.lower().str.contains("shreya"))
        )
        if shreya_dup.any():
            te_rolls = set(df[df["Class"].astype(str).str.upper() == "TE"]["RollNo"].astype(str))
            if "12" not in te_rolls:
                df.loc[shreya_dup, "RollNo"] = "12"

        # Strictly enforce unique (Class, RollNo)
        df = df.drop_duplicates(subset=["Class", "RollNo"], keep="first").reset_index(drop=True)
        return df

    def _load(self) -> pd.DataFrame:
        target = (
            Config.EXCEL_FILE if os.path.exists(Config.EXCEL_FILE)
            else (Config.PARENT_EXCEL_FILE if os.path.exists(Config.PARENT_EXCEL_FILE) else None)
        )

        if target and os.path.exists(target):
            try:
                df = pd.read_excel(target, dtype={"RollNo": str})
                df["RollNo"] = df["RollNo"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                if "Date" not in df.columns:
                    df["Date"] = "-"
                df = df[~df["Name"].str.lower().isin(["change intent", "change the intent", "python users"])]
                df = self._enforce_primary_key(df)
                return df
            except Exception as e:
                print(f"[Roster Warning]: Error reading Excel file ({e}). Rebuilding roster.")

        df = pd.DataFrame(Config.DEFAULT_ROSTER)
        df = self._enforce_primary_key(df)
        self._save_df(df)
        return df

    def _save_df(self, df: pd.DataFrame) -> None:
        try:
            df.to_excel(Config.EXCEL_FILE, index=False)
            if os.path.exists(os.path.dirname(Config.PARENT_EXCEL_FILE)) and os.path.abspath(Config.PARENT_EXCEL_FILE) != os.path.abspath(Config.EXCEL_FILE):
                df.to_excel(Config.PARENT_EXCEL_FILE, index=False)
        except PermissionError:
            print("\n[WARNING] Could not save Excel file: file is currently open in Excel.\n")

    def save(self) -> None:
        with self.lock:
            self._save_df(self.df)

    def is_roll_taken_in_class(
        self, roll_no: str, class_name: str, exclude_name: Optional[str] = None
    ) -> Tuple[bool, Optional[str]]:
        """
        Checks if roll_no is already taken by another student in class_name.
        Roll number serves as a unique primary key within each class, but
        identical roll numbers are allowed across different classes.
        Returns (is_taken, holder_name).
        """
        clean_cls = class_name.upper().strip()
        norm_roll = self.normalize_roll(roll_no)
        same_class = self.df[self.df["Class"].astype(str).str.upper() == clean_cls]

        matches = same_class[same_class["RollNo"].apply(self.normalize_roll) == norm_roll]
        if matches.empty:
            return False, None

        for _, row in matches.iterrows():
            row_name = str(row["Name"]).strip()
            if exclude_name:
                ex_clean = exclude_name.lower().strip()
                r_clean = row_name.lower().strip()
                if (r_clean == ex_clean or 
                    r_clean.split()[0] == ex_clean.split()[0] or 
                    ex_clean in r_clean or 
                    r_clean in ex_clean):
                    continue
            return True, row_name

        return False, None

    def find_matches(self, name: str, class_filter: Optional[str] = None) -> pd.DataFrame:
        pool = self.df
        if class_filter:
            pool = pool[pool["Class"].astype(str).str.upper() == class_filter.upper()]

        name_clean = name.lower().strip()
        stored_names = pool["Name"].astype(str).str.lower().str.strip()

        exact = pool[stored_names == name_clean]
        if not exact.empty:
            return exact

        first_word = stored_names.str.split().str[0]
        first_match = pool[first_word == name_clean]
        if not first_match.empty:
            return first_match

        contains = pool[stored_names.str.contains(name_clean, na=False)]
        return contains

    def mark_attendance(
        self, name: str, roll_no: Optional[str] = None, class_name: Optional[str] = None
    ) -> str:
        with self.lock:
            if not name or len(name.strip()) < 2:
                return "Could not identify a valid student name to mark attendance."

            clean_name = name.strip().title()
            matches = self.find_matches(clean_name, class_filter=class_name)
            today_str = datetime.now().strftime("%Y-%m-%d %H:%M")

            if not matches.empty:
                idx = matches.index[0]
                target_cls = (class_name or str(self.df.loc[idx, "Class"])).upper().strip()
                self.df.loc[idx, "Status"] = "Present"
                self.df.loc[idx, "Date"] = today_str
                if class_name and str(class_name) != "N/A":
                    self.df.loc[idx, "Class"] = target_cls

                if roll_no and str(roll_no) != "N/A":
                    norm_roll = self.normalize_roll(roll_no)
                    is_taken, holder_name = self.is_roll_taken_in_class(norm_roll, target_cls, exclude_name=clean_name)
                    if not is_taken:
                        self.df.loc[idx, "RollNo"] = norm_roll

                rec = self.df.loc[idx]
                self._save_df(self.df)
                return (
                    f"Thank you {rec['Name']}! You are marked present for class {rec['Class']} "
                    f"(Roll No: {rec['RollNo']})."
                )

            # Strict guard: only add to Excel if valid class and roll number are present
            clean_cls = (class_name or "").upper().strip()
            clean_roll = self.normalize_roll(roll_no or "")

            if clean_cls not in Config.DEFAULT_CLASSES or not clean_roll.isdigit():
                return (
                    f"Student '{clean_name}' was not found in the roster. "
                    f"To enroll, please state your name, roll number, and class (FE, SE, TE, BE)."
                )

            # Primary key guard: check if clean_roll is already registered in this class
            is_taken, holder_name = self.is_roll_taken_in_class(clean_roll, clean_cls, exclude_name=clean_name)
            if is_taken:
                return (
                    f"Roll number {clean_roll} is already registered to {holder_name} in class {clean_cls}. "
                    f"Roll numbers must be unique within each class."
                )

            new_row = {
                "Name": clean_name,
                "RollNo": clean_roll,
                "Class": clean_cls,
                "Status": "Present",
                "Date": today_str
            }
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
            self._save_df(self.df)
            return (
                f"Thank you {clean_name}! You have been registered in class {clean_cls} "
                f"with roll number {clean_roll} and marked present."
            )

    def register_student(self, name: str, roll_no: str, class_name: str) -> str:
        with self.lock:
            if not name or not roll_no or not class_name:
                return "Registration requires complete student information: Name, Roll Number, and Class."

            clean_cls = class_name.upper().strip()
            if clean_cls not in Config.DEFAULT_CLASSES:
                return f"Class '{class_name}' is not recognized. Allowed classes are FE, SE, TE, BE."

            norm_roll = self.normalize_roll(roll_no)
            clean_name = name.strip().title()
            today_str = datetime.now().strftime("%Y-%m-%d %H:%M")

            # Primary key guard: within clean_cls, roll_no must be unique
            is_taken, holder_name = self.is_roll_taken_in_class(norm_roll, clean_cls, exclude_name=clean_name)
            if is_taken:
                return (
                    f"Roll number {norm_roll} is already registered to {holder_name} in class {clean_cls}. "
                    f"Roll numbers must be unique within each class."
                )

            # 1. Existing student enrolled in this class updating details
            name_match = self.find_matches(clean_name, class_filter=clean_cls)
            if not name_match.empty:
                idx = name_match.index[0]
                self.df.loc[idx, "Status"] = "Present"
                self.df.loc[idx, "RollNo"] = norm_roll
                self.df.loc[idx, "Date"] = today_str
                self._save_df(self.df)
                return f"Thank you {self.df.loc[idx, 'Name']}! Roster updated with roll number {norm_roll} and marked present for class {clean_cls}."

            # 2. Existing student with same roll number in this class
            same_class = self.df[self.df["Class"].astype(str).str.upper() == clean_cls]
            roll_match = same_class[same_class["RollNo"].apply(self.normalize_roll) == norm_roll]
            if not roll_match.empty:
                idx = roll_match.index[0]
                self.df.loc[idx, "Status"] = "Present"
                self.df.loc[idx, "Date"] = today_str
                self._save_df(self.df)
                return f"Thank you {self.df.loc[idx, 'Name']} (Roll {norm_roll})! Marked present for class {clean_cls}."

            # 3. New student registration
            new_row = {
                "Name": clean_name,
                "RollNo": norm_roll,
                "Class": clean_cls,
                "Status": "Present",
                "Date": today_str
            }
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
            self._save_df(self.df)
            return (
                f"Registration successful! Welcome {clean_name}, registered in class "
                f"{clean_cls} with roll number {norm_roll} and marked present."
            )

    def toggle_student_status(self, name: str, new_status: str) -> bool:
        with self.lock:
            matches = self.find_matches(name)
            if matches.empty:
                return False
            idx = matches.index[0]
            self.df.loc[idx, "Status"] = new_status
            if new_status == "Present":
                self.df.loc[idx, "Date"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            self._save_df(self.df)
            return True

    def delete_student(self, name: str, roll_no: Optional[str] = None, class_name: Optional[str] = None) -> str:
        """Remove a student from the roster by name (with optional roll/class filter)."""
        with self.lock:
            if not name or len(name.strip()) < 2:
                return "Please provide a valid student name to remove."

            matches = self.find_matches(name, class_filter=class_name)

            # If roll_no specified, narrow down further
            if roll_no and not matches.empty:
                roll_match = matches[matches["RollNo"].astype(str) == str(roll_no)]
                if not roll_match.empty:
                    matches = roll_match

            if matches.empty:
                return f"No student named '{name}' found in the roster."

            removed_name = self.df.loc[matches.index[0], 'Name']
            removed_class = self.df.loc[matches.index[0], 'Class']
            self.df = self.df.drop(index=matches.index[0]).reset_index(drop=True)
            self._save_df(self.df)
            return f"Student '{removed_name}' (Class {removed_class}) has been removed from the roster."

    def get_class_summary(self, class_name: str) -> str:
        """
        Returns summary for the SPECIFIC class requested (FE, SE, TE, BE),
        or overall summary if ALL is requested.
        """
        cls_clean = class_name.upper().strip()
        if cls_clean in ("ALL", "OVERALL", "TOTAL"):
            return self.get_overall_summary()

        filtered = self.df[self.df["Class"].astype(str).str.upper() == cls_clean]
        if filtered.empty:
            return f"No student records found for class {cls_clean}."

        present_df = filtered[filtered["Status"] == "Present"]
        absent_df = filtered[filtered["Status"] == "Absent"]

        present_list = [f"{r['Name']} (Roll {r['RollNo']})" for _, r in present_df.iterrows()]
        absent_list = [f"{r['Name']} (Roll {r['RollNo']})" for _, r in absent_df.iterrows()]

        present_str = ", ".join(present_list) if present_list else "None"
        absent_str = ", ".join(absent_list) if absent_list else "None"

        return (
            f"Class {cls_clean} Attendance Summary:\n"
            f"Total Enrolled: {len(filtered)} | Present: {len(present_df)} | Absent: {len(absent_df)}\n"
            f"Present Students: {present_str}\n"
            f"Absent Students: {absent_str}"
        )

    def get_overall_summary(self) -> str:
        lines = ["Overall Attendance Summary across all classes:"]
        for cls in Config.DEFAULT_CLASSES:
            sub = self.df[self.df["Class"].astype(str).str.upper() == cls]
            p = (sub["Status"] == "Present").sum()
            a = (sub["Status"] == "Absent").sum()
            lines.append(f"• Class {cls}: {p} Present, {a} Absent (Total: {len(sub)})")
        return "\n".join(lines)

    def get_attendance_count(self, class_name: str) -> str:
        cls_clean = class_name.upper().strip()
        if cls_clean in ("ALL", "OVERALL"):
            tot = len(self.df)
            pres = (self.df["Status"] == "Present").sum()
            abs_cnt = tot - pres
            return f"Across all classes, {pres} students are present and {abs_cnt} are absent out of {tot} enrolled."

        filtered = self.df[self.df["Class"].astype(str).str.upper() == cls_clean]
        if filtered.empty:
            return f"No records found for class {cls_clean}."

        present = (filtered["Status"] == "Present").sum()
        absent = (filtered["Status"] == "Absent").sum()
        total = len(filtered)
        return (
            f"In class {cls_clean}, {present} students are present and {absent} are absent "
            f"out of {total} enrolled students."
        )

    def lookup_student(self, name: str) -> str:
        matches = self.find_matches(name)
        if matches.empty:
            return f"Sorry, could not find any student named '{name}' in the records."

        rec = matches.iloc[0]
        return (
            f"Student Record: {rec['Name']} is enrolled in class {rec['Class']} "
            f"with roll number {rec['RollNo']}. Current status: {rec['Status']}."
        )

    def get_roster_data(self) -> Dict[str, Any]:
        records = self.df.fillna("").to_dict(orient="records")
        total = len(self.df)
        present = int((self.df["Status"] == "Present").sum()) if total else 0
        absent = total - present
        rate = round((present / total * 100), 1) if total else 0.0

        classes_data = {}
        for cls in Config.DEFAULT_CLASSES:
            sub = self.df[self.df["Class"].astype(str).str.upper() == cls]
            c_tot = len(sub)
            c_pres = int((sub["Status"] == "Present").sum())
            classes_data[cls] = {
                "total": c_tot,
                "present": c_pres,
                "absent": c_tot - c_pres,
                "rate": round((c_pres / c_tot * 100), 1) if c_tot else 0.0
            }

        return {
            "students": records,
            "stats": {
                "total": total,
                "present": present,
                "absent": absent,
                "rate": rate,
                "classes": classes_data
            }
        }


# =============================================================================
# 6. APPLICATION CONTROLLER
# =============================================================================

class VoxSheetApp:
    def __init__(self):
        self.voice = VoiceService()
        self.extractor = EntityExtractor()
        self.repo = RosterRepository()

    def process_sentence(self, text: str) -> Tuple[str, Intent, ExtractedEntities]:
        # If user mentions location, politely reject location and guide them to name, roll no, and class
        if EntityExtractor.has_location(text, self.extractor._nlp):
            msg = "Please don't tell location. I want your name, roll no, and class for registration."
            return msg, Intent.LOCATION_REJECTED, ExtractedEntities()

        intent = IntentClassifier.classify(text)
        entities = self.extractor.extract(text)

        if intent == Intent.REGISTER:
            if not entities.name or not entities.roll_no or not entities.class_name:
                missing = entities.missing_fields(("name", "roll_no", "class_name"))
                msg = f"Registration needs complete info. Please provide: {', '.join(missing)}."
            else:
                msg = self.repo.register_student(entities.name, entities.roll_no, entities.class_name)

        elif intent == Intent.MARK_ATTENDANCE:
            if not entities.name:
                msg = "Could not identify student name to mark attendance. Please specify the student's name."
            else:
                msg = self.repo.mark_attendance(entities.name, entities.roll_no, entities.class_name)

        elif intent == Intent.SUMMARY:
            # If user explicitly asked for a class (e.g. "summary of te"), give THAT class only!
            if entities.class_name:
                msg = self.repo.get_class_summary(entities.class_name)
            else:
                # Ask specifically or provide overall summary
                msg = "Which class summary would you like: FE, SE, TE, or BE? (Or say 'overall summary')."

        elif intent == Intent.QUERY:
            cls = entities.class_name or "ALL"
            msg = self.repo.get_attendance_count(cls)

        elif intent == Intent.LOOKUP:
            if not entities.name:
                msg = "Please specify a student name to look up."
            else:
                msg = self.repo.lookup_student(entities.name)

        else:
            msg = "Sorry, I could not understand what you'd like to do. Please try again."

        return msg, intent, entities


# Global singleton instance for the HTTP server
app_instance = VoxSheetApp()


# =============================================================================
# 7. MODERN WEB DASHBOARD (HTML5 / CSS3 / JAVASCRIPT)
# =============================================================================

HTML_DASHBOARD = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>VoxSheet — Voice-Activated Attendance System</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #090d16;
      --card-bg: rgba(18, 24, 38, 0.75);
      --card-border: rgba(255, 255, 255, 0.08);
      --accent: #6366f1;
      --accent-glow: rgba(99, 102, 241, 0.35);
      --cyan: #06b6d4;
      --cyan-glow: rgba(6, 182, 212, 0.35);
      --emerald: #10b981;
      --emerald-glow: rgba(16, 185, 129, 0.3);
      --rose: #f43f5e;
      --rose-glow: rgba(244, 63, 94, 0.3);
      --text: #f1f5f9;
      --text-muted: #94a3b8;
      --text-dim: #64748b;
      --radius: 18px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      background-color: var(--bg);
      color: var(--text);
      font-family: 'Plus Jakarta Sans', sans-serif;
      min-height: 100vh;
      overflow-x: hidden;
      background-image: 
        radial-gradient(circle at 15% 15%, rgba(99, 102, 241, 0.15) 0%, transparent 40%),
        radial-gradient(circle at 85% 20%, rgba(6, 182, 212, 0.12) 0%, transparent 40%),
        radial-gradient(circle at 50% 85%, rgba(16, 185, 129, 0.08) 0%, transparent 50%);
      background-attachment: fixed;
    }

    /* Top Navigation */
    header {
      position: sticky;
      top: 0;
      z-index: 50;
      backdrop-filter: blur(16px);
      background: rgba(9, 13, 22, 0.75);
      border-bottom: 1px solid var(--card-border);
      padding: 1rem 2.5rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 12px;
      font-weight: 800;
      font-size: 1.45rem;
      letter-spacing: -0.5px;
    }

    .brand-icon {
      width: 42px;
      height: 42px;
      background: linear-gradient(135deg, var(--accent), var(--cyan));
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 4px 20px var(--accent-glow);
    }

    .brand span {
      background: linear-gradient(90deg, #fff, #cbd5e1);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
    }

    .badge-nlp {
      font-size: 0.72rem;
      background: rgba(99, 102, 241, 0.18);
      color: #818cf8;
      border: 1px solid rgba(99, 102, 241, 0.35);
      padding: 4px 10px;
      border-radius: 20px;
      font-weight: 600;
      letter-spacing: 0.5px;
      text-transform: uppercase;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 12px;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 9px 16px;
      border-radius: 12px;
      font-size: 0.86rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
      border: 1px solid transparent;
      outline: none;
    }

    .btn-secondary {
      background: rgba(255, 255, 255, 0.05);
      color: var(--text);
      border-color: var(--card-border);
    }

    .btn-secondary:hover {
      background: rgba(255, 255, 255, 0.1);
      transform: translateY(-2px);
    }

    /* Voice Mute / Unmute Button */
    .btn-voice-toggle {
      background: rgba(99, 102, 241, 0.15);
      border-color: rgba(99, 102, 241, 0.35);
      color: #a5b4fc;
    }

    .btn-voice-toggle.muted {
      background: rgba(244, 63, 94, 0.15);
      border-color: rgba(244, 63, 94, 0.35);
      color: #fda4af;
    }

    .btn-excel {
      background: linear-gradient(135deg, #059669, #10b981);
      color: white;
      box-shadow: 0 4px 15px rgba(16, 185, 129, 0.25);
    }

    .btn-excel:hover {
      box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
      transform: translateY(-2px);
    }

    /* Main Container */
    main {
      max-width: 1400px;
      margin: 0 auto;
      padding: 2.5rem;
      display: grid;
      grid-template-columns: 440px 1fr;
      gap: 2rem;
    }

    @media (max-width: 1080px) {
      main {
        grid-template-columns: 1fr;
      }
    }

    /* Glass Cards */
    .glass-card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: var(--radius);
      backdrop-filter: blur(20px);
      box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.5);
      padding: 1.75rem;
      transition: border-color 0.3s ease;
    }

    .glass-card:hover {
      border-color: rgba(255, 255, 255, 0.14);
    }

    /* Voice Center Card */
    .voice-card {
      display: flex;
      flex-direction: column;
      align-items: center;
      text-align: center;
      position: relative;
      overflow: hidden;
    }

    .voice-card::before {
      content: '';
      position: absolute;
      top: -50px;
      width: 200px;
      height: 200px;
      background: var(--accent-glow);
      filter: blur(80px);
      border-radius: 50%;
      pointer-events: none;
    }

    .mic-button-wrapper {
      position: relative;
      margin: 2.2rem 0 1.5rem 0;
    }

    .mic-ring {
      position: absolute;
      top: -15px;
      left: -15px;
      right: -15px;
      bottom: -15px;
      border-radius: 50%;
      border: 2px dashed rgba(99, 102, 241, 0.4);
      animation: rotateRing 15s linear infinite;
    }

    @keyframes rotateRing {
      to { transform: rotate(360deg); }
    }

    .mic-btn {
      width: 110px;
      height: 110px;
      border-radius: 50%;
      background: linear-gradient(135deg, #4f46e5, #06b6d4);
      border: none;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 0 35px var(--accent-glow);
      transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
      position: relative;
      z-index: 2;
    }

    .mic-btn:hover {
      transform: scale(1.08);
      box-shadow: 0 0 50px rgba(6, 182, 212, 0.5);
    }

    .mic-btn.listening {
      background: linear-gradient(135deg, #e11d48, #f43f5e);
      box-shadow: 0 0 50px var(--rose-glow);
      animation: micPulse 1.2s ease-in-out infinite;
    }

    @keyframes micPulse {
      0%, 100% { transform: scale(1); box-shadow: 0 0 30px var(--rose-glow); }
      50% { transform: scale(1.1); box-shadow: 0 0 60px rgba(244, 63, 94, 0.6); }
    }

    .mic-btn svg {
      width: 48px;
      height: 48px;
      fill: white;
    }

    /* Soundwave Visualizer */
    .soundwave {
      display: flex;
      align-items: center;
      gap: 5px;
      height: 28px;
      margin-bottom: 1rem;
    }

    .soundwave .bar {
      width: 4px;
      height: 6px;
      background: var(--text-dim);
      border-radius: 4px;
      transition: height 0.15s ease, background 0.2s ease;
    }

    .listening .soundwave .bar {
      background: linear-gradient(to top, var(--accent), var(--cyan));
      animation: soundwaveAnim 0.8s ease-in-out infinite alternate;
    }

    .listening .soundwave .bar:nth-child(2) { animation-delay: 0.1s; }
    .listening .soundwave .bar:nth-child(3) { animation-delay: 0.2s; }
    .listening .soundwave .bar:nth-child(4) { animation-delay: 0.3s; }
    .listening .soundwave .bar:nth-child(5) { animation-delay: 0.4s; }
    .listening .soundwave .bar:nth-child(6) { animation-delay: 0.25s; }
    .listening .soundwave .bar:nth-child(7) { animation-delay: 0.15s; }

    @keyframes soundwaveAnim {
      0% { height: 6px; }
      100% { height: 28px; }
    }

    .status-text {
      font-size: 0.95rem;
      font-weight: 600;
      color: var(--text-muted);
      min-height: 24px;
    }

    /* Natural Language Input Box */
    .nl-input-box {
      width: 100%;
      margin-top: 1.5rem;
      display: flex;
      gap: 8px;
    }

    .nl-input {
      flex: 1;
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--card-border);
      padding: 12px 16px;
      border-radius: 12px;
      color: white;
      font-size: 0.9rem;
      font-family: inherit;
      outline: none;
      transition: border-color 0.2s;
    }

    .nl-input:focus {
      border-color: var(--accent);
      background: rgba(255, 255, 255, 0.08);
    }

    .btn-send {
      background: var(--accent);
      color: white;
      border: none;
      border-radius: 12px;
      padding: 0 18px;
      cursor: pointer;
      font-weight: 600;
      font-size: 0.88rem;
      transition: background 0.2s;
    }

    .btn-send:hover {
      background: #4f46e5;
    }

    /* Live NLP Parsing Result */
    .nlp-breakdown {
      width: 100%;
      margin-top: 1.5rem;
      padding: 14px;
      background: rgba(0, 0, 0, 0.35);
      border-radius: 12px;
      border: 1px dashed var(--card-border);
      text-align: left;
    }

    .nlp-title {
      font-size: 0.72rem;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: var(--cyan);
      font-weight: 700;
      margin-bottom: 8px;
      display: flex;
      justify-content: space-between;
    }

    .nlp-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }

    .tag {
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.75rem;
      padding: 4px 8px;
      border-radius: 6px;
      background: rgba(255, 255, 255, 0.06);
      border: 1px solid rgba(255, 255, 255, 0.1);
    }

    .tag-intent { color: #a78bfa; border-color: rgba(167, 139, 250, 0.3); }
    .tag-name { color: #34d399; border-color: rgba(52, 211, 153, 0.3); }
    .tag-roll { color: #38bdf8; border-color: rgba(56, 189, 248, 0.3); }
    .tag-class { color: #fbbf24; border-color: rgba(251, 191, 36, 0.3); }

    .nlp-tips {
      margin-top: 8px;
      font-size: 0.74rem;
      color: var(--text-dim);
      line-height: 1.4;
    }

    /* System Audio Feedback */
    .speech-box {
      margin-top: 1.5rem;
      background: rgba(0, 0, 0, 0.4);
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 12px;
      padding: 14px;
      display: flex;
      flex-direction: column;
      gap: 6px;
      text-align: left;
    }

    .speech-header {
      font-size: 0.72rem;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: var(--text-dim);
      font-weight: 700;
    }

    .speech-bubble {
      font-size: 0.9rem;
      color: #cbd5e1;
      line-height: 1.4;
      white-space: pre-line;
    }

    /* Right Section: Stats & Roster */
    .right-section {
      display: flex;
      flex-direction: column;
      gap: 1.75rem;
    }

    /* Stats Grid */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 1rem;
    }

    @media (max-width: 768px) {
      .stats-grid {
        grid-template-columns: repeat(2, 1fr);
      }
    }

    .stat-card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: var(--radius);
      padding: 1.25rem 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 6px;
      position: relative;
      overflow: hidden;
    }

    .stat-label {
      font-size: 0.78rem;
      font-weight: 600;
      color: var(--text-dim);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .stat-val {
      font-size: 1.85rem;
      font-weight: 800;
      letter-spacing: -0.5px;
    }

    .val-total { color: #f8fafc; }
    .val-present { color: var(--emerald); }
    .val-absent { color: var(--rose); }
    .val-rate { color: var(--cyan); }

    /* Class Filter Tabs */
    .table-header-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.25rem;
      flex-wrap: wrap;
      gap: 12px;
    }

    .class-tabs {
      display: flex;
      background: rgba(255, 255, 255, 0.05);
      padding: 4px;
      border-radius: 12px;
      gap: 4px;
    }

    .tab-btn {
      background: transparent;
      border: none;
      color: var(--text-muted);
      padding: 7px 16px;
      border-radius: 8px;
      font-size: 0.84rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }

    .tab-btn.active {
      background: var(--accent);
      color: white;
      box-shadow: 0 2px 10px var(--accent-glow);
    }

    .search-input {
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--card-border);
      padding: 8px 14px;
      border-radius: 10px;
      color: white;
      font-size: 0.85rem;
      font-family: inherit;
      width: 220px;
      outline: none;
    }

    .search-input:focus {
      border-color: var(--cyan);
    }

    /* Roster Table */
    .table-container {
      overflow-x: auto;
      border-radius: 12px;
      border: 1px solid var(--card-border);
    }

    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.88rem;
    }

    th {
      background: rgba(15, 23, 42, 0.65);
      padding: 12px 18px;
      color: var(--text-muted);
      font-size: 0.76rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border-bottom: 1px solid var(--card-border);
    }

    td {
      padding: 14px 18px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
      color: #e2e8f0;
    }

    tr:hover td {
      background: rgba(255, 255, 255, 0.02);
    }

    .badge-status {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 5px 12px;
      border-radius: 20px;
      font-size: 0.75rem;
      font-weight: 700;
      letter-spacing: 0.3px;
      cursor: pointer;
      user-select: none;
      transition: transform 0.15s ease;
    }

    .badge-status:hover {
      transform: scale(1.05);
    }

    .status-present {
      background: rgba(16, 185, 129, 0.15);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.35);
    }

    .status-absent {
      background: rgba(244, 63, 94, 0.15);
      color: #fb7185;
      border: 1px solid rgba(244, 63, 94, 0.35);
    }

    .timestamp {
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.78rem;
      color: var(--text-dim);
    }

    .btn-delete {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 5px 11px;
      border-radius: 8px;
      font-size: 0.75rem;
      font-weight: 600;
      cursor: pointer;
      border: 1px solid rgba(244, 63, 94, 0.4);
      background: rgba(244, 63, 94, 0.12);
      color: #fb7185;
      transition: all 0.2s ease;
    }

    .btn-delete:hover {
      background: rgba(244, 63, 94, 0.28);
      transform: scale(1.05);
    }

    /* Toast Notification */
    .toast {
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: rgba(18, 24, 38, 0.95);
      border: 1px solid var(--accent);
      color: white;
      padding: 14px 20px;
      border-radius: 14px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.6);
      font-size: 0.9rem;
      font-weight: 500;
      display: flex;
      align-items: center;
      gap: 10px;
      opacity: 0;
      transform: translateY(20px);
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 100;
      pointer-events: none;
    }

    .toast.show {
      opacity: 1;
      transform: translateY(0);
    }
  </style>
</head>
<body>

  <!-- Top Navigation -->
  <header>
    <div class="brand">
      <div class="brand-icon">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
          <path d="M19 10v2a7 7 0 0 1-14 0v-2"/>
          <line x1="12" y1="19" x2="12" y2="23"/>
          <line x1="8" y1="23" x2="16" y2="23"/>
        </svg>
      </div>
      <div>
        <span>VoxSheet</span>
      </div>
      <div class="badge-nlp">NLP Attendance System</div>
    </div>

    <div class="header-actions">
      <!-- Voice Audio Mute/Unmute Toggle -->
      <button class="btn btn-voice-toggle" id="btnVoiceToggle" onclick="toggleVoiceAudio()" title="Click to mute or unmute verbal audio">
        <span id="voiceToggleIcon">🔊</span>
        <span id="voiceToggleText">Voice: ON</span>
      </button>

      <!-- Instant Class Summary -->
      <button class="btn btn-secondary" onclick="speakSummary()">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polygon points="11 5 6 9 2 9 2 15 6 15 11 19 11 5"/>
          <path d="M19.07 4.93a10 10 0 0 1 0 14.14M15.54 8.46a5 5 0 0 1 0 7.07"/>
        </svg>
        Summary Report
      </button>

      <!-- Export to Excel -->
      <a href="/api/download-excel" class="btn btn-excel" download>
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Export Excel
      </a>
    </div>
  </header>

  <!-- Main Body -->
  <main>
    <!-- Left Column: Voice Assistant & NLP Inspector -->
    <div class="glass-card voice-card">
      <h2 style="font-size: 1.25rem; font-weight: 700; margin-bottom: 4px;">Voice & Text Assistant</h2>
      <p style="font-size: 0.85rem; color: var(--text-muted);">Speak via mic or type your natural command below</p>

      <div class="mic-button-wrapper" id="micWrapper">
        <div class="mic-ring"></div>
        <button class="mic-btn" id="micBtn" onclick="toggleVoiceInput()" title="Click to speak (Optional)">
          <svg viewBox="0 0 24 24">
            <path d="M12 14c1.66 0 3-1.34 3-3V5c0-1.66-1.34-3-3-3S9 3.34 9 5v6c0 1.66 1.34 3 3 3z"/>
            <path d="M17 11c0 2.76-2.24 5-5 5s-5-2.24-5-5H5c0 3.53 2.61 6.43 6 6.92V21h2v-3.08c3.39-.49 6-3.39 6-6.92h-2z"/>
          </svg>
        </button>
      </div>

      <div class="soundwave">
        <div class="bar"></div>
        <div class="bar"></div>
        <div class="bar"></div>
        <div class="bar"></div>
        <div class="bar"></div>
        <div class="bar"></div>
        <div class="bar"></div>
      </div>

      <div class="status-text" id="statusText">Tap microphone or type command</div>

      <!-- Natural Language Text Entry (Type without speaking) -->
      <div class="nl-input-box">
        <input type="text" id="textInput" class="nl-input" placeholder="e.g. summary of te, mark Diksha present, count in fe..." onkeydown="if(event.key==='Enter') sendTextMessage()">
        <button class="btn-send" onclick="sendTextMessage()">Send</button>
      </div>

      <!-- Live NLP Breakdown -->
      <div class="nlp-breakdown">
        <div class="nlp-title">
          <span>NLP Entity Extraction</span>
          <span style="color: var(--text-dim);" id="nlpPipeline">spaCy + RegEx</span>
        </div>
        <div class="nlp-tags" id="nlpTags">
          <span class="tag tag-intent" id="currentIntentTag">Intent: Ready</span>
          <span class="tag tag-name" id="currentNameTag">Name: —</span>
          <span class="tag tag-roll" id="currentRollTag">Roll: —</span>
          <span class="tag tag-class" id="currentClassTag">Class: —</span>
        </div>
        <div class="nlp-tips">
          💡 Try typing or speaking: <i>"summary of te"</i>, <i>"summary of fe"</i>, or <i>"mark Rahul Patil present"</i>
        </div>
      </div>

      <!-- System Response Display -->
      <div class="speech-box" style="width: 100%;">
        <div class="speech-header">System Response</div>
        <div class="speech-bubble" id="speechResponse">"Welcome to VoxSheet! Voice and text attendance engine is ready."</div>
      </div>
    </div>

    <!-- Right Column: Stats & Student Roster -->
    <div class="right-section">
      <!-- Top Metrics -->
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-label">Total Enrolled</div>
          <div class="stat-val val-total" id="statTotal">0</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">Present Today</div>
          <div class="stat-val val-present" id="statPresent">0</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">Absent Today</div>
          <div class="stat-val val-absent" id="statAbsent">0</div>
        </div>
        <div class="stat-card">
          <div class="stat-label">Attendance Rate</div>
          <div class="stat-val val-rate" id="statRate">0%</div>
        </div>
      </div>

      <!-- Live Roster Table Card -->
      <div class="glass-card" style="padding: 1.5rem;">
        <div class="table-header-row">
          <div class="class-tabs">
            <button class="tab-btn active" onclick="filterClass('ALL', this)">All</button>
            <button class="tab-btn" onclick="filterClass('FE', this)">FE</button>
            <button class="tab-btn" onclick="filterClass('SE', this)">SE</button>
            <button class="tab-btn" onclick="filterClass('TE', this)">TE</button>
            <button class="tab-btn" onclick="filterClass('BE', this)">BE</button>
          </div>

          <input type="text" id="searchInput" class="search-input" placeholder="Search name or roll..." oninput="renderTable()">
        </div>

        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Roll No</th>
                <th>Student Name</th>
                <th>Class</th>
                <th>Status (Click to toggle)</th>
                <th>Recorded At</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody id="rosterTableBody">
              <!-- Dynamically populated rows -->
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <!-- Floating Toast -->
  <div class="toast" id="toastMessage">Attendance recorded successfully!</div>

  <script>
    let rosterData = [];
    let activeFilter = 'ALL';
    let isListening = false;
    let webSpeechRecognizer = null;
    let voiceAudioEnabled = true;

    function toggleVoiceAudio() {
      voiceAudioEnabled = !voiceAudioEnabled;
      const btn = document.getElementById('btnVoiceToggle');
      const icon = document.getElementById('voiceToggleIcon');
      const text = document.getElementById('voiceToggleText');

      if (voiceAudioEnabled) {
        btn.classList.remove('muted');
        icon.innerText = '🔊';
        text.innerText = 'Voice: ON';
        showToast('Voice audio enabled');
      } else {
        btn.classList.add('muted');
        icon.innerText = '🔇';
        text.innerText = 'Voice: MUTED';
        showToast('Voice audio muted (Silent Mode)');
      }
    }

    // Initialize Web Speech API if supported in browser
    if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      webSpeechRecognizer = new SpeechRecognition();
      webSpeechRecognizer.continuous = false;
      webSpeechRecognizer.interimResults = false;
      webSpeechRecognizer.lang = 'en-US';

      webSpeechRecognizer.onresult = function(event) {
        const transcript = event.results[0][0].transcript;
        document.getElementById('textInput').value = transcript;
        showToast("Heard: \"" + transcript + "\"");
        processNLPInput(transcript);
        stopListeningVisuals();
      };

      webSpeechRecognizer.onerror = function(event) {
        console.warn("Browser STT Error:", event.error);
        stopListeningVisuals();
        if (event.error === 'no-speech') {
          showToast('No speech detected. Please try again or type your command.');
        } else if (event.error === 'network') {
          showToast('Mic needs internet for speech recognition. Please type your command instead.');
        } else {
          showToast('Microphone error: ' + event.error + '. Please type your command.');
        }
      };

      webSpeechRecognizer.onend = function() {
        stopListeningVisuals();
      };
    }

    async function loadRoster() {
      try {
        const res = await fetch('/api/roster');
        const data = await res.json();
        rosterData = data.students || [];

        document.getElementById('statTotal').innerText = data.stats.total;
        document.getElementById('statPresent').innerText = data.stats.present;
        document.getElementById('statAbsent').innerText = data.stats.absent;
        document.getElementById('statRate').innerText = data.stats.rate + '%';

        renderTable();
      } catch (err) {
        console.error("Failed to load roster:", err);
      }
    }

    function renderTable() {
      const tbody = document.getElementById('rosterTableBody');
      const search = document.getElementById('searchInput').value.toLowerCase().trim();

      const filtered = rosterData.filter(s => {
        const matchesClass = (activeFilter === 'ALL') || (s.Class && s.Class.toUpperCase() === activeFilter);
        const matchesSearch = !search || 
          (s.Name && s.Name.toLowerCase().includes(search)) || 
          (s.RollNo && String(s.RollNo).includes(search));
        return matchesClass && matchesSearch;
      });

      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-dim); padding: 2.5rem;">No students found for this filter.</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(s => {
        const isPresent = (s.Status === 'Present');
        const badgeClass = isPresent ? 'status-present' : 'status-absent';
        const nextStatus = isPresent ? 'Absent' : 'Present';
        const safeName = (s.Name || '').replace(/'/g, "\\'");
        return `
          <tr>
            <td style="font-weight: 700; color: var(--cyan); font-family: 'JetBrains Mono', monospace;">#${s.RollNo || '—'}</td>
            <td style="font-weight: 600;">${s.Name || '—'}</td>
            <td><span style="background: rgba(255,255,255,0.06); padding: 3px 8px; border-radius: 6px; font-weight: 700; font-size: 0.78rem;">${s.Class || '—'}</span></td>
            <td>
              <span class="badge-status ${badgeClass}" onclick="toggleStatus('${safeName}', '${nextStatus}')">
                <span style="width: 6px; height: 6px; border-radius: 50%; background: currentColor;"></span>
                ${s.Status || 'Absent'}
              </span>
            </td>
            <td class="timestamp">${s.Date || '—'}</td>
            <td>
              <button class="btn-delete" onclick="deleteStudent('${safeName}', '${s.RollNo || ''}', '${s.Class || ''}')" title="Remove this student">
                🗑 Remove
              </button>
            </td>
          </tr>
        `;
      }).join('');
    }

    function filterClass(cls, btn) {
      activeFilter = cls;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderTable();
    }

    async function toggleStatus(name, newStatus) {
      try {
        const res = await fetch('/api/toggle-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: name, status: newStatus })
        });
        const result = await res.json();
        if (result.success) {
          showToast(`Updated ${name} to ${newStatus}`);
          loadRoster();
        }
      } catch (err) {
        console.error(err);
      }
    }

    async function deleteStudent(name, rollNo, className) {
      if (!confirm(`Remove "${name}" (Class ${className}, Roll #${rollNo}) from the roster? This cannot be undone.`)) return;
      try {
        const res = await fetch('/api/delete-student', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: name, roll_no: rollNo, class_name: className })
        });
        const result = await res.json();
        showToast(result.message || 'Student removed.');
        loadRoster();
      } catch (err) {
        console.error(err);
        showToast('Error removing student.');
      }
    }

    function toggleVoiceInput() {
      if (isListening) {
        stopListeningVisuals();
        if (webSpeechRecognizer) webSpeechRecognizer.stop();
        return;
      }

      isListening = true;
      document.getElementById('micBtn').classList.add('listening');
      document.getElementById('micWrapper').parentElement.classList.add('listening');
      document.getElementById('statusText').innerText = "Listening... please speak now";

      if (webSpeechRecognizer) {
        try {
          webSpeechRecognizer.start();
        } catch (e) {
          stopListeningVisuals();
          showToast('Could not start microphone. Please type your command instead.');
        }
      } else {
        stopListeningVisuals();
        showToast('Speech recognition not supported in this browser. Please type your command.');
      }
    }

    function stopListeningVisuals() {
      isListening = false;
      document.getElementById('micBtn').classList.remove('listening');
      document.getElementById('micWrapper').parentElement.classList.remove('listening');
      document.getElementById('statusText').innerText = "Tap microphone or type command";
    }

    async function triggerServerMic() {
      showToast("Listening through server microphone...");
      try {
        const res = await fetch('/api/voice-record', { 
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ speak: voiceAudioEnabled })
        });
        const data = await res.json();
        stopListeningVisuals();

        if (data.transcript) {
          document.getElementById('textInput').value = data.transcript;
        }
        updateNLPBreakdown(data.intent, data.entities);
        document.getElementById('speechResponse').innerText = `"${data.message}"`;
        showToast(data.message);
        loadRoster();
      } catch (err) {
        stopListeningVisuals();
        console.error(err);
      }
    }

    function sendTextMessage() {
      const input = document.getElementById('textInput');
      const text = input.value.trim();
      if (!text) return;
      processNLPInput(text);
      input.value = '';
    }

    async function processNLPInput(text) {
      document.getElementById('statusText').innerText = "Processing Natural Language...";
      try {
        const res = await fetch('/api/process', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ text: text, speak: voiceAudioEnabled })
        });
        const data = await res.json();
        document.getElementById('statusText').innerText = "Done!";
        updateNLPBreakdown(data.intent, data.entities);
        document.getElementById('speechResponse').innerText = `"${data.message}"`;
        showToast(data.message);
        loadRoster();
      } catch (err) {
        console.error(err);
      }
    }

    function updateNLPBreakdown(intent, entities) {
      const tagEl = document.getElementById('currentIntentTag');
      if (tagEl && intent) {
        const label = intent.replace('_', ' ').toUpperCase();
        tagEl.innerText = "Intent: " + label;
      }
      if (entities) {
        document.getElementById('currentNameTag').innerText = "Name: " + (entities.name || '—');
        document.getElementById('currentRollTag').innerText = "Roll: " + (entities.roll_no || '—');
        document.getElementById('currentClassTag').innerText = "Class: " + (entities.class || '—');
      }
    }

    async function speakSummary() {
      const targetClass = (activeFilter === 'ALL') ? 'ALL' : activeFilter;
      showToast("Fetching summary for " + targetClass + "...");
      try {
        const res = await fetch('/api/summary?class=' + targetClass + '&speak=' + voiceAudioEnabled);
        const data = await res.json();
        document.getElementById('speechResponse').innerText = data.summary;
        showToast("Summary generated for " + targetClass);
      } catch (err) {
        console.error(err);
      }
    }

    function showToast(msg) {
      const t = document.getElementById('toastMessage');
      t.innerText = msg;
      t.classList.add('show');
      setTimeout(() => { t.classList.remove('show'); }, 3800);
    }

    // Initial Load & polling
    loadRoster();
    setInterval(loadRoster, 6000);
  </script>
</body>
</html>
"""


# =============================================================================
# 8. MULTI-THREADED HTTP REQUEST HANDLER
# =============================================================================

class VoxSheetRequestHandler(http.server.BaseHTTPRequestHandler):
    """Zero-dependency HTTP server handling REST endpoints and serving the Web Dashboard."""

    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} - {args[1]}\n")

    def send_json(self, data: Dict[str, Any], status: int = 200) -> None:
        payload = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        query = urllib.parse.parse_qs(parsed_url.query)

        # Serve Web Dashboard
        if path in ("/", "/index.html"):
            payload = HTML_DASHBOARD.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return

        # API: Return Roster & Stats
        if path == "/api/roster":
            data = app_instance.repo.get_roster_data()
            self.send_json(data)
            return

        # API: Download Excel File
        if path == "/api/download-excel":
            excel_path = Config.EXCEL_FILE
            if not os.path.exists(excel_path):
                app_instance.repo.save()

            if os.path.exists(excel_path):
                with open(excel_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", 'attachment; filename="students.xlsx"')
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_error(404, "Excel file not found.")
            return

        # API: Class Summary
        if path == "/api/summary":
            cls = query.get("class", ["TE"])[0]
            should_speak = query.get("speak", ["true"])[0].lower() in ("true", "1")
            summary_msg = app_instance.repo.get_class_summary(cls)
            if should_speak:
                threading.Thread(target=app_instance.voice.speak, args=(summary_msg,), daemon=True).start()
            self.send_json({"summary": summary_msg})
            return

        self.send_error(404, "Endpoint not found.")

    def do_POST(self) -> None:
        path = self.path
        content_len = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_len).decode("utf-8") if content_len else "{}"
        try:
            payload = json.loads(body)
        except Exception:
            payload = {}

        # API: Natural Language Processing
        if path == "/api/process":
            user_text = payload.get("text", "").strip()
            should_speak = payload.get("speak", True)
            msg, intent, entities = app_instance.process_sentence(user_text)
            if should_speak:
                threading.Thread(target=app_instance.voice.speak, args=(msg,), daemon=True).start()
            self.send_json({
                "success": True,
                "message": msg,
                "intent": intent.value if hasattr(intent, "value") else str(intent),
                "entities": entities.to_dict()
            })
            return

        # API: Voice Recording Turn (Microphone)
        if path == "/api/voice-record":
            should_speak = payload.get("speak", True)
            spoken_text = app_instance.voice.listen()
            if not spoken_text:
                self.send_json({
                    "success": False,
                    "transcript": "",
                    "message": "Sorry, no speech was detected. Please try speaking again.",
                    "intent": "unknown",
                    "entities": {}
                })
                return

            msg, intent, entities = app_instance.process_sentence(spoken_text)
            if should_speak:
                threading.Thread(target=app_instance.voice.speak, args=(msg,), daemon=True).start()
            self.send_json({
                "success": True,
                "transcript": spoken_text,
                "message": msg,
                "intent": intent.value if hasattr(intent, "value") else str(intent),
                "entities": entities.to_dict()
            })
            return

        # API: Toggle Status
        if path == "/api/toggle-status":
            name = payload.get("name", "")
            new_status = payload.get("status", "Present")
            success = app_instance.repo.toggle_student_status(name, new_status)
            self.send_json({"success": success})
            return

        # API: Delete / Remove Student
        if path == "/api/delete-student":
            name = payload.get("name", "").strip()
            roll_no = payload.get("roll_no", None)
            class_name = payload.get("class_name", None)
            msg = app_instance.repo.delete_student(name, roll_no, class_name)
            success = "removed" in msg.lower()
            self.send_json({"success": success, "message": msg})
            return

        self.send_error(404, "Endpoint not found.")


# =============================================================================
# 9. SERVER RUNNER & ENTRYPOINT
# =============================================================================

def start_server():
    port = Config.PORT
    handler = VoxSheetRequestHandler
    socketserver.TCPServer.allow_reuse_address = True

    try:
        httpd = socketserver.ThreadingTCPServer(("", port), handler)
    except OSError:
        port = 5055
        httpd = socketserver.ThreadingTCPServer(("", port), handler)

    url = f"http://localhost:{port}"
    print("\n" + "=" * 65)
    print(f"  [ONLINE] VOXSHEET WEB DASHBOARD READY!")
    print(f"  [URL]   {url}")
    print("=" * 65 + "\n")

    threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down VoxSheet Web Server...")
        httpd.server_close()


if __name__ == "__main__":
    start_server()