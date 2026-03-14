import streamlit as st
import pandas as pd
import numpy as np
import io, re, unicodedata
import warnings
from datetime import datetime
from rapidfuzz import fuzz

# Suppress warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Client vs Arkan Comparator – v1.7.0", page_icon="🧮", layout="wide")

# ---------- THEME ----------
st.markdown("""
<style>
:root { --brand-1:#0ea5e9; --brand-2:#22c55e; --brand-3:#a78bfa; }
.block-container { padding-top: 1.5rem; }
.header-box {
  background: linear-gradient(120deg, var(--brand-1), var(--brand-3));
  color: white; padding: 16px 20px; border-radius: 16px;
  box-shadow: 0 10px 25px rgba(0,0,0,.12); margin-bottom: 16px;
}
.badge { display:inline-block; padding:4px 10px; font-size:12px;
  background: rgba(255,255,255,.15); border-radius:999px; margin-left:8px; }
.stButton>button, .stDownloadButton>button { color:white; border:0; padding:.6rem 1rem;
  border-radius:12px; font-weight:600; }
.stButton>button { background: linear-gradient(90deg, var(--brand-2), var(--brand-1)); }
.stDownloadButton>button { background: linear-gradient(90deg, var(--brand-1), var(--brand-3)); }
.small { font-size:12px; opacity:.85; }
.error-box {
  background: linear-gradient(90deg, #ef4444, #dc2626);
  color: white; padding: 12px 16px; border-radius: 8px; margin: 10px 0;
  font-size: 14px;
}
.warning-box {
  background: linear-gradient(90deg, #f59e0b, #f97316);
  color: white; padding: 12px 16px; border-radius: 8px; margin: 10px 0;
  font-size: 14px;
}
.success-box {
  background: linear-gradient(90deg, var(--brand-2), #16a34a);
  color: white; padding: 12px 16px; border-radius: 8px; margin: 10px 0;
  font-size: 14px;
}
</style>
<div class="header-box">
  <h2 style="margin:0;">🧮 Client vs Arkan Comparator <span class="badge">v1.7.0 - Enhanced</span></h2>
  <p style="margin:.25rem 0 0 0;opacity:.9">
    Fixed encoding issues • Enhanced file reading • Improved regex patterns • Better error handling • Optimized performance
  </p>
</div>
""", unsafe_allow_html=True)

# ---------- OPTIONS ----------
st.sidebar.title("⚙️ Options")
strip_leading_zeros = st.sidebar.checkbox("Strip leading zeros in numeric reference (HTL-WBD only)", value=False)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Date parsing")
date_mode_choices = ["Auto (detect)","YMD (YYYY-MM-DD)","DMY (DD/MM/YYYY)","MDY (MM/DD/YYYY)"]
client_date_mode = st.sidebar.selectbox("Client date order", date_mode_choices, index=0)
arkan_date_mode  = st.sidebar.selectbox("Arkan date order",  date_mode_choices, index=0)

st.sidebar.markdown("---")
st.sidebar.subheader("🏨 Hotel matching")
mode_hotel = st.sidebar.radio("Hotel match mode", ["Strict (exact after cleaning)","Smart (tokens + fuzzy)"], index=1)
fuzzy_threshold_hotel = st.sidebar.slider("Hotel fuzzy threshold", 0, 100, 92, 1)
jaccard_threshold = st.sidebar.slider("Hotel token Jaccard threshold", 0.0, 1.0, 0.65, 0.05)
ignore_tokens_default = "hotel,hotels,resort,resorts,residence,residences,suite,suites,inn,lodge,tower,towers,apartment,apartments,aparthotel,by,the,of,at,an,ihg,marriott,hilton"
ignore_tokens_input = st.sidebar.text_input("Hotel ignore tokens (comma-separated)", value=ignore_tokens_default)
location_tokens_default = "makkah,mecca,makka,madina,madinah,medina,riyadh,jeddah,khobar,al,aziziyah,aziziah,alaziziah,saudi,ksa,dubai,uae,doha,qatar,jabal,omar"
location_tokens_input = st.sidebar.text_input("Hotel ignore location tokens", value=location_tokens_default)

st.sidebar.markdown("---")
st.sidebar.subheader("🧑‍🤝‍🧑 Guest matching")
mode_guest = st.sidebar.radio("Guest match mode", ["Exact (normalized)","Smart (fuzzy on tokens)"], index=0)
fuzzy_threshold_guest = st.sidebar.slider("Guest fuzzy threshold", 0, 100, 95, 1)
compare_guest_even_if_hotel_mismatch = st.sidebar.checkbox("Also compute guest match even if hotel mismatches (for diagnostics)", value=True)

# ---------- ENHANCED FILE READING ----------
def read_any_enhanced(uploaded):
    """Enhanced file reader with multiple fallback strategies and better encoding handling"""
    name = uploaded.name.lower()
    errors_encountered = []
    
    try:
        if name.endswith(".csv"):
            # Try different CSV encoding strategies
            for encoding in ['utf-8', 'utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']:
                try:
                    uploaded.seek(0)
                    df = pd.read_csv(uploaded, dtype=str, encoding=encoding)
                    return df, f"✅ CSV read with {encoding} encoding"
                except Exception as e:
                    errors_encountered.append(f"{encoding}: {str(e)}")
                    continue
            
        elif name.endswith(('.xls', '.xlsb')):
            # Strategy 1: Try xlrd for legacy formats
            try:
                uploaded.seek(0)
                df = pd.read_excel(uploaded, dtype=str, engine="xlrd")
                return df, f"✅ Excel read with xlrd engine"
            except Exception as e:
                errors_encountered.append(f"xlrd: {str(e)}")
                
            # Strategy 2: Try python engine as fallback
            try:
                uploaded.seek(0)
                df = pd.read_excel(uploaded, dtype=str, engine="python")
                return df, f"⚠️ Excel read with python engine (fallback)"
            except Exception as e:
                errors_encountered.append(f"python engine: {str(e)}")
                
        elif name.endswith('.xlsx'):
            # Strategy 1: Try openpyxl
            try:
                uploaded.seek(0)
                df = pd.read_excel(uploaded, dtype=str, engine="openpyxl")
                return df, f"✅ Excel read with openpyxl engine"
            except Exception as e:
                errors_encountered.append(f"openpyxl: {str(e)}")
                
            # Strategy 2: Try python engine as fallback
            try:
                uploaded.seek(0)
                df = pd.read_excel(uploaded, dtype=str, engine="python")
                return df, f"⚠️ Excel read with python engine (fallback)"
            except Exception as e:
                errors_encountered.append(f"python engine: {str(e)}")
        
        # Final fallback: auto-detect engine
        try:
            uploaded.seek(0)
            df = pd.read_excel(uploaded, dtype=str)
            return df, f"⚠️ Excel read with auto-detected engine"
        except Exception as e:
            errors_encountered.append(f"auto-engine: {str(e)}")
            
    except Exception as e:
        errors_encountered.append(f"general error: {str(e)}")
    
    # If all strategies fail
    error_msg = f"❌ All reading strategies failed for {name}:\n" + "\n".join(f"• {err}" for err in errors_encountered)
    error_msg += f"\n\n💡 Suggestions:\n• Ensure file isn't corrupted\n• Try saving as .xlsx in Excel\n• Convert to CSV format with UTF-8 encoding"
    raise ValueError(error_msg)

# ---------- FIXED REGEX AND NORMALIZATION ----------
# Fixed Arabic diacritics removal with proper Unicode ranges
AR_DIAC = re.compile(r"[\u0617-\u061A\u064B-\u0652\u0670]")

def ar_norm_fixed(s: str) -> str:
    """Fixed Arabic normalization with proper Unicode handling"""
    if not isinstance(s, str): 
        return ""
    
    # Remove diacritics
    s = AR_DIAC.sub("", s)
    
    # Remove tatweel (Arabic kashida)
    s = s.replace("ـ", "")
    
    # Normalize Arabic characters
    replacements = {
        "أ": "ا", "إ": "ا", "آ": "ا",  # Alef variations
        "ى": "ي",                      # Alef maksura to yeh
        "ة": "ه",                      # Teh marbuta to heh
        "ؤ": "و",                      # Waw with hamza
        "ئ": "ي",                      # Yeh with hamza
        "ء": ""                        # Remove standalone hamza
    }
    
    for old, new in replacements.items():
        s = s.replace(old, new)
    
    return s

def ascii_fold_fixed(s: str) -> str:
    """Enhanced ASCII folding with better Unicode handling"""
    if not isinstance(s, str): 
        return ""
    
    try:
        # Normalize to decomposed form and remove accents
        normalized = unicodedata.normalize("NFKD", s)
        ascii_str = normalized.encode("ascii", "ignore").decode("ascii")
        return ascii_str
    except Exception:
        # Fallback: remove non-ASCII characters
        return re.sub(r'[^\x00-\x7F]', '', s)

# Remove zero-width & control characters - FIXED regex
ZW_RE = re.compile(r"[\u200B-\u200D\uFEFF\u2060]")

# ----- FIXED TITLE PATTERNS -----
# Fixed regex patterns with proper escaping
PAT_EN = re.compile(
    r"(?i)^(?:mr|mister|mrs|ms|miss|mx|sir|madam|ma'?am|master|dr|prof|eng|engr|arch|capt|cpt|maj|lt|lt\.?\s*col|col|gen|sgt|cdr|cmdr|adm|rev|fr|pastor|imam|rabbi|hrh|h\.?e\.?|h\.?h\.?|hon|rt\s*hon|lord|lady|prince|princess|king|queen|emir|m\.|mme|mlle|sr\.?|sra\.?|srta\.?|sig\.?|sig\.?ra|sig\.?na|dott\.?|dott\.?ssa|ing\.?|herr|frau|don|doña|shri|shree|sri|smt|kumari|bpk|ibu|encik|en\.?|puan|pn\.?|cik|tuan|datuk|dato'?|datin|tun|tunku)[\s./&\-]+"
)

PAT_AR = re.compile(
    r"^(?:أ\.?د\.?|أستاذ(?:ة)?|أ\.|د\.?|دكتور(?:ة)?|م\.?|مهندس(?:ة)?|سيد(?:ة)?|آنسة|مدام|حضرة|شيخ(?:ة)?|إمام|حاج(?:ة)?|سعادة|معالي|سمو)[\s./&\-]+"
)

def strip_titles_series_fixed(s: pd.Series) -> pd.Series:
    """Fixed title stripping with proper regex patterns"""
    s = s.fillna("").astype(str)
    
    def _strip_once(x):
        if not isinstance(x, str):
            return ""
        x = PAT_EN.sub("", x)
        x = PAT_AR.sub("", x)
        return x.strip()
    
    # Apply title stripping up to 3 times to handle nested titles
    for _ in range(3):
        s_new = s.apply(_strip_once)
        if s_new.equals(s):
            break
        s = s_new
    
    return s

# ----- ENHANCED GUEST NORMALIZATION -----
def guest_clean_base_fixed(x: str) -> str:
    """Enhanced guest name cleaning with better Unicode handling"""
    if not isinstance(x, str) or not x.strip():
        return ""
    
    try:
        # Step 1: Remove zero-width characters
        s = ZW_RE.sub("", x)
        
        # Step 2: ASCII folding (handle accents)
        s = ascii_fold_fixed(s)
        
        # Step 3: Arabic normalization
        s = ar_norm_fixed(s)
        
        # Step 4: Case folding (better than lower())
        s = s.casefold()
        
        # Step 5: Clean whitespace
        s = s.strip()
        
        # Step 6: Replace punctuation with spaces
        s = re.sub(r"[^\w\s]", " ", s)
        
        # Step 7: Normalize whitespace
        s = re.sub(r"\s+", " ", s).strip()
        
        return s
    except Exception:
        # Fallback for any encoding issues
        return re.sub(r"[^\w\s]", " ", str(x).lower().strip())

def guest_tokens_fixed(x: str) -> list:
    """Extract meaningful tokens from guest name"""
    cleaned = guest_clean_base_fixed(x)
    if not cleaned:
        return []
    
    tokens = [t for t in cleaned.split() if t and len(t) > 1]  # Ignore single characters
    return sorted(set(tokens))  # Remove duplicates and sort for consistency

# ----- ENHANCED HOTEL NORMALIZATION -----
def hotel_clean_base_fixed(x: str) -> str:
    """Enhanced hotel name cleaning"""
    if not isinstance(x, str) or not x.strip():
        return ""
    
    try:
        s = x.strip()
        s = ascii_fold_fixed(s)
        s = ar_norm_fixed(s)
        s = re.sub(r"[^\w\s]", " ", s)
        s = re.sub(r"\s+", " ", s).strip().lower()
        return s
    except Exception:
        return re.sub(r"[^\w\s]", " ", str(x).lower().strip())

def hotel_tokens_fixed(x: str, ignore_tokens: set, location_tokens: set) -> list:
    """Extract meaningful hotel tokens"""
    s = hotel_clean_base_fixed(x)
    if not s:
        return []
    
    tokens = [t for t in s.split() if t and not t.isdigit()]
    
    filtered_tokens = []
    for token in tokens:
        if token in ignore_tokens or token in location_tokens:
            continue
        if len(token) > 1:  # Ignore single characters
            filtered_tokens.append(token)
    
    return sorted(set(filtered_tokens))

def hotels_match_enhanced(h1: str, h2: str, ignore_tokens: set, location_tokens: set, 
                         fuzz_thr: int, jac_thr: float, mode: str, alias_pairs: set):
    """Enhanced hotel matching with better error handling"""
    
    try:
        # Check aliases first
        c_norm = hotel_clean_base_fixed(h1)
        a_norm = hotel_clean_base_fixed(h2)
        
        if (c_norm, a_norm) in alias_pairs or (a_norm, c_norm) in alias_pairs:
            return True, 100, 1.0, c_norm, a_norm
        
        # Strict mode
        if mode.startswith("Strict"):
            match = c_norm == a_norm
            return match, 100 if match else 0, 1.0 if match else 0.0, c_norm, a_norm
        
        # Smart mode with tokens
        t1 = set(hotel_tokens_fixed(h1, ignore_tokens, location_tokens))
        t2 = set(hotel_tokens_fixed(h2, ignore_tokens, location_tokens))
        
        # Handle empty tokens
        if not t1 and not t2:
            return True, 100, 1.0, "", ""
        if not t1 or not t2:
            return False, 0, 0.0, " ".join(sorted(t1)), " ".join(sorted(t2))
        
        # Calculate metrics
        intersection = len(t1 & t2)
        union = len(t1 | t2)
        jaccard = intersection / union if union > 0 else 0.0
        
        # Fuzzy matching on token strings
        t1_str = " ".join(sorted(t1))
        t2_str = " ".join(sorted(t2))
        fuzzy_score = fuzz.token_set_ratio(t1_str, t2_str) if t1_str and t2_str else 0
        
        # Match criteria
        match = (fuzzy_score >= fuzz_thr) or (jaccard >= jac_thr)
        
        return match, fuzzy_score, jaccard, t1_str, t2_str
        
    except Exception:
        return False, 0, 0.0, "", ""

# -------- ENHANCED DATE PARSING --------
# Fixed regex patterns
RE_YMD_START = re.compile(r"^\s*\d{4}[-/]")
RE_DMY_MDY = re.compile(r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})")

def parse_series_to_date_enhanced(s: pd.Series, mode: str) -> pd.Series:
    """Enhanced date parsing with better error handling"""
    
    if s.empty:
        return pd.Series(dtype='datetime64[ns]')
    
    s_str = s.astype(str).str.strip()
    
    try:
        if mode.startswith("YMD"):
            return pd.to_datetime(s_str, errors="coerce", yearfirst=True).dt.date
        elif mode.startswith("DMY"):
            return pd.to_datetime(s_str, errors="coerce", dayfirst=True).dt.date
        elif mode.startswith("MDY"):
            return pd.to_datetime(s_str, errors="coerce", dayfirst=False).dt.date
        
        # Auto mode - enhanced detection
        def parse_auto(x: str):
            if not isinstance(x, str) or not x.strip():
                return pd.NaT
            
            x = x.strip()
            
            # Try YYYY-MM-DD or ISO format first
            if RE_YMD_START.match(x) or "T" in x:
                return pd.to_datetime(x, errors="coerce", yearfirst=True)
            
            # Try DD/MM/YYYY vs MM/DD/YYYY detection
            match = RE_DMY_MDY.match(x)
            if match:
                try:
                    a, b, c = match.groups()
                    a_int, b_int = int(a), int(b)
                    
                    # If first number > 12, it must be day
                    if a_int > 12:
                        return pd.to_datetime(x, errors="coerce", dayfirst=True)
                    # If second number > 12, first must be month
                    elif b_int > 12:
                        return pd.to_datetime(x, errors="coerce", dayfirst=False)
                    # Default to DMY for ambiguous cases
                    else:
                        return pd.to_datetime(x, errors="coerce", dayfirst=True)
                        
                except (ValueError, TypeError):
                    pass
            
            # Final fallback
            return pd.to_datetime(x, errors="coerce")
        
        return s_str.apply(parse_auto).dt.date
        
    except Exception:
        # Ultimate fallback
        return pd.to_datetime(s_str, errors="coerce").dt.date

def fmt_date_only_series_fixed(s: pd.Series) -> pd.Series:
    """Enhanced date formatting with better error handling"""
    def format_date(d):
        try:
            if pd.isna(d):
                return ""
            # Convert to pandas datetime if it's a date object
            dt = pd.to_datetime(d, errors="coerce")
            if pd.isna(dt):
                return ""
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return ""
    
    return s.apply(format_date)

# ---------- UI: Upload ----------
st.markdown("""
<div class="success-box">
📋 <strong>Enhanced Features:</strong> Better file reading • Fixed encoding issues • Improved regex patterns • Enhanced error handling
</div>
""", unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    client_file = st.file_uploader("📤 Upload CLIENT unified file (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="client")
with c2:
    arkan_file = st.file_uploader("📤 Upload ARKAN sheet (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="arkan")

alias_file = st.file_uploader("📚 (Optional) Hotel alias mapping CSV (columns: client,arkan)", type=["csv"], key="alias")

if client_file and arkan_file:
    try:
        # Read files with enhanced error handling
        with st.spinner("Reading Client file..."):
            client_raw, client_note = read_any_enhanced(client_file)
        
        with st.spinner("Reading Arkan file..."):
            arkan_raw, arkan_note = read_any_enhanced(arkan_file)
        
        # Display file reading status
        col_status1, col_status2 = st.columns(2)
        with col_status1:
            st.markdown(f'<div class="success-box">Client: {client_note}</div>', unsafe_allow_html=True)
        with col_status2:
            st.markdown(f'<div class="success-box">Arkan: {arkan_note}</div>', unsafe_allow_html=True)
        
        # ---- Read alias mapping ----
        alias_pairs = set()
        if alias_file is not None:
            try:
                alias_df, alias_note = read_any_enhanced(alias_file)
                for _, r in alias_df.iterrows():
                    c = hotel_clean_base_fixed(str(r.get("client", "")))
                    a = hotel_clean_base_fixed(str(r.get("arkan", "")))
                    if c and a:
                        alias_pairs.add((c, a))
                st.success(f"✅ Loaded hotel alias mappings: {len(alias_pairs)} pair(s).")
            except Exception as e:
                st.warning(f"⚠️ Failed to read alias mapping: {e}")

        st.success(f"📊 Loaded Client: {len(client_raw):,} rows • Arkan: {len(arkan_raw):,} rows")

        # ---------- ENHANCED COLUMN MAPPING ----------
        st.subheader("🧭 Column mapping")
        st.caption("Enhanced normalization: Unicode handling • Fixed regex • Better date parsing • Improved token matching")

        def detect_column_enhanced(candidates, cols):
            """Enhanced column detection with fuzzy matching"""
            if not cols:
                return None
                
            cols_lower = {str(c).strip().lower(): c for c in cols}
            candidates_lower = [c.lower() for c in candidates]
            
            # Exact match first
            for cand in candidates_lower:
                if cand in cols_lower:
                    return cols_lower[cand]
            
            # Partial match
            for cand in candidates_lower:
                for col_lower, col_orig in cols_lower.items():
                    if cand in col_lower or col_lower in cand:
                        return col_orig
            
            # Fuzzy match as last resort
            for cand in candidates_lower:
                for col_lower, col_orig in cols_lower.items():
                    if fuzz.partial_ratio(cand, col_lower) > 80:
                        return col_orig
            
            return cols[0] if cols else None

        def detect_column_prioritized_enhanced(priority_list, fallback_list, cols):
            """Enhanced prioritized column detection"""
            result = detect_column_enhanced(priority_list, cols)
            if result:
                return result
            return detect_column_enhanced(fallback_list, cols)

        # Enhanced column detection
        hotel_aliases = ["accommodation","accomodation","accomm","accom","hotel name","hotel","property"]
        client_expected = {
            "Booking Reference": ["booking reference","booking ref","booking code","client reference","dotw ref. #","dotw ref #","reference"],
            "Hotel Name": hotel_aliases,
            "Guest Name": ["guest name","guest","holder","lead"],
            "Arrival Date": ["arrival date","check-in date","check in","arrival"],
            "Departure Date": ["departure date","check-out date","check out","departure"],
        }
        
        arkan_expected = {
            "JoinKey": {
                "priority": ["clientreference", "client reference", "client ref"],
                "fallback": ["bookingnumber","booking number","reference","reservation no","booking ref","booking code","dotw ref. #","dotw ref #"]
            },
            "HotelName": hotel_aliases,
            "GuestName": ["guestname","guest name","guest","holder","lead"],
            "ArrivalDate": ["arrivaldate","arrival date","check-in date","check in","arrival"],
            "DepartureDate": ["departuredate","departure date","check-out date","check out","departure"],
        }

        cmaps = {}
        amaps = {}
        cols_client = client_raw.columns.tolist()
        cols_arkan = arkan_raw.columns.tolist()

        c3, c4 = st.columns(2)
        with c3:
            st.markdown("**Client columns**")
            for k, cands in client_expected.items():
                default = detect_column_enhanced(cands, cols_client)
                default_idx = cols_client.index(default) if default in cols_client else 0
                cmaps[k] = st.selectbox(f"Client {k}", cols_client, index=default_idx, key=f"client_{k}")

        with c4:
            st.markdown("**Arkan columns**")
            join_default = detect_column_prioritized_enhanced(
                arkan_expected["JoinKey"]["priority"], 
                arkan_expected["JoinKey"]["fallback"], 
                cols_arkan
            )
            join_idx = cols_arkan.index(join_default) if join_default in cols_arkan else 0
            amaps["JoinKey"] = st.selectbox("Arkan Reference", cols_arkan, index=join_idx, key="arkan_join")
            
            for label, cands in [("HotelName", arkan_expected["HotelName"]),
                                 ("GuestName", arkan_expected["GuestName"]),
                                 ("ArrivalDate", arkan_expected["ArrivalDate"]),
                                 ("DepartureDate", arkan_expected["DepartureDate"])]:
                default = detect_column_enhanced(cands, cols_arkan)
                default_idx = cols_arkan.index(default) if default in cols_arkan else 0
                amaps[label] = st.selectbox(f"Arkan {label}", cols_arkan, index=default_idx, key=f"arkan_{label}")

        # ---------- ENHANCED DATA PROCESSING ----------
        with st.spinner("Processing and normalizing data..."):
            
            # Create normalized DataFrames
            client = pd.DataFrame({
                "BookingRef": client_raw[cmaps["Booking Reference"]].fillna(""),
                "Hotel": client_raw[cmaps["Hotel Name"]].fillna(""),
                "Guest": client_raw[cmaps["Guest Name"]].fillna(""),
                "Arr": client_raw[cmaps["Arrival Date"]].fillna(""),
                "Dep": client_raw[cmaps["Departure Date"]].fillna(""),
            })
            
            arkan = pd.DataFrame({
                "BookingRef": arkan_raw[amaps["JoinKey"]].fillna(""),
                "Hotel": arkan_raw[amaps["HotelName"]].fillna(""),
                "Guest": arkan_raw[amaps["GuestName"]].fillna(""),
                "Arr": arkan_raw[amaps["ArrivalDate"]].fillna(""),
                "Dep": arkan_raw[amaps["DepartureDate"]].fillna(""),
            })

            # Enhanced reference normalization
            def conditional_ref_series_enhanced(s: pd.Series) -> pd.Series:
                """Enhanced reference normalization with better regex"""
                HTL_WBD_RE = re.compile(r"(?i)^\s*HTL-WBD-(\d+)\s*$")
                s = s.fillna("").astype(str)
                
                result = []
                for val in s:
                    try:
                        raw = str(val).strip()
                        match = HTL_WBD_RE.match(raw)
                        if match:
                            num = match.group(1)
                            if strip_leading_zeros:
                                num = re.sub(r"^0+", "", num) or "0"
                            result.append(num)
                        else:
                            result.append(raw)
                    except Exception:
                        result.append(str(val).strip())
                
                return pd.Series(result, index=s.index, dtype="object")

            client["BookingRef_norm"] = conditional_ref_series_enhanced(client["BookingRef"])
            arkan["BookingRef_norm"] = conditional_ref_series_enhanced(arkan["BookingRef"])

            # Enhanced guest normalization with title removal
            client["Guest_clean"] = strip_titles_series_fixed(client["Guest"]).apply(guest_clean_base_fixed)
            arkan["Guest_clean"] = strip_titles_series_fixed(arkan["Guest"]).apply(guest_clean_base_fixed)

            # Enhanced date parsing
            client["Arr_norm"] = parse_series_to_date_enhanced(client["Arr"], client_date_mode)
            client["Dep_norm"] = parse_series_to_date_enhanced(client["Dep"], client_date_mode)
            arkan["Arr_norm"] = parse_series_to_date_enhanced(arkan["Arr"], arkan_date_mode)
            arkan["Dep_norm"] = parse_series_to_date_enhanced(arkan["Dep"], arkan_date_mode)

        # ---------- ENHANCED MATCHING LOGIC ----------
        with st.spinner("Performing matching analysis..."):
            
            # Merge on normalized reference
            merged = client.merge(
                arkan, 
                how="left", 
                left_on="BookingRef_norm", 
                right_on="BookingRef_norm", 
                suffixes=("_Client", "_Arkan")
            )
            merged["Exists_in_Arkan"] = ~merged["Hotel_Arkan"].isna()

            # Prepare token sets for hotel matching
            ignore_tokens = set([t.strip().lower() for t in ignore_tokens_input.split(",") if t.strip()])
            location_tokens = set([t.strip().lower() for t in location_tokens_input.split(",") if t.strip()])

            # Enhanced hotel matching
            hotel_results = []
            for idx, row in merged.iterrows():
                h1 = row.get("Hotel_Client", "")
                h2 = row.get("Hotel_Arkan", "")
                
                if pd.isna(h1) or pd.isna(h2) or not row["Exists_in_Arkan"]:
                    hotel_results.append((False, 0, 0.0, "", ""))
                else:
                    try:
                        result = hotels_match_enhanced(
                            str(h1), str(h2), ignore_tokens, location_tokens,
                            fuzzy_threshold_hotel, jaccard_threshold, mode_hotel, alias_pairs
                        )
                        hotel_results.append(result)
                    except Exception:
                        hotel_results.append((False, 0, 0.0, "", ""))

            # Unpack hotel matching results
            hotel_matches, hotel_fuzzy, hotel_jaccard, hotel_tokens_client, hotel_tokens_arkan = zip(*hotel_results)
            
            merged["Hotel_Match"] = list(hotel_matches)
            merged["Hotel_Fuzzy"] = list(hotel_fuzzy)
            merged["Hotel_Jaccard"] = list(hotel_jaccard)
            merged["Hotel_Tokens_Client"] = list(hotel_tokens_client)
            merged["Hotel_Tokens_Arkan"] = list(hotel_tokens_arkan)

            # Enhanced guest matching
            def guest_match_enhanced(g1: str, g2: str):
                """Enhanced guest matching with better error handling"""
                try:
                    g1_clean = guest_clean_base_fixed(str(g1) if pd.notna(g1) else "")
                    g2_clean = guest_clean_base_fixed(str(g2) if pd.notna(g2) else "")
                    
                    if not g1_clean or not g2_clean:
                        return False, 0
                    
                    if mode_guest.startswith("Exact"):
                        match = g1_clean == g2_clean
                        score = 100 if match else fuzz.token_set_ratio(g1_clean, g2_clean)
                        return match, score
                    else:
                        # Smart fuzzy matching on tokens
                        t1 = guest_tokens_fixed(g1_clean)
                        t2 = guest_tokens_fixed(g2_clean)
                        
                        if not t1 or not t2:
                            return False, 0
                        
                        score = fuzz.token_set_ratio(" ".join(t1), " ".join(t2))
                        match = score >= fuzzy_threshold_guest
                        return match, score
                        
                except Exception:
                    return False, 0

            # Apply guest matching
            guest_results = []
            for idx, row in merged.iterrows():
                g1 = row.get("Guest_Client", "")
                g2 = row.get("Guest_Arkan", "")
                match, score = guest_match_enhanced(g1, g2)
                guest_results.append((match, score))

            guest_matches, guest_scores = zip(*guest_results)
            merged["Guest_Match_Raw"] = list(guest_matches)
            merged["Guest_Fuzzy"] = list(guest_scores)

            # Apply guest matching logic based on settings
            if compare_guest_even_if_hotel_mismatch:
                merged["Guest_Match"] = merged["Guest_Match_Raw"]
            else:
                merged["Guest_Match"] = np.where(
                    merged["Hotel_Match"], 
                    merged["Guest_Match_Raw"], 
                    False
                )

            # Enhanced date matching
            merged["Arrival_Match"] = np.where(
                merged["Guest_Match"] & merged["Exists_in_Arkan"],
                merged["Arr_norm_Client"] == merged["Arr_norm_Arkan"],
                False
            )
            
            merged["Departure_Match"] = np.where(
                merged["Guest_Match"] & merged["Exists_in_Arkan"],
                merged["Dep_norm_Client"] == merged["Dep_norm_Arkan"],
                False
            )
            
            merged["Dates_Match"] = merged["Arrival_Match"] & merged["Departure_Match"]

            # Enhanced status determination
            def determine_status(row):
                """Enhanced status determination with clearer categories"""
                try:
                    if not row["Exists_in_Arkan"]:
                        return "❌ Missing in Arkan"
                    elif not row["Hotel_Match"]:
                        return "🏨 Found by Reference (Hotel mismatch)"
                    elif not row["Guest_Match"]:
                        return "🧑 Match: Hotel only (Guest mismatch)"
                    elif not row["Dates_Match"]:
                        return "📅 Match: Hotel & Guest (Dates mismatch)"
                    else:
                        return "✅ Full Match"
                except Exception:
                    return "❓ Unknown Status"

            merged["Status"] = merged.apply(determine_status, axis=1)

            # Format dates for display
            merged["Arr_Client_out"] = fmt_date_only_series_fixed(merged.get("Arr_norm_Client", pd.Series()))
            merged["Dep_Client_out"] = fmt_date_only_series_fixed(merged.get("Dep_norm_Client", pd.Series()))
            merged["Arr_Arkan_out"] = fmt_date_only_series_fixed(merged.get("Arr_norm_Arkan", pd.Series()))
            merged["Dep_Arkan_out"] = fmt_date_only_series_fixed(merged.get("Dep_norm_Arkan", pd.Series()))

        # ---------- ENHANCED REPORTING ----------
        with st.spinner("Generating reports..."):
            
            # Define desired columns for output
            desired_cols = [
                "BookingRef_norm",
                "BookingRef_Client", "BookingRef_Arkan",
                "Exists_in_Arkan", "Hotel_Match", "Hotel_Fuzzy", "Hotel_Jaccard",
                "Guest_Match", "Guest_Match_Raw", "Guest_Fuzzy",
                "Arrival_Match", "Departure_Match", "Dates_Match", "Status",
                "Hotel_Client", "Hotel_Arkan", "Hotel_Tokens_Client", "Hotel_Tokens_Arkan",
                "Guest_Client", "Guest_Arkan", "Guest_clean_Client", "Guest_clean_Arkan",
                "Arr_Client_out", "Arr_Arkan_out", "Dep_Client_out", "Dep_Arkan_out"
            ]

            # Add cleaned guest columns to merged DataFrame
            if "Guest_clean_Client" not in merged.columns:
                merged["Guest_clean_Client"] = client["Guest_clean"]
            if "Guest_clean_Arkan" not in merged.columns:
                merged["Guest_clean_Arkan"] = arkan["Guest_clean"]

            # Select available columns
            available_cols = [c for c in desired_cols if c in merged.columns]
            report = merged[available_cols].copy()

            # Generate analysis reports
            full_matches = report[report["Status"] == "✅ Full Match"].copy()
            differences = report[report["Status"] != "✅ Full Match"].copy()

            # Find records only in Arkan
            client_refs = set(client["BookingRef_norm"].astype(str).str.strip())
            arkan_only = arkan[~arkan["BookingRef_norm"].astype(str).str.strip().isin(client_refs)].copy()

            # Enhanced summary with percentages
            total_client = len(client)
            found_in_arkan = int(merged["Exists_in_Arkan"].sum())
            full_match_count = len(full_matches)
            partial_count = len(differences)
            arkan_only_count = len(arkan_only)

            summary = pd.DataFrame({
                "Metric": [
                    "Total Client Rows",
                    "Found in Arkan", 
                    "Full Matches",
                    "Partial/Issues",
                    "Only in Arkan",
                    "Match Rate (%)",
                    "Coverage Rate (%)"
                ],
                "Count": [
                    total_client,
                    found_in_arkan,
                    full_match_count,
                    partial_count,
                    arkan_only_count,
                    f"{(full_match_count/total_client*100):.1f}%" if total_client > 0 else "0.0%",
                    f"{(found_in_arkan/total_client*100):.1f}%" if total_client > 0 else "0.0%"
                ]
            })

            # Status breakdown
            status_breakdown = report["Status"].value_counts().reset_index()
            status_breakdown.columns = ["Status", "Count"]
            status_breakdown["Percentage"] = (status_breakdown["Count"] / total_client * 100).round(1)

        # ---------- ENHANCED UI DISPLAY ----------
        st.subheader("📊 Enhanced Analysis Summary")
        
        col_sum1, col_sum2 = st.columns(2)
        with col_sum1:
            st.dataframe(summary, use_container_width=True)
        with col_sum2:
            st.dataframe(status_breakdown, use_container_width=True)

        # Quality metrics
        if found_in_arkan > 0:
            hotel_match_rate = (merged["Hotel_Match"].sum() / found_in_arkan * 100)
            guest_match_rate = (merged["Guest_Match"].sum() / found_in_arkan * 100)
            date_match_rate = (merged["Dates_Match"].sum() / found_in_arkan * 100)
            
            st.markdown(f"""
            <div class="success-box">
            📈 <strong>Quality Metrics:</strong> 
            Hotel Match: {hotel_match_rate:.1f}% • 
            Guest Match: {guest_match_rate:.1f}% • 
            Date Match: {date_match_rate:.1f}%
            </div>
            """, unsafe_allow_html=True)

        # Enhanced previews with better formatting
        st.subheader("✅ Full Matches (preview)")
        if len(full_matches) > 0:
            display_cols = [c for c in ["BookingRef_norm", "Hotel_Client", "Guest_Client", 
                           "Arr_Client_out", "Dep_Client_out", "Status"] if c in full_matches.columns]
            st.dataframe(full_matches[display_cols].head(100), use_container_width=True)
        else:
            st.info("No full matches found.")

        st.subheader("⚠️ Differences / Partials (preview)")
        if len(differences) > 0:
            display_cols = [c for c in ["BookingRef_norm", "Status", "Hotel_Match", "Guest_Match", 
                           "Hotel_Client", "Hotel_Arkan", "Hotel_Fuzzy"] if c in differences.columns]
            st.dataframe(differences[display_cols].head(100), use_container_width=True)
        else:
            st.success("All records are full matches!")

        # Enhanced download section
        st.subheader("💾 Download Reports")
        
        try:
            # Create Excel report with enhanced formatting
            out_excel = io.BytesIO()
            with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
                # Summary sheet
                summary.to_excel(writer, sheet_name="Summary", index=False)
                status_breakdown.to_excel(writer, sheet_name="Status Breakdown", index=False)
                
                # Main reports
                full_matches.to_excel(writer, sheet_name="Full Matches", index=False)
                differences.to_excel(writer, sheet_name="Differences", index=False)
                arkan_only.to_excel(writer, sheet_name="Only in Arkan", index=False)
                
                # Add complete report for analysis
                report.to_excel(writer, sheet_name="Complete Analysis", index=False)
            
            out_excel.seek(0)
            
            st.download_button(
                "📊 Download Enhanced Comparison Report (Excel)",
                out_excel.getvalue(),
                file_name=f"Enhanced_Client_vs_Arkan_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error creating Excel report: {e}")

    except Exception as e:
        st.error(f"❌ Critical Error: {e}")
        st.markdown("""
        <div class="error-box">
        <strong>Error Details:</strong><br>
        • Check that your files are not corrupted<br>
        • Ensure files contain the expected columns<br>
        • Try converting to CSV format if using old Excel files<br>
        • Contact support if the issue persists
        </div>
        """, unsafe_allow_html=True)

else:
    st.info("🔼 ارفع ملف العميل وملف أركان للبدء في المقارنة المحسنة. الإصدار الجديد يحل مشاكل الترميز ويحسن دقة المطابقة.")

# ---------- DEBUG SECTION ----------
with st.sidebar:
    with st.expander("🔧 Debug Information", expanded=False):
        st.write("**Available Libraries:**")
        try:
            import xlrd
            st.success(f"✅ xlrd {xlrd.__version__}")
        except ImportError:
            st.error("❌ xlrd not available")
        
        try:
            import openpyxl
            st.success(f"✅ openpyxl {openpyxl.__version__}")
        except ImportError:
            st.error("❌ openpyxl not available")
        
        try:
            from rapidfuzz import fuzz
            st.success("✅ rapidfuzz available")
        except ImportError:
            st.error("❌ rapidfuzz not available")
        
        st.write(f"**Pandas:** {pd.__version__}")
        st.write("**Fixed Issues:**")
        st.write("• Unicode encoding problems")
        st.write("• Regex pattern errors")
        st.write("• File reading fallbacks")
        st.write("• Date parsing enhancement")
        st.write("• Error handling improvement")

st.markdown("<br>", unsafe_allow_html=True)
st.caption("© Client vs Arkan Comparator v1.7.0 • Enhanced & Fixed • Better accuracy & reliability")