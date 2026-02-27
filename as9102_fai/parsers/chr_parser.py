import pandas as pd
import re
import logging


logger = logging.getLogger(__name__)
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class FaiCharacteristic:
    id: str
    feature_name: str
    description: str  # The requirement/tolerance string
    actual: str
    nominal: str
    upper_tol: str
    lower_tol: str
    type: str
    unit: str = ""
    group1: str = ""
    # Calypso metadata (optional; present in some exports)
    idsymbol: str = ""
    mmc: str = ""
    datumaid: str = ""
    datumbid: str = ""
    datumcid: str = ""
    # Marks where the row originated. Calypso-imported rows are "calypso".
    # Rows synthesized by the parser (thread Go/No-Go, minor dia, etc.) are "derived".
    source: str = "calypso"
    is_thread: bool = False
    is_attribute: bool = False
    # Original note/comment text from the CHR export (if present).
    comment: str = ""
    
class ChrParser:
    def __init__(self):
        # Regex to identify thread features from ID, Feature Name, or Comments
        self.thread_patterns = [
            r'\bM\d+(?:x\d+)?\b',  # Metric M6, M6x1
            r'\bUN\b', r'\bUNC\b', r'\bUNF\b', r'\bUNEF\b',
            r'\bUNJ\b', r'\bUNJC\b', r'\bUNJF\b',
            r'\bNPT\b', r'\bNPTF\b',
            r'\bBSP\b', r'\bBSPT\b',
            r'\bSTI\b',
            r'\bACME\b', r'\bSTUB ACME\b',
            r'\bWHITWORTH\b'
        ]
        self.thread_regex = re.compile('|'.join(self.thread_patterns), re.IGNORECASE)

    def parse_file(self, file_path: str) -> List[FaiCharacteristic]:
        try:
            # Read the file. The example shows a header row.
            # We use on_bad_lines='skip' to be robust against malformed lines
            df = pd.read_csv(file_path, sep='\t', encoding='utf-8', on_bad_lines='skip')
        except Exception as e:
            logger.exception("Error reading CHR file")
            return []

        characteristics = []
        
        # Normalize column names to lowercase and strip whitespace
        df.columns = [c.lower().strip() for c in df.columns]

        def _infer_unit_column_name() -> str:
            """Best-effort discovery of a unit column in Calypso exports."""
            direct_candidates = [
                "unit",
                "uom",
                "unitofmeasure",
                "unitofmeasurement",
                "unit of measure",
                "unit of measurement",
            ]
            for name in direct_candidates:
                if name in df.columns:
                    return name
            for c in df.columns:
                c_low = str(c).lower()
                if c_low == "uom":
                    return c
                if "unit" in c_low and ("measure" in c_low or "measurement" in c_low):
                    return c
            return ""

        unit_col = _infer_unit_column_name()

        def _infer_group1_column_name() -> str:
            candidates = [
                "group1",
                "group 1",
                "group",
                "grp1",
                "grp 1",
            ]
            for name in candidates:
                if name in df.columns:
                    return name
            for c in df.columns:
                c_low = str(c).lower().strip()
                if c_low.replace(" ", "") == "group1":
                    return c
            return ""

        group1_col = _infer_group1_column_name()
        
        # Check for essential columns. 
        # Based on attachment: 'id', 'actual', 'nominal', 'uppertol', 'lowertol', 'type'
        # 'featureid' and 'comment' are also useful.
        
        for index, row in df.iterrows():
            # Helper to safely get string value
            def get_val(col):
                val = row.get(col, '')
                return str(val) if not pd.isna(val) else ''

            char_id = get_val('id')
            feature_id = get_val('featureid')
            comment = get_val('comment')
            type_str = get_val('type')
            unit_str = get_val(unit_col) if unit_col else ""
            group1_str = get_val(group1_col) if group1_col else ""
            idsymbol_str = get_val('idsymbol')
            mmc_str = get_val('mmc')
            datumaid_str = get_val('datumaid')
            datumbid_str = get_val('datumbid')
            datumcid_str = get_val('datumcid')
            
            # Detect Thread
            is_thread = self._is_thread(char_id) or self._is_thread(feature_id) or self._is_thread(comment)
            
            # Format Requirement String
            nominal = row.get('nominal', 0)
            upper = row.get('uppertol', 0)
            lower = row.get('lowertol', 0)
            
            requirement_text = self._format_requirement(nominal, upper, lower, type_str)
            
            char = FaiCharacteristic(
                id=char_id,
                feature_name=feature_id,
                description=requirement_text,
                actual=get_val('actual'),
                nominal=get_val('nominal'),
                upper_tol=get_val('uppertol'),
                lower_tol=get_val('lowertol'),
                type=type_str,
                unit=unit_str,
                group1=group1_str,
                idsymbol=idsymbol_str,
                mmc=mmc_str,
                datumaid=datumaid_str,
                datumbid=datumbid_str,
                datumcid=datumcid_str,
                source="calypso",
                is_thread=is_thread,
                comment=comment,
            )
            characteristics.append(char)

        # Post-process for threads (add Go/No Go and Minor Dia if missing)
        final_characteristics = self._expand_threads(characteristics)
        
        return final_characteristics

    def _is_thread(self, text: str) -> bool:
        if not text: return False
        return bool(self.thread_regex.search(text))

    def _format_requirement(self, nominal, upper, lower, type_str) -> str:
        def strip_leading_zero(val: float) -> str:
            """Format number without leading zero before decimal (e.g., 0.1234 -> .1234)"""
            if val == 0:
                return ".0000"
            formatted = f"{abs(val):.4f}"
            if formatted.startswith("0."):
                formatted = formatted[1:]  # Remove leading 0
            return formatted

        try:
            nom_val = float(nominal)
            up_val = float(upper)
            low_val = float(lower)
        except (ValueError, TypeError):
            return str(nominal)

        # Basic Dimension (999 tolerance) - just show the value without brackets/BASIC
        if abs(up_val) >= 990 and abs(low_val) >= 990:
            return strip_leading_zero(nom_val)
        
        # Position/GD&T tolerance where nominal is 0 and upper > 0 -> show as "MAX"
        if nom_val == 0 and up_val > 0 and low_val == 0:
            return f"{strip_leading_zero(up_val)} MAX"
        
        # Bilateral Tolerance (equal + and -)
        if abs(up_val) == abs(low_val) and up_val != 0:
            return f"{strip_leading_zero(nom_val)} +/- {strip_leading_zero(up_val)}"
        
        # Unilateral / Unequal Tolerance
        if up_val != 0 or low_val != 0:
            up_str = f"+{strip_leading_zero(up_val)}"
            low_str = f"-{strip_leading_zero(abs(low_val))}" if low_val <= 0 else f"+{strip_leading_zero(low_val)}"
            return f"{strip_leading_zero(nom_val)} {up_str}/{low_str}"
            
        # If tolerances are 0, check type. 
        # Flatness, Perpendicularity etc usually are "MAX" or just the tolerance value.
        # But in the file, Flatness has uppertol 0.004, lowertol 0.0.
        # Wait, the logic above catches uppertol != 0.
        # Example: Flatness 0.004, 0.0 -> "0.0000 +0.0040 / 0.0000" which is weird for GD&T.
        # GD&T usually is just the tolerance zone.
        
        if 'flatness' in type_str.lower() or 'profile' in type_str.lower() or 'position' in type_str.lower():
             # For GD&T, the 'nominal' is often 0 in Calypso for the deviation, 
             # but here we see nominal 0.0000 for Flatness.
             # The requirement is the upper tolerance.
             if nom_val == 0 and up_val > 0:
                 return f"{strip_leading_zero(up_val)} MAX"

        return strip_leading_zero(nom_val)

    def _expand_threads(self, chars: List[FaiCharacteristic]) -> List[FaiCharacteristic]:
        expanded = []
        
        for char in chars:
            expanded.append(char)
            
            if char.is_thread and not char.is_attribute:
                # 1. Add Go/No Go Check
                go_no_go = FaiCharacteristic(
                    id=f"{char.id}_GNG",
                    feature_name=f"{char.feature_name} Go/No Go",
                    description="Thread Attribute",
                    actual="", # User to fill or default to Pass
                    nominal="Pass",
                    upper_tol="",
                    lower_tol="",
                    type="Attribute",
                    unit=char.unit,
                    group1=char.group1,
                    source="derived",
                    is_attribute=True,
                    is_thread=True
                )
                expanded.append(go_no_go)
                
                # 2. Add Minor Diameter Check
                # Check if a "Minor" feature already exists for this thread.
                found_minor = False
                
                # If the current char description says "Minor", don't add another.
                if "minor" in char.feature_name.lower() or "minor" in char.id.lower():
                    found_minor = True

                if not found_minor:
                    # Scan other chars to see if they are the minor dia for this thread
                    for existing in chars:
                        if existing == char: continue
                        
                        is_minor = "minor" in existing.feature_name.lower() or "minor" in existing.id.lower()
                        if is_minor:
                            # Check if it relates to this thread
                            # e.g. if char.id is "Thread1", does existing.id contain "Thread1"?
                            # Or if feature_name matches
                            if (char.id and char.id.lower() in existing.id.lower()) or \
                               (char.feature_name and char.feature_name.lower() in existing.feature_name.lower()):
                                 found_minor = True
                                 break

                if not found_minor:
                    minor_dia = FaiCharacteristic(
                        id=f"{char.id}_MIN",
                        feature_name=f"{char.feature_name} Minor Dia",
                        description="Minor Diameter Check",
                        actual="", 
                        nominal="",
                        upper_tol="",
                        lower_tol="",
                        type="Attribute", 
                        unit=char.unit,
                        group1=char.group1,
                        source="derived",
                        is_attribute=True,
                        is_thread=True
                    )
                    expanded.append(minor_dia)
                
        return expanded
