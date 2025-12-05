from difflib import SequenceMatcher
import spacy
import re

nlp = spacy.load("en_core_web_sm")

COMMON_PREFIXES = [r"\bm/s\b", r"\bm s\b", r"\bms\b"]

def normalize_name(name: str) -> str:
    """Normalize a company/person name for comparison."""
    if not name:
        return ""

    s = name.lower()
    # replace & with and
    s = s.replace("&", " and ")

    # remove common prefixes like m/s
    for p in COMMON_PREFIXES:
        s = re.sub(p, " ", s)

    s = re.sub(r"\.", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # run through spaCy to keep alphabetic tokens and preserve token order
    doc = nlp(s)
    tokens = [tok.text for tok in doc if tok.text.strip() != ""]
    normalized = " ".join(tokens)
    return normalized

def token_set_similarity(a: str, b: str) -> float:
    """Token set overlap ratio (Jaccard-like)."""
    sa = set(a.split())
    sb = set(b.split())
    if not sa and not sb:
        return 1.0
    inter = sa.intersection(sb)
    union = sa.union(sb)
    return len(inter) / len(union)

def sequence_similarity(a: str, b: str) -> float:
    """Character-level similarity via SequenceMatcher."""
    return SequenceMatcher(None, a, b).ratio()

def is_same_name(name1: str, name2: str,
                 token_set_threshold: float = 0.85,
                 seq_threshold: float = 0.88) -> dict:
    
    norm1 = normalize_name(name1)
    norm2 = normalize_name(name2)

    # Exact match after normalization
    if norm1 == norm2 and norm1 != "":
        return {"same": True, "reason": "exact_normalized_match", "scores": None, "norm1": norm1, "norm2": norm2}

    # Token-set similarity (order independent)
    ts_sim = token_set_similarity(norm1, norm2)
    if ts_sim >= token_set_threshold:
        return {"same": True, "reason": f"token_set_similarity >= {token_set_threshold:.2f}",
                "scores": {"token_set": ts_sim}, "norm1": norm1, "norm2": norm2}

    # Sequence similarity (order + characters)
    seq_sim = sequence_similarity(norm1, norm2)
    if seq_sim >= seq_threshold:
        return {"same": True, "reason": f"sequence_similarity >= {seq_threshold:.2f}",
                "scores": {"sequence": seq_sim, "token_set": ts_sim},
                "norm1": norm1, "norm2": norm2}

    # If both short and share all tokens (rare case)
    if norm1.split() == [] and norm2.split() == []:
        return {"same": True, "reason": "both_empty_after_normalization", "scores": None, "norm1": norm1, "norm2": norm2}

    # Otherwise not same
    return {"same": False, "reason": "below_thresholds", "scores": {"token_set": ts_sim, "sequence": seq_sim},
            "norm1": norm1, "norm2": norm2}

def description_contains_category(category: str, description: str,
                                  max_concat=4,
                                  partial_match_threshold=0.5,
                                  seq_threshold=0.60) -> dict:
    """
    Improved containment test:
      - normalized substring
      - token-set subset
      - concat-token match (handles 'EAR TH' -> 'EARTH')
      - compact string match
      - partial token-proportion match (>= partial_match_threshold)
      - relaxed sequence similarity fallback (>= seq_threshold)

    Returns same dict shape as before.
    """
    norm_cat = normalize_name(category)
    norm_desc = normalize_name(description)

    if not norm_cat:
        return {
            "contains": False,
            "reason": "empty_normalized_category",
            "norm_category": norm_cat,
            "norm_description": norm_desc,
            "cat_tokens": [],
            "desc_tokens": norm_desc.split()
        }

    cat_tokens = norm_cat.split()
    desc_tokens = norm_desc.split()

    # 1) normalized substring
    if norm_cat in norm_desc:
        return {"contains": True, "reason": "normalized_substring_match",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens}

    # 2) full token-set subset
    if set(cat_tokens).issubset(set(desc_tokens)):
        return {"contains": True, "reason": "token_set_subset_match",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens}

    # helper: can ct be built by concatenating up to max_concat desc tokens?
    def ct_matches_by_concat(ct, desc_tokens, max_concat):
        for i in range(len(desc_tokens)):
            concat = desc_tokens[i]
            if concat == ct:
                return True
            for j in range(i+1, min(i+max_concat, len(desc_tokens))):
                concat += desc_tokens[j]
                if concat == ct:
                    return True
        return False

    # 3) cat tokens match by concatenating consecutive desc tokens (handles splits like 'EAR TH' -> 'EARTH')
    all_matched = True
    for ct in cat_tokens:
        if ct in desc_tokens:
            continue
        if ct_matches_by_concat(ct, desc_tokens, max_concat):
            continue
        all_matched = False
        break
    if all_matched:
        return {"contains": True, "reason": "concat_token_match",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens}

    # 4) compact string match
    compact_cat = "".join(cat_tokens)
    compact_desc = "".join(desc_tokens)
    if compact_cat in compact_desc:
        return {"contains": True, "reason": "compact_string_match",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens}

    # 5) partial token proportion match
    # Count how many category tokens appear (directly or via concatenation) in description
    matched_count = 0
    for ct in cat_tokens:
        if ct in desc_tokens or ct_matches_by_concat(ct, desc_tokens, max_concat):
            matched_count += 1

    proportion = matched_count / max(1, len(cat_tokens))
    if proportion >= partial_match_threshold:
        return {"contains": True,
                "reason": f"partial_token_proportion_match >= {partial_match_threshold} ({matched_count}/{len(cat_tokens)})",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens,
                "matched_count": matched_count, "proportion": proportion}

    # 6) relaxed sequence similarity fallback (handles OCR noise and small re-ordering)
    seq_sim = SequenceMatcher(None, norm_cat, norm_desc).ratio()
    if seq_sim >= seq_threshold:
        return {"contains": True,
                "reason": f"sequence_similarity >= {seq_threshold:.2f}",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens,
                "sequence_similarity": seq_sim}

    # 7) token-set similarity fallback (kept for compatibility)
    ts_sim = token_set_similarity(norm_cat, norm_desc)
    if ts_sim >= 0.75:
        return {"contains": True,
                "reason": f"token_set_similarity_high (>=0.75) ({ts_sim:.2f})",
                "norm_category": norm_cat, "norm_description": norm_desc,
                "cat_tokens": cat_tokens, "desc_tokens": desc_tokens,
                "score": ts_sim}

    # no match
    return {"contains": False, "reason": "no_match_below_thresholds",
            "norm_category": norm_cat, "norm_description": norm_desc,
            "cat_tokens": cat_tokens, "desc_tokens": desc_tokens,
            "matched_count": matched_count, "proportion": proportion,
            "sequence_similarity": seq_sim, "score": ts_sim}
