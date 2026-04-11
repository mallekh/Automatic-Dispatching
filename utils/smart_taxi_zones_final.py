"""
smart_taxi_zones.py  — Grand Tunis Taxi Grouping Engine

STEP 1 — Same exact destination (case-insensitive):
         Agents sharing the EXACT same address → closed taxi(s) of max 4.
         These taxis are NEVER mixed with others.
         Remainder after filling 4s → open for Step 2.

STEP 2 — Same zone (A–K), free packing:
         All remaining agents in a zone are sorted by address then packed
         into taxis of 4 freely. No block constraints — agents from
         different addresses within the same zone CAN share a taxi.
         Only full taxis (4) assigned here; partial remainder → Step 3.

STEP 3 — Zone-match merging (ZONE_MATCH_RULES):
         Leftover agents from partner zones pooled and packed freely.
         Sorted by zone then address, chunked by 4.
"""

import pandas as pd

MAX_PER_TAXI = 4

ZONES: dict = {
    "A": ["borj cedria","hammam chat","hammam lif","ezzahra","yasminette","ben arous","bir bay","bir el bey"],
    "B": ["rades","medina jedida","madina jedida","megrine","mornag","boumhal","medina jadida","madina jadida"],
    "D": ["mhamdia","fouchena","mourouj 1","mourouj 2","mourouj 3","mourouj 4",
          "mourouj 5","mourouj 6","mourouj","naasen","ibn sina","ibnou sina","wardia","ouardia","kabaria",
          "cite hlel","saida","el mourouj","montfleury","belle vue","bellevue","cite elfath","elfath","cite olympique","olympique"],
    "E": ["lafayette","passage","centre ville","cite khadhra","avenue madrid","madrid tunis"],
    "F": ["mornageya","diar ben mahmoud","agba","manouba","mannouba","kobbet ennhass","danden","khaznadar","10 decembre","monoprix 10",
          "bardo","sidi hsin","sidi hssine","zahrouni","zouhour","beb saadoun","omran",
          "tebourba","tborba","ksar said","cite ezzouhour"],
    "G": ["jdaida","bjeoua","oued lil","mnihla","bassatine","hay bassatine",
          "omran superieur","tahrir","ettahrir","intilaka","tadhamon","tadhamen","thadhamen","ettadhamen",
          "manar","jardin lmanzah","nasseer","zayatine",
          "manzah 1","manzah 2","manzah 3","manzah 4","manzah 5","manzah 6",
          "menzah 1","menzah 2","menzah 3","menzah 4","menzah 5","menzah 6",
          "menzah 7","menzah","manzah","ennasr"],
    "H": ["ariana centre","ariana beb lahdid","ariana superieur",
          "ariana nouvelle","ariana","cite la gazelle","cite etaamir","cite ettaamir"],
    "I": ["borj louzir","cite sahha","hedi nouira","cite ennasim","nkhilette",
          "ariana soghra","ghazela","nour jaafer","jaafer","raoued","cite chaker",
          "borj touil","kalaat landalos","kalaat el andalous","slimene","souk el khorda","behi ladram"],
    "J": ["gammarth","marsa","sidi bou said","bhar lazreg","soukra parc",
          "soukra","sidi salah","chotrana","diar soukra"],
    "K": ["la goulette","kram","le kram","carthage","jardin de carthage",
          "ain zaghouen","aouina","dar fadhal",
          "lac 1","lac 2","les berges du lac","berges du lac","sidi daoud"],
}

ZONE_MATCH_RULES: dict = {
    "A": ["B","E"],
    "B": ["A","E"],
    "D": ["E"],
    "E": ["A","B","D"],
    "F": ["G","H"],
    "G": ["F","H"],
    "H": ["F","G","I"],
    "I": ["H"],
    "J": ["K"],
    "K": ["J"],
}

def _build_keyword_map():
    pairs = [(z, kw.lower()) for z, kws in ZONES.items() for kw in kws]
    return sorted(pairs, key=lambda x: len(x[1]), reverse=True)

_KEYWORD_MAP = _build_keyword_map()

# ── Address normalization ─────────────────────────────────────────────────────
# Maps spelling variants → canonical key for Step 1 same-destination grouping.
# Only the KEY matters for grouping; the original address is kept for display.
ADDRESS_ALIASES: dict = {
    "thadhamen":        "tadhamen",
    "ettadhamen":       "tadhamen",
    "tadhamon":         "tadhamen",
    "18 janvier ettadhamen": "tadhamen",
    "mannouba":         "manouba",
    "sidi hssine":      "sidi hsin",
    "sidi hassin":      "sidi hsin",
    "ibnou sina":       "ibn sina",
    "madina jedida":    "medina jedida",
    "madina jadida":    "medina jedida",
    "le kram":          "kram",
    "denden":           "danden",
    "hammam-lif":       "hammam lif",
    "hammam chatt":     "hammam chat",
    "cite lkhadhra":    "cite khadhra",
    "lkhadhra":         "cite khadhra",
    "el khadhra":       "cite khadhra",
    "cite el khadhra":  "cite khadhra",
    "berges du lac":    "les berges du lac",
    # Section = neighbourhood in Manouba area → Zone F
    "section":          "manouba",
    "behi ladram":      "borj louzir",
}

def _normalize_addr(address: str) -> str:
    """Return canonical address key for Step 1 grouping (case-insensitive)."""
    import unicodedata
    raw = str(address).strip().lower()
    # Strip accents for alias lookup
    plain = ''.join(c for c in unicodedata.normalize('NFD', raw) if unicodedata.category(c) != 'Mn')
    return ADDRESS_ALIASES.get(plain, ADDRESS_ALIASES.get(raw, raw))


def detect_zone(address: str) -> str:
    import unicodedata
    if not address or str(address).strip() == "": return "UNASSIGNED"
    raw = str(address).strip().lower()
    # Strip accents
    plain = ''.join(c for c in unicodedata.normalize('NFD', raw) if unicodedata.category(c) != 'Mn')
    # Check alias first
    normalized = _normalize_addr(address)
    if normalized not in (raw, plain):
        zone = _detect_zone_raw(normalized)
        if zone != "UNASSIGNED":
            return zone
    # Try accent-stripped version
    zone = _detect_zone_raw(plain)
    if zone != "UNASSIGNED":
        return zone
    return _detect_zone_raw(raw)

def _detect_zone_raw(lower: str) -> str:
    for zone, kw in _KEYWORD_MAP:
        if kw in lower: return zone
    return "UNASSIGNED"

def extract_area_label(address: str) -> str:
    if not address or str(address).strip() == "": return "Unassigned"
    lower = str(address).lower()
    for _, kw in _KEYWORD_MAP:
        if kw in lower: return kw.title()
    return "Unassigned"

def detect_corridor(zone: str) -> str:
    m = {"A":"SOUTH (A-B-E)","B":"SOUTH (A-B-E)","D":"SOUTH (D-E)",
         "E":"SOUTH (A-B-D-E)","F":"BARDO (F-G-H)","G":"BARDO (F-G-H)",
         "H":"ARIANA (F-H-I)","I":"ARIANA (H-I)",
         "J":"COASTAL (J-K)","K":"COASTAL (J-K)"}
    return m.get(zone, "UNASSIGNED")

def detect_micro_zone(address: str) -> str:
    return detect_zone(address)

def _chunks(lst, size):
    return [lst[i:i+size] for i in range(0, len(lst), size)]


def assign_taxi_groups(df: pd.DataFrame, max_per_taxi: int = MAX_PER_TAXI) -> pd.DataFrame:
    """
    STEP 1 — Same destination (normalized, case-insensitive, GLOBAL):
              All agents with the same normalized address are grouped first,
              regardless of zone. Full taxis of 4 are closed immediately.
              Remainders (1-3) stay as atomic blocks, assigned to their zone pool.

    STEP 2 — Same zone, bin-pack blocks:
              Atomic blocks from the same zone are packed together (FFD).
              Full bins → assigned as taxis. Partial bins → Step 3.

    STEP 3 — Zone-match merging (direct partners only, no transitivity):
              Leftover agents from directly-matched zones pooled and packed.
    """
    df = df.copy()
    df["Taxi_Group"] = ""
    counter = 1
    zone_order = list(ZONES.keys()) + ["UNASSIGNED"]

    # Pre-compute normalized address key for every row
    norm_keys = {idx: _normalize_addr(df.at[idx, "Adresse"]) for idx in df.index}

    # ── STEP 1: GLOBAL same-destination grouping ─────────────────────────────
    # Build global address → [indices] map
    global_addr_map: dict[str, list] = {}
    for idx in df.index:
        global_addr_map.setdefault(norm_keys[idx], []).append(idx)

    # zone → list of atomic blocks remaining after Step 1
    zone_blocks: dict[str, list] = {}
    assigned_step1: set = set()

    for key, idxs in global_addr_map.items():
        # Close full taxis of 4 immediately
        while len(idxs) >= max_per_taxi:
            chunk = idxs[:max_per_taxi]
            df.loc[chunk, "Taxi_Group"] = f"Taxi_{counter}"
            assigned_step1.update(chunk)
            counter += 1
            idxs = idxs[max_per_taxi:]
        # Remainder → atomic block placed in its zone pool
        if idxs:
            zone = df.at[idxs[0], "Zone"]
            zone_blocks.setdefault(zone, []).append(idxs)

    # ── STEP 2: bin-pack blocks within the same zone (FFD) ───────────────────
    # zone_leftover_blocks: zone → list of atomic blocks (each block = list of indices)
    zone_leftover_blocks: dict[str, list] = {}

    for zone_name in zone_order:
        blocks = zone_blocks.get(zone_name, [])
        if not blocks:
            continue

        blocks.sort(key=len, reverse=True)
        bins:      list = []   # list of lists of indices
        bin_sizes: list = []
        for block in blocks:
            placed = False
            for i in range(len(bins)):
                if bin_sizes[i] + len(block) <= max_per_taxi:
                    bins[i].extend(block)
                    bin_sizes[i] += len(block)
                    placed = True
                    break
            if not placed:
                bins.append(list(block))
                bin_sizes.append(len(block))

        for bin_idx, bin_indices in enumerate(bins):
            if bin_sizes[bin_idx] == max_per_taxi:
                df.loc[bin_indices, "Taxi_Group"] = f"Taxi_{counter}"
                counter += 1
            else:
                # Keep as atomic block for Step 3 — do NOT flatten
                zone_leftover_blocks.setdefault(zone_name, []).append(bin_indices)

    # ── STEP 3: zone-cluster merging ────────────────────────────────────────────
    # Find groups of mutually-compatible zones (cliques in the match graph),
    # pool all their leftover agents, sort by zone+address, then chunk into taxis.
    # Same-destination integrity is already preserved: Step 1 ensured same-dest
    # agents are in the same Step-2 bin, so they appear consecutively when sorted.

    def zones_compatible(zset):
        zlist = list(zset)
        return all(
            z1 == z2 or z2 in ZONE_MATCH_RULES.get(z1, [])
            for z1 in zlist for z2 in zlist
        )

    # Build zone clusters: merge connected allowed-partner zones into pools
    remaining_zones = set(zone_leftover_blocks.keys())
    processed: set = set()
    clusters: list = []   # each cluster = list of zone names

    for zone in zone_order:
        if zone not in remaining_zones or zone in processed:
            continue
        # Expand cluster: zone + all direct partners that are also in remaining_zones
        cluster = {zone}
        for partner in ZONE_MATCH_RULES.get(zone, []):
            if partner in remaining_zones:
                candidate = cluster | {partner}
                if zones_compatible(candidate):
                    cluster.add(partner)
        clusters.append(sorted(cluster, key=lambda z: zone_order.index(z) if z in zone_order else 99))
        processed.update(cluster)

    # For each cluster, pool all agents, sort zone-first then by address, chunk into taxis
    for cluster in clusters:
        pool = [
            idx
            for zone in cluster
            for block in zone_leftover_blocks.get(zone, [])
            for idx in block
        ]
        if not pool:
            continue
        # Sort: by zone order first (keeps same-zone agents together), then by address
        pool.sort(key=lambda i: (
            zone_order.index(df.at[i, "Zone"]) if df.at[i, "Zone"] in zone_order else 99,
            norm_keys[i]
        ))
        for chunk in _chunks(pool, max_per_taxi):
            df.loc[chunk, "Taxi_Group"] = f"Taxi_{counter}"
            counter += 1

    return df
