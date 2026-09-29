#!/usr/bin/env python3
"""
Parser Autonor.xlsx fra Autologik og producerer cars.json i det format,
som Davids dashboard forventer.

Input : Autonor.xlsx (modtaget som mail-vedhæftning hver nat kl. 23:22)
Output: cars.json    (commit'es til repo'et autonor-cars-data)

Formathistorik fra Autologik:
 - t.o.m. 24/8 2026: header i række 3, Stelnr. i kolonne A, datoer som datoceller.
 - fra 24/8 2026:    header i række 1, Stelnr i kolonne B, datoer som tekst "DD-MM-YYYY".
 - fra 9/9 2026:     ny kolonne "Reference" i B → Stelnr rykket til C.

Parseren er derfor HEADER-STYRET: den finder header-rækken ved at lede efter
"Stelnr" i de første rækker og slår derefter kolonnerne op på navn. Så tåler
den at Autologik flytter/tilføjer kolonner igen.

Dedupliceringslogik:
 - Hver bil identificeres entydigt af sit stelnummer (VIN).
 - Hvis samme VIN optræder flere gange i samme fil, beholdes den første forekomst.
 - Biler uden gyldigt VIN (tom eller < 10 tegn) ignoreres.

Exit-koder:
 0 = OK
 2 = forkert brug
 3 = header-række (Stelnr) ikke fundet → formatet er ændret igen, workflow'et fejler synligt
"""
from __future__ import annotations
import json
import re
import sys
from datetime import datetime, date
from pathlib import Path
from openpyxl import load_workbook


# Kolonne-opslag: JSON-felt → kandidat-headere (lowercase, "starter med"-match).
# Første header i listen der matcher, vinder.
COLUMN_ALIASES: dict[str, list[str]] = {
    "vin":           ["stelnr"],
    "model":         ["model"],
    "maerke":        ["mærke", "maerke"],
    "regnr":         ["regnr", "reg.nr", "reg nr"],
    "bookingnr":     ["bookingnr", "booking"],
    "fra":           ["navn, afhentningsadresse", "afhentningsadresse", "afhentning", "fra"],
    "til":           ["navn, leveringsadresse", "leveringsadresse", "til"],
    "klardato":      ["klardato"],
    "koerselsdato":  ["kørselsdato", "koerselsdato"],
    "leveringsdato": ["forventet leveringsdato", "forventet levering", "leveringsdato"],
    "track":         ["link til t&t", "track and trace", "track", "t&t"],
}

HEADER_SEARCH_ROWS = 15   # hvor mange rækker vi leder efter header-rækken i


def _norm(v) -> str:
    """Normaliser en celleværdi til sammenligning: str, lowercase, ét mellemrum."""
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip().lower()


def find_header_row(ws) -> tuple[int, dict[str, int]] | None:
    """Find header-rækken (den der indeholder 'Stelnr') og returnér
    (1-baseret rækkenr, {json-felt: 0-baseret kolonneindeks})."""
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, values_only=True), 1):
        cells = [_norm(v) for v in row]
        if not any(c.startswith("stelnr") for c in cells):
            continue

        mapping: dict[str, int] = {}
        for field, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                hit = next((i for i, c in enumerate(cells) if c and c.startswith(alias)), None)
                if hit is not None and hit not in mapping.values():
                    mapping[field] = hit
                    break
        return row_idx, mapping
    return None


def to_iso_date(v) -> str | None:
    """Datocelle eller tekst ('DD-MM-YYYY', 'DD.MM.YYYY', 'DD/MM/YYYY', 'YYYY-MM-DD') → 'YYYY-MM-DD'."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r"^(\d{1,2})[-./](\d{1,2})[-./](\d{4})", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        try:
            return date(y, mo, d).isoformat()
        except ValueError:
            return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return m.group(0)
    return None


def first_line(v) -> str:
    """'Autonor ApS\\nEnergivej 23' → 'Autonor ApS' (kun navnet, ikke adressen)."""
    if v is None:
        return ""
    return str(v).replace("\r", "\n").split("\n")[0].strip()


def parse_excel(xlsx_path: Path) -> list[dict]:
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    found = find_header_row(ws)
    if found is None:
        raise RuntimeError(
            f"Kunne ikke finde header-rækken (ingen kolonne der starter med 'Stelnr') "
            f"i de første {HEADER_SEARCH_ROWS} rækker af ark '{ws.title}'. "
            f"Autologik har formentlig ændret formatet igen."
        )
    header_row, col = found
    missing = [f for f in ("vin", "model", "leveringsdato") if f not in col]
    if missing:
        raise RuntimeError(f"Header-række {header_row} fundet, men mangler kolonner: {missing}. Fundet: {col}")

    print(f"Header i række {header_row} · kolonner: "
          + ", ".join(f"{k}={chr(65 + i)}" for k, i in sorted(col.items(), key=lambda kv: kv[1])),
          file=sys.stderr)

    def get(row, field):
        i = col.get(field)
        return row[i] if i is not None and i < len(row) else None

    cars: list[dict] = []
    seen: set[str] = set()

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        vin_raw = get(row, "vin")
        if not isinstance(vin_raw, str):
            continue
        vin = vin_raw.strip().upper()
        if len(vin) < 10:
            continue
        if vin in seen:
            continue
        seen.add(vin)

        # Autologik skriver fx "VW-P" / "MERCEDES-P" (P = personbil, V = varebil) — suffikset fjernes
        maerke = re.sub(r"-[PV]$", "", first_line(get(row, "maerke")).strip(), flags=re.IGNORECASE)
        model = first_line(get(row, "model"))
        # Dashboardet viser kun `model` — så vi lægger mærket foran når vi har det
        # ("VOLVO XC40" i stedet for bare "XC40"). Gamle filer havde ikke mærke.
        model_vist = f"{maerke} {model}".strip() if maerke and not model.upper().startswith(maerke.upper()) else model

        cars.append(
            {
                "vin": vin,
                "model": model_vist,
                "maerke": maerke or None,
                "regnr": first_line(get(row, "regnr")) or None,
                "bookingnr": first_line(get(row, "bookingnr")) or None,
                "fra": first_line(get(row, "fra")),
                "til": first_line(get(row, "til")),
                "klardato": to_iso_date(get(row, "klardato")),
                "koerselsdato": to_iso_date(get(row, "koerselsdato")),
                "leveringsdato": to_iso_date(get(row, "leveringsdato")),
                "track": first_line(get(row, "track")) or None,
            }
        )

    return cars


def build_payload(cars: list[dict]) -> dict:
    return {
        "lastUpdated": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source": "autologik-mail",
        "count": len(cars),
        "cars": cars,
    }


def main() -> int:
    if len(sys.argv) < 2:
        print("Brug: parse_autonor.py <input.xlsx> [output.json]", file=sys.stderr)
        return 2

    xlsx = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("cars.json")

    try:
        cars = parse_excel(xlsx)
    except RuntimeError as e:
        print(f"FEJL: {e}", file=sys.stderr)
        return 3

    payload = build_payload(cars)
    out.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"OK — {len(cars)} unikke biler skrevet til {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
