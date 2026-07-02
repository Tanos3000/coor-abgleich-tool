"""
Abgleich-Engine: Vergleicht die C16-Buchungsexport-Datei (Coor) mit der
eigenen Eingangsrechnungen-Excel und markiert Abweichungen gelb.

Annahmen (siehe README.md fuer Details):
- C16-Datei: flache Tabelle, Spalten u.a. 'Belegnr.', 'Belegnr.2', 'Istbetrag'.
- Eigene Datei: verschachtelte Bloecke. Kopfzeile pro Rechnung (Spalte B =
  Gewerke-Nr gefuellt) mit 'Ext. Re. Nr 1' (Spalte F) und 'Ext. Re. Nr 2'
  (Spalte I). Darunter eine Unterueberschrift ('Nummer' in Spalte C) und
  danach 1..n Zahlungszeilen mit 'Zahlung Netto' (Spalte F) und
  'Zahlung Brutto' (Spalte G).
- Verknuepfung: Ext. Re. Nr 1/2 <-> Belegnr./Belegnr.2 (exakt, sowie ueber
  den Praefix vor dem ersten '/', da Ratenzahlungen in der C16 oft als
  'BASISNR/laufendenr' auftauchen).
- Abgleich je einzelner Zahlungszeile (nicht Summenebene): jeder
  Zahlung-Brutto (bzw. ersatzweise Netto) -Betrag muss unter dem
  zugehoerigen Beleg-Key in der C16 exakt (Toleranz 0.02) auftauchen.
- Bei Abweichung wird NUR die Betrag-Zelle (Netto/Brutto) der betroffenen
  Zahlungszeile gelb markiert.
- Automatische Korrektur erfolgt nur im eindeutigen Fall: genau eine
  Zahlungszeile im Rechnungsblock UND genau ein noch nicht verbrauchter
  C16-Betrag unter demselben Key. In allen anderen Faellen (mehrere
  Zahlungen, kein Key gefunden, mehrdeutig) wird NICHT automatisch
  korrigiert - die Zeile bleibt gelb markiert und im Log als
  "manuelle Pruefung noetig" aufgefuehrt.
"""

from dataclasses import dataclass, field

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
TOLERANCE = 0.02  # Euro

# Obergrenze fuer automatische Korrekturen: Ein eindeutiger Fall (1 offene
# Zahlung + 1 freier C16-Betrag) wird nur automatisch korrigiert, wenn die
# Differenz zum urspruenglichen Betrag plausibel klein ist (Tippfehler,
# Rundung). Bei grossen Abweichungen (z.B. 6970 -> 300) handelt es sich
# vermutlich um zwei unterschiedliche Zahlungen, die zufaellig unter
# demselben Beleg-Key stehen - das darf NICHT automatisch verbucht werden,
# sondern muss als Abweichung gelb markiert und manuell geprueft werden.
AUTO_KORREKTUR_MAX_ABS_DIFF = 50.0     # Euro
AUTO_KORREKTUR_MAX_REL_DIFF = 0.20     # 20 %


def norm_key(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def prefix_key(v):
    """Praefix vor dem ersten '/' - fasst Ratenzahlungen wie
    'U18260100008/26-04' zu 'U18260100008' zusammen."""
    if v is None:
        return None
    s = str(v).strip()
    if "/" in s:
        return s.split("/")[0]
    return s


@dataclass
class C16Row:
    row: int
    belegnr: str
    belegnr2: str
    betrag: float
    kreditor: str = None


@dataclass
class PaymentLine:
    row: int
    nummer: str
    netto: float
    brutto: float
    status: str = "OFFEN"          # OK / ABWEICHUNG / KEIN_BELEG / KORRIGIERT
    matched_betrag: float = None
    note: str = ""


@dataclass
class InvoiceBlock:
    header_row: int
    auftrag: str
    beschreibung: str
    ext1: str
    ext2: str
    payments: list = field(default_factory=list)


def load_c16(path):
    """Laedt die C16-Datei. Sucht das Sheet, das eine Spalte 'Belegnr.'
    hat (Sheetname enthaelt sonst z.B. das Objektkuerzel, ist also nicht
    fix)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    header = None
    for sheet in wb.worksheets:
        hdr = [c.value for c in sheet[1]]
        if "Belegnr." in hdr and "Istbetrag" in hdr:
            ws = sheet
            header = hdr
            break
    if ws is None:
        raise ValueError(
            "In der C16-Datei wurde kein Sheet mit den Spalten "
            "'Belegnr.' und 'Istbetrag' gefunden. Ist das wirklich die "
            "C16-Exportdatei?"
        )
    col = {h: i + 1 for i, h in enumerate(header) if h}
    rows = []
    for r in range(2, ws.max_row + 1):
        belegnr = norm_key(ws.cell(r, col["Belegnr."]).value)
        belegnr2 = norm_key(ws.cell(r, col["Belegnr.2"]).value) if "Belegnr.2" in col else None
        betrag = ws.cell(r, col["Istbetrag"]).value
        kreditor = ws.cell(r, col["Debitor/Kreditor-Name"]).value if "Debitor/Kreditor-Name" in col else None
        if belegnr is None and belegnr2 is None and betrag is None:
            continue
        rows.append(C16Row(row=r, belegnr=belegnr, belegnr2=belegnr2, betrag=betrag, kreditor=kreditor))
    return rows


def build_indexes(c16_rows):
    exact_idx = {}
    prefix_idx = {}
    for row in c16_rows:
        for k in (row.belegnr, row.belegnr2):
            if not k:
                continue
            exact_idx.setdefault(k, []).append(row)
            pk = prefix_key(k)
            prefix_idx.setdefault(pk, []).append(row)
    return exact_idx, prefix_idx


def _find_own_sheet(wb):
    for sheet in wb.worksheets:
        hdr = [c.value for c in sheet[1]]
        if "Ext. Re. Nr 1" in hdr:
            return sheet
    return wb.worksheets[0]


def load_own_blocks(path):
    """Parst die verschachtelte Struktur der eigenen Datei in
    InvoiceBlock-Objekte mit ihren PaymentLine-Kindern."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _find_own_sheet(wb)
    blocks = []
    current = None
    r = 2
    max_row = ws.max_row
    while r <= max_row:
        col_b = ws.cell(r, 2).value  # Gewerke-Nr
        col_c = ws.cell(r, 3).value  # Auftragsnummer / 'Nummer' / Zahlungs-Nr
        if col_b not in (None, ""):
            current = InvoiceBlock(
                header_row=r,
                auftrag=ws.cell(r, 3).value,
                beschreibung=ws.cell(r, 4).value,
                ext1=norm_key(ws.cell(r, 6).value),
                ext2=norm_key(ws.cell(r, 9).value),
            )
            blocks.append(current)
        elif col_c == "Nummer":
            pass  # Unterueberschrift, keine Daten
        elif col_c not in (None, ""):
            # echte Zahlungszeile: hat immer eine 'Nummer' in Spalte C
            if current is not None:
                netto = ws.cell(r, 6).value
                brutto = ws.cell(r, 7).value
                current.payments.append(
                    PaymentLine(row=r, nummer=col_c, netto=netto, brutto=brutto)
                )
        else:
            # Spalte B und C beide leer: kommt in der Vorlage manchmal vor,
            # wenn die Kopfzeile eines Blocks ueber zwei physische Zeilen
            # verteilt ist (Ext.Re.Nr1/Re.Datum/Re.Eingang landen dann
            # nicht in derselben Zeile wie Gewerke-Nr/Auftragsnummer).
            # In dem Fall die fehlenden Kopf-Werte in den aktuellen Block
            # nachtragen, statt sie faelschlich als Zahlungszeile zu lesen.
            if current is not None and current.ext1 is None:
                val_f = ws.cell(r, 6).value
                if val_f is not None and not isinstance(val_f, (int, float)):
                    current.ext1 = norm_key(val_f)
            if current is not None and current.ext2 is None:
                val_i = ws.cell(r, 9).value
                if val_i is not None and not isinstance(val_i, (int, float)):
                    current.ext2 = norm_key(val_i)
        r += 1
    return wb, ws, blocks


def compare(c16_rows, blocks):
    """Fuehrt den eigentlichen Abgleich durch und setzt status/matched_betrag/
    note auf jeder PaymentLine. Gibt eine Zusammenfassung (dict) zurueck."""
    exact_idx, prefix_idx = build_indexes(c16_rows)

    summary = {
        "bloecke": len(blocks),
        "zahlungen": sum(len(b.payments) for b in blocks),
        "ok": 0,
        "abweichung": 0,
        "kein_beleg": 0,
        "auto_korrigiert": 0,
        "manuelle_pruefung": 0,
    }

    for b in blocks:
        pool = []
        seen = set()
        for k in (b.ext1, b.ext2):
            if not k:
                continue
            for src in (exact_idx.get(k, []), prefix_idx.get(prefix_key(k), [])):
                for c in src:
                    if c.row not in seen:
                        pool.append(c)
                        seen.add(c.row)
        used = [False] * len(pool)

        unmatched_payments = []
        for p in b.payments:
            amt = p.brutto if p.brutto is not None else p.netto
            if amt is None:
                continue
            if not pool:
                p.status = "KEIN_BELEG"
                p.note = "Kein Beleg mit dieser Ext.Re.Nr in der C16 gefunden"
                summary["kein_beleg"] += 1
                unmatched_payments.append(p)
                continue
            found_idx = None
            for i, c in enumerate(pool):
                if used[i]:
                    continue
                if c.betrag is not None and abs(float(c.betrag) - float(amt)) < TOLERANCE:
                    found_idx = i
                    break
            if found_idx is not None:
                used[found_idx] = True
                p.status = "OK"
                p.matched_betrag = pool[found_idx].betrag
                summary["ok"] += 1
            else:
                p.status = "ABWEICHUNG"
                p.note = "Betrag nicht in C16 unter diesem Beleg gefunden"
                summary["abweichung"] += 1
                unmatched_payments.append(p)

        # Eindeutige Auto-Korrektur: genau 1 offene Zahlung + genau 1 freier C16-Betrag,
        # UND die Differenz ist plausibel klein (siehe AUTO_KORREKTUR_MAX_*).
        free = [c for i, c in enumerate(pool) if not used[i]]
        auto_korrigiert = False
        if len(unmatched_payments) == 1 and len(free) == 1 and free[0].betrag is not None:
            p = unmatched_payments[0]
            old = p.brutto if p.brutto is not None else p.netto
            new = free[0].betrag
            diff = abs(float(new) - float(old))
            rel = diff / abs(float(old)) if old else float("inf")
            if diff <= AUTO_KORREKTUR_MAX_ABS_DIFF or rel <= AUTO_KORREKTUR_MAX_REL_DIFF:
                was_kein_beleg = p.status == "KEIN_BELEG"
                p.status = "KORRIGIERT"
                p.matched_betrag = new
                p.note = f"Automatisch korrigiert: {old} -> {new} (C16 Zeile {free[0].row})"
                summary["auto_korrigiert"] += 1
                if was_kein_beleg:
                    summary["kein_beleg"] -= 1
                else:
                    summary["abweichung"] -= 1
                auto_korrigiert = True
            else:
                p.note = (
                    f"{p.note} (1 moeglicher C16-Betrag {new} gefunden, aber Differenz "
                    f"zu {old} ist zu gross fuer automatische Korrektur - bitte manuell pruefen)"
                )
        if not auto_korrigiert:
            for p in unmatched_payments:
                if p.status in ("ABWEICHUNG", "KEIN_BELEG"):
                    summary["manuelle_pruefung"] += 1

    return summary


def write_marked_copy(src_path, out_path, ws_title, blocks):
    """Schreibt eine Kopie der Originaldatei, in der die Betrag-Zellen
    abweichender Zahlungszeilen gelb markiert sind (Original bleibt
    unveraendert erhalten)."""
    wb = openpyxl.load_workbook(src_path)  # ohne data_only, um Formeln/Format zu erhalten
    ws = wb[ws_title]
    for b in blocks:
        for p in b.payments:
            if p.status in ("ABWEICHUNG", "KEIN_BELEG"):
                for col in (6, 7):  # F = Netto, G = Brutto
                    cell = ws.cell(p.row, col)
                    if cell.value is not None:
                        cell.fill = YELLOW_FILL
                        cell.comment = Comment(p.note, "Abgleich-Tool")
    wb.save(out_path)


def write_corrected_copy(src_path, out_path, ws_title, blocks):
    """Schreibt eine Kopie, in der eindeutig korrigierbare Betraege
    ueberschrieben wurden. Alles, was manuelle Pruefung braucht, bleibt
    unveraendert, aber weiterhin gelb markiert."""
    wb = openpyxl.load_workbook(src_path)
    ws = wb[ws_title]
    for b in blocks:
        for p in b.payments:
            if p.status == "KORRIGIERT":
                if p.brutto is not None:
                    ws.cell(p.row, 7).value = p.matched_betrag
                if p.netto is not None and p.brutto is None:
                    ws.cell(p.row, 6).value = p.matched_betrag
                cell = ws.cell(p.row, 7 if p.brutto is not None else 6)
                cell.comment = Comment(p.note, "Abgleich-Tool")
            elif p.status in ("ABWEICHUNG", "KEIN_BELEG"):
                for col in (6, 7):
                    cell = ws.cell(p.row, col)
                    if cell.value is not None:
                        cell.fill = YELLOW_FILL
                        cell.comment = Comment(p.note + " (manuelle Pruefung noetig)", "Abgleich-Tool")
    wb.save(out_path)


def run_abgleich(c16_path, own_path, out_dir):
    """High-level Einstiegspunkt fuer die GUI. Gibt (summary, marked_path,
    corrected_path, log_lines) zurueck."""
    import os

    c16_rows = load_c16(c16_path)
    wb_own, ws_own, blocks = load_own_blocks(own_path)
    summary = compare(c16_rows, blocks)

    base = os.path.splitext(os.path.basename(own_path))[0]
    marked_path = os.path.join(out_dir, f"{base}_markiert.xlsx")
    corrected_path = os.path.join(out_dir, f"{base}_korrigiert.xlsx")

    write_marked_copy(own_path, marked_path, ws_own.title, blocks)
    write_corrected_copy(own_path, corrected_path, ws_own.title, blocks)

    log_lines = [
        f"Rechnungsbloecke geprueft: {summary['bloecke']}",
        f"Zahlungszeilen geprueft: {summary['zahlungen']}",
        f"  OK (uebereinstimmend): {summary['ok']}",
        f"  Automatisch korrigiert: {summary['auto_korrigiert']}",
        f"  Abweichung (Betrag nicht gefunden): {summary['abweichung']}",
        f"  Kein passender Beleg in C16: {summary['kein_beleg']}",
        f"  -> insgesamt gelb markiert / manuelle Pruefung: {summary['abweichung'] + summary['kein_beleg']}",
    ]
    return summary, marked_path, corrected_path, log_lines
