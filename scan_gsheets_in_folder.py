import os, unicodedata, argparse, csv, sys, time
from typing import Dict, List, Set
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

def normalize(s: str) -> str:
    if not s: return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.lower()

def col_number_to_a1(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def get_creds():
    creds = None
    if os.path.exists("token.json"):
        try:
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        except Exception:
            creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w", encoding="utf-8") as token:
            token.write(creds.to_json())
    return creds

def list_spreadsheets(drive):
    files = []
    page_token = None
    while True:
        resp = drive.files().list(
            q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
            fields="nextPageToken, files(id,name)",
            pageSize=1000,
            pageToken=page_token,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files

def scan_spreadsheet(sheets, spreadsheet_id: str, local_paths: List[str], needle_norm: str, hits: List[dict]):
    meta = sheets.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="properties/title,sheets(properties(title,gridProperties(rowCount,columnCount)))"
    ).execute()

    ss_title = meta["properties"]["title"]
    sh_props = [s["properties"] for s in meta.get("sheets", [])]

    ranges = []
    grid_sizes = []
    for p in sh_props:
        title = p["title"]
        rows = max(1, int(p.get("gridProperties", {}).get("rowCount", 1000)))
        cols = max(1, int(p.get("gridProperties", {}).get("columnCount", 26)))
        last_col = col_number_to_a1(cols)
        ranges.append(f"'{title}'!A1:{last_col}{rows}")
        grid_sizes.append((title, rows, cols))

    if not ranges:
        return

    resp = sheets.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id,
        ranges=ranges,
        majorDimension="ROWS"
    ).execute()

    for (title, _rows, _cols), vr in zip(grid_sizes, resp.get("valueRanges", [])):
        values = vr.get("values", [])
        for r_i, row in enumerate(values, start=1):
            for c_i, val in enumerate(row, start=1):
                if val is None: 
                    continue
                if needle_norm in normalize(str(val)):
                    addr = f"{col_number_to_a1(c_i)}{r_i}"
                    hits.append({
                        "LocalPaths": " | ".join(local_paths),
                        "SpreadsheetTitle": ss_title,
                        "SpreadsheetId": spreadsheet_id,
                        "Sheet": title,
                        "Cell": addr,
                        "Value": str(val)
                    })

def main():
    ap = argparse.ArgumentParser(description="Scan Google Sheets par correspondance de nom (sans lire les fichiers .gsheet).")
    ap.add_argument("-root", required=True, help="Dossier racine (.gsheet) à parcourir.")
    ap.add_argument("-name", required=True, help="Nom à chercher, ex: \"Flauger Stéphane\".")
    ap.add_argument("-out", default="scan_gsheets_hits.csv", help="CSV de sortie.")
    args = ap.parse_args()

    needle_norm = normalize(args.name)

    # 1) Récupérer tous les .gsheet et construire l'ensemble des basenames locaux
    name_to_paths: Dict[str, List[str]] = {}
    for root, _dirs, files in os.walk(args.root):
        for fn in files:
            if fn.lower().endswith(".gsheet"):
                full = os.path.join(root, fn)
                base = os.path.splitext(os.path.basename(full))[0]
                name_to_paths.setdefault(base, []).append(full)

    basenames: Set[str] = set(name_to_paths.keys())
    if not basenames:
        print(f"Aucun fichier .gsheet trouvé sous: {args.root}")
        sys.exit(0)

    print(f"Local .gsheet trouvés: {sum(len(v) for v in name_to_paths.values())}  —  Noms uniques: {len(basenames)}")

    # 2) Résoudre via Drive par NOM de spreadsheet
    creds = get_creds()
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    remote_files = list_spreadsheets(drive)
    name_to_ids: Dict[str, List[str]] = {}
    for f in remote_files:
        nm = f["name"]
        if nm in basenames:
            name_to_ids.setdefault(nm, []).append(f["id"])

    # 3) Associer IDs <-> chemins locaux
    id_to_paths: Dict[str, List[str]] = {}
    unresolved = []
    for nm, paths in name_to_paths.items():
        ids = name_to_ids.get(nm, [])
        if not ids:
            unresolved.append(nm)
        else:
            for sid in ids:
                id_to_paths.setdefault(sid, []).extend(paths)

    if unresolved:
        print("ATTENTION — noms locaux sans équivalent Drive (aucun spreadsheet de ce nom) :")
        for nm in sorted(unresolved):
            print(" -", nm)

    if not id_to_paths:
        print("Aucun ID résolu par nom → rien à scanner.")
        sys.exit(0)

    # 4) Scan des cellules
    hits: List[dict] = []
    scanned = 0
    for sid, paths in id_to_paths.items():
        try:
            scan_spreadsheet(sheets, sid, paths, needle_norm, hits)
        except Exception as e:
            print(f"[WARN] Impossible de scanner {sid} — {e}")
        scanned += 1
        time.sleep(0.1)

    # 5) Sortie
    print("\n=== RÉSULTATS ===")
    for h in hits:
        print(f"- {h['SpreadsheetTitle']} [{h['SpreadsheetId']}] | {h['Sheet']}!{h['Cell']} → {h['Value']}")
        print(f"  Local: {h['LocalPaths']}")
    print(f"\nTotal Google Sheets scannés: {scanned} — Occurrences trouvées: {len(hits)}")

    out_path = os.path.abspath(args.out)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["SpreadsheetTitle","SpreadsheetId","Sheet","Cell","Value","LocalPaths"])
        w.writeheader()
        for h in hits:
            w.writerow(h)
    print(f"CSV écrit: {out_path}")

if __name__ == "__main__":
    main()
