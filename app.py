"""
Audit-Assistent Qualitaetsmanagement (Webversion)
Oberflaeche: Streamlit Community Cloud
Sprachmodell: OpenAI-kompatible Schnittstelle, voreingestellt xAI (Grok)

Drei Ebenen: Auditprogramm, Einzelaudit in fuenf Schritten, Massnahmenverfolgung.
Dialektisches Verfahren: Konstruktionsteam baut den Konformitaetsnachweis,
Falsifikationsteam greift ihn an, der Auditor entscheidet als Richter.
"""

import io
import json
import re
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
import streamlit as st
from docx import Document
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Einstellungen aus den Secrets
# ---------------------------------------------------------------------------
APP_NAME = st.secrets.get("APP_NAME", "Audit-Assistent Qualitätsmanagement")
ORGANISATION = st.secrets.get("ORGANISATION", "")
API_BASE = str(st.secrets.get("API_BASE", "https://api.x.ai/v1")).rstrip("/")
API_KEY = str(st.secrets.get("API_KEY", "")).strip()
PROVIDER = st.secrets.get("PROVIDER_NAME", "xAI (Grok)")
MODEL_PRO = st.secrets.get("MODEL_PRO", "grok-4.3")
MODEL_CONTRA = st.secrets.get("MODEL_CONTRA", "grok-4.20")
TIMEOUT = float(st.secrets.get("TIMEOUT", 600))
MAX_TOKENS = int(st.secrets.get("MAX_TOKENS", 3000))
STREAM = bool(st.secrets.get("STREAM", True))

MAX_ZEICHEN = int(st.secrets.get("MAX_ZEICHEN", 20000))
NORMEN_ORDNER = Path(__file__).parent / "normen"
LOGO = Path(__file__).parent / "logo.png"

st.set_page_config(page_title=APP_NAME, layout="wide")
if LOGO.exists():
    st.logo(str(LOGO), size="large")

HINWEISE = f"""
- Diese Anwendung unterstützt die Urteilsbildung, sie trifft keine Auditfeststellungen.
  Einstufung und Schlussfolgerung verantwortet immer der Auditor.
- Ihre hochgeladenen Unterlagen werden zur Auswertung an den Dienst {PROVIDER} übermittelt.
  Laden Sie deshalb keine vertraulichen Originalunterlagen und keine Personendaten hoch,
  solange das mit Ihrer IT und dem Datenschutz nicht geklärt ist.
- Die KI kann sich irren und Nachweise falsch zuordnen. Jedes Zitat wird gegen die Quelle
  geprüft und entsprechend gekennzeichnet. Prüfen Sie Feststellungen trotzdem selbst.
- Ihr Arbeitsstand liegt nur in dieser Sitzung. Sichern Sie ihn links über
  "Arbeitsstand sichern", sonst geht er beim Schliessen verloren.
"""

# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------
SPRACHREGEL = ("\n\nSPRACHE: Schreibe alle Inhalte ausschliesslich auf Deutsch in Schweizer "
               "Rechtschreibung (ss statt ß). Zitate aus Nachweisen übernimmst du wörtlich "
               "in ihrer Originalsprache.")
NORMREGEL = ("\n\nNORMBEZUG: Beziehe dich nur auf die im Auditkriterium genannte Norm und "
             "Abschnittsnummer. Erfinde keine weiteren Normabschnitte oder Normzitate.")
JSONREGEL = "\n\nAUSGABE: Antworte ausschliesslich mit gültigem JSON, ohne Vor- und Nachtext."

PRO_SYSTEM = """Du bist das KONSTRUKTIONSTEAM eines internen Audits.
Baue einen Konformitätsnachweis als Assurance Case (Aussage, Argument, Nachweis).
Ordne jede Teilaussage einem Prüfpunkt des Auditkriteriums zu.
Regeln
1. Nutze ausschliesslich die bereitgestellten Nachweise (Dokumente, Interviewnotizen,
   Beobachtungen, Leistungsdaten).
2. Jeder Nachweis enthält den exakten Quellennamen und ein WÖRTLICHES Zitat (höchstens 40 Wörter).
3. Prüfe nicht nur, OB etwas vorliegt, sondern ob der Inhalt die Aussage trägt.
4. Was nicht belegt ist, trägst du als Annahme ein. Verstecke keine Lücken.
Schema
{"hauptaussage": "...", "teilaussagen": [{"id": "T1", "pruefpunkt": "...", "aussage": "...",
 "argument": "...", "nachweise": [{"dokument": "...", "zitat": "..."}], "annahmen": ["..."]}]}"""

CONTRA_SYSTEM = """Du bist das FALSIFIKATIONSTEAM eines internen Audits.
Du erhältst einen Assurance Case und die Nachweise. Dein Ziel ist, ihn zu widerlegen.
Einwandtypen
"widerlegend": Nachweise belegen das Gegenteil der Aussage.
"untergrabend": Der Nachweis ist unzuverlässig, veraltet, nicht freigegeben oder unvollständig.
"unterlaufend": Der Nachweis stimmt, stützt die Aussage aber inhaltlich nicht.
"ungestuetzte_annahme": Eine Annahme trägt die Aussage, ohne belegt zu sein.
Achte besonders auf Widersprüche zwischen Vorgabe, Aufzeichnung, Aussagen aus Interviews
und Beobachtungen.
Regeln
1. Erfinde nichts. Zitate müssen wörtlich aus den Nachweisen stammen.
2. Fehlt ein Nachweis ganz, lass die Liste "nachweise" leer und begründe das.
3. Schlage für jeden Einwand eine konkrete Prüfung vor Ort vor (Interview, Stichprobe, Beobachtung).
Schema
{"einwaende": [{"id": "E1", "ziel": "T1", "typ": "...", "begruendung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}], "schwere": "hoch|mittel|gering",
 "pruefung_vor_ort": "..."}]}"""

ERWIDERUNG_SYSTEM = """Du bist das KONSTRUKTIONSTEAM. Erwidere auf jeden Einwand sachlich.
Gestehe Einwände zu, wenn die Nachweise sie stützen. Nutze nur wörtliche Zitate.
Schema
{"erwiderungen": [{"einwand": "E1", "zugestanden": true, "erwiderung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}]}]}"""

BERICHT_SYSTEM = """Du formulierst Auditfeststellungen nach ISO 19011 (Anforderung mit
Normabschnitt, objektiver Nachweis, Feststellung). Die Einstufung hat der menschliche Auditor
bereits getroffen, du darfst sie NICHT ändern. Formuliere sachlich, prüfbar, ohne Schuldzuweisung.
Schema
{"feststellungen": [{"einwand": "E1", "einstufung": "...", "abschnitt": "...", "anforderung": "...",
 "objektiver_nachweis": "...", "feststellung": "..."}]}"""

URTEILE = ["Ausgeräumt", "Bestätigt: Abweichung", "Bestätigt: Verbesserungspotenzial",
           "Offen: vor Ort prüfen"]
TYPEN = {"widerlegend": "widerlegend", "untergrabend": "untergrabend",
         "unterlaufend": "unterlaufend", "ungestuetzte_annahme": "ungestützte Annahme"}
RISIKO = ["hoch", "mittel", "gering"]
STATUS = ["geplant", "in Arbeit", "abgeschlossen"]
MSTATUS = ["offen", "in Umsetzung", "umgesetzt", "wirksam bestätigt"]


# ---------------------------------------------------------------------------
# Passwortschutz
# ---------------------------------------------------------------------------
def anmeldung() -> bool:
    if st.session_state.get("auth_ok"):
        return True
    st.title(APP_NAME)
    st.write("Bitte Zugangspasswort eingeben.")
    with st.expander("Hinweise zur Nutzung", expanded=True):
        st.markdown(HINWEISE)
    pw = st.text_input("Passwort", type="password")
    if st.button("Anmelden"):
        if pw and pw == st.secrets.get("APP_PASSWORD", ""):
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("Das Passwort stimmt nicht.")
    return False


if not anmeldung():
    st.stop()

if "audits" not in st.session_state:
    st.session_state.audits = []


# ---------------------------------------------------------------------------
# Normkataloge und Dateien
# ---------------------------------------------------------------------------
@st.cache_data
def normkataloge_laden() -> dict:
    kataloge = {}
    if NORMEN_ORDNER.exists():
        for datei in sorted(NORMEN_ORDNER.glob("*.json")):
            try:
                k = json.loads(datei.read_text(encoding="utf-8"))
                kataloge[k.get("norm", datei.stem)] = k
            except (json.JSONDecodeError, OSError):
                pass
    return kataloge


def datei_lesen(hochgeladen) -> str:
    daten = hochgeladen.getvalue()
    name = hochgeladen.name.lower()
    if name.endswith(".pdf"):
        return "\n".join((s.extract_text() or "") for s in PdfReader(io.BytesIO(daten)).pages)
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(daten))
        teile = [p.text for p in doc.paragraphs]
        for tab in doc.tables:
            for zeile in tab.rows:
                teile.append(" | ".join(z.text for z in zeile.cells))
        return "\n".join(teile)
    return daten.decode("utf-8", errors="ignore")


# ---------------------------------------------------------------------------
# Sprachmodell
# ---------------------------------------------------------------------------
class ModellFehler(Exception):
    pass


def json_aus_text(text: str):
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    anfang, ende = text.find("{"), text.rfind("}")
    if anfang >= 0 and ende > anfang:
        try:
            return json.loads(text[anfang:ende + 1])
        except json.JSONDecodeError:
            pass
    return {"fehler": "Keine gültige JSON-Antwort", "rohtext": text[:2000]}


def llm_json(modell: str, system: str, nutzer: str, audit: dict):
    """Anfrage an die OpenAI-kompatible Schnittstelle, bei Bedarf als Datenstrom.

    Streaming hält die Verbindung offen. Ohne das brechen langsame Gateways
    die Anfrage mit einem Fehler 504 ab, bevor das Modell fertig ist.
    """
    system = system + NORMREGEL + SPRACHREGEL + JSONREGEL
    nutzlast = {"model": modell, "temperature": 0.2, "max_tokens": MAX_TOKENS,
                "stream": STREAM, "response_format": {"type": "json_object"},
                "messages": [{"role": "system", "content": system},
                             {"role": "user", "content": nutzer}]}
    kopf = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    def senden(mit_format: bool):
        daten = dict(nutzlast)
        if not mit_format:
            daten.pop("response_format", None)
        return requests.post(f"{API_BASE}/chat/completions", headers=kopf, json=daten,
                             timeout=(15, TIMEOUT), stream=STREAM)

    try:
        r = senden(True)
        if r.status_code == 400:
            r.close()
            r = senden(False)   # Modell kennt das erzwungene JSON-Format nicht
    except requests.exceptions.Timeout:
        raise ModellFehler("Zeitüberschreitung. Weniger Nachweise je Lauf wählen oder "
                           "MAX_ZEICHEN in den Secrets verkleinern.")
    except requests.RequestException as f:
        raise ModellFehler(f"Verbindung fehlgeschlagen: {f}")

    if r.status_code != 200:
        text = r.text[:400]
        if r.status_code in (502, 503, 504):
            raise ModellFehler(
                f"Der Dienst hat die Anfrage abgebrochen (HTTP {r.status_code}). "
                "Das passiert bei zu langen Anfragen. Laden Sie weniger Nachweise, "
                "oder verkleinern Sie MAX_TOKENS und MAX_ZEICHEN in den Secrets.")
        raise ModellFehler(f"HTTP {r.status_code}: {text}")

    if STREAM:
        teile = []
        try:
            for zeile in r.iter_lines(decode_unicode=True):
                if not zeile or not zeile.startswith("data:"):
                    continue
                nutzdaten = zeile[5:].strip()
                if nutzdaten == "[DONE]":
                    break
                try:
                    stueck = json.loads(nutzdaten)
                except json.JSONDecodeError:
                    continue
                delta = (stueck.get("choices") or [{}])[0].get("delta", {})
                teile.append(delta.get("content") or "")
        except requests.RequestException as f:
            raise ModellFehler(f"Die Antwort wurde unterbrochen: {f}")
        inhalt = "".join(teile).strip()
    else:
        nachricht = r.json().get("choices", [{}])[0].get("message", {})
        inhalt = (nachricht.get("content") or "").strip()

    audit.setdefault("protokoll", []).append(
        {"zeit": datetime.now().isoformat(timespec="seconds"), "modell": modell,
         "system": system[:200], "zeichen_eingabe": len(nutzer), "ausgabe": inhalt[:4000]})
    if not inhalt:
        raise ModellFehler("Das Modell hat nichts zurückgegeben. Bitte erneut versuchen.")
    return json_aus_text(inhalt)


# ---------------------------------------------------------------------------
# Logik
# ---------------------------------------------------------------------------
def normalisieren(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def zitate_pruefen(nachweise, quellen):
    for n in nachweise or []:
        z = normalisieren(n.get("zitat", ""))
        n["verifiziert"] = bool(z) and z in normalisieren(quellen.get(n.get("dokument", ""), ""))
    return nachweise


def quellen_sammeln(audit: dict) -> dict:
    quellen = dict(audit.get("dokumente", {}))
    for i, n in enumerate(audit.get("notizen", []), start=1):
        quellen[f"{n['typ']} {i} ({n.get('quelle', 'ohne Angabe')}, {n.get('datum', '')})"] = n.get("text", "")
    return quellen


def korpus(quellen: dict) -> str:
    text = "".join(f"\n=== NACHWEIS: {n} ===\n{i}\n" for n, i in quellen.items())
    if len(text) > MAX_ZEICHEN:
        st.warning("Sehr viel Material, der Text wurde gekürzt. Weniger Nachweise je Lauf wählen.")
    return text[:MAX_ZEICHEN]


def kriterium_text(katalog: dict, abschnitt: dict) -> str:
    punkte = "\n".join(f"- {p}" for p in abschnitt.get("pruefpunkte", []))
    return (f"Norm: {katalog['norm']}, Abschnitt {abschnitt['nr']} {abschnitt['titel']}\n\n"
            f"Anforderung\n{abschnitt['anforderung']}\n\nPrüfpunkte\n{punkte}")


def analyse_durchfuehren(audit, katalog, abschnitt, mit_erwiderung=True):
    quellen = quellen_sammeln(audit)
    kriterium = kriterium_text(katalog, abschnitt)
    material = korpus(quellen)
    fall = llm_json(MODEL_PRO, PRO_SYSTEM,
                    f"Auditkriterium\n{kriterium}\n\nNachweise\n{material}", audit)
    for t in fall.get("teilaussagen", []):
        zitate_pruefen(t.get("nachweise"), quellen)
    einw = llm_json(MODEL_CONTRA, CONTRA_SYSTEM,
                    f"Auditkriterium\n{kriterium}\n\nAssurance Case\n"
                    f"{json.dumps(fall, ensure_ascii=False)}\n\nNachweise\n{material}", audit)
    praefix = abschnitt["nr"].replace(".", "_")
    for x in einw.get("einwaende", []):
        x["id"] = f"{praefix}-{x.get('id', 'E')}"
        x["abschnitt"] = abschnitt["nr"]
        zitate_pruefen(x.get("nachweise"), quellen)
    erw = {"erwiderungen": []}
    if mit_erwiderung and einw.get("einwaende"):
        erw = llm_json(MODEL_PRO, ERWIDERUNG_SYSTEM,
                       f"Einwände\n{json.dumps(einw, ensure_ascii=False)}\n\n"
                       f"Nachweise\n{material}", audit)
        for x in erw.get("erwiderungen", []):
            zitate_pruefen(x.get("nachweise"), quellen)
    audit.setdefault("analysen", {})[abschnitt["nr"]] = {
        "titel": abschnitt["titel"], "kriterium": kriterium, "fall": fall,
        "einwaende": einw, "erwiderungen": erw, "erwiderungsrunde": mit_erwiderung,
        "modelle": {"pro": MODEL_PRO, "contra": MODEL_CONTRA},
        "zeit": datetime.now().isoformat(timespec="seconds")}


def nachweise_anzeigen(nachweise):
    for n in nachweise or []:
        marke = "Zitat belegt" if n.get("verifiziert") else "Zitat NICHT in der Quelle gefunden"
        st.markdown(f"> {n.get('zitat', '')}  \n*{n.get('dokument', '')}* · {marke}")


def bericht_markdown(audit: dict) -> str:
    z = [f"# Auditbericht · {audit.get('titel', '')}", "",
         f"**Organisation** {ORGANISATION}  ",
         f"**Auditierter Bereich oder Prozess** {audit.get('prozess', '')}  ",
         f"**Auditziel** {audit.get('ziel', '')}  ",
         f"**Umfang** {audit.get('umfang', '')}  ",
         f"**Auditkriterien** {audit.get('norm', '')}, Abschnitte {', '.join(audit.get('kriterien', []))}  ",
         f"**Auditor** {audit.get('auditor', '')}  ",
         f"**Auditierte** {audit.get('auditierte', '')}  ",
         f"**Termin** {audit.get('termin', '')}  ",
         f"**Bericht erstellt** {date.today():%d.%m.%Y}", "", "## Geprüfte Nachweise"]
    z += [f"- {n}" for n in audit.get("dokumente", {})]
    z += [f"- {n['typ']} mit {n.get('quelle', '')} am {n.get('datum', '')}"
          for n in audit.get("notizen", [])]
    z += ["", "## Feststellungen"]
    fs = audit.get("feststellungen", {}).get("feststellungen", [])
    if not fs:
        z.append("Keine Feststellungen.")
    for f in fs:
        z += [f"### {f.get('einwand')} · {f.get('einstufung')} (Abschnitt {f.get('abschnitt', '')})",
              f"**Anforderung** {f.get('anforderung')}  ",
              f"**Objektiver Nachweis** {f.get('objektiver_nachweis')}  ",
              f"**Feststellung** {f.get('feststellung')}", ""]
    z += ["## Ausgeräumte Zweifel (Nachweis der Prüftiefe)"]
    ausger = [f"- {i}: {u['begruendung']}" for i, u in audit.get("urteile", {}).items()
              if u.get("urteil") == "Ausgeräumt"]
    z += ausger or ["Keine."]
    z += ["", "## Massnahmen"]
    for m in audit.get("massnahmen", []):
        z.append(f"- {m['beschreibung']} · verantwortlich {m['verantwortlich']} · "
                 f"Termin {m['termin']} · Status {m['status']}")
    z += ["", "*KI-gestützter Entwurf nach dialektischem Verfahren. Einstufungen und "
          "Schlussfolgerungen verantwortet der Auditor.*"]
    return "\n".join(z)


# ---------------------------------------------------------------------------
# Oberflaeche
# ---------------------------------------------------------------------------
kataloge = normkataloge_laden()
audits = st.session_state.audits

st.title(APP_NAME)
if ORGANISATION:
    st.caption(f"{ORGANISATION} · Internes Audit planen, durchführen und auswerten")

with st.sidebar:
    bereich = st.radio("Bereich", ["1 · Auditprogramm", "2 · Einzelaudit", "3 · Massnahmen"],
                       label_visibility="collapsed")
    st.divider()
    st.subheader("Arbeitsstand")
    st.caption("Diese Anwendung speichert nichts dauerhaft. Sichern Sie Ihren Stand als Datei "
               "und laden Sie ihn beim nächsten Mal wieder hoch.")
    st.download_button("Arbeitsstand sichern",
                       json.dumps(audits, ensure_ascii=False, indent=1),
                       file_name=f"auditstand_{date.today()}.json", use_container_width=True)
    wieder = st.file_uploader("Arbeitsstand laden", type=["json"], key="restore")
    if wieder is not None and st.button("Übernehmen", use_container_width=True):
        try:
            st.session_state.audits = json.loads(wieder.getvalue().decode("utf-8"))
            st.success("Arbeitsstand geladen.")
            st.rerun()
        except (json.JSONDecodeError, UnicodeDecodeError):
            st.error("Die Datei konnte nicht gelesen werden.")
    st.divider()
    with st.expander("Hinweise zur Nutzung"):
        st.markdown(HINWEISE)
    with st.expander("Technische Angaben"):
        st.caption(f"Dienst {PROVIDER}\n\nKonstruktionsteam {MODEL_PRO}\n\n"
                   f"Falsifikationsteam {MODEL_CONTRA}")
        schnell = st.checkbox("Schnellmodus ohne Erwiderungsrunde", value=False)
        if not API_KEY:
            st.error("Kein API-Schlüssel in den Secrets hinterlegt.")

# ---------------------------------------------------------------- 1 Programm
if bereich.startswith("1"):
    st.subheader("Auditprogramm")
    st.caption("ISO 9001 Abschnitt 9.2 verlangt ein geplantes Programm, das Bedeutung und "
               "Risiko der Prozesse sowie frühere Ergebnisse berücksichtigt.")
    if not kataloge:
        st.error("Kein Normkatalog gefunden. Legen Sie eine JSON-Datei im Ordner 'normen' ab.")
        st.stop()

    with st.expander("Neues Audit ins Programm aufnehmen", expanded=not audits):
        with st.form("neu"):
            s1, s2 = st.columns(2)
            titel = s1.text_input("Bezeichnung", "Prozessaudit Einkauf")
            prozess = s2.text_input("Prozess oder Bereich", "Einkauf")
            risiko = s1.selectbox("Risikoeinstufung", RISIKO)
            termin = s2.date_input("Geplanter Termin", date.today())
            auditor = s1.text_input("Auditor")
            auditierte = s2.text_input("Auditierte Personen oder Funktion")
            norm = st.selectbox("Norm", list(kataloge))
            liste = {f"{a['nr']} {a['titel']}": a["nr"] for a in kataloge[norm]["abschnitte"]}
            wahl = st.multiselect("Zu prüfende Normabschnitte", list(liste))
            ziel = st.text_area("Auditziel",
                                "Feststellung der Konformität und Wirksamkeit des Prozesses.")
            umfang = st.text_area("Umfang und Grenzen",
                                  "Standort, Zeitraum, betrachtete Tätigkeiten.")
            if st.form_submit_button("Audit anlegen"):
                audits.append({"id": date.today().strftime("%Y%m%d") + "-" + uuid.uuid4().hex[:4],
                               "titel": titel, "prozess": prozess, "risiko": risiko,
                               "termin": str(termin), "auditor": auditor,
                               "auditierte": auditierte, "norm": norm,
                               "kriterien": [liste[w] for w in wahl], "ziel": ziel,
                               "umfang": umfang, "status": "geplant", "dokumente": {},
                               "notizen": [], "analysen": {}, "urteile": {},
                               "feststellungen": {}, "massnahmen": [], "protokoll": []})
                st.success("Audit angelegt.")
                st.rerun()

    if audits:
        st.markdown("**Jahresübersicht**")
        st.dataframe([{"ID": a["id"], "Bezeichnung": a["titel"], "Prozess": a.get("prozess", ""),
                       "Risiko": a.get("risiko", ""), "Termin": a.get("termin", ""),
                       "Auditor": a.get("auditor", ""),
                       "Abschnitte": ", ".join(a.get("kriterien", [])),
                       "Status": a.get("status", "")} for a in audits],
                     use_container_width=True)
        st.markdown("**Abdeckung der Normabschnitte**")
        norm_wahl = st.selectbox("Norm für die Abdeckungsprüfung", list(kataloge), key="abd")
        alle = [a["nr"] for a in kataloge[norm_wahl]["abschnitte"]]
        geplant = {nr for a in audits if a.get("norm") == norm_wahl for nr in a.get("kriterien", [])}
        offen = [nr for nr in alle if nr not in geplant]
        s1, s2 = st.columns(2)
        s1.metric("Abgedeckte Abschnitte", f"{len(alle) - len(offen)} / {len(alle)}")
        s2.metric("Noch offen", len(offen))
        if offen:
            st.info("Noch nicht im Programm enthalten: " + ", ".join(offen))
        else:
            st.success("Alle Abschnitte sind im Programm abgedeckt.")
    else:
        st.info("Noch keine Audits im Programm.")

# ---------------------------------------------------------------- 2 Einzelaudit
elif bereich.startswith("2"):
    if not audits:
        st.warning("Legen Sie zuerst im Auditprogramm ein Audit an.")
        st.stop()
    namen = {f"{a['termin']} · {a['titel']} ({a['id']})": a for a in audits}
    with st.sidebar:
        st.subheader("Audit")
        audit = namen[st.selectbox("Audit", list(namen), label_visibility="collapsed")]
    katalog = kataloge.get(audit.get("norm"))
    if not katalog:
        st.error("Der Normkatalog dieses Audits wurde nicht gefunden.")
        st.stop()
    abschnitte = {a["nr"]: a for a in katalog["abschnitte"]}

    t1, t2, t3, t4, t5 = st.tabs(["1 Planung", "2 Vorbereitung", "3 Durchführung",
                                  "4 Urteil", "5 Bericht und Massnahmen"])

    with t1:
        st.caption("Auditplan nach ISO 19011. Grundlage für die Einladung der auditierten Stelle.")
        with st.form("plan"):
            s1, s2 = st.columns(2)
            audit["titel"] = s1.text_input("Bezeichnung", audit.get("titel", ""))
            audit["prozess"] = s2.text_input("Prozess oder Bereich", audit.get("prozess", ""))
            audit["auditor"] = s1.text_input("Auditor (Richter)", audit.get("auditor", ""))
            audit["auditierte"] = s2.text_input("Auditierte", audit.get("auditierte", ""))
            audit["termin"] = str(s1.date_input(
                "Termin", date.fromisoformat(audit.get("termin", str(date.today())))))
            audit["status"] = s2.selectbox("Status", STATUS,
                                           index=STATUS.index(audit.get("status", "geplant")))
            audit["ziel"] = st.text_area("Auditziel", audit.get("ziel", ""))
            audit["umfang"] = st.text_area("Umfang und Grenzen", audit.get("umfang", ""))
            liste = {f"{a['nr']} {a['titel']}": a["nr"] for a in katalog["abschnitte"]}
            vor = [k for k, v in liste.items() if v in audit.get("kriterien", [])]
            neu = st.multiselect("Auditkriterien (Normabschnitte)", list(liste), default=vor)
            if st.form_submit_button("Auditplan speichern"):
                audit["kriterien"] = [liste[n] for n in neu]
                st.success("Gespeichert.")
        if audit.get("kriterien"):
            plan = (f"# Auditplan {audit['titel']}\n\nTermin {audit['termin']}\n"
                    f"Prozess {audit['prozess']}\nAuditor {audit['auditor']}\n"
                    f"Auditierte {audit['auditierte']}\n\nZiel\n{audit['ziel']}\n\n"
                    f"Umfang\n{audit['umfang']}\n\nKriterien\n{audit['norm']}, Abschnitte "
                    + ", ".join(audit["kriterien"]))
            st.download_button("Auditplan herunterladen", plan,
                               file_name=f"auditplan_{audit['id']}.md")

    with t2:
        st.caption("Dokumentenprüfung. Die Agenten arbeiten jeden gewählten Normabschnitt "
                   "einzeln ab.")
        neue = st.file_uploader("Nachweisdokumente hinzufügen",
                                type=["pdf", "docx", "txt", "md"],
                                accept_multiple_files=True, key=f"up_{audit['id']}")
        if neue and st.button("Dokumente übernehmen"):
            for f in neue:
                audit.setdefault("dokumente", {})[f.name] = datei_lesen(f)
            st.success(f"{len(neue)} Dokument(e) übernommen.")
            st.rerun()
        for name in list(audit.get("dokumente", {})):
            s1, s2 = st.columns([6, 1])
            s1.write(name)
            if s2.button("entfernen", key=f"del_{name}"):
                del audit["dokumente"][name]
                st.rerun()
        st.divider()
        if not audit.get("kriterien"):
            st.warning("Bitte zuerst im Auditplan Normabschnitte wählen.")
        elif not quellen_sammeln(audit):
            st.warning("Bitte zuerst Nachweise hochladen oder Notizen erfassen.")
        else:
            offen = [nr for nr in audit["kriterien"] if nr not in audit.get("analysen", {})]
            st.write("Noch nicht analysiert: " + (", ".join(offen) if offen else "keine"))
            s1, s2 = st.columns(2)
            einzel = s1.selectbox("Abschnitt", audit["kriterien"])
            if s1.button("Diesen Abschnitt analysieren"):
                try:
                    with st.status(f"Abschnitt {einzel} wird geprüft", expanded=True) as stt:
                        analyse_durchfuehren(audit, katalog, abschnitte[einzel], not schnell)
                        stt.update(label=f"Abschnitt {einzel} fertig", state="complete")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))
            if s2.button("Alle offenen Abschnitte analysieren", disabled=not offen):
                try:
                    with st.status("Sammellauf läuft", expanded=True) as stt:
                        for nr in offen:
                            st.write(f"Abschnitt {nr}")
                            analyse_durchfuehren(audit, katalog, abschnitte[nr], not schnell)
                        stt.update(label="Sammellauf fertig", state="complete")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))
        for nr, erg in audit.get("analysen", {}).items():
            with st.expander(f"Ergebnis Abschnitt {nr} {erg.get('titel', '')}"):
                fall = erg.get("fall", {})
                if fall.get("fehler"):
                    st.warning(fall["fehler"])
                st.markdown(f"**Hauptaussage** {fall.get('hauptaussage', '')}")
                for t in fall.get("teilaussagen", []):
                    st.markdown(f"**{t.get('id')}** {t.get('aussage')}")
                    st.caption(f"Prüfpunkt · {t.get('pruefpunkt', '')}")
                    nachweise_anzeigen(t.get("nachweise"))
                    for a in t.get("annahmen", []):
                        st.markdown(f"Annahme ohne Beleg · {a}")
                if st.button(f"Analyse {nr} verwerfen", key=f"rm_{nr}"):
                    del audit["analysen"][nr]
                    st.rerun()

    with t3:
        st.caption("Interviews, Beobachtungen und Leistungsdaten. Diese Notizen gehen als "
                   "gleichwertige Nachweise in die Analyse ein und decken Widersprüche "
                   "zwischen Vorgabe und gelebter Praxis auf.")
        with st.form("notiz", clear_on_submit=True):
            s1, s2, s3 = st.columns(3)
            typ = s1.selectbox("Art", ["Interviewnotiz", "Beobachtung", "Leistungsdaten"])
            quelle = s2.text_input("Quelle (Person, Arbeitsplatz, Kennzahl)")
            datum = s3.date_input("Datum", date.today())
            text = st.text_area("Notiz, möglichst wörtlich festhalten", height=150)
            if st.form_submit_button("Notiz speichern") and text.strip():
                audit.setdefault("notizen", []).append(
                    {"typ": typ, "quelle": quelle, "datum": str(datum), "text": text})
                st.success("Gespeichert. Für die Berücksichtigung die Analyse erneut laufen lassen.")
                st.rerun()
        for i, n in enumerate(audit.get("notizen", [])):
            with st.container(border=True):
                st.markdown(f"**{n['typ']}** · {n.get('quelle', '')} · {n.get('datum', '')}")
                st.write(n["text"])
                if st.button("löschen", key=f"nd_{i}"):
                    audit["notizen"].pop(i)
                    st.rerun()

    with t4:
        alle_einw = [e for erg in audit.get("analysen", {}).values()
                     for e in erg.get("einwaende", {}).get("einwaende", [])]
        if not alle_einw:
            st.info("Noch keine Einwände. Führen Sie zuerst eine Analyse durch.")
        else:
            erw = {x.get("einwand"): x for erg in audit.get("analysen", {}).values()
                   for x in erg.get("erwiderungen", {}).get("erwiderungen", [])}
            st.caption("Sie entscheiden über jeden Einwand und begründen das. "
                       "Ohne Begründung entsteht kein Bericht.")
            for e in alle_einw:
                eid = e["id"]
                alt = audit.get("urteile", {}).get(eid, {})
                with st.container(border=True):
                    st.markdown(f"#### {eid} · {TYPEN.get(e.get('typ'), e.get('typ'))} "
                                f"· Schwere {e.get('schwere')}")
                    st.caption(f"Abschnitt {e.get('abschnitt')} · gegen {e.get('ziel')}")
                    links, rechts = st.columns(2)
                    with links:
                        st.markdown("**Einwand Falsifikationsteam**")
                        st.write(e.get("begruendung"))
                        nachweise_anzeigen(e.get("nachweise"))
                    with rechts:
                        st.markdown("**Erwiderung Konstruktionsteam**")
                        kurz = erw.get(eid.split("-", 1)[-1]) or erw.get(eid)
                        if kurz:
                            st.write(("zugestanden · " if kurz.get("zugestanden") else "")
                                     + str(kurz.get("erwiderung")))
                            nachweise_anzeigen(kurz.get("nachweise"))
                        else:
                            st.write("keine Erwiderung")
                    st.caption(f"Vorschlag Prüfung vor Ort · {e.get('pruefung_vor_ort')}")
                    idx = URTEILE.index(alt["urteil"]) if alt.get("urteil") in URTEILE else None
                    urteil = st.radio("Ihr Urteil", URTEILE, key=f"u_{eid}", index=idx,
                                      horizontal=True)
                    begr = st.text_area("Begründung (Pflicht)", value=alt.get("begruendung", ""),
                                        key=f"b_{eid}", height=80)
                    audit.setdefault("urteile", {})[eid] = {
                        "urteil": urteil, "begruendung": begr, "einwand": e}
            fehlend = [i for i, u in audit.get("urteile", {}).items()
                       if not (u.get("urteil") and u.get("begruendung", "").strip())]
            if fehlend:
                st.warning("Ohne begründetes Urteil: " + ", ".join(fehlend))
            if st.button("Urteile festschreiben und Feststellungen entwerfen",
                         type="primary", disabled=bool(fehlend)):
                relevant = [u for u in audit["urteile"].values() if u["urteil"] != "Ausgeräumt"]
                try:
                    with st.spinner("Feststellungen werden formuliert"):
                        audit["feststellungen"] = llm_json(
                            MODEL_PRO, BERICHT_SYSTEM,
                            f"Auditkriterien {audit['norm']}, Abschnitte "
                            f"{', '.join(audit['kriterien'])}\n\nEntscheidungen des Auditors\n"
                            f"{json.dumps(relevant, ensure_ascii=False)}", audit)
                    st.success("Entwurf erstellt, weiter im Reiter Bericht.")
                except ModellFehler as f:
                    st.error(str(f))

    with t5:
        if not audit.get("feststellungen"):
            st.info("Der Bericht entsteht, sobald alle Einwände beurteilt sind.")
        else:
            st.markdown(bericht_markdown(audit))
            st.divider()
            st.subheader("Massnahmen")
            with st.form("mass", clear_on_submit=True):
                s1, s2, s3 = st.columns(3)
                besch = s1.text_input("Massnahme")
                verantw = s2.text_input("Verantwortlich")
                term = s3.date_input("Termin", date.today() + timedelta(days=30))
                if st.form_submit_button("Massnahme hinzufügen") and besch.strip():
                    audit.setdefault("massnahmen", []).append(
                        {"id": uuid.uuid4().hex[:6], "beschreibung": besch,
                         "verantwortlich": verantw, "termin": str(term), "status": "offen"})
                    st.rerun()
            for m in audit.get("massnahmen", []):
                s1, s2 = st.columns([5, 2])
                s1.write(f"{m['beschreibung']} · {m['verantwortlich']} · bis {m['termin']}")
                m["status"] = s2.selectbox("Status", MSTATUS, index=MSTATUS.index(m["status"]),
                                           key=f"ms_{m['id']}", label_visibility="collapsed")
            st.divider()
            trail = {k: v for k, v in audit.items() if k != "dokumente"}
            trail["nachweise"] = list(audit.get("dokumente", {}))
            s1, s2 = st.columns(2)
            s1.download_button("Auditbericht (Markdown)", bericht_markdown(audit),
                               file_name=f"auditbericht_{audit['id']}.md",
                               use_container_width=True)
            s2.download_button("Audit Trail (JSON)",
                               json.dumps(trail, ensure_ascii=False, indent=2),
                               file_name=f"audit_trail_{audit['id']}.json",
                               use_container_width=True)

# ---------------------------------------------------------------- 3 Massnahmen
else:
    st.subheader("Massnahmenverfolgung über alle Audits")
    zeilen = [{"Audit": a["titel"], "Massnahme": m["beschreibung"],
               "Verantwortlich": m["verantwortlich"], "Termin": m["termin"],
               "Status": m["status"],
               "überfällig": "ja" if (m["status"] in ("offen", "in Umsetzung")
                                      and m["termin"] < str(date.today())) else ""}
              for a in audits for m in a.get("massnahmen", [])]
    if not zeilen:
        st.info("Noch keine Massnahmen erfasst.")
    else:
        s1, s2, s3 = st.columns(3)
        s1.metric("Massnahmen gesamt", len(zeilen))
        s2.metric("Offen", sum(1 for z in zeilen if z["Status"] in ("offen", "in Umsetzung")))
        s3.metric("Überfällig", sum(1 for z in zeilen if z["überfällig"]))
        st.dataframe(zeilen, use_container_width=True)
        st.caption("Status ändern Sie im jeweiligen Audit im Reiter Bericht und Massnahmen.")
