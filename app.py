"""
Audit-Assistent Qualitaetsmanagement (Webversion)
Oberflaeche: Streamlit Community Cloud
Sprachmodell: OpenAI-kompatible Schnittstelle (xAI, Public AI oder anderer Anbieter)

Ablauf: Das System bereitet das interne Audit weitgehend selbsttaetig vor und prueft es.
Konstruktionsteam baut den Konformitaetsnachweis, Falsifikationsteam greift ihn an,
Massnahmenteam leitet Verbesserungen ab. Der Auditor entscheidet als Richter ueber
jeden Zweifel. Diese Entscheidung bleibt beim Menschen (ISO 19011).
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
PROVIDER = st.secrets.get("PROVIDER_NAME", "")
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
- Das System bereitet vor und prüft, es entscheidet nicht. Einstufung und Schlussfolgerung
  jeder Feststellung verantwortet der Auditor.
- Ihre Unterlagen werden zur Auswertung an den Dienst {PROVIDER or 'des eingestellten Anbieters'}
  übermittelt. Laden Sie keine vertraulichen Originalunterlagen und keine Personendaten hoch,
  solange das mit Ihrer IT und dem Datenschutz nicht geklärt ist.
- Jedes Zitat wird gegen die Quelle geprüft und gekennzeichnet. Nicht bestätigte Zitate
  sind ein Warnzeichen und dürfen nicht in einen Bericht übernommen werden.
- Der Arbeitsstand liegt nur in dieser Sitzung. Sichern Sie ihn links als Datei.
"""

# ---------------------------------------------------------------------------
# Rollen der Agenten
# ---------------------------------------------------------------------------
SPRACHREGEL = ("\n\nSPRACHE: Schreibe alle Inhalte ausschliesslich auf Deutsch in Schweizer "
               "Rechtschreibung (ss statt ß). Zitate aus Nachweisen übernimmst du wörtlich.")
NORMREGEL = ("\n\nNORMBEZUG: Beziehe dich nur auf die genannte Norm und Abschnittsnummer. "
             "Erfinde keine weiteren Normabschnitte und keine Normzitate.")
JSONREGEL = "\n\nAUSGABE: Antworte ausschliesslich mit gültigem JSON, ohne Vor- und Nachtext."

SCOPING_SYSTEM = """Du bist das VORBEREITUNGSTEAM eines internen Audits.
Du erhältst die Liste der Normabschnitte und die vorhandenen Nachweise der Organisation.
Deine Aufgabe ist der Auditplan.
Regeln
1. Wähle nur Normabschnitte, zu denen die Nachweise tatsächlich etwas hergeben.
2. Höchstens sechs Abschnitte, nach Bedeutung geordnet.
3. Begründe jede Wahl in einem Satz mit Bezug auf die Nachweise.
Schema
{"kriterien": ["7.2"], "begruendungen": {"7.2": "..."}, "ziel": "...", "umfang": "...",
 "fehlende_nachweise": ["..."]}"""

PRO_SYSTEM = """Du bist das KONSTRUKTIONSTEAM eines internen Audits.
Baue einen Konformitätsnachweis als Assurance Case (Aussage, Argument, Nachweis).
Ordne jede Teilaussage einem Prüfpunkt des Auditkriteriums zu.
Regeln
1. Nutze ausschliesslich die bereitgestellten Nachweise.
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
Achte besonders auf Widersprüche zwischen Vorgabe, Aufzeichnung, Aussagen aus Interviews,
Beobachtungen und Leistungsdaten.
Regeln
1. Erfinde nichts. Zitate müssen wörtlich aus den Nachweisen stammen.
2. Fehlt ein Nachweis ganz, lass die Liste "nachweise" leer und begründe das.
3. Schlage für jeden Einwand eine konkrete Prüfung vor Ort vor.
Schema
{"einwaende": [{"id": "E1", "ziel": "T1", "typ": "...", "begruendung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}], "schwere": "hoch|mittel|gering",
 "pruefung_vor_ort": "..."}]}"""

ERWIDERUNG_SYSTEM = """Du bist das KONSTRUKTIONSTEAM. Erwidere auf jeden Einwand sachlich.
Gestehe Einwände zu, wenn die Nachweise sie stützen. Nutze nur wörtliche Zitate.
Schema
{"erwiderungen": [{"einwand": "E1", "zugestanden": true, "erwiderung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}]}]}"""

BERICHT_SYSTEM = """Du formulierst Auditfeststellungen nach ISO 19011, also Anforderung mit
Normabschnitt, objektiver Nachweis und Feststellung. Die Einstufung hat der Auditor bereits
getroffen, du darfst sie NICHT ändern. Formuliere sachlich, prüfbar, ohne Schuldzuweisung
und ohne Namen einzelner Personen.
Schreibe zusätzlich eine Zusammenfassung von höchstens acht Sätzen, die beschreibt, was
geprüft wurde und wo die Schwerpunkte der Feststellungen liegen. Bewerte darin nicht.
Schema
{"zusammenfassung": "...", "feststellungen": [{"einwand": "E1", "einstufung": "...",
 "abschnitt": "...", "anforderung": "...", "objektiver_nachweis": "...", "feststellung": "..."}]}"""

MASSNAHMEN_SYSTEM = """Du bist das MASSNAHMENTEAM. Zu jeder bestätigten Feststellung leitest du
einen Verbesserungsvorschlag nach ISO 9001 Abschnitt 10.2 ab.
Regeln
1. Trenne Sofortmassnahme (beseitigt die Folge) und Korrekturmassnahme (beseitigt die Ursache).
2. Nenne eine Ursachenhypothese, klar als Hypothese erkennbar, die der Auditor prüfen muss.
3. Schlage als Verantwortung eine Rolle oder Funktion vor, niemals einen Personennamen.
4. Nenne eine realistische Frist in Tagen und einen konkreten Wirksamkeitsnachweis.
Schema
{"massnahmen": [{"feststellung": "E1", "ursache_hypothese": "...", "sofortmassnahme": "...",
 "korrekturmassnahme": "...", "verantwortlich_rolle": "...", "frist_tage": 30,
 "wirksamkeitsnachweis": "..."}]}"""

URTEILE = ["Ausgeräumt", "Bestätigt: Abweichung", "Bestätigt: Verbesserungspotenzial",
           "Offen: vor Ort prüfen"]
TYPEN = {"widerlegend": "widerlegend", "untergrabend": "untergrabend",
         "unterlaufend": "unterlaufend", "ungestuetzte_annahme": "ungestützte Annahme"}
RISIKO = ["hoch", "mittel", "gering"]
STATUS = ["geplant", "in Arbeit", "abgeschlossen"]
MSTATUS = ["offen", "in Umsetzung", "umgesetzt", "wirksam bestätigt"]


# ---------------------------------------------------------------------------
# Zugang
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

st.session_state.setdefault("audits", [])


# ---------------------------------------------------------------------------
# Dateien und Normkataloge
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
    """Anfrage an die OpenAI-kompatible Schnittstelle, bei Bedarf als Datenstrom."""
    system = system + NORMREGEL + SPRACHREGEL + JSONREGEL
    basis = {"model": modell, "temperature": 0.2, "max_tokens": MAX_TOKENS, "stream": STREAM,
             "response_format": {"type": "json_object"},
             "messages": [{"role": "system", "content": system},
                          {"role": "user", "content": nutzer}]}
    kopf = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    def senden(mit_format: bool):
        daten = dict(basis)
        if not mit_format:
            daten.pop("response_format", None)
        return requests.post(f"{API_BASE}/chat/completions", headers=kopf, json=daten,
                             timeout=(15, TIMEOUT), stream=STREAM)

    try:
        r = senden(True)
        if r.status_code == 400:
            r.close()
            r = senden(False)
    except requests.exceptions.Timeout:
        raise ModellFehler("Zeitüberschreitung. Weniger Nachweise je Lauf verwenden.")
    except requests.RequestException as f:
        raise ModellFehler(f"Verbindung fehlgeschlagen: {f}")

    if r.status_code != 200:
        if r.status_code in (502, 503, 504):
            raise ModellFehler(f"Der Dienst hat abgebrochen (HTTP {r.status_code}). "
                               "Weniger Nachweise laden oder MAX_TOKENS und MAX_ZEICHEN "
                               "in den Secrets verkleinern.")
        raise ModellFehler(f"HTTP {r.status_code}: {r.text[:400]}")

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
        inhalt = (r.json().get("choices", [{}])[0].get("message", {}).get("content") or "").strip()

    audit.setdefault("protokoll", []).append(
        {"zeit": datetime.now().isoformat(timespec="seconds"), "modell": modell,
         "rolle": system[:60], "zeichen_eingabe": len(nutzer), "ausgabe": inhalt[:4000]})
    if not inhalt:
        raise ModellFehler("Das Modell hat nichts zurückgegeben. Bitte erneut versuchen.")
    return json_aus_text(inhalt)


# ---------------------------------------------------------------------------
# Nachweise und Pruefung
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
    return text[:MAX_ZEICHEN]


def kriterium_text(katalog: dict, abschnitt: dict) -> str:
    punkte = "\n".join(f"- {p}" for p in abschnitt.get("pruefpunkte", []))
    return (f"Norm: {katalog['norm']}, Abschnitt {abschnitt['nr']} {abschnitt['titel']}\n\n"
            f"Anforderung\n{abschnitt['anforderung']}\n\nPrüfpunkte\n{punkte}")


def plan_vorschlagen(audit: dict, katalog: dict):
    quellen = quellen_sammeln(audit)
    liste = "\n".join(f"- {a['nr']} {a['titel']}: {a['anforderung'][:160]}"
                      for a in katalog["abschnitte"])
    probe = "".join(f"\n=== {n} ===\n{t[:1500]}\n" for n, t in quellen.items())[:MAX_ZEICHEN]
    return llm_json(MODEL_PRO, SCOPING_SYSTEM,
                    f"Norm {katalog['norm']}\n\nNormabschnitte\n{liste}\n\n"
                    f"Auditierter Prozess: {audit.get('prozess', '')}\n\n"
                    f"Vorhandene Nachweise\n{probe}", audit)


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
        "titel": abschnitt["titel"], "kriterium": kriterium, "fall": fall, "einwaende": einw,
        "erwiderungen": erw, "modelle": {"pro": MODEL_PRO, "contra": MODEL_CONTRA},
        "zeit": datetime.now().isoformat(timespec="seconds")}


def alle_einwaende(audit: dict) -> list:
    return [e for erg in audit.get("analysen", {}).values()
            for e in erg.get("einwaende", {}).get("einwaende", [])]


def kennzahlen(audit: dict) -> dict:
    nachweise = [n for erg in audit.get("analysen", {}).values()
                 for t in erg.get("fall", {}).get("teilaussagen", [])
                 for n in t.get("nachweise", [])]
    nachweise += [n for e in alle_einwaende(audit) for n in e.get("nachweise", [])]
    urteile = [u.get("urteil") for u in audit.get("urteile", {}).values()]
    return {"abschnitte": len(audit.get("analysen", {})),
            "einwaende": len(alle_einwaende(audit)),
            "zitate": len(nachweise),
            "zitate_belegt": sum(1 for n in nachweise if n.get("verifiziert")),
            "abweichungen": sum(1 for u in urteile if u == "Bestätigt: Abweichung"),
            "potenziale": sum(1 for u in urteile if u == "Bestätigt: Verbesserungspotenzial"),
            "offen": sum(1 for u in urteile if u == "Offen: vor Ort prüfen"),
            "ausgeraeumt": sum(1 for u in urteile if u == "Ausgeräumt")}


def nachweise_anzeigen(nachweise):
    for n in nachweise or []:
        marke = "Zitat belegt" if n.get("verifiziert") else "Zitat NICHT in der Quelle gefunden"
        st.markdown(f"> {n.get('zitat', '')}  \n*{n.get('dokument', '')}* · {marke}")


# ---------------------------------------------------------------------------
# Bericht
# ---------------------------------------------------------------------------
def bericht_bloecke(audit: dict) -> dict:
    k = kennzahlen(audit)
    fs = audit.get("feststellungen", {}).get("feststellungen", [])
    kopf = [("Organisation", ORGANISATION or "nicht angegeben"),
            ("Auditierter Bereich oder Prozess", audit.get("prozess", "")),
            ("Auditziel", audit.get("ziel", "")),
            ("Umfang und Grenzen", audit.get("umfang", "")),
            ("Auditkriterien", f"{audit.get('norm', '')}, Abschnitte "
                               + ", ".join(audit.get("kriterien", []))),
            ("Auditor", audit.get("auditor", "")),
            ("Auditierte", audit.get("auditierte", "")),
            ("Audittermin", audit.get("termin", "")),
            ("Berichtsdatum", f"{date.today():%d.%m.%Y}")]
    quellen = ([f"Dokument · {n}" for n in audit.get("dokumente", {})]
               + [f"{n['typ']} · {n.get('quelle', '')} · {n.get('datum', '')}"
                  for n in audit.get("notizen", [])])
    return {"kennzahlen": k, "kopf": kopf, "quellen": quellen, "feststellungen": fs,
            "zusammenfassung": audit.get("feststellungen", {}).get("zusammenfassung", ""),
            "fazit": audit.get("fazit", ""),
            "ausgeraeumt": [(i, u["begruendung"]) for i, u in audit.get("urteile", {}).items()
                            if u.get("urteil") == "Ausgeräumt"],
            "massnahmen": audit.get("massnahmen", [])}


def bericht_markdown(audit: dict) -> str:
    b = bericht_bloecke(audit)
    k = b["kennzahlen"]
    z = [f"# Auditbericht · {audit.get('titel', '')}", ""]
    z += [f"**{name}** {wert}  " for name, wert in b["kopf"]]
    z += ["", "## Zusammenfassung", b["zusammenfassung"] or "nicht erstellt", ""]
    if b["fazit"]:
        z += ["### Fazit des Auditors", b["fazit"], ""]
    z += ["## Kennzahlen der Prüfung", "",
          "| Grösse | Wert |", "| --- | --- |",
          f"| Geprüfte Normabschnitte | {k['abschnitte']} |",
          f"| Geprüfte Zweifel insgesamt | {k['einwaende']} |",
          f"| Davon ausgeräumt | {k['ausgeraeumt']} |",
          f"| Abweichungen | {k['abweichungen']} |",
          f"| Verbesserungspotenziale | {k['potenziale']} |",
          f"| Vor Ort nachzuprüfen | {k['offen']} |",
          f"| Zitate gegen die Quelle bestätigt | {k['zitate_belegt']} von {k['zitate']} |", ""]
    z += ["## Geprüfte Nachweise"] + [f"- {q}" for q in b["quellen"]] + [""]
    z += ["## Feststellungen"]
    if not b["feststellungen"]:
        z.append("Keine Feststellungen.")
    for f in b["feststellungen"]:
        z += [f"### {f.get('einwand')} · {f.get('einstufung')}",
              f"**Normabschnitt** {f.get('abschnitt', '')}  ",
              f"**Anforderung** {f.get('anforderung')}  ",
              f"**Objektiver Nachweis** {f.get('objektiver_nachweis')}  ",
              f"**Feststellung** {f.get('feststellung')}", ""]
    z += ["## Massnahmenplan"]
    if not b["massnahmen"]:
        z.append("Keine Massnahmen erfasst.")
    else:
        z += ["", "| Feststellung | Massnahme | Verantwortlich | Termin | Status |",
              "| --- | --- | --- | --- | --- |"]
        for m in b["massnahmen"]:
            z.append(f"| {m.get('feststellung', '')} | {m.get('beschreibung', '')} | "
                     f"{m.get('verantwortlich', '')} | {m.get('termin', '')} | "
                     f"{m.get('status', '')} |")
        z.append("")
        for m in b["massnahmen"]:
            if m.get("ursache_hypothese") or m.get("wirksamkeitsnachweis"):
                z += [f"**{m.get('feststellung', '')}**  ",
                      f"Ursachenhypothese · {m.get('ursache_hypothese', '')}  ",
                      f"Sofortmassnahme · {m.get('sofortmassnahme', '')}  ",
                      f"Wirksamkeitsnachweis · {m.get('wirksamkeitsnachweis', '')}", ""]
    z += ["## Ausgeräumte Zweifel (Nachweis der Prüftiefe)"]
    z += [f"- {i}: {g}" for i, g in b["ausgeraeumt"]] or ["Keine."]
    z += ["", "---", "", "Erstellt mit einem dialektischen Prüfverfahren. Zwei getrennte "
          "Agententeams haben den Konformitätsnachweis konstruiert und angegriffen. "
          "Über jeden verbliebenen Zweifel hat der Auditor entschieden und dies begründet. "
          "Einstufung und Schlussfolgerung verantwortet der Auditor."]
    return "\n".join(z)


def bericht_html(audit: dict) -> str:
    b = bericht_bloecke(audit)
    k = b["kennzahlen"]

    def schutz(t):
        return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    kopf = "".join(f"<tr><th>{schutz(n)}</th><td>{schutz(w)}</td></tr>" for n, w in b["kopf"])
    kacheln = "".join(
        f'<div class="kachel"><div class="zahl">{w}</div><div class="text">{t}</div></div>'
        for t, w in [("Normabschnitte", k["abschnitte"]), ("Geprüfte Zweifel", k["einwaende"]),
                     ("Abweichungen", k["abweichungen"]), ("Potenziale", k["potenziale"]),
                     ("Vor Ort offen", k["offen"]),
                     ("Zitate belegt", f"{k['zitate_belegt']}/{k['zitate']}")])
    farbe = {"Bestätigt: Abweichung": "abw", "Bestätigt: Verbesserungspotenzial": "pot",
             "Offen: vor Ort prüfen": "off"}
    fest = "".join(
        f'<div class="fest {farbe.get(f.get("einstufung", ""), "")}">'
        f'<h3>{schutz(f.get("einwand"))} · {schutz(f.get("einstufung"))}</h3>'
        f'<p><b>Normabschnitt</b> {schutz(f.get("abschnitt", ""))}</p>'
        f'<p><b>Anforderung</b> {schutz(f.get("anforderung"))}</p>'
        f'<p><b>Objektiver Nachweis</b> {schutz(f.get("objektiver_nachweis"))}</p>'
        f'<p><b>Feststellung</b> {schutz(f.get("feststellung"))}</p></div>'
        for f in b["feststellungen"]) or "<p>Keine Feststellungen.</p>"
    mass = "".join(
        f"<tr><td>{schutz(m.get('feststellung', ''))}</td><td>{schutz(m.get('beschreibung', ''))}"
        f"<br><small>Ursachenhypothese: {schutz(m.get('ursache_hypothese', ''))}<br>"
        f"Wirksamkeitsnachweis: {schutz(m.get('wirksamkeitsnachweis', ''))}</small></td>"
        f"<td>{schutz(m.get('verantwortlich', ''))}</td><td>{schutz(m.get('termin', ''))}</td>"
        f"<td>{schutz(m.get('status', ''))}</td></tr>" for m in b["massnahmen"])
    quellen = "".join(f"<li>{schutz(q)}</li>" for q in b["quellen"])
    ausger = "".join(f"<li><b>{schutz(i)}</b> {schutz(g)}</li>" for i, g in b["ausgeraeumt"])
    fazit = (f"<h2>Fazit des Auditors</h2><p>{schutz(b['fazit'])}</p>") if b["fazit"] else ""
    return f"""<!doctype html><html lang="de"><head><meta charset="utf-8">
<title>Auditbericht {schutz(audit.get('titel', ''))}</title><style>
body{{font-family:Segoe UI,Helvetica,Arial,sans-serif;max-width:900px;margin:2rem auto;
padding:0 1.5rem;color:#1a1a1a;line-height:1.55}}
h1{{font-size:1.7rem;border-bottom:3px solid #1f4e79;padding-bottom:.4rem;color:#1f4e79}}
h2{{font-size:1.2rem;margin-top:2rem;color:#1f4e79}}
table{{border-collapse:collapse;width:100%;margin:.6rem 0}}
th,td{{border:1px solid #d5dde5;padding:.45rem .6rem;text-align:left;vertical-align:top;
font-size:.92rem}}
th{{background:#f0f4f8;width:16rem}}
.kacheln{{display:flex;flex-wrap:wrap;gap:.6rem;margin:1rem 0}}
.kachel{{flex:1 1 8rem;background:#f0f4f8;border-radius:8px;padding:.7rem;text-align:center}}
.zahl{{font-size:1.6rem;font-weight:700;color:#1f4e79}}
.text{{font-size:.78rem;color:#4a5a68}}
.fest{{border-left:5px solid #99a;background:#fafbfc;padding:.7rem 1rem;margin:.8rem 0;
border-radius:0 6px 6px 0}}
.fest.abw{{border-color:#b3261e}} .fest.pot{{border-color:#e08b00}} .fest.off{{border-color:#5a6b7a}}
.fest h3{{margin:.2rem 0 .5rem;font-size:1rem}} .fest p{{margin:.25rem 0;font-size:.92rem}}
small{{color:#55626e}} footer{{margin-top:2.5rem;font-size:.82rem;color:#55626e;
border-top:1px solid #d5dde5;padding-top:.8rem}}
@media print{{body{{margin:0}} .fest{{break-inside:avoid}}}}
</style></head><body>
<h1>Auditbericht · {schutz(audit.get('titel', ''))}</h1>
<table>{kopf}</table>
<h2>Zusammenfassung</h2><p>{schutz(b['zusammenfassung']) or 'nicht erstellt'}</p>
{fazit}
<h2>Kennzahlen der Prüfung</h2><div class="kacheln">{kacheln}</div>
<h2>Feststellungen</h2>{fest}
<h2>Massnahmenplan</h2>
<table><tr><th>Feststellung</th><th>Massnahme</th><th>Verantwortlich</th><th>Termin</th>
<th>Status</th></tr>{mass or '<tr><td colspan="5">Keine Massnahmen erfasst.</td></tr>'}</table>
<h2>Geprüfte Nachweise</h2><ul>{quellen}</ul>
<h2>Ausgeräumte Zweifel</h2><ul>{ausger or '<li>Keine.</li>'}</ul>
<footer>Erstellt mit einem dialektischen Prüfverfahren. Zwei getrennte Agententeams haben den
Konformitätsnachweis konstruiert und angegriffen. Über jeden verbliebenen Zweifel hat der
Auditor entschieden und dies begründet. Einstufung und Schlussfolgerung verantwortet der
Auditor.</footer></body></html>"""


# ---------------------------------------------------------------------------
# Oberflaeche
# ---------------------------------------------------------------------------
kataloge = normkataloge_laden()
audits = st.session_state.audits

st.title(APP_NAME)
st.caption((ORGANISATION + " · " if ORGANISATION else "")
           + "Internes Audit vorbereiten, prüfen, beurteilen und verbessern")

with st.sidebar:
    bereich = st.radio("Bereich", ["1 · Auditprogramm", "2 · Einzelaudit", "3 · Massnahmen"],
                       label_visibility="collapsed")
    st.divider()
    st.subheader("Arbeitsstand")
    st.caption("Nichts wird dauerhaft gespeichert. Stand als Datei sichern und später laden.")
    st.download_button("Arbeitsstand sichern", json.dumps(audits, ensure_ascii=False, indent=1),
                       file_name=f"auditstand_{date.today()}.json", use_container_width=True)
    wieder = st.file_uploader("Arbeitsstand laden", type=["json"], key="restore")
    if wieder is not None and st.button("Übernehmen", use_container_width=True):
        try:
            st.session_state.audits = json.loads(wieder.getvalue().decode("utf-8"))
            st.rerun()
        except (json.JSONDecodeError, UnicodeDecodeError):
            st.error("Die Datei konnte nicht gelesen werden.")
    st.divider()
    with st.expander("Hinweise zur Nutzung"):
        st.markdown(HINWEISE)
    with st.expander("Technik"):
        st.caption(f"Konstruktionsteam {MODEL_PRO}\n\nFalsifikationsteam {MODEL_CONTRA}")
        schnell = st.checkbox("Schnellmodus ohne Erwiderungsrunde", value=False)
        if not API_KEY:
            st.error("Kein API-Schlüssel in den Secrets hinterlegt.")

# ---------------------------------------------------------------- 1 Programm
if bereich.startswith("1"):
    st.subheader("Auditprogramm")
    st.caption("ISO 9001 Abschnitt 9.2 verlangt ein geplantes Programm nach Bedeutung und "
               "Risiko der Prozesse. Legen Sie hier die Audits des Jahres an.")
    if not kataloge:
        st.error("Kein Normkatalog gefunden. Legen Sie eine JSON-Datei im Ordner normen ab.")
        st.stop()
    with st.form("neu"):
        s1, s2, s3 = st.columns(3)
        titel = s1.text_input("Bezeichnung", "Prozessaudit Einkauf")
        prozess = s2.text_input("Prozess oder Bereich", "Einkauf")
        risiko = s3.selectbox("Risiko", RISIKO)
        auditor = s1.text_input("Auditor")
        auditierte = s2.text_input("Auditierte Funktion")
        termin = s3.date_input("Termin", date.today())
        norm = st.selectbox("Norm", list(kataloge))
        st.caption("Die zu prüfenden Normabschnitte schlägt das System später anhand Ihrer "
                   "Nachweise vor. Sie können sie jederzeit ändern.")
        if st.form_submit_button("Audit anlegen", type="primary"):
            audits.append({"id": date.today().strftime("%Y%m%d") + "-" + uuid.uuid4().hex[:4],
                           "titel": titel, "prozess": prozess, "risiko": risiko,
                           "termin": str(termin), "auditor": auditor, "auditierte": auditierte,
                           "norm": norm, "kriterien": [], "ziel": "", "umfang": "",
                           "status": "geplant", "dokumente": {}, "notizen": [],
                           "analysen": {}, "urteile": {}, "feststellungen": {},
                           "massnahmen": [], "protokoll": [], "fazit": ""})
            st.success("Audit angelegt. Weiter im Bereich Einzelaudit.")
            st.rerun()

    if audits:
        st.markdown("**Jahresübersicht**")
        st.dataframe([{"ID": a["id"], "Bezeichnung": a["titel"], "Prozess": a.get("prozess", ""),
                       "Risiko": a.get("risiko", ""), "Termin": a.get("termin", ""),
                       "Auditor": a.get("auditor", ""),
                       "Abschnitte": ", ".join(a.get("kriterien", [])) or "noch offen",
                       "Abweichungen": kennzahlen(a)["abweichungen"],
                       "Status": a.get("status", "")} for a in audits],
                     use_container_width=True)
        st.markdown("**Abdeckung der Normabschnitte**")
        norm_wahl = st.selectbox("Norm", list(kataloge), key="abd")
        alle = [a["nr"] for a in kataloge[norm_wahl]["abschnitte"]]
        geplant = {nr for a in audits if a.get("norm") == norm_wahl
                   for nr in a.get("kriterien", [])}
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
    namen = {f"{a['termin']} · {a['titel']}": a for a in audits}
    with st.sidebar:
        st.subheader("Audit")
        audit = namen[st.selectbox("Audit", list(namen), label_visibility="collapsed")]
    katalog = kataloge.get(audit.get("norm"))
    if not katalog:
        st.error("Der Normkatalog dieses Audits wurde nicht gefunden.")
        st.stop()
    abschnitte = {a["nr"]: a for a in katalog["abschnitte"]}
    k = kennzahlen(audit)

    with st.sidebar:
        st.divider()
        st.subheader("Fortschritt")
        for name, fertig in [("Nachweise vorhanden", bool(quellen_sammeln(audit))),
                             ("Auditplan steht", bool(audit.get("kriterien"))),
                             ("Prüfung gelaufen", bool(audit.get("analysen"))),
                             ("Urteile vollständig", k["einwaende"] > 0 and
                              k["einwaende"] == sum(1 for u in audit.get("urteile", {}).values()
                                                    if u.get("urteil"))),
                             ("Bericht erstellt", bool(audit.get("feststellungen")))]:
            st.markdown(f"{name} · {'erledigt' if fertig else 'offen'}")

    t1, t2, t3, t4 = st.tabs(["1 Nachweise", "2 Prüfung", "3 Urteil des Auditors",
                              "4 Bericht und Massnahmen"])

    # ---------------- 1 Nachweise
    with t1:
        st.caption("Alles, worauf sich das Audit stützt. Dokumente sind der Anfang, "
                   "erst Interviews, Beobachtungen und Leistungsdaten zeigen die gelebte Praxis.")
        neue = st.file_uploader("Dokumente hinzufügen (PDF, Word, Text)",
                                type=["pdf", "docx", "txt", "md"], accept_multiple_files=True,
                                key=f"up_{audit['id']}")
        if neue and st.button("Dokumente übernehmen", type="primary"):
            for f in neue:
                audit.setdefault("dokumente", {})[f.name] = datei_lesen(f)
            st.rerun()
        for name in list(audit.get("dokumente", {})):
            s1, s2 = st.columns([6, 1])
            s1.write(name)
            if s2.button("entfernen", key=f"del_{name}"):
                del audit["dokumente"][name]
                st.rerun()
        st.divider()
        st.markdown("**Interviews, Beobachtungen und Leistungsdaten**")
        with st.form("notiz", clear_on_submit=True):
            s1, s2, s3 = st.columns(3)
            typ = s1.selectbox("Art", ["Interviewnotiz", "Beobachtung", "Leistungsdaten"])
            quelle = s2.text_input("Quelle (Rolle, Arbeitsplatz, Kennzahl)")
            datum = s3.date_input("Datum", date.today())
            text = st.text_area("Notiz, möglichst wörtlich festhalten", height=120)
            if st.form_submit_button("Notiz speichern") and text.strip():
                audit.setdefault("notizen", []).append(
                    {"typ": typ, "quelle": quelle, "datum": str(datum), "text": text})
                st.rerun()
        for i, n in enumerate(audit.get("notizen", [])):
            with st.container(border=True):
                st.markdown(f"**{n['typ']}** · {n.get('quelle', '')} · {n.get('datum', '')}")
                st.write(n["text"])
                if st.button("löschen", key=f"nd_{i}"):
                    audit["notizen"].pop(i)
                    st.rerun()

    # ---------------- 2 Pruefung
    with t2:
        if not quellen_sammeln(audit):
            st.warning("Bitte zuerst Nachweise erfassen.")
        else:
            st.markdown("**Automatischer Durchlauf**")
            st.caption("Das System schlägt die passenden Normabschnitte vor, baut je Abschnitt "
                       "den Konformitätsnachweis, greift ihn an und legt Ihnen die verbliebenen "
                       "Zweifel zur Entscheidung vor.")
            if st.button("Audit vorbereiten und prüfen", type="primary",
                         use_container_width=True):
                try:
                    with st.status("Die Agententeams arbeiten", expanded=True) as stt:
                        if not audit.get("kriterien"):
                            st.write("Vorbereitungsteam wählt die Normabschnitte")
                            vor = plan_vorschlagen(audit, katalog)
                            gueltig = [n for n in vor.get("kriterien", []) if n in abschnitte]
                            audit["kriterien"] = gueltig or list(abschnitte)[:3]
                            audit["ziel"] = vor.get("ziel", audit.get("ziel", ""))
                            audit["umfang"] = vor.get("umfang", audit.get("umfang", ""))
                            audit["planbegruendung"] = vor.get("begruendungen", {})
                            audit["fehlende_nachweise"] = vor.get("fehlende_nachweise", [])
                            st.write("Gewählt: " + ", ".join(audit["kriterien"]))
                        for nr in audit["kriterien"]:
                            if nr in audit.get("analysen", {}):
                                continue
                            st.write(f"Abschnitt {nr} {abschnitte[nr]['titel']}")
                            analyse_durchfuehren(audit, katalog, abschnitte[nr], not schnell)
                        audit["status"] = "in Arbeit"
                        stt.update(label="Prüfung abgeschlossen", state="complete")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))

            if audit.get("analysen"):
                s1, s2, s3 = st.columns(3)
                s1.metric("Geprüfte Abschnitte", k["abschnitte"])
                s2.metric("Gefundene Zweifel", k["einwaende"])
                s3.metric("Zitate belegt", f"{k['zitate_belegt']} / {k['zitate']}")
            if audit.get("fehlende_nachweise"):
                st.info("Das Vorbereitungsteam vermisst: "
                        + "; ".join(audit["fehlende_nachweise"]))

            with st.expander("Auditplan ansehen und ändern"):
                liste = {f"{a['nr']} {a['titel']}": a["nr"] for a in katalog["abschnitte"]}
                vorauswahl = [t for t, v in liste.items() if v in audit.get("kriterien", [])]
                neu = st.multiselect("Normabschnitte", list(liste), default=vorauswahl)
                audit["ziel"] = st.text_area("Auditziel", audit.get("ziel", ""))
                audit["umfang"] = st.text_area("Umfang und Grenzen", audit.get("umfang", ""))
                if st.button("Auditplan übernehmen"):
                    audit["kriterien"] = [liste[n] for n in neu]
                    st.rerun()
                for nr, grund in (audit.get("planbegruendung") or {}).items():
                    st.caption(f"{nr} · {grund}")

            for nr, erg in audit.get("analysen", {}).items():
                with st.expander(f"Konformitätsnachweis Abschnitt {nr} {erg.get('titel', '')}"):
                    fall = erg.get("fall", {})
                    st.markdown(f"**Hauptaussage** {fall.get('hauptaussage', '')}")
                    for t in fall.get("teilaussagen", []):
                        st.markdown(f"**{t.get('id')}** {t.get('aussage')}")
                        st.caption(f"Prüfpunkt · {t.get('pruefpunkt', '')}")
                        nachweise_anzeigen(t.get("nachweise"))
                        for a in t.get("annahmen", []):
                            st.markdown(f"Annahme ohne Beleg · {a}")
                    if st.button(f"Abschnitt {nr} erneut prüfen", key=f"rm_{nr}"):
                        del audit["analysen"][nr]
                        st.rerun()

    # ---------------- 3 Urteil
    with t3:
        einwaende = alle_einwaende(audit)
        if not einwaende:
            st.info("Noch keine Zweifel. Führen Sie zuerst die Prüfung durch.")
        else:
            erw = {x.get("einwand"): x for erg in audit.get("analysen", {}).values()
                   for x in erg.get("erwiderungen", {}).get("erwiderungen", [])}
            st.caption("Hier endet die Automatik. Sie entscheiden über jeden Zweifel und "
                       "begründen das. Diese Begründung ist Teil des Auditnachweises.")
            nur_offen = st.checkbox("Nur noch unbeurteilte anzeigen", value=False)
            for e in einwaende:
                eid = e["id"]
                alt = audit.get("urteile", {}).get(eid, {})
                if nur_offen and alt.get("urteil"):
                    continue
                with st.container(border=True):
                    st.markdown(f"#### {eid} · {TYPEN.get(e.get('typ'), e.get('typ'))} "
                                f"· Schwere {e.get('schwere')}")
                    st.caption(f"Normabschnitt {e.get('abschnitt')} · "
                               f"richtet sich gegen {e.get('ziel')}")
                    links, rechts = st.columns(2)
                    with links:
                        st.markdown("**Zweifel des Falsifikationsteams**")
                        st.write(e.get("begruendung"))
                        nachweise_anzeigen(e.get("nachweise"))
                    with rechts:
                        st.markdown("**Erwiderung des Konstruktionsteams**")
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
                                        key=f"b_{eid}", height=70)
                    audit.setdefault("urteile", {})[eid] = {"urteil": urteil,
                                                            "begruendung": begr, "einwand": e}
            fehlend = [i for i, u in audit.get("urteile", {}).items()
                       if not (u.get("urteil") and u.get("begruendung", "").strip())]
            if fehlend:
                st.warning("Ohne begründetes Urteil: " + ", ".join(fehlend))
            if st.button("Urteile festschreiben, Bericht und Massnahmen erzeugen",
                         type="primary", disabled=bool(fehlend), use_container_width=True):
                relevant = [u for u in audit["urteile"].values() if u["urteil"] != "Ausgeräumt"]
                try:
                    with st.status("Bericht und Massnahmen werden erstellt", expanded=True) as stt:
                        st.write("Feststellungen werden formuliert")
                        audit["feststellungen"] = llm_json(
                            MODEL_PRO, BERICHT_SYSTEM,
                            f"Auditkriterien {audit['norm']}, Abschnitte "
                            f"{', '.join(audit['kriterien'])}\n\nEntscheidungen des Auditors\n"
                            f"{json.dumps(relevant, ensure_ascii=False)}", audit)
                        fs = audit["feststellungen"].get("feststellungen", [])
                        if fs:
                            st.write("Massnahmenteam leitet Verbesserungen ab")
                            vor = llm_json(MODEL_CONTRA, MASSNAHMEN_SYSTEM,
                                           "Bestätigte Feststellungen\n"
                                           f"{json.dumps(fs, ensure_ascii=False)}", audit)
                            vorhanden = {m.get("feststellung") for m in audit.get("massnahmen", [])}
                            for m in vor.get("massnahmen", []):
                                if m.get("feststellung") in vorhanden:
                                    continue
                                try:
                                    tage = int(m.get("frist_tage", 30))
                                except (TypeError, ValueError):
                                    tage = 30
                                audit.setdefault("massnahmen", []).append({
                                    "id": uuid.uuid4().hex[:6],
                                    "feststellung": m.get("feststellung", ""),
                                    "beschreibung": m.get("korrekturmassnahme", ""),
                                    "sofortmassnahme": m.get("sofortmassnahme", ""),
                                    "ursache_hypothese": m.get("ursache_hypothese", ""),
                                    "wirksamkeitsnachweis": m.get("wirksamkeitsnachweis", ""),
                                    "verantwortlich": m.get("verantwortlich_rolle", ""),
                                    "termin": str(date.today() + timedelta(days=tage)),
                                    "status": "offen", "audit": audit["id"]})
                        stt.update(label="Fertig, weiter im Reiter Bericht", state="complete")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))

    # ---------------- 4 Bericht und Massnahmen
    with t4:
        if not audit.get("feststellungen"):
            st.info("Der Bericht entsteht, sobald Sie alle Zweifel beurteilt haben.")
        else:
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Abweichungen", k["abweichungen"])
            s2.metric("Potenziale", k["potenziale"])
            s3.metric("Vor Ort offen", k["offen"])
            s4.metric("Ausgeräumt", k["ausgeraeumt"])
            audit["fazit"] = st.text_area(
                "Fazit des Auditors (erscheint im Bericht)", value=audit.get("fazit", ""),
                height=110, help="Ihre Gesamteinschätzung. Die Maschine schreibt sie nicht.")
            st.divider()
            st.markdown("**Massnahmenplan, vom System vorgeschlagen und von Ihnen zu bestätigen**")
            st.caption("Die Ursachen sind Hypothesen. Prüfen Sie sie, bevor Sie die Massnahme "
                       "freigeben, und tragen Sie die verantwortliche Stelle ein.")
            if audit.get("massnahmen"):
                tabelle = st.data_editor(
                    audit["massnahmen"], num_rows="dynamic", use_container_width=True,
                    key=f"me_{audit['id']}",
                    column_config={
                        "id": None, "audit": None,
                        "feststellung": st.column_config.TextColumn("Feststellung", width="small"),
                        "beschreibung": st.column_config.TextColumn("Korrekturmassnahme",
                                                                    width="large"),
                        "sofortmassnahme": st.column_config.TextColumn("Sofortmassnahme"),
                        "ursache_hypothese": st.column_config.TextColumn("Ursachenhypothese"),
                        "wirksamkeitsnachweis": st.column_config.TextColumn("Wirksamkeitsnachweis"),
                        "verantwortlich": st.column_config.TextColumn("Verantwortlich"),
                        "termin": st.column_config.TextColumn("Termin", width="small"),
                        "status": st.column_config.SelectboxColumn("Status", options=MSTATUS)})
                if st.button("Massnahmen übernehmen"):
                    for zeile in tabelle:
                        zeile.setdefault("id", uuid.uuid4().hex[:6])
                        zeile.setdefault("audit", audit["id"])
                        zeile.setdefault("status", "offen")
                    audit["massnahmen"] = list(tabelle)
                    st.success("Gespeichert.")
            else:
                st.caption("Keine Massnahmen, weil keine Feststellung bestätigt wurde.")
            st.divider()
            s1, s2, s3 = st.columns(3)
            s1.download_button("Bericht als Webseite (HTML)", bericht_html(audit),
                               file_name=f"auditbericht_{audit['id']}.html",
                               mime="text/html", use_container_width=True)
            s2.download_button("Bericht als Text (Markdown)", bericht_markdown(audit),
                               file_name=f"auditbericht_{audit['id']}.md",
                               use_container_width=True)
            trail = {kk: vv for kk, vv in audit.items() if kk != "dokumente"}
            trail["nachweise"] = list(audit.get("dokumente", {}))
            s3.download_button("Audit Trail (JSON)",
                               json.dumps(trail, ensure_ascii=False, indent=2),
                               file_name=f"audit_trail_{audit['id']}.json",
                               use_container_width=True)
            st.caption("Die HTML-Fassung lässt sich im Browser über Drucken als PDF speichern.")
            with st.expander("Bericht ansehen"):
                st.markdown(bericht_markdown(audit))

# ---------------------------------------------------------------- 3 Massnahmen
else:
    st.subheader("Massnahmen über alle Audits")
    zeilen = [{"Audit": a["titel"], "Feststellung": m.get("feststellung", ""),
               "Massnahme": m.get("beschreibung", ""),
               "Verantwortlich": m.get("verantwortlich", ""), "Termin": m.get("termin", ""),
               "Status": m.get("status", ""),
               "überfällig": "ja" if (m.get("status") in ("offen", "in Umsetzung")
                                      and str(m.get("termin", "")) < str(date.today())) else ""}
              for a in audits for m in a.get("massnahmen", [])]
    if not zeilen:
        st.info("Noch keine Massnahmen. Sie entstehen im Einzelaudit nach Ihrem Urteil.")
    else:
        s1, s2, s3 = st.columns(3)
        s1.metric("Massnahmen gesamt", len(zeilen))
        s2.metric("Offen", sum(1 for z in zeilen if z["Status"] in ("offen", "in Umsetzung")))
        s3.metric("Überfällig", sum(1 for z in zeilen if z["überfällig"]))
        st.dataframe(zeilen, use_container_width=True)
        st.caption("Status ändern Sie im jeweiligen Audit im Reiter Bericht und Massnahmen.")
