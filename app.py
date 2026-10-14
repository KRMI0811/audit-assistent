"""
Audit-Assistent Qualitaetsmanagement (Webversion)
Internes Audit nach ISO 9001 entlang des PDCA-Zyklus.

Plan   Auditprogramm und Auditvorbereitung
Do     Sammeln objektiver Nachweise, Gespraech und Beobachtung vor Ort
Check  Dialektische Pruefung durch zwei Agententeams, Urteil des Auditors, Bericht
Act    Korrekturmassnahmen und Nachverfolgung der Wirksamkeit

Die Agenten bereiten vor und pruefen. Ueber jede Feststellung entscheidet der Auditor.
"""

import base64
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
APP_NAME = st.secrets.get("APP_NAME", "Audix")
APP_CLAIM = st.secrets.get("APP_CLAIM", "Dein KI-Assistent für das interne Audit")
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
AVATAR = Path(__file__).parent / "audix.png"
# Ordner für eigene Figuren der Teams. Liegt dort eine Datei mit passendem Namen,
# etwa team_pro.png oder tiefe_standard.png, wird sie statt der gezeichneten
# Ersatzgrafik verwendet. Fehlt der Ordner, läuft alles unverändert weiter.
BILDER = Path(__file__).parent / "bilder"

st.set_page_config(page_title=APP_NAME, page_icon=str(AVATAR) if AVATAR.exists() else None,
                   layout="wide")


AUDIX_START = """Hallo, ich bin **Audix** und begleite dich durch dein internes Audit.

Mein Trick ist, dass ich nicht nur eine Meinung habe, sondern zwei. Ein Team von mir baut aus
deinen Unterlagen den bestmöglichen Nachweis, dass eine Normforderung erfüllt ist. Ein zweites
Team greift genau diesen Nachweis an und sucht nach Lücken. Was danach an Zweifeln übrig bleibt,
lege ich dir vor. **Entscheiden musst du**, denn nach ISO 19011 trägt der Auditor die
Verantwortung für jede Feststellung, nicht die Maschine.

Aus diesen Zweifeln baue ich dir ausserdem den **Gesprächsleitfaden für den Termin vor Ort**.
Denn ein Dokument zeigt mir nur, was beschrieben ist. Ob danach auch gearbeitet wird, siehst
allein du.

**So läuft es ab**

1. Im **Auditprogramm** legst du die Audits des Jahres an.
2. Im **Einzelaudit** führst du eines davon durch. Dieser Teil folgt dem PDCA-Zyklus, also
   also Plan, Do, Check und Act. Jede Phase hat ihre eigenen Aufgaben.
3. Unter **Massnahmen** verfolgst du über alle Audits hinweg, ob die vereinbarten
   Massnahmen wirksam waren.

Oben rechts findest du bei jeder Aufgabe den Knopf **Audix fragen**. Dort erkläre ich dir, was
an dieser Stelle zu tun ist, und dort kannst du mir auch direkt Fragen stellen. Ich kenne deinen
Arbeitsstand, du musst mir also nicht erklären, wo du gerade stehst. Fangen wir an."""

AUDIX_HILFE = {
    "programm": """**Hier fängt alles an.**

ISO 9001 Abschnitt 9.2 verlangt ein Auditprogramm, das Häufigkeit, Methoden und
Verantwortlichkeiten festlegt und sich nach Bedeutung und Risiko der Prozesse sowie nach
früheren Ergebnissen richtet.

**Was du jetzt tust**
- Lege für jeden Prozess, den du dieses Jahr prüfst, ein Audit an.
- Stufe die Bedeutung und das Risiko ein, denn danach richtet sich die Prüftiefe.
- Die Normabschnitte kannst du gleich wählen oder leer lassen, dann schlage ich sie dir später
  anhand deiner Unterlagen vor.

**Worauf ich achte**
Ohne bestätigte Unparteilichkeit lasse ich dich kein Audit anlegen. Niemand darf die eigene
Arbeit auditieren, das ist eine harte Normforderung. Unten zeige ich dir ausserdem, welche
Normabschnitte im Jahresprogramm noch gar nicht vorkommen.""",

    "vorbereitung": """**Plan. Erst der Plan, dann das Team.**

**Was du jetzt tust**
- Fülle den Auditplan aus. Daraus wird die Einladung an den Fachbereich.
- Stelle dein Auditteam zusammen. Unten legst du fest, wie viele Falsifikationsteams angreifen
  und ob das Konstruktionsteam erwidern darf. Zwei Teams kannst du nicht abwählen, denn ohne
  Angriff auf den Nachweis bleibt nur eine Meinung übrig.
- Klapp ein Team auf, gib ihm einen Namen und wähle ein Muster für das Rollenverständnis. Den
  Mustertext kannst du danach frei überschreiben.

**Wann ein zweites Falsifikationsteam lohnt**
Wenn du zwei sehr verschiedene Blickwinkel brauchst, etwa eines auf die Belastbarkeit der
Unterlagen und eines auf die gelebte Praxis. Gib jedem dann wirklich einen eigenen Schwerpunkt.
Ohne eigenen Schwerpunkt findet das zweite Team dasselbe wie das erste, und du bezahlst nur
mit Wartezeit.
- Schau dir unten die gewählten Normabschnitte an. Zu jedem zeige ich dir, was die Norm
  verlangt, worauf du achten solltest und welche Unterlagen du anfordern musst.

**Mein Tipp**
Lade dir den Auditplan mit Checkliste herunter und schicke ihn vorab an den Fachbereich. Dann
liegen die richtigen Unterlagen schon bereit, wenn du kommst.""",

    "durchfuehrung": """**Plan. Jetzt wird gesammelt.**

**Was du jetzt tust**
- Lade die Dokumente hoch, also Verfahrensanweisungen, Listen, Protokolle, Formulare.
- Erfasse deine Notizen aus Gesprächen und Beobachtungen, möglichst wörtlich.
- Trage Leistungsdaten ein, etwa Reklamationen, Ausschuss oder Termintreue.

**Warum das wichtig ist**
Dokumente zeigen mir nur die Vorschrift. Erst wenn du mir erzählst, was die Leute sagen und was
du gesehen hast, finde ich die interessanten Widersprüche zwischen Vorgabe und gelebter Praxis.
Genau dort stecken die Feststellungen, die wirklich etwas verändern.

**Achtung**
Wenn du nach einer Prüfung noch Notizen ergänzt, musst du den betreffenden Abschnitt erneut
prüfen lassen. Sonst kenne ich sie nicht.""",

    "pruefung": """**Do. Jetzt arbeiten meine Teams gegeneinander.**

**Was passiert, wenn du auf Prüfung starten klickst**
1. Mein Vorbereitungsteam wählt die Normabschnitte, falls du das noch nicht getan hast.
2. Das Konstruktionsteam baut je Abschnitt den bestmöglichen Konformitätsnachweis.
3. Jedes eingesetzte Falsifikationsteam greift diesen Nachweis an.
4. Das Konstruktionsteam darf einmal erwidern, sofern du die Erwiderungsrunde eingeschaltet hast.

Oben siehst du die Besetzung und kannst sie hier noch ändern. Konstruktionsteam und
Falsifikationsteam laufen immer, alles andere entscheidest du. Die Anzeige sagt dir vorher,
wie viele Modellaufrufe der Lauf kostet.

Das dauert einige Minuten. Du kannst den Fortschritt am Balken mitverfolgen.

**Wie du die Ergebnisse liest**
Jeder Prüfpunkt bekommt eine Ampel. Grün heisst, es gibt einen belegten Nachweis und keinen
schweren Zweifel. Gelb heisst, ein Nachweis ist da, aber es steht ein Einwand dagegen. Rot
heisst, es fehlt ein belastbarer Nachweis.

**Wichtig**
Die Ampel ist mein Befund, nicht deine Feststellung. Klapp jeden Prüfpunkt auf, dann siehst du
die ganze Kette von der Normforderung über die Zitate bis zur Prognose fürs externe Audit.""",

    "vorort": """**Do. Papier ist geduldig, der Betrieb ist es nicht.**

**Warum es diesen Schritt gibt**
Bis hierher habe ich nur gelesen. Ich sehe, was beschrieben ist, nicht was getan wird. Ob im
Lager, in der Werkstatt oder am Empfang wirklich so gearbeitet wird, wie es im Handbuch steht,
stellst nur du fest, durch Beobachtung und Gespräch. Genau dafür bereite ich dir hier das
Gespräch vor.

**Was ich dir liefere**
Zu jedem Zweifel aus der Prüfung bekommst du
- eine offene Frage, die nach der Praxis fragt und nicht nach dem Dokument,
- zwei Nachfragen für den Fall, dass die erste Antwort zu glatt kommt,
- eine konkrete Stichprobe, die du im Termin ziehen kannst,
- die Merkmale, an denen du gelebte Praxis von reinem Papier unterscheidest,
- und die Gefährdung, die entsteht, wenn die Vorgabe nicht gelebt wird.

**Was du jetzt tust**
- Erzeuge die Fragen und lade dir den Leitfaden als Word-Datei herunter. Nimm ihn mit.
- Trage nach dem Termin ein, was du gesehen und gehört hast, möglichst wörtlich.
- Setze das Ergebnis je Frage. Das erscheint danach beim passenden Zweifel unter Check bei der Aufgabe
  Urteil und geht als objektiver Nachweis in den Bericht.

**Mein Tipp**
Lies die Antwort nicht vor und lass sie dir nicht vorformulieren. Der Wert dieses Schrittes
steckt in dem, was dir jemand unaufgefordert erzählt, und in dem, was du mit eigenen Augen siehst.

**Normbezug**
ISO 19011 verlangt objektive Nachweise aus Befragung, Beobachtung und Durchsicht von
Aufzeichnungen. Eine Dokumentenprüfung allein deckt nur einen Teil davon ab.""",

    "urteil": """**Hier endet meine Arbeit und deine beginnt.**

**Was du jetzt tust**
Du entscheidest über jeden Zweifel und begründest das. Vier Urteile stehen dir zur Verfügung.

- **Ausgeräumt** · Der Einwand trifft nicht zu, die Unterlagen reichen aus.
- **Bestätigt: Abweichung** · Eine Anforderung ist nicht erfüllt.
- **Bestätigt: Verbesserungspotenzial** · Kein Normverstoss, aber eine Schwäche.
- **Offen: vor Ort prüfen** · Aus den Unterlagen allein nicht entscheidbar.

**Warum die Begründung Pflicht ist**
Sie ist Teil deines Auditnachweises. Ein ausgeräumter Zweifel ist genauso wertvoll wie eine
Abweichung, denn er belegt, wie tief du geprüft hast. Im Bericht erscheinen beide.

Wenn du fertig bist, schreibe ich dir daraus die Feststellungen und schlage Massnahmen vor.""",

    "bericht": """**Check. Der Bericht geht an die Leitung.**

**Was du hier vorfindest**
Ein Formular nach dem Aufbau eines Auditberichts mit allen Kapiteln. Die meisten Felder habe
ich aus deinen Daten vorbelegt, etwa die eingereichten Unterlagen, die Erfüllungsübersicht und
die Hinweise. Bei sieben Feldern weiss ich nichts, dort steht, was hineingehört.

**Was du jetzt tust**
- Geh die Felder durch und überschreibe, was nicht passt.
- Schreib dein Fazit und die Begründung in 6.2 selbst. Das ist deine Kernaussage als Auditor,
  die kann ich dir nicht abnehmen.
- Lade den Bericht als Word-Datei herunter und arbeite dort weiter.

**Normbezug**
ISO 9001 verlangt, dass die Ergebnisse an die zuständige Leitung berichtet werden und dass der
Bericht als dokumentierte Information aufbewahrt wird, also Abschnitte 9.2 und 7.5.""",

    "massnahmen": """**Act. Ein Audit ist erst dann etwas wert, wenn Fehler behoben werden.**

**Was ich vorbereitet habe**
Zu jeder bestätigten Abweichung habe ich eine Ursachenhypothese, eine Sofortmassnahme, eine
Korrekturmassnahme, eine verantwortliche Rolle, eine Frist und einen Wirksamkeitsnachweis
vorgeschlagen. Das ist die Struktur nach ISO 9001 Abschnitt 10.2.

**Was du jetzt tust**
- Prüfe meine Ursachen mit den Prozessverantwortlichen. Es sind Hypothesen, keine Analysen.
  Eine Ursachenanalyse ohne Gespräch mit den Betroffenen ist nicht belastbar.
- Setze die richtige Person oder Funktion als Verantwortung ein.
- Passe die Fristen an eure Realität an.

Vergiss den Unterschied nicht. Die Sofortmassnahme beseitigt die Folge, die Korrekturmassnahme
die Ursache. Nur die zweite verhindert, dass der Fehler wiederkommt.""",

    "nachverfolgung": """**Act. Hier schliesst sich der Kreis.**

ISO 9001 verlangt, dass du die Wirksamkeit der ergriffenen Massnahmen prüfst und das
dokumentierst. Genau das machst du hier.

**Was du jetzt tust**
- Behalte die offenen Massnahmen im Blick, überfällige markiere ich dir.
- Wähle unten eine Massnahme aus und halte fest, was du geprüft hast und mit welchem Ergebnis.
- Steht eine Massnahme auf nicht wirksam, dann war die Ursache falsch bestimmt. Dann beginnt
  die Analyse von vorn.

**Mein Tipp**
Nimm diese Übersicht in die Managementbewertung mit. Die Frage, wie viele Massnahmen fristgerecht
und wirksam abgeschlossen wurden, sagt mehr über euer QM-System aus als die Zahl der
Abweichungen."""}


AUDIX_CHAT_SYSTEM = """Du bist Audix, ein freundlicher und fachlich präziser Begleiter für
interne Audits nach ISO 9001. Du sprichst die auditierende Person mit Du an.

Was du kannst
- Normanforderungen von ISO 9001 und Auditmethodik nach ISO 19011 erklären.
- Erklären, wie diese Anwendung funktioniert und was an der aktuellen Stelle zu tun ist.
- Beim Formulieren von Feststellungen, Auditfragen und Korrekturmassnahmen helfen.
- Einschätzen, welche Nachweise für eine Anforderung üblicherweise gebraucht werden.

Deine Regeln
1. Du triffst keine Auditfeststellungen und stufst nichts ein. Das ist Sache des Auditors.
   Wenn jemand das von dir verlangt, sagst du freundlich, dass du nur vorbereiten kannst.
2. Du erfindest keine Normzitate und keine Abschnittsnummern. Wenn du unsicher bist, sagst du
   das offen und verweist auf den lizenzierten Normtext.
3. Du antwortest kurz, höchstens fünf Sätze oder eine kurze Liste. Keine langen Vorreden.
4. Du beziehst dich auf den Arbeitsstand, der dir im Kontext mitgegeben wird, wenn er zur Frage
   passt. Erfinde keine Daten, die dort nicht stehen.
5. Du schreibst auf Deutsch in Schweizer Rechtschreibung, also ss statt ß."""


def llm_text(modell: str, system: str, verlauf: list) -> str:
    """Freie Textantwort fuer den Chat. Gibt bei Fehlern eine lesbare Meldung zurueck."""
    nutzlast = {"model": modell, "temperature": 0.4, "max_tokens": 700, "stream": False,
                "messages": [{"role": "system", "content": system}] + verlauf}
    kopf = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    try:
        r = requests.post(f"{API_BASE}/chat/completions", headers=kopf, json=nutzlast,
                          timeout=(15, 180))
    except requests.exceptions.Timeout:
        return "Da habe ich zu lange gebraucht. Stell die Frage bitte noch einmal."
    except requests.RequestException as f:
        return f"Ich komme gerade nicht an mein Sprachmodell heran ({f})."
    if r.status_code != 200:
        return (f"Mein Sprachmodell antwortet nicht (HTTP {r.status_code}). "
                "Bitte später erneut versuchen.")
    inhalt = (r.json().get("choices", [{}])[0].get("message", {}).get("content") or "").strip()
    return inhalt or "Dazu fällt mir gerade nichts ein. Frag mich bitte anders."


def audix_kontext(schluessel: str) -> str:
    """Beschreibt Audix, wo der Auditor gerade steht und woran er arbeitet."""
    phasen = {"programm": "Auditprogramm des Jahres (Phase Plan)",
              "vorbereitung": "Auditvorbereitung und Auditteam (Phase Plan)",
              "durchfuehrung": "Nachweise zusammenstellen (Phase Plan)",
              "pruefung": "Dialektische Prüfung durch die Agententeams (Phase Do)",
              "vorort": "Gesprächsleitfaden und Beobachtungen vor Ort (Phase Do)",
              "urteil": "Urteil des Auditors über die Zweifel (Phase Check)",
              "bericht": "Auditbericht (Phase Check)",
              "massnahmen": "Korrekturmassnahmen (Phase Act)",
              "nachverfolgung": "Nachverfolgung der Wirksamkeit (Phase Act)"}
    zeilen = [f"Die Person befindet sich hier: {phasen.get(schluessel, schluessel)}."]
    a = st.session_state.get("aktuelles_audit")
    if isinstance(a, dict):
        kz = kennzahlen(a)
        def mz(anzahl, eins, viele):
            return f"{anzahl} {eins if anzahl == 1 else viele}"

        zeilen += [
            f"Aktuelles Audit: {a.get('titel', '')}, Prozess {a.get('prozess', '')}.",
            f"Norm {a.get('norm', '')}, geprüfte Abschnitte "
            f"{', '.join(a.get('kriterien', [])) or 'noch keine'}.",
            "Stand: " + ", ".join([
                mz(len(a.get("dokumente", {})), "Dokument", "Dokumente"),
                mz(len(a.get("notizen", [])), "Notiz", "Notizen"),
                mz(kz["abschnitte"], "analysierter Abschnitt", "analysierte Abschnitte"),
                mz(kz["einwaende"], "Zweifel", "Zweifel") + f", davon {kz['beurteilt']} beurteilt"]) + ".",
            f"Bestätigt sind {mz(kz['abweichungen'], 'Abweichung', 'Abweichungen')} und "
            f"{mz(kz['potenziale'], 'Verbesserungspotenzial', 'Verbesserungspotenziale')}, "
            f"{kz['offen']} Punkte bleiben für die Prüfung vor Ort offen.",
            f"Leitfaden vor Ort: {kz['fragen']} Fragen vorbereitet, "
            f"{kz['fragen_geprueft']} davon vor Ort bearbeitet, bei {kz['nicht_gelebt']} "
            "wurde die Praxis nicht oder nur teilweise bestätigt."]
        unbeurteilt = [e for e in alle_einwaende(a)
                       if not a.get("urteile", {}).get(e["id"], {}).get("urteil")]
        if unbeurteilt:
            zeilen.append("Noch unbeurteilte Zweifel: " + "; ".join(
                f"{e['id']} zu Abschnitt {e.get('abschnitt', '')}: "
                f"{str(e.get('begruendung', ''))[:180]}" for e in unbeurteilt[:5]))
    else:
        zeilen.append("Es ist noch kein einzelnes Audit ausgewählt.")
    return "\n".join(zeilen)


VORSCHLAEGE = {
    "programm": ["Wie oft muss ich welchen Prozess auditieren?",
                 "Was gehört in ein Auditprogramm?"],
    "vorbereitung": ["Welche Unterlagen soll ich anfordern?",
                     "Wie formuliere ich gute Auditfragen?"],
    "durchfuehrung": ["Was frage ich im Eröffnungsgespräch?",
                      "Was ist ein objektiver Nachweis?"],
    "pruefung": ["Was bedeutet die gelbe Ampel?",
                 "Worin unterscheiden sich die Einwandtypen?"],
    "vorort": ["Woran erkenne ich, ob ein Prozess wirklich gelebt wird?",
               "Wie ziehe ich eine sinnvolle Stichprobe?"],
    "urteil": ["Abweichung oder Verbesserungspotenzial, wie entscheide ich?",
               "Wie formuliere ich eine Feststellung sauber?"],
    "bericht": ["Was gehört in die Begründung in 6.2?",
                "Wesentlich oder geringfügig, was ist der Unterschied?"],
    "massnahmen": ["Wie finde ich die wirkliche Ursache?",
                   "Korrektur oder Korrekturmassnahme, was ist was?"],
    "nachverfolgung": ["Wie prüfe ich die Wirksamkeit einer Massnahme?",
                       "Was gehört davon in die Managementbewertung?"]}


def audix_antworten(frage: str, schluessel: str):
    """Haengt Frage und Antwort an den Chatverlauf dieser Phase."""
    verlauf = st.session_state.setdefault("audix_chat", {}).setdefault(schluessel, [])
    verlauf.append({"role": "user", "content": frage})
    nachrichten = [{"role": "system", "content": "Arbeitsstand\n" + audix_kontext(schluessel)}]
    nachrichten += verlauf[-8:]
    with st.spinner("Audix denkt nach"):
        antwort = llm_text(MODEL_PRO, AUDIX_CHAT_SYSTEM, nachrichten)
    verlauf.append({"role": "assistant", "content": antwort})


def audix_hilfe(schluessel: str, knopftext: str = "Audix fragen"):
    """Aufklappbarer Begleiter mit Erklaerung zur Phase und Chat."""
    with st.popover(knopftext, use_container_width=False):
        bild, inhalt = st.columns([1, 4])
        with bild:
            if AVATAR.exists():
                st.image(str(AVATAR), use_container_width=True)
        with inhalt:
            st.markdown(f"#### {APP_NAME}")
            st.caption(APP_CLAIM)
        erklaerung, chat = st.tabs(["Was hier zu tun ist", "Audix fragen"])
        with erklaerung:
            st.markdown(AUDIX_HILFE.get(schluessel, ""))
        with chat:
            verlauf = st.session_state.get("audix_chat", {}).get(schluessel, [])
            if not verlauf:
                st.caption("Frag mich alles zur Norm, zur Auditmethodik oder zu dieser "
                           "Anwendung. Ich kenne deinen aktuellen Arbeitsstand.")
                for i, vorschlag in enumerate(VORSCHLAEGE.get(schluessel, [])):
                    if st.button(vorschlag, key=f"vs_{schluessel}_{i}",
                                 use_container_width=True):
                        audix_antworten(vorschlag, schluessel)
                        st.rerun()
            for n in verlauf:
                with st.chat_message("user" if n["role"] == "user" else "assistant",
                                     avatar=(str(AVATAR) if n["role"] == "assistant"
                                             and AVATAR.exists() else None)):
                    st.markdown(n["content"])
            frage = st.text_area("Deine Frage", key=f"frage_{schluessel}", height=70,
                                 placeholder="Zum Beispiel: Welche Nachweise brauche ich "
                                             "für Abschnitt 7.2?", label_visibility="collapsed")
            s1, s2 = st.columns([3, 1])
            if s1.button("Fragen", key=f"senden_{schluessel}", type="primary",
                         use_container_width=True, disabled=not API_KEY):
                if frage.strip():
                    audix_antworten(frage.strip(), schluessel)
                    st.rerun()
            if verlauf and s2.button("Neu", key=f"reset_{schluessel}",
                                     use_container_width=True):
                st.session_state["audix_chat"][schluessel] = []
                st.rerun()
            if not API_KEY:
                st.caption("Ohne hinterlegten API-Schlüssel kann ich nicht antworten.")


@st.dialog("Willkommen bei Audix", width="large")
def audix_begruessung():
    bild, inhalt = st.columns([1, 3])
    with bild:
        if AVATAR.exists():
            st.image(str(AVATAR), use_container_width=True)
    with inhalt:
        st.markdown(AUDIX_START)
    if st.button("Los geht es", type="primary", use_container_width=True):
        st.session_state.begruesst = True
        st.rerun()


HINWEISE = f"""
- Audix bereitet vor und prüft, er entscheidet nicht. Jede Feststellung und ihre
  Einstufung verantwortet der Auditor.
- Ihre Unterlagen werden zur Auswertung an den Dienst {PROVIDER or 'des eingestellten Anbieters'}
  übermittelt. Laden Sie keine vertraulichen Originalunterlagen und keine Personendaten hoch,
  solange das mit Ihrer IT und dem Datenschutz nicht geklärt ist.
- Jedes Zitat wird gegen die Quelle geprüft. Ein nicht auffindbares Zitat ist ein Warnzeichen
  und darf nicht in einen Bericht übernommen werden.
- Im Chat von Audix können Fragen gestellt werden. Auch dort gilt, dass seine Auskünfte den
  lizenzierten Normtext nicht ersetzen und keine Auditfeststellung darstellen.
- Der Arbeitsstand liegt nur in dieser Sitzung. Sichern Sie ihn links als Datei.
"""

LEGENDE = """
**Audix erklärt die Ampel**

🟢 **belegt** · Zu diesem Prüfpunkt gibt es einen Nachweis, dessen Zitat wörtlich in Ihren
Unterlagen wiedergefunden wurde, und es steht kein schwerer Zweifel dagegen.

🟡 **mit Zweifel** · Es gibt einen Nachweis, aber das Falsifikationsteam hat einen Einwand
dagegen. Typisch ist, dass das Dokument zwar existiert, die Anforderung inhaltlich aber nicht
abdeckt.

🔴 **ohne Nachweis** · Es wurde kein belastbarer Nachweis gefunden, oder ein Nachweis spricht
gegen die Erfüllung. Das sind die Stellen, an denen Abweichungen entstehen.

**Was die Begriffe bedeuten**

*Zitat belegt* heisst nur, dass der zitierte Satz wirklich so in Ihrem Dokument steht. Es heisst
nicht, dass die Normanforderung erfüllt ist. Das entscheiden Sie.

*Zitat nicht auffindbar* heisst, dass die KI einen Satz zitiert hat, den es in Ihren Unterlagen
so nicht gibt. Solche Nachweise sind wertlos und dürfen nie in den Bericht.

*Behauptung ohne Nachweis* heisst, dass das Konstruktionsteam etwas annehmen musste, weil in den
Unterlagen nichts dazu steht. Genau das sind die Punkte, die Sie vor Ort nachfragen sollten.
"""

# ---------------------------------------------------------------------------
# Rollen der Agenten
# ---------------------------------------------------------------------------
SPRACHREGEL = ("\n\nSPRACHE: Schreibe alle Inhalte ausschliesslich auf Deutsch in Schweizer "
               "Rechtschreibung (ss statt ß). Zitate übernimmst du wörtlich.")
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
4. Nenne, welche Nachweise für ein belastbares Audit fehlen.
Schema
{"kriterien": ["7.2"], "begruendungen": {"7.2": "..."}, "ziel": "...", "umfang": "...",
 "fehlende_nachweise": ["..."]}"""

PRO_SYSTEM = """Du bist das KONSTRUKTIONSTEAM eines internen Audits.
Baue einen Konformitätsnachweis als Assurance Case (Aussage, Argument, Nachweis).
Bilde für JEDEN Prüfpunkt des Auditkriteriums genau eine Teilaussage und übernimm den
Wortlaut des Prüfpunkts unverändert in das Feld pruefpunkt.
Regeln
1. Nutze ausschliesslich die bereitgestellten Nachweise.
2. Jeder Nachweis enthält den exakten Quellennamen und ein WÖRTLICHES Zitat (höchstens 40 Wörter).
3. Prüfe nicht nur, OB etwas vorliegt, sondern ob der Inhalt die Anforderung inhaltlich abdeckt.
4. Was nicht belegt ist, trägst du als Annahme ein. Verstecke keine Lücken.
5. Schreibe die Aussage in einem kurzen, klaren Satz ohne Fachjargon.
Schema
{"hauptaussage": "...", "teilaussagen": [{"id": "T1", "pruefpunkt": "...", "aussage": "...",
 "argument": "...", "nachweise": [{"dokument": "...", "zitat": "..."}], "annahmen": ["..."]}]}"""

CONTRA_SYSTEM = """Du bist das FALSIFIKATIONSTEAM eines internen Audits.
Du erhältst einen Assurance Case und die Nachweise. Dein Ziel ist, ihn zu widerlegen.
Einwandtypen
"widerlegend": Nachweise belegen das Gegenteil der Aussage.
"untergrabend": Der Nachweis ist unzuverlässig, veraltet, nicht freigegeben oder unvollständig.
"unterlaufend": Der Nachweis stimmt, deckt die Anforderung inhaltlich aber nicht ab.
"ungestuetzte_annahme": Eine Annahme trägt die Aussage, ohne belegt zu sein.
Achte besonders auf Widersprüche zwischen Vorgabe, Aufzeichnung, Aussagen aus Interviews,
Beobachtungen und Leistungsdaten.
Regeln
1. Erfinde nichts. Zitate müssen wörtlich aus den Nachweisen stammen.
2. Fehlt ein Nachweis ganz, lass die Liste nachweise leer und begründe das.
3. Schreibe die Begründung in höchstens drei klaren Sätzen.
4. Schlage für jeden Einwand zwei bis drei konkrete Prüfungen vor Ort vor.
5. Schätze das Risiko im externen Zertifizierungsaudit ein. Hauptabweichung bedeutet, dass eine
   Normanforderung systematisch oder vollständig nicht erfüllt ist. Nebenabweichung bedeutet
   einen Einzelfall oder eine formale Lücke. Kein Risiko bedeutet, dass ein externer Auditor
   dies nicht beanstanden würde.
Schema
{"einwaende": [{"id": "E1", "ziel": "T1", "typ": "...", "begruendung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}], "schwere": "hoch|mittel|gering",
 "pruefung_vor_ort": ["...", "..."],
 "zertifizierungsrisiko": {"klassifizierung": "Hauptabweichung|Nebenabweichung|kein Risiko",
 "prognose": "...", "konsequenz": "...", "dringlichkeit": "sehr hoch|hoch|mittel|gering"}}]}"""

ERWIDERUNG_SYSTEM = """Du bist das KONSTRUKTIONSTEAM. Erwidere auf jeden Einwand sachlich.
Gestehe Einwände zu, wenn die Nachweise sie stützen. Nutze nur wörtliche Zitate.
Schema
{"erwiderungen": [{"einwand": "E1", "zugestanden": true, "erwiderung": "...",
 "nachweise": [{"dokument": "...", "zitat": "..."}]}]}"""

BERICHT_SYSTEM = """Du formulierst Auditfeststellungen nach ISO 19011, also Anforderung mit
Normabschnitt, objektiver Nachweis und Feststellung. Die Einstufung hat der Auditor bereits
getroffen, du darfst sie NICHT ändern. Formuliere sachlich, prüfbar, ohne Schuldzuweisung
und ohne Namen einzelner Personen.
Schreibe zusätzlich eine Zusammenfassung von höchstens acht Sätzen, die beschreibt, was geprüft
wurde und wo die Schwerpunkte liegen. Bewerte darin nicht.
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

VORORT_SYSTEM = """Du bist das LEITFADENTEAM eines internen Audits.
Die Dokumentenprüfung ist gelaufen. Du erhältst den Konformitätsnachweis, die Einwände des
Falsifikationsteams, die Erwiderungen und, falls vorhanden, die Urteile des Auditors.
Deine Aufgabe ist der Gesprächsleitfaden für den Termin vor Ort.
Eine Unterlage beweist nur, dass etwas beschrieben ist. Ob danach gearbeitet wird, zeigt sich
erst im Gespräch, in der Beobachtung und an der Stichprobe. Genau dorthin zielen deine Fragen.
Regeln
1. Zu jedem Einwand, den der Auditor nicht ausgeräumt hat, mindestens eine Frage. Ergänze
   höchstens zwei weitere Fragen je Abschnitt zu Punkten, die sich aus Unterlagen grundsätzlich
   nicht beantworten lassen, etwa Verständnis, Zuständigkeit im Alltag oder Umgang mit Ausnahmen.
2. Stelle offene Fragen. Keine Ja-Nein-Fragen, keine Suggestivfragen und keine Fragen, welche
   die gewünschte Antwort schon enthalten.
3. Frage nach der Praxis, nicht nach dem Papier. Nicht "Gibt es eine Verfahrensanweisung",
   sondern "Zeigen Sie mir am letzten Fall, wie Sie vorgegangen sind".
4. Nenne zu jeder Frage eine konkrete Stichprobe mit Zeitraum und Umfang, die der Auditor im
   Termin ziehen kann, und was er sich dabei zeigen oder vorführen lassen soll.
5. Beschreibe getrennt, woran man erkennt, dass der Prozess gelebt wird, und woran man erkennt,
   dass er nur auf dem Papier steht.
6. Beschreibe die Gefährdung konkret für diesen Prozess, also was im Betrieb, beim Produkt oder
   beim Kunden schiefgeht, wenn die Vorgabe nicht gelebt wird. Keine allgemeinen Formeln wie
   Imageschaden oder Qualitätsverlust.
7. Nenne als Gesprächspartner eine Rolle oder Funktion, niemals einen Personennamen.
8. Du bewertest nicht, stufst nichts ein und formulierst keine Feststellung.
Schema
{"fragen": [{"id": "F1", "einwand": "E1", "thema": "...", "frage": "...",
 "nachfragen": ["...", "..."], "stichprobe": "...", "beobachten": "...",
 "gelebt_wenn": "...", "papier_wenn": "...", "gefaehrdung": "...", "wen_fragen": "..."}]}"""

AGENTEN = [
    ("vorbereitung", "Vorbereitungsteam", "Plan",
     "Liest die Nachweise, wählt die passenden Normabschnitte, schreibt Auditziel und Umfang "
     "und nennt fehlende Nachweise.", SCOPING_SYSTEM, "Konstruktionsmodell"),
    ("pro", "Konstruktionsteam", "Do",
     "Baut je Normabschnitt den bestmöglichen Konformitätsnachweis und markiert unbelegte "
     "Annahmen offen. Es ist bewusst der Anwalt der Organisation.", PRO_SYSTEM,
     "Konstruktionsmodell"),
    ("contra", "Falsifikationsteam", "Do",
     "Greift diesen Nachweis an und sucht widerlegende, untergrabende und unterlaufende "
     "Nachweise sowie ungestützte Annahmen. Es ist bewusst der Gegenspieler.", CONTRA_SYSTEM,
     "Falsifikationsmodell"),
    ("erwiderung", "Erwiderung des Konstruktionsteams", "Do",
     "Antwortet auf jeden Einwand und gesteht zu, was die Nachweise stützen.",
     ERWIDERUNG_SYSTEM, "Konstruktionsmodell"),
    ("vorort", "Leitfadenteam", "Do",
     "Übersetzt die offenen Zweifel in Fragen für das Gespräch vor Ort, nennt Stichprobe, "
     "Beobachtungsmerkmale und die Gefährdung, wenn die Vorgabe nicht gelebt wird.",
     VORORT_SYSTEM, "Falsifikationsmodell"),
    ("bericht", "Berichtsteam", "Check",
     "Formuliert die Feststellungen aus den Entscheidungen des Auditors. Es darf die "
     "Einstufung nicht ändern.", BERICHT_SYSTEM, "Konstruktionsmodell"),
    ("massnahmen", "Massnahmenteam", "Act",
     "Leitet je bestätigter Feststellung Ursachenhypothese, Sofort- und Korrekturmassnahme, "
     "Verantwortung, Frist und Wirksamkeitsnachweis ab.", MASSNAHMEN_SYSTEM,
     "Falsifikationsmodell")]

AGENT_STANDARD = {s: (n, z) for s, n, _, z, _, _ in AGENTEN}

# Zusaetzliche Falsifikationsteams. Sie arbeiten mit derselben Grundanweisung, aber mit
# eigenem Namen und eigenem Schwerpunkt. Damit laesst sich der Nachweis aus mehreren
# Richtungen angreifen, ohne das dialektische Grundprinzip zu veraendern.
ZUSATZ_CONTRA = [("contra2", "Falsifikationsteam 2"), ("contra3", "Falsifikationsteam 3")]
AGENT_STANDARD.update({s: (n, "Zweite Angriffsrichtung auf denselben Nachweis.")
                       for s, n in ZUSATZ_CONTRA})

# Pflichtbesetzung. Ohne diese beiden Rollen gibt es kein dialektisches Verfahren,
# sondern nur eine Meinung. Sie sind daher nicht abwaehlbar.
PFLICHTTEAMS = ["pro", "contra"]

# Anklickbare Muster fuer das Rollenverstaendnis. Je Team mehrere Profile.
MUSTERROLLEN = {
    "vorbereitung": [
        ("Risikoorientiert", "Du wählst die Normabschnitte nach dem Risiko für Kunde und "
                             "Produkt aus und begründest jede Wahl mit dem möglichen Schaden."),
        ("Nachweisnah", "Du wählst nur Abschnitte, zu denen die vorliegenden Unterlagen "
                        "wirklich etwas hergeben, und benennst Lücken schonungslos.")],
    "pro": [
        ("Erfahrener Prozesskenner", "Du bist ein erfahrener Prozessverantwortlicher und baust "
                                     "den Nachweis so stark, wie die Unterlagen es hergeben. Du "
                                     "beschönigst nichts und markierst jede Annahme offen."),
        ("Zertifizierungserfahren", "Du kennst externe Zertifizierungsaudits und formulierst den "
                                    "Nachweis so, wie ihn ein externer Auditor sehen möchte, also "
                                    "mit Vorgabe, Aufzeichnung und Wirksamkeitsbeleg."),
        ("Am Einzelfall", "Du belegst jede Aussage am konkreten Einzelfall und verweist auf "
                          "Datum, Vorgang und Dokument, nie auf allgemeine Regelungen allein.")],
    "contra": [
        ("Schwerpunkt Dokumentenlage", "Du prüfst vor allem die Belastbarkeit der Unterlagen. "
                                       "Veraltete Stände, fehlende Freigaben, widersprüchliche "
                                       "Fassungen und Aufzeichnungen ohne Datum sind dein Thema."),
        ("Schwerpunkt gelebte Praxis", "Du misstraust jeder Vorgabe, die nur beschrieben ist. Du "
                                       "suchst nach Belegen dafür, dass die Regelung im Alltag "
                                       "tatsächlich angewendet und ihre Wirkung geprüft wird."),
        ("Schwerpunkt Risiko", "Du kommst aus einer stark regulierten Branche und legst "
                               "besonderen Wert auf Rückverfolgbarkeit, Freigaben und den "
                               "Umgang mit Abweichungen und Ausnahmen.")],
    "contra2": [
        ("Schwerpunkt gelebte Praxis", "Du misstraust jeder Vorgabe, die nur beschrieben ist. Du "
                                       "suchst nach Belegen dafür, dass die Regelung im Alltag "
                                       "tatsächlich angewendet und ihre Wirkung geprüft wird."),
        ("Schwerpunkt Schnittstellen", "Du suchst Brüche zwischen Bereichen, also Übergaben ohne "
                                       "klare Zuständigkeit, doppelte Zuständigkeiten und "
                                       "Informationen, die unterwegs verloren gehen.")],
    "contra3": [
        ("Schwerpunkt Kennzahlen", "Du vergleichst Vorgaben mit Leistungsdaten und fragst, ob "
                                   "Reklamationen, Termintreue oder Ausschuss das bestätigen, "
                                   "was die Unterlagen behaupten."),
        ("Schwerpunkt Risiko", "Du kommst aus einer stark regulierten Branche und legst "
                               "besonderen Wert auf Rückverfolgbarkeit, Freigaben und den "
                               "Umgang mit Abweichungen und Ausnahmen.")],
    "erwiderung": [
        ("Selbstkritisch", "Du gestehst jeden Einwand zu, den die Unterlagen stützen, und "
                           "verteidigst nur, was du wörtlich belegen kannst."),
        ("Sachlich abwägend", "Du trennst sauber zwischen formalen Mängeln und inhaltlichen "
                              "Lücken und sagst bei jedem Einwand, welcher Fall vorliegt.")],
    "vorort": [
        ("Gesprächsführung", "Du formulierst Fragen, die zum Erzählen einladen, und vermeidest "
                             "alles, was nach Prüfungssituation klingt."),
        ("Stichprobenlogik", "Du denkst in Stichproben und benennst immer Zeitraum, Umfang und "
                             "Auswahlkriterium, damit die Prüfung nachvollziehbar bleibt.")],
    "bericht": [
        ("Nüchtern und prüfbar", "Du schreibst knapp, sachlich und ohne Wertung, jede "
                                 "Feststellung mit Anforderung, Nachweis und Abweichung."),
        ("Adressat Leitung", "Du schreibst so, dass die Geschäftsleitung ohne Fachjargon "
                             "versteht, worum es geht und was auf dem Spiel steht.")],
    "massnahmen": [
        ("Ursachenorientiert", "Du bohrst bei der Ursache nach und gibst dich nicht mit "
                               "Schulung oder Sensibilisierung als Korrekturmassnahme zufrieden."),
        ("Umsetzbar", "Du schlägst Massnahmen vor, die eine kleine Organisation mit knappen "
                      "Ressourcen auch wirklich umsetzen kann.")]}


def agent_name(audit: dict, schluessel: str) -> str:
    eigen = (audit.get("agenten") or {}).get(schluessel, {}).get("name", "").strip()
    return eigen or AGENT_STANDARD.get(schluessel, ("", ""))[0]


def agent_rolle(audit: dict, schluessel: str) -> str:
    return (audit.get("agenten") or {}).get(schluessel, {}).get("rolle", "")


def besetzung(audit: dict) -> dict:
    """Welche Teams in der Pruefung laufen. Die beiden Pflichtteams sind immer dabei."""
    gesetzt = audit.setdefault("besetzung", {})
    gesetzt.setdefault("falsifikationsteams", 1)
    gesetzt.setdefault("erwiderung", True)
    gesetzt["falsifikationsteams"] = min(3, max(1, int(gesetzt["falsifikationsteams"])))
    return gesetzt


def contra_schluessel(audit: dict) -> list:
    """Die Schluessel der aktiven Falsifikationsteams, immer mindestens einer."""
    anzahl = besetzung(audit)["falsifikationsteams"]
    return ["contra"] + [s for s, _ in ZUSATZ_CONTRA[:anzahl - 1]]


# ---------------------------------------------------------------------------
# Kleine Grafiken. Sie stecken als SVG im Programm, es sind keine Bilddateien
# noetig. bild() macht daraus ein Bild, das sich in Markdown einbetten laesst.
# ---------------------------------------------------------------------------
def bild(svg: str, breite: int = 160) -> str:
    daten = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f'<img src="data:image/svg+xml;base64,{daten}" width="{breite}">'


def bildpfad(name: str):
    """Sucht eine eigene Bilddatei im Ordner bilder. Fehlt sie, wird gezeichnet."""
    for endung in (".png", ".webp", ".jpg"):
        pfad = BILDER / f"{name}{endung}"
        if pfad.exists():
            return pfad
    return None


# Teams, die sich eine Figur teilen. Wer erwidert, baut auch, und jedes zusaetzliche
# Falsifikationsteam ist dieselbe Rolle. So genuegen zwei eigene Figuren.
FIGUR_ERSATZ = {"erwiderung": "pro", "contra2": "contra", "contra3": "contra"}


def figurpfad(schluessel: str):
    """Figur eines Teams. Fehlt sie, wird die Figur der verwandten Rolle genommen."""
    pfad = bildpfad(f"team_{schluessel}")
    if pfad is None and schluessel in FIGUR_ERSATZ:
        pfad = bildpfad(f"team_{FIGUR_ERSATZ[schluessel]}")
    return pfad


@st.cache_data(show_spinner=False)
def figur_eingebettet(pfad: str, kante: int = 160) -> str:
    """Verkleinert eine Figur einmalig und gibt sie als eingebettete Bilddaten zurueck."""
    from PIL import Image as PilBild
    with PilBild.open(pfad) as datei:
        klein = datei.convert("RGBA")
        klein.thumbnail((kante, kante))
        puffer = io.BytesIO()
        klein.save(puffer, format="PNG")
    return base64.b64encode(puffer.getvalue()).decode("ascii")


def _tiefe_svg(angriffe: int, erwiderung: bool) -> str:
    """Schema der Prueftiefe. Links baut ein Team, rechts greifen ein oder zwei an.

    Liegen eigene Figuren für das bauende und das angreifende Team im Ordner bilder,
    werden sie statt der Kreise eingesetzt. Zwei Dateien genügen also für alle drei
    Prüftiefen.
    """
    figuren = {}
    for seite, schluessel in (("links", "pro"), ("rechts", "contra")):
        pfad = bildpfad(f"team_{schluessel}")
        if pfad:
            try:
                figuren[seite] = figur_eingebettet(str(pfad))
            except Exception:
                pass
    t = ['<svg xmlns="http://www.w3.org/2000/svg" '
         'xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 200 124">',
         '<defs><marker id="s" markerWidth="8" markerHeight="8" refX="7" refY="4" '
         'orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#6b7280"/></marker>'
         '<marker id="b" markerWidth="8" markerHeight="8" refX="7" refY="4" '
         'orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#3b7ddd"/></marker></defs>']
    links_y, rechts = (50, [50]) if angriffe == 1 else (50, [24, 76])

    def kreis(seite, cx, cy, farbe, glyphe):
        daten = figuren.get(seite)
        if daten:
            t.append(f'<image xlink:href="data:image/png;base64,{daten}" x="{cx - 24}" '
                     f'y="{cy - 24}" width="48" height="48"/>')
            return
        t.append(f'<circle cx="{cx}" cy="{cy}" r="20" fill="{farbe}"/>')
        t.append(glyphe)

    kreis("links", 38, links_y, "#3b7ddd",
          f'<path d="M28 {links_y} H48 M38 {links_y - 10} V{links_y + 10}" stroke="#fff" '
          'stroke-width="3.6" stroke-linecap="round"/>')
    for y in rechts:
        kreis("rechts", 162, y, "#cf5a52",
              f'<path d="M154 {y - 8} L170 {y + 8} M170 {y - 8} L154 {y + 8}" stroke="#fff" '
              'stroke-width="3.6" stroke-linecap="round"/>')
    if angriffe == 1:
        t.append(f'<path d="M138 {40 if erwiderung else 50} H66" stroke="#6b7280" '
                 'stroke-width="2.4" marker-end="url(#s)"/>')
        if erwiderung:
            t.append('<path d="M62 62 H134" stroke="#3b7ddd" stroke-width="2.4" '
                     'stroke-dasharray="6 4" marker-end="url(#b)"/>')
    else:
        t.append('<path d="M140 30 L66 44" stroke="#6b7280" stroke-width="2.4" '
                 'marker-end="url(#s)"/>')
        t.append('<path d="M140 70 L66 56" stroke="#6b7280" stroke-width="2.4" '
                 'marker-end="url(#s)"/>')
        if erwiderung:
            t.append('<path d="M62 50 H130" stroke="#3b7ddd" stroke-width="2.4" '
                     'stroke-dasharray="6 4" marker-end="url(#b)"/>')
    beschriftung = "greift an" if angriffe == 1 else "greifen an"
    unterkante = max(rechts + [links_y]) + 38
    t.append(f'<text x="38" y="{unterkante}" font-size="13" fill="#6b7280" '
             'text-anchor="middle" font-family="sans-serif">baut</text>')
    t.append(f'<text x="162" y="{unterkante}" font-size="13" fill="#6b7280" '
             f'text-anchor="middle" font-family="sans-serif">{beschriftung}</text>')
    return "".join(t) + "</svg>"


def _symbol(farbe: str, inhalt: str) -> str:
    return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 44 44">'
            f'<circle cx="22" cy="22" r="21" fill="{farbe}"/>{inhalt}</svg>')


_KREUZ = ('<path d="M14 14 L30 30 M30 14 L14 30" stroke="#fff" stroke-width="3.4" '
          'stroke-linecap="round"/>')
_BLATT = ('<rect x="13" y="11" width="18" height="22" rx="2.5" fill="none" stroke="#fff" '
          'stroke-width="2.4"/><path d="M17 18 H27 M17 23 H27 M17 28 H23" stroke="#fff" '
          'stroke-width="2" stroke-linecap="round"/>')

TEAM_BILD = {
    "vorbereitung": _symbol("#5b7fa6", _BLATT),
    "pro": _symbol("#3b7ddd", '<path d="M13 22 H31 M22 13 V31" stroke="#fff" stroke-width="3.6" '
                              'stroke-linecap="round"/>'),
    "contra": _symbol("#cf5a52", _KREUZ),
    "contra2": _symbol("#d9827b", _KREUZ),
    "contra3": _symbol("#e3a19b", _KREUZ),
    "erwiderung": _symbol("#3b7ddd",
                          '<rect x="9" y="12" width="17" height="12" rx="3.5" fill="none" '
                          'stroke="#fff" stroke-width="2.2"/>'
                          '<rect x="19" y="21" width="17" height="12" rx="3.5" fill="#3b7ddd" '
                          'stroke="#fff" stroke-width="2.2"/>'),
    "vorort": _symbol("#4a9d5f",
                      '<circle cx="20" cy="19" r="8" fill="none" stroke="#fff" '
                      'stroke-width="2.6"/><path d="M26 25 L32 31" stroke="#fff" '
                      'stroke-width="3" stroke-linecap="round"/>'),
    "bericht": _symbol("#d9a431", _BLATT),
    "massnahmen": _symbol("#cf5a52",
                          '<path d="M32 22 A10 10 0 1 1 25 12.6" fill="none" stroke="#fff" '
                          'stroke-width="2.8" stroke-linecap="round"/>'
                          '<path d="M24 7 L28 14 L21 15.5 z" fill="#fff"/>')}

# Die drei Prüftiefen. Schluessel, Name, Zahl der Falsifikationsteams, Erwiderung,
# was passiert, wofuer es gedacht ist.
PRUEFTIEFEN = [
    ("schnell", "Schnell", 1, False,
     "Ein Team baut, ein Team greift an.",
     "Für den ersten Überblick. Sie sehen die Zweifel ohne Gegenrede."),
    ("standard", "Standard", 1, True,
     "Ein Angriff, dann darf das bauende Team erwidern.",
     "Die empfohlene Einstellung für das normale interne Audit."),
    ("gruendlich", "Gründlich", 2, True,
     "Zwei Teams greifen aus verschiedenen Richtungen an, dann die Erwiderung.",
     "Für kritische Prozesse. Findet mehr, dauert länger.")]


def team_grafik(schluessel: str, breite: int = 46):
    pfad = figurpfad(schluessel)
    if pfad:
        st.image(str(pfad), width=breite)
    else:
        st.markdown(bild(TEAM_BILD.get(schluessel, TEAM_BILD["pro"]), breite),
                    unsafe_allow_html=True)


def tiefe_grafik(schluessel: str, angriffe: int, erwiderung: bool, breite: int = 170):
    pfad = bildpfad(f"tiefe_{schluessel}")
    if pfad:
        st.image(str(pfad), width=breite)
    else:
        st.markdown(bild(_tiefe_svg(angriffe, erwiderung), breite), unsafe_allow_html=True)


def tiefe_aktuell(audit: dict) -> str:
    g = besetzung(audit)
    for schluessel, _, anzahl, erw, _, _ in PRUEFTIEFEN:
        if g["falsifikationsteams"] == anzahl and bool(g["erwiderung"]) == erw:
            return schluessel
    return "eigene"


URTEILE = ["Ausgeräumt", "Bestätigt: Abweichung", "Bestätigt: Verbesserungspotenzial",
           "Offen: vor Ort prüfen"]
TYPEN = {"widerlegend": "Nachweise sprechen dagegen",
         "untergrabend": "Nachweis ist nicht belastbar",
         "unterlaufend": "Nachweis deckt die Anforderung nicht ab",
         "ungestuetzte_annahme": "Behauptung ohne Nachweis"}
VORORT_ERGEBNIS = ["noch nicht geprüft", "in der Praxis bestätigt", "teilweise bestätigt",
                   "nicht bestätigt", "nicht anwendbar"]

# Der Auditablauf als Regelkreis. Je Phase die Aufgaben in der Reihenfolge, in der
# sie bearbeitet werden. Die Nummer ist zugleich der Schluessel des Inhaltsblocks.
PHASEN = [
    ("Plan", "#3b7ddd", "Das Audit vorbereiten und die Grundlage zusammentragen.",
     [(1, "Vorbereitung", "Auditziel, Umfang, Normabschnitte und Auditteam festlegen."),
      (2, "Nachweise", "Unterlagen, Gesprächsnotizen und Leistungsdaten zusammenstellen.")]),
    ("Do", "#4a9d5f", "Das Audit durchführen und die Nachweise erheben.",
     [(3, "Prüfung", "Die Agententeams bauen den Konformitätsnachweis und greifen ihn an."),
      (4, "Vor Ort", "Gespräch führen, beobachten, Stichproben ziehen und festhalten.")]),
    ("Check", "#d9a431", "Die Nachweise bewerten und das Ergebnis dokumentieren.",
     [(5, "Urteil", "Über jeden Zweifel entscheiden und die Entscheidung begründen."),
      (6, "Bericht", "Feststellungen und Ergebnis nachvollziehbar dokumentieren.")]),
    ("Act", "#cf5a52", "Verbessern und die Wirkung nachhalten.",
     [(7, "Massnahmen", "Verbesserungen festlegen, zuordnen und Wirksamkeit nachverfolgen.")])]

SCHRITT_PHASE = {nr: name for name, _, _, aufgaben in PHASEN for nr, _, _ in aufgaben}
SCHRITT_NAME = {nr: titel for _, _, _, aufgaben in PHASEN for nr, titel, _ in aufgaben}
RISIKO = ["hoch", "mittel", "gering"]
STATUS = ["geplant", "in Vorbereitung", "durchgeführt", "berichtet", "abgeschlossen"]
MSTATUS = ["offen", "in Umsetzung", "umgesetzt", "wirksam bestätigt", "nicht wirksam"]


# ---------------------------------------------------------------------------
# Zugang
# ---------------------------------------------------------------------------
def anmeldung() -> bool:
    if st.session_state.get("auth_ok"):
        return True
    rand_l, mitte, rand_r = st.columns([1, 2, 1])
    with mitte:
        if LOGO.exists():
            bild_l, bild_m, bild_r = st.columns([1, 2, 1])
            with bild_m:
                st.image(str(LOGO), use_container_width=True)
        else:
            st.title(APP_NAME)
        st.markdown(
            f"<p style='text-align:center;font-size:1.1rem;color:#4a5a68;margin-top:-.5rem'>"
            f"{APP_CLAIM}</p>", unsafe_allow_html=True)
        st.write("")
        pw = st.text_input("Zugangspasswort", type="password",
                           placeholder="Passwort eingeben")
        if st.button("Anmelden", type="primary", use_container_width=True):
            if pw and pw == st.secrets.get("APP_PASSWORD", ""):
                st.session_state.auth_ok = True
                st.rerun()
            else:
                st.error("Das Passwort stimmt nicht.")
        with st.expander("Was Audix tut und was nicht"):
            st.markdown(HINWEISE)
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


def llm_json(modell: str, system: str, nutzer: str, audit: dict, rolle: str = ""):
    """Anfrage an die OpenAI-kompatible Schnittstelle, bei Bedarf als Datenstrom."""
    name = agent_name(audit, rolle)
    eigene_rolle = agent_rolle(audit, rolle).strip()
    if name:
        system = f"Dein Name in diesem Audit ist {name}.\n" + system
    if eigene_rolle:
        system += ("\n\nROLLENVERSTÄNDNIS, das der Auditor für dich festgelegt hat:\n"
                   + eigene_rolle)
    zusatz = (audit.get("prompt_zusatz") or {}).get(rolle, "").strip()
    if zusatz:
        system += ("\n\nZUSÄTZLICHE ANWEISUNG DES AUDITORS. Sie ergänzt die Regeln oben und "
                   "hebt sie nicht auf:\n" + zusatz)
    system = system + NORMREGEL + SPRACHREGEL + JSONREGEL
    basis = {"model": modell, "temperature": 0.2, "max_tokens": MAX_TOKENS, "stream": STREAM,
             "response_format": {"type": "json_object"},
             "messages": [{"role": "system", "content": system},
                          {"role": "user", "content": nutzer}]}
    kopf = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    start = datetime.now()

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
            raise ModellFehler(f"Der Dienst hat abgebrochen (HTTP {r.status_code}). Weniger "
                               "Nachweise laden oder MAX_TOKENS und MAX_ZEICHEN verkleinern.")
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
        {"zeit": start.isoformat(timespec="seconds"), "rolle": rolle, "modell": modell,
         "dauer_sekunden": round((datetime.now() - start).total_seconds()),
         "zeichen_eingabe": len(nutzer), "zeichen_ausgabe": len(inhalt),
         "zusatzanweisung": zusatz, "ausgabe": inhalt[:6000]})
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
    return "".join(f"\n=== NACHWEIS: {n} ===\n{i}\n" for n, i in quellen.items())[:MAX_ZEICHEN]


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
                    f"Vorhandene Nachweise\n{probe}", audit, "vorbereitung")


def analyse_durchfuehren(audit, katalog, abschnitt, mit_erwiderung=True, melden=None):
    quellen = quellen_sammeln(audit)
    kriterium = kriterium_text(katalog, abschnitt)
    material = korpus(quellen)
    nr = abschnitt["nr"]
    if melden:
        melden(f"Konstruktionsteam prüft Abschnitt {nr}")
    fall = llm_json(MODEL_PRO, PRO_SYSTEM,
                    f"Auditkriterium\n{kriterium}\n\nNachweise\n{material}", audit, "pro")
    for t in fall.get("teilaussagen", []):
        zitate_pruefen(t.get("nachweise"), quellen)
    praefix = nr.replace(".", "_")
    einw = {"einwaende": []}
    for i, schluessel in enumerate(contra_schluessel(audit)):
        teamname = agent_name(audit, schluessel)
        if melden:
            melden(f"{teamname} greift Abschnitt {nr} an")
        teil = llm_json(MODEL_CONTRA, CONTRA_SYSTEM,
                        f"Auditkriterium\n{kriterium}\n\nAssurance Case\n"
                        f"{json.dumps(fall, ensure_ascii=False)}\n\nNachweise\n{material}",
                        audit, schluessel)
        for x in teil.get("einwaende", []):
            # Die zusaetzlichen Teams bekommen ein eigenes Kuerzel, damit sich die
            # Nummern der Einwaende nicht gegenseitig ueberschreiben.
            kennung = "" if i == 0 else f"T{i + 1}-"
            x["id"] = f"{praefix}-{kennung}{x.get('id', 'E')}"
            x["abschnitt"] = nr
            x["team"] = teamname
            zitate_pruefen(x.get("nachweise"), quellen)
            einw["einwaende"].append(x)
    erw = {"erwiderungen": []}
    if mit_erwiderung and einw.get("einwaende"):
        if melden:
            melden(f"Konstruktionsteam erwidert zu Abschnitt {nr}")
        erw = llm_json(MODEL_PRO, ERWIDERUNG_SYSTEM,
                       f"Einwände\n{json.dumps(einw, ensure_ascii=False)}\n\n"
                       f"Nachweise\n{material}", audit, "erwiderung")
        for x in erw.get("erwiderungen", []):
            zitate_pruefen(x.get("nachweise"), quellen)
    audit.setdefault("analysen", {})[nr] = {
        "titel": abschnitt["titel"], "anforderung": abschnitt["anforderung"],
        "pruefpunkte": abschnitt.get("pruefpunkte", []), "kriterium": kriterium,
        "fall": fall, "einwaende": einw, "erwiderungen": erw,
        "modelle": {"pro": MODEL_PRO, "contra": MODEL_CONTRA},
        "teams": [agent_name(audit, s) for s in contra_schluessel(audit)],
        "zeit": datetime.now().isoformat(timespec="seconds")}


def alle_einwaende(audit: dict) -> list:
    return [e for erg in audit.get("analysen", {}).values()
            for e in erg.get("einwaende", {}).get("einwaende", [])]


# ---------------------------------------------------------------------------
# Leitfaden fuer das Gespraech vor Ort
# ---------------------------------------------------------------------------
def leitfaden_erzeugen(audit: dict, abschnitt_nr: str):
    """Leitet aus der Pruefung eines Abschnitts Fragen fuer den Termin vor Ort ab."""
    erg = audit.get("analysen", {}).get(abschnitt_nr)
    if not erg:
        return
    urteile = audit.get("urteile", {})
    einwaende = []
    for e in erg.get("einwaende", {}).get("einwaende", []):
        einwaende.append({**e, "urteil_des_auditors":
                          urteile.get(e["id"], {}).get("urteil", "noch nicht beurteilt")})
    auftrag = (f"Normabschnitt {abschnitt_nr} {erg.get('titel', '')}\n"
               f"Anforderung\n{erg.get('anforderung', '')}\n\n"
               f"Prüfpunkte\n" + "\n".join(f"- {p}" for p in erg.get("pruefpunkte", [])) +
               "\n\nKonformitätsnachweis\n"
               f"{json.dumps(erg.get('fall', {}), ensure_ascii=False)}\n\n"
               f"Einwände mit Urteil des Auditors\n"
               f"{json.dumps({'einwaende': einwaende}, ensure_ascii=False)}\n\n"
               f"Erwiderungen\n{json.dumps(erg.get('erwiderungen', {}), ensure_ascii=False)}\n\n"
               f"Auditierter Prozess {audit.get('prozess', '')}")
    ergebnis = llm_json(MODEL_CONTRA, VORORT_SYSTEM, auftrag, audit, "vorort")
    praefix = abschnitt_nr.replace(".", "_")
    fragen = []
    for i, fr in enumerate(ergebnis.get("fragen", []), start=1):
        fr["id"] = f"{praefix}-F{i}"
        fr["abschnitt"] = abschnitt_nr
        fragen.append(fr)
    audit.setdefault("leitfaden", {})[abschnitt_nr] = {
        "titel": erg.get("titel", ""), "fragen": fragen,
        "zeit": datetime.now().isoformat(timespec="seconds")}


def alle_fragen(audit: dict) -> list:
    return [f for nr in sorted(audit.get("leitfaden", {}))
            for f in audit["leitfaden"][nr].get("fragen", [])]


def beobachtung(audit: dict, frage_id: str) -> dict:
    return audit.get("beobachtungen", {}).get(frage_id, {})


def beobachtung_zu_einwand(audit: dict, einwand_id: str) -> list:
    """Alle Fragen samt Beobachtung, die zu einem bestimmten Einwand gehoeren."""
    treffer = []
    kurz = einwand_id.split("-", 1)[-1]
    for f in alle_fragen(audit):
        bezug = str(f.get("einwand", ""))
        if bezug and (bezug == einwand_id or bezug.split("-", 1)[-1] == kurz):
            treffer.append((f, beobachtung(audit, f["id"])))
    return treffer


def vorort_befunde(audit: dict) -> list:
    """Die vor Ort erfassten Beobachtungen in einer flachen, berichtsfaehigen Liste."""
    befunde = []
    for f in alle_fragen(audit):
        b = beobachtung(audit, f["id"])
        text = str(b.get("text", "")).strip()
        ergebnis = b.get("ergebnis", VORORT_ERGEBNIS[0])
        if not text and ergebnis == VORORT_ERGEBNIS[0]:
            continue
        befunde.append({"abschnitt": f.get("abschnitt", ""), "frage": f.get("frage", ""),
                        "einwand": f.get("einwand", ""), "ergebnis": ergebnis,
                        "beobachtung": text})
    return befunde


def ampel_teilaussage(t: dict, einwaende: list, urteile: dict):
    """Gibt (Zeichen, Kurztext) fuer eine Teilaussage zurueck."""
    belegt = any(n.get("verifiziert") for n in t.get("nachweise", []))
    offen = [e for e in einwaende if e.get("ziel") == t.get("id")
             and urteile.get(e.get("id"), {}).get("urteil") != "Ausgeräumt"]
    schwer = [e for e in offen if e.get("typ") in ("widerlegend", "untergrabend")
              or e.get("schwere") == "hoch"]
    if not belegt or schwer:
        return "🔴", "ohne belastbaren Nachweis"
    if offen or t.get("annahmen"):
        return "🟡", "Nachweis vorhanden, Zweifel offen"
    return "🟢", "belegt"


def ampel_abschnitt(erg: dict, urteile: dict):
    einwaende = erg.get("einwaende", {}).get("einwaende", [])
    zeichen = [ampel_teilaussage(t, einwaende, urteile)[0]
               for t in erg.get("fall", {}).get("teilaussagen", [])]
    if not zeichen:
        return "🔴", "Kein Nachweis erarbeitet"
    gruen, gelb, rot = zeichen.count("🟢"), zeichen.count("🟡"), zeichen.count("🔴")
    if rot:
        lage = "Mögliche Abweichung, mindestens ein Prüfpunkt ohne Nachweis"
    elif gelb:
        lage = "Nachweis unvollständig, offene Zweifel"
    else:
        lage = "Anforderung durchgängig belegt"
    return ("🔴" if rot else "🟡" if gelb else "🟢"), f"{lage} ({gruen} belegt, {gelb} mit Zweifel, {rot} ohne Nachweis)"


def kennzahlen(audit: dict) -> dict:
    nachweise = [n for erg in audit.get("analysen", {}).values()
                 for t in erg.get("fall", {}).get("teilaussagen", [])
                 for n in t.get("nachweise", [])]
    nachweise += [n for e in alle_einwaende(audit) for n in e.get("nachweise", [])]
    urteile = [u.get("urteil") for u in audit.get("urteile", {}).values()]
    mass = audit.get("massnahmen", [])
    fragen = alle_fragen(audit)
    return {"abschnitte": len(audit.get("analysen", {})),
            "einwaende": len(alle_einwaende(audit)),
            "fragen": len(fragen),
            "fragen_geprueft": len(vorort_befunde(audit)),
            "nicht_gelebt": sum(1 for b in vorort_befunde(audit)
                                if b["ergebnis"] in ("nicht bestätigt", "teilweise bestätigt")),
            "beurteilt": sum(1 for u in urteile if u),
            "zitate": len(nachweise),
            "zitate_belegt": sum(1 for n in nachweise if n.get("verifiziert")),
            "abweichungen": sum(1 for u in urteile if u == "Bestätigt: Abweichung"),
            "potenziale": sum(1 for u in urteile if u == "Bestätigt: Verbesserungspotenzial"),
            "offen": sum(1 for u in urteile if u == "Offen: vor Ort prüfen"),
            "ausgeraeumt": sum(1 for u in urteile if u == "Ausgeräumt"),
            "massnahmen": len(mass),
            "massnahmen_offen": sum(1 for m in mass
                                    if m.get("status") in ("offen", "in Umsetzung"))}


def phasen_schaltflaeche(erledigt: dict) -> int:
    """Der Regelkreis als Schaltflaeche. Gibt die gewaehlte Aufgabe zurueck."""
    phase = st.session_state.setdefault("phase", "Plan")
    schritt = st.session_state.setdefault("schritt", 1)
    if SCHRITT_PHASE.get(schritt) != phase:
        schritt = next(a[0] for n, _, _, a in PHASEN if n == phase)
        st.session_state.schritt = schritt
    farben = {n: f for n, f, _, _ in PHASEN}
    aufgabe_schluessel = ", ".join(f".st-key-auf_{n} button" for n in SCHRITT_NAME)
    aufgabe_innen = ", ".join(f".st-key-auf_{n} button div, .st-key-auf_{n} button span"
                              for n in SCHRITT_NAME)

    regeln = [
        # Die vier Phasenfelder. Nur das gewaehlte Feld ist gefuellt.
        ".st-key-ph_Plan button, .st-key-ph_Do button, .st-key-ph_Check button, "
        ".st-key-ph_Act button {height:5rem; border-radius:0.8rem; border:1px solid "
        "transparent; background-color:transparent; box-shadow:none;}",
        # Die Aufgaben als flache, linksbuendige Zeilen
        aufgabe_schluessel + " {justify-content:flex-start; text-align:left; "
        "background-color:transparent; border:none; box-shadow:none; padding:0.3rem 0.6rem;}",
        # Streamlit zentriert die Beschriftung in zwei inneren Huellen. Beide umstellen.
        aufgabe_innen + " {justify-content:flex-start; text-align:left; width:100%;}"]
    for name, farbe, _, aufgaben in PHASEN:
        if name == phase:
            regeln.append(f".st-key-ph_{name} button, .st-key-ph_{name} button:hover, "
                          f".st-key-ph_{name} button:focus {{background-color:{farbe}; "
                          f"border-color:{farbe};}}")
            regeln.append(f".st-key-ph_{name} button p {{color:#ffffff; font-weight:700;}}")
        else:
            regeln.append(f".st-key-ph_{name} button:hover {{background-color:{farbe}1f; "
                          f"border-color:{farbe};}}")
    aktiv_farbe = farben.get(phase, "#3b7ddd")
    regeln.append(f".st-key-auf_{schritt} button {{background-color:{aktiv_farbe}1f; "
                  f"border-left:4px solid {aktiv_farbe};}}")
    regeln.append(f".st-key-auf_{schritt} button p {{font-weight:700;}}")
    st.markdown("<style>" + "".join(regeln) + "</style>", unsafe_allow_html=True)

    st.caption("Der Auditablauf als Regelkreis. Wählen Sie eine Phase.")
    for reihe in (PHASEN[:2], PHASEN[2:]):
        spalten = st.columns(2)
        for spalte, (name, _, _, aufgaben) in zip(spalten, reihe):
            with spalte:
                wort = "Aufgabe" if len(aufgaben) == 1 else "Aufgaben"
                fertig = sum(1 for nr, _, _ in aufgaben if erledigt.get(nr))
                zusatz = (" · alle erledigt" if fertig == len(aufgaben)
                          else f" · {fertig} erledigt" if fertig else "")
                if st.button(f"{name}  \n{len(aufgaben)} {wort}{zusatz}", key=f"ph_{name}",
                             use_container_width=True):
                    st.session_state.phase = name
                    st.session_state.schritt = aufgaben[0][0]
                    st.rerun()
    st.divider()

    name, _, zweck, aufgaben = next(p for p in PHASEN if p[0] == phase)
    st.markdown(f"**{name} · Aufgabenübersicht**")
    st.caption(zweck)
    for nr, titel, text in aufgaben:
        marke = "✓" if erledigt.get(nr) else "○"
        if st.button(f"{marke}  {nr}  {titel}", key=f"auf_{nr}", use_container_width=True):
            st.session_state.schritt = nr
            st.rerun()
        st.caption(text)
    st.divider()
    return schritt


def besetzung_feld(audit: dict):
    """Die drei Pruefttiefen als Karten, darunter bei Bedarf die Feineinstellung."""
    g = besetzung(audit)
    gewaehlt = tiefe_aktuell(audit)
    st.markdown("**Wer prüft**")
    st.caption("Ein Team baut den Nachweis, mindestens ein zweites greift ihn an. Diese beiden "
               "sind immer dabei, denn darauf beruht das Verfahren.")
    st.markdown(
        "<style>.st-key-tiefe_schnell button, .st-key-tiefe_standard button, "
        ".st-key-tiefe_gruendlich button {border-radius:0.6rem;}</style>",
        unsafe_allow_html=True)
    spalten = st.columns(3)
    for spalte, (schluessel, name, anzahl, erw, was, wofuer) in zip(spalten, PRUEFTIEFEN):
        with spalte:
            with st.container(border=True):
                tiefe_grafik(schluessel, anzahl, erw, 170)
                st.markdown(f"**{name}**" + (" ✓" if schluessel == gewaehlt else ""))
                st.caption(was)
                st.caption(wofuer)
                aufrufe = anzahl + 1 + (1 if erw else 0)
                st.caption(f"{aufrufe} Modellaufrufe je Normabschnitt")
                if st.button("Diese Tiefe wählen" if schluessel != gewaehlt else "Gewählt",
                             key=f"tiefe_{schluessel}", use_container_width=True,
                             type="primary" if schluessel == gewaehlt else "secondary",
                             disabled=schluessel == gewaehlt):
                    g["falsifikationsteams"] = anzahl
                    g["erwiderung"] = erw
                    st.session_state[f"bes_anzahl_{audit['id']}"] = anzahl
                    st.session_state[f"bes_erwiderung_{audit['id']}"] = erw
                    st.rerun()

    with st.expander("Besetzung von Hand anpassen"):
        s1, s2 = st.columns([3, 2])
        g["falsifikationsteams"] = s1.radio(
            "Wie viele Teams greifen den Nachweis an", [1, 2, 3],
            index=g["falsifikationsteams"] - 1, horizontal=True,
            key=f"bes_anzahl_{audit['id']}")
        g["erwiderung"] = s2.checkbox(
            "Das bauende Team darf erwidern", value=bool(g["erwiderung"]),
            key=f"bes_erwiderung_{audit['id']}")
        abschnitte_anzahl = max(1, len(audit.get("kriterien") or []))
        je_abschnitt = 1 + g["falsifikationsteams"] + (1 if g["erwiderung"] else 0)
        st.caption(f"Das ergibt {je_abschnitt * abschnitte_anzahl} Modellaufrufe bei "
                   f"{abschnitte_anzahl} gewählten Normabschnitten. Rechnen Sie mit etwa einer "
                   "Minute je Aufruf.")


def team_feld(audit: dict, schluessel: str, standardname: str, zweck: str, prompt: str,
              modell: str):
    """Ein Teamkarte mit Name, anklickbaren Musterrollen und Zusatzanweisung."""
    eintrag = audit.setdefault("agenten", {}).setdefault(schluessel, {"name": "", "rolle": ""})
    rollenfeld = f"ar_{audit['id']}_{schluessel}"
    if rollenfeld not in st.session_state:
        st.session_state[rollenfeld] = eintrag.get("rolle", "")
    kopf_bild, kopf_text = st.columns([1, 6], vertical_alignment="center")
    with kopf_bild:
        team_grafik(schluessel, 64)
    with kopf_text:
        st.caption(zweck)
        st.caption(f"Arbeitet mit dem {modell}.")

    muster = MUSTERROLLEN.get(schluessel, [])
    if muster:
        st.caption("Muster für den Schwerpunkt. Anklicken, dann steht der Text unten im Feld "
                   "und lässt sich dort weiter anpassen.")
        spalten = st.columns(len(muster) + 1)
        for spalte, (titel, text) in zip(spalten, muster):
            if spalte.button(titel, key=f"mu_{audit['id']}_{schluessel}_{titel}",
                             use_container_width=True):
                st.session_state[rollenfeld] = text
                eintrag["rolle"] = text
                st.rerun()
        if spalten[-1].button("Feld leeren", key=f"ml_{audit['id']}_{schluessel}",
                              use_container_width=True):
            st.session_state[rollenfeld] = ""
            eintrag["rolle"] = ""
            st.rerun()

    eintrag["name"] = st.text_input("Name dieses Teams", value=eintrag.get("name", ""),
                                    key=f"an_{audit['id']}_{schluessel}",
                                    placeholder=standardname)
    eintrag["rolle"] = st.text_area(
        "Schwerpunkt dieses Teams", height=100, key=rollenfeld,
        placeholder="Beispiel: Du bist ein erfahrener Auditor aus der Medizintechnik und legst "
                    "besonderen Wert auf Rückverfolgbarkeit.")
    audit.setdefault("prompt_zusatz", {})[schluessel] = st.text_area(
        "Zusätzliche Anweisung für dieses Audit",
        value=audit["prompt_zusatz"].get(schluessel, ""), height=80,
        key=f"pz_{audit['id']}_{schluessel}",
        placeholder="Beispiel: Achte besonders auf Fristen und auf Dokumente ohne Freigabevermerk.")
    with st.expander("Grundanweisung dieses Teams anzeigen"):
        st.code(prompt, language="text")


def teamliste(audit: dict) -> list:
    """Alle Teams in Arbeitsreihenfolge, zugeschaltete Falsifikationsteams eingereiht."""
    aktiv = contra_schluessel(audit)
    liste = []
    for schluessel, standardname, phase, zweck, prompt, modell in AGENTEN:
        liste.append((schluessel, standardname, phase, zweck, prompt, modell))
        if schluessel == "contra":
            for zusatz, zusatzname in ZUSATZ_CONTRA:
                if zusatz in aktiv:
                    liste.append((zusatz, zusatzname, phase,
                                  "Greift denselben Nachweis aus einer anderen Richtung an. "
                                  "Geben Sie ihm einen eigenen Schwerpunkt, sonst findet er "
                                  "dasselbe wie das erste Team.", prompt, modell))
    return liste


def team_daten(audit: dict, schluessel: str):
    """Die Beschreibungsdaten eines Teams, auch der zugeschalteten."""
    for eintrag in teamliste(audit):
        if eintrag[0] == schluessel:
            return eintrag
    return (schluessel, schluessel, "", "", "", "")


@st.dialog("Team anpassen", width="large")
def team_fenster(audit: dict, schluessel: str):
    _, standardname, _, zweck, prompt, modell = team_daten(audit, schluessel)
    team_feld(audit, schluessel, standardname, zweck, prompt, modell)
    if st.button("Fertig", type="primary", use_container_width=True):
        st.rerun()


def team_zeile(audit: dict, schluessel: str, zusatz: str = ""):
    """Eine ruhige Zeile je Team. Das Anpassen passiert in einem Fenster darüber."""
    _, standardname, _, _, _, _ = team_daten(audit, schluessel)
    rolle = agent_rolle(audit, schluessel).strip()
    muster = next((t for t, text in MUSTERROLLEN.get(schluessel, []) if text == rolle), "")
    beschreibung = muster or ("eigener Schwerpunkt" if rolle else "Standardverhalten")
    if zusatz:
        beschreibung += f" · {zusatz}"
    with st.container(border=True):
        s1, s2, s3 = st.columns([1, 6, 2], vertical_alignment="center")
        with s1:
            team_grafik(schluessel, 46)
        s2.markdown(f"**{agent_name(audit, schluessel)}**  \n{beschreibung}")
        if s3.button("anpassen", key=f"tz_{audit['id']}_{schluessel}",
                     use_container_width=True):
            team_fenster(audit, schluessel)


def nachweise_anzeigen(nachweise, leer_hinweis="Kein Nachweis angegeben."):
    if not nachweise:
        st.caption(leer_hinweis)
    for n in nachweise or []:
        marke = ("🟢 Zitat im Dokument gefunden" if n.get("verifiziert")
                 else "🔴 Zitat NICHT im Dokument gefunden, nicht verwendbar")
        st.markdown(f"> {n.get('zitat', '')}  \n*Quelle · {n.get('dokument', '')}* · {marke}")


# ---------------------------------------------------------------------------
# Bericht
# ---------------------------------------------------------------------------
def bericht_bloecke(audit: dict) -> dict:
    k = kennzahlen(audit)
    kopf = [("Organisation", ORGANISATION or "nicht angegeben"),
            ("Auditierter Bereich oder Prozess", audit.get("prozess", "")),
            ("Auditziel", audit.get("ziel", "")),
            ("Umfang und Grenzen", audit.get("umfang", "")),
            ("Auditkriterien", f"{audit.get('norm', '')}, Abschnitte "
                               + ", ".join(audit.get("kriterien", []))),
            ("Auditor", audit.get("auditor", "")),
            ("Unparteilichkeit bestätigt", "ja" if audit.get("unparteilich") else "nicht bestätigt"),
            ("Auditierte", audit.get("auditierte", "")),
            ("Audittermin", audit.get("termin", "")),
            ("Berichtsdatum", f"{date.today():%d.%m.%Y}")]
    quellen = ([f"Dokument · {n}" for n in audit.get("dokumente", {})]
               + [f"{n['typ']} · {n.get('quelle', '')} · {n.get('datum', '')}"
                  for n in audit.get("notizen", [])])
    return {"kennzahlen": k, "kopf": kopf, "quellen": quellen,
            "feststellungen": audit.get("feststellungen", {}).get("feststellungen", []),
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
    z += ["## Ergebnis je Normabschnitt", "", "| Abschnitt | Ergebnis |", "| --- | --- |"]
    for nr, erg in audit.get("analysen", {}).items():
        zeichen, text = ampel_abschnitt(erg, audit.get("urteile", {}))
        z.append(f"| {nr} {erg.get('titel', '')} | {zeichen} {text} |")
    z += ["", "## Kennzahlen der Prüfung", "", "| Grösse | Wert |", "| --- | --- |",
          f"| Geprüfte Normabschnitte | {k['abschnitte']} |",
          f"| Geprüfte Zweifel insgesamt | {k['einwaende']} |",
          f"| Davon ausgeräumt | {k['ausgeraeumt']} |",
          f"| Abweichungen | {k['abweichungen']} |",
          f"| Verbesserungspotenziale | {k['potenziale']} |",
          f"| Vor Ort nachzuprüfen | {k['offen']} |",
          f"| Zitate in der Quelle gefunden | {k['zitate_belegt']} von {k['zitate']} |", ""]
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
        z += ["", "| Feststellung | Korrekturmassnahme | Verantwortlich | Termin | Status |",
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
    z += ["## Prüfung vor Ort"]
    befunde = vorort_befunde(audit)
    if not befunde:
        z.append("Keine Beobachtungen vor Ort erfasst. Die Aussagen stützen sich auf die "
                 "eingereichten Unterlagen.")
    else:
        z += ["", "| Abschnitt | Frage | Beobachtung | Ergebnis |", "| --- | --- | --- | --- |"]
        for bef in befunde:
            z.append(f"| {bef['abschnitt']} | {bef['frage']} | "
                     f"{bef['beobachtung'] or 'keine Notiz'} | {bef['ergebnis']} |")
        z.append("")
    z += ["## Ausgeräumte Zweifel (Nachweis der Prüftiefe)"]
    z += [f"- {i}: {g}" for i, g in b["ausgeraeumt"]] or ["Keine."]
    z += ["", "---", "", f"Erstellt mit {APP_NAME}, einem dialektischen Prüfverfahren. Zwei getrennte "
          "Agententeams haben den Konformitätsnachweis konstruiert und angegriffen. Über jeden "
          "verbliebenen Zweifel hat der Auditor entschieden und dies begründet. Einstufung und "
          "Schlussfolgerung verantwortet der Auditor."]
    return "\n".join(z)


def bericht_html(audit: dict) -> str:
    b = bericht_bloecke(audit)
    k = b["kennzahlen"]

    def schutz(t):
        return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    kopf = "".join(f"<tr><th>{schutz(n)}</th><td>{schutz(w)}</td></tr>" for n, w in b["kopf"])
    kacheln = "".join(
        f'<div class="kachel"><div class="zahl">{w}</div><div class="text">{t}</div></div>'
        for t, w in [("Normabschnitte", k["abschnitte"]), ("Geprüfte Zweifel", k["einwaende"]),
                     ("Abweichungen", k["abweichungen"]), ("Potenziale", k["potenziale"]),
                     ("Vor Ort offen", k["offen"]),
                     ("Zitate belegt", f"{k['zitate_belegt']}/{k['zitate']}")])
    klasse = {"🟢": "gruen", "🟡": "gelb", "🔴": "rot"}
    uebersicht = "".join(
        f'<tr><td>{schutz(nr)} {schutz(erg.get("titel", ""))}</td>'
        f'<td class="{klasse[ampel_abschnitt(erg, audit.get("urteile", {}))[0]]}">'
        f'{schutz(ampel_abschnitt(erg, audit.get("urteile", {}))[1])}</td></tr>'
        for nr, erg in audit.get("analysen", {}).items())
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
    fazit = f"<h2>Fazit des Auditors</h2><p>{schutz(b['fazit'])}</p>" if b["fazit"] else ""
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
td.gruen{{background:#e8f5e9}} td.gelb{{background:#fff8e1}} td.rot{{background:#fdecea}}
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
<h2>Ergebnis je Normabschnitt</h2>
<table><tr><th>Abschnitt</th><th>Ergebnis der Nachweisprüfung</th></tr>{uebersicht}</table>
<h2>Kennzahlen der Prüfung</h2><div class="kacheln">{kacheln}</div>
<h2>Feststellungen</h2>{fest}
<h2>Massnahmenplan</h2>
<table><tr><th>Feststellung</th><th>Massnahme</th><th>Verantwortlich</th><th>Termin</th>
<th>Status</th></tr>{mass or '<tr><td colspan="5">Keine Massnahmen erfasst.</td></tr>'}</table>
<h2>Geprüfte Nachweise</h2><ul>{quellen}</ul>
<h2>Ausgeräumte Zweifel</h2><ul>{ausger or '<li>Keine.</li>'}</ul>
<footer>Erstellt mit {schutz(APP_NAME)}, einem dialektischen Prüfverfahren. Zwei getrennte Agententeams haben den
Konformitätsnachweis konstruiert und angegriffen. Über jeden verbliebenen Zweifel hat der
Auditor entschieden und dies begründet. Einstufung und Schlussfolgerung verantwortet der
Auditor.</footer></body></html>"""


# ---------------------------------------------------------------------------
# Berichtsformular nach dem Muster eines Auditkurzberichts, auf ISO 9001 bezogen
# Aufbau (schluessel, Beschriftung, Hilfetext, Zeilen im Eingabefeld)
# ---------------------------------------------------------------------------
BERICHT_FELDER = [
    ("ziel_audit", "1.1 Zielsetzung des Audits",
     "Wozu wird auditiert. Wird aus dem Auditplan vorbelegt.", 100),
    ("dienstleistung", "1.2 Auditierter Prozess oder Bereich",
     "Welcher Prozess, welche Tätigkeiten und welche Schnittstellen.", 80),
    ("institution", "1.3 Organisation und Ansprechpartner",
     "Organisationseinheit, Anzahl Mitarbeitende, Prozessverantwortliche.", 80),
    ("zeitplan", "1.4 Auditzeitplan",
     "Datum, Uhrzeit, Ablauf, vor Ort oder remote.", 70),
    ("standort", "1.5 Geprüfter Standort",
     "Nur bei mehreren Standorten oder virtuellen Standorten ausfüllen.", 60),
    ("unterlagen", "1.6 Eingereichte Unterlagen",
     "Wird aus den erfassten Nachweisen vorbelegt.", 120),
    ("haftung", "1.7 Haftungsausschluss",
     "Standardtext. Die Prüfung beruht auf einer Stichprobe.", 90),
    ("verteiler", "1.8 Verteiler",
     "Wer den Bericht erhält. ISO 9001 verlangt Bericht an die zuständige Leitung.", 80),
    ("geltungsbereich", "2 Geltungsbereich",
     "Welcher Teil des QM-Systems geprüft wurde.", 80),
    ("begriffe", "3 Begriffe",
     "Definition von Abweichung, Hinweis und Empfehlung.", 110),
    ("veraenderungen_org", "4.1 Wichtige Veränderungen in der Organisation",
     "Organisatorische und personelle Änderungen seit dem letzten Audit.", 80),
    ("veraenderungen_prozess", "4.2 Veränderungen am Prozess oder Angebot",
     "Neue oder geänderte Prozesse, Produkte oder Dienstleistungen.", 80),
    ("kennzahlen_entwicklung", "4.3 Entwicklung der relevanten Kennzahlen",
     "Mengen, Reklamationen, Termintreue, Ausschuss und ähnliche Daten.", 80),
    ("erledigung_vorjahr", "4.4 Erledigungsnachweise früherer Korrekturmassnahmen",
     "Stand der Massnahmen aus früheren Audits und Bewertung ihrer Wirksamkeit.", 90),
    ("umgang_hinweise", "4.5 Umgang mit Hinweisen aus früheren Berichten",
     "Welche Hinweise wurden aufgegriffen, welche bleiben offen.", 80),
    ("qualitaetsinitiativen", "4.6 Eigene Qualitätsinitiativen",
     "Projekte, Evaluationen und Verbesserungen aus eigener Initiative.", 80),
    ("selbstbewertung", "4.7 Selbstbewertung der Organisation",
     "Wie die Organisation ihren eigenen Erfüllungsgrad einschätzt.", 80),
    ("regelkreis_fuehrung", "5.1.1 Regelkreis Führung und Qualitätsmanagement",
     "Normabschnitte 4, 5, 6, 9.3 und 10. Wird aus den Befunden vorbelegt.", 110),
    ("regelkreis_leistung", "5.1.2 Regelkreis Leistungserbringung und Unterstützung",
     "Normabschnitte 7 und 8. Wird aus den Befunden vorbelegt.", 110),
    ("gespraeche", "5.2 Eröffnungs- und Abschlussgespräch",
     "Wird aus den erfassten Gesprächsnotizen vorbelegt.", 100),
    ("erfuellung", "6.1 Erfüllung der geprüften Normabschnitte",
     "Wird aus der Ampelübersicht vorbelegt.", 120),
    ("begruendung", "6.2 Begründung des Auditergebnisses",
     "Warum das Ergebnis so ausfällt. Das ist Ihre Kernaussage als Auditor.", 120),
    ("hinweise_text", "6.3.3 Hinweise",
     "Entwicklungspotenziale und Risiken ohne Abweichungscharakter. "
     "Wird aus den Verbesserungspotenzialen vorbelegt.", 110),
    ("empfehlungen_text", "6.3.4 Empfehlungen",
     "Freiwillige Anregungen mit Nutzen für die Organisation.", 100),
    ("schlusswort", "6.4 Schlusswort",
     "Abschliessende Würdigung und Hinweis auf die Grundlagen der Bewertung.", 120),
    ("naechste_pruefung", "7 Planung der nächsten Überprüfung",
     "Termin, Schwerpunkte und Standorte des nächsten Audits.", 80),
    ("antrag", "8 Antrag an die Leitung",
     "Was der Auditor der Leitung empfiehlt, etwa Freigabe mit oder ohne Auflagen.", 100),
]


def feld_vorbelegung(audit: dict, schluessel: str) -> str:
    """Erzeugt den Vorschlagstext eines Berichtsfeldes aus den vorhandenen Daten."""
    k = kennzahlen(audit)
    urteile = audit.get("urteile", {})
    analysen = audit.get("analysen", {})

    if schluessel == "ziel_audit":
        basis = audit.get("ziel", "").strip()
        return (basis + "\n" if basis else "") + (
            "Das Audit verfolgt folgende Ziele:\n"
            f"- Prüfung der Konformität mit {audit.get('norm', 'der Norm')}, Abschnitte "
            + ", ".join(audit.get("kriterien", [])) + "\n"
            "- Beurteilung der Wirksamkeit des Prozesses und des Qualitätsmanagementsystems\n"
            "- Bewertung der Massnahmen aus früheren Audits\n"
            "- Ermittlung von Verbesserungspotenzialen\n"
            "- Berichterstattung an die zuständige Leitung")
    if schluessel == "dienstleistung":
        return audit.get("prozess", "")
    if schluessel == "institution":
        return (f"{ORGANISATION}\nAuditierte Funktion: {audit.get('auditierte', '')}\n"
                "Anzahl Mitarbeitende im geprüften Bereich: ")
    if schluessel == "zeitplan":
        return f"Audittermin: {audit.get('termin', '')}\nDurchführung: vor Ort"
    if schluessel == "standort":
        return "Nicht zutreffend, Einzelstandort."
    if schluessel == "unterlagen":
        zeilen = ["Folgende Nachweise wurden ausgewertet:"]
        zeilen += [f"- Dokument: {n}" for n in audit.get("dokumente", {})]
        zeilen += [f"- {n['typ']}: {n.get('quelle', '')} vom {n.get('datum', '')}"
                   for n in audit.get("notizen", [])]
        return "\n".join(zeilen)
    if schluessel == "haftung":
        return ("Das Audit beruht auf einem Stichprobenverfahren der verfügbaren Informationen. "
                f"Die Dokumentenanalyse wurde durch das KI-gestützte Prüfverfahren {APP_NAME} "
                "mit zwei "
                "getrennt arbeitenden Agententeams unterstützt. Jedes von der KI angeführte "
                "Zitat wurde maschinell gegen die Quelle geprüft. Die Bewertung der Nachweise "
                "und alle Schlussfolgerungen verantwortet der Auditor.")
    if schluessel == "verteiler":
        return ("Diesen Bericht erhalten:\n- die oberste Leitung\n"
                f"- die Leitung des auditierten Bereichs ({audit.get('auditierte', '')})\n"
                "- die Qualitätsmanagementbeauftragte Person\n"
                f"- der Auditor ({audit.get('auditor', '')})")
    if schluessel == "geltungsbereich":
        return (f"Geprüft wurde der Prozess {audit.get('prozess', '')} gegen "
                f"{audit.get('norm', '')}, Abschnitte "
                + ", ".join(audit.get("kriterien", [])) + ". "
                + (audit.get("umfang", "") or ""))
    if schluessel == "begriffe":
        return ("Die verwendeten Begriffe orientieren sich an ISO 9000 und ISO 19011.\n"
                "Abweichung (Nichtkonformität): Nichterfüllung einer Anforderung. "
                "Eine wesentliche Abweichung liegt vor, wenn eine Anforderung systematisch oder "
                "vollständig nicht erfüllt ist. Eine geringfügige Abweichung ist ein Einzelfall "
                "oder eine formale Lücke ohne Systemversagen.\n"
                "Hinweis: Entwicklungspotenzial oder Risiko ohne Abweichungscharakter.\n"
                "Empfehlung: freiwillige Anregung ohne Verpflichtung zur Umsetzung.")
    if schluessel in ("regelkreis_fuehrung", "regelkreis_leistung"):
        fuehrung = ("4", "5", "6", "9", "10")
        zeilen = []
        for nr, erg in analysen.items():
            ist_fuehrung = nr.split(".")[0] in fuehrung
            if (schluessel == "regelkreis_fuehrung") != ist_fuehrung:
                continue
            zeichen, lage = ampel_abschnitt(erg, urteile)
            zeilen.append(f"- {nr} {erg.get('titel', '')}: {lage}")
        return "\n".join(zeilen) or "In diesem Regelkreis wurde kein Abschnitt geprüft."
    if schluessel == "gespraeche":
        zeilen = [f"{n['typ']} am {n.get('datum', '')} mit {n.get('quelle', '')}\n{n.get('text', '')}"
                  for n in audit.get("notizen", [])
                  if n.get("typ") in ("Eröffnungsgespräch", "Abschlussgespräch")]
        return "\n\n".join(zeilen) or ("Eröffnungs- und Abschlussgespräch wurden durchgeführt. "
                                       "Inhalte bitte ergänzen.")
    if schluessel == "erfuellung":
        zeilen = []
        for nr, erg in analysen.items():
            zeichen, lage = ampel_abschnitt(erg, urteile)
            wort = {"🟢": "erfüllt", "🟡": "mit Einschränkung erfüllt",
                    "🔴": "nicht erfüllt"}[zeichen]
            zeilen.append(f"- {nr} {erg.get('titel', '')}: {wort}")
        zeilen.append("")
        zeilen.append(f"Geprüft wurden {k['abschnitte']} Normabschnitte. Das Prüfverfahren hat "
                      f"{k['einwaende']} begründete Zweifel erzeugt. Davon wurden "
                      f"{k['ausgeraeumt']} ausgeräumt, {k['abweichungen']} als Abweichung und "
                      f"{k['potenziale']} als Verbesserungspotenzial bestätigt, "
                      f"{k['offen']} bleiben zur Prüfung vor Ort offen.")
        return "\n".join(zeilen)
    if schluessel == "begruendung":
        if k["abweichungen"] == 0:
            return ("Die geprüften Anforderungen sind erfüllt. Für jede Anforderung lagen "
                    "Nachweise vor, deren Aussagekraft im Prüfverfahren gezielt angegriffen "
                    "wurde. Die aufgeworfenen Zweifel konnten mit den vorliegenden Nachweisen "
                    "ausgeräumt werden.")
        zahl = k["abweichungen"]
        wort = "eine Abweichung" if zahl == 1 else f"{zahl} Abweichungen"
        return (f"Es wurde {wort} festgestellt." if zahl == 1
                else f"Es wurden {wort} festgestellt.") + (
                " Die Begründung je Abweichung ist in Abschnitt 6.3 dargestellt. Bitte ergänzen "
                "Sie hier Ihre zusammenfassende Beurteilung der Systemreife, also ob es sich um "
                "Einzelfälle oder um eine systematische Schwäche handelt.")
    if schluessel == "hinweise_text":
        zeilen = ["Hinweise zeigen Entwicklungspotenziale und Risiken auf. Sie sind auf "
                  "Relevanz zu prüfen und dienen der fortlaufenden Verbesserung.", ""]
        for i, u in urteile.items():
            if u.get("urteil") == "Bestätigt: Verbesserungspotenzial":
                e = u.get("einwand", {})
                zeilen.append(f"- {i} (Abschnitt {e.get('abschnitt', '')}): "
                              f"{e.get('begruendung', '')[:400]}")
        if len(zeilen) == 2:
            zeilen.append("- Keine Hinweise.")
        return "\n".join(zeilen)
    if schluessel == "empfehlungen_text":
        return ("Empfehlungen sollen einen Nutzen für die Organisation stiften und stellen keine "
                "Abweichung dar. Die Umsetzung liegt im Ermessen der Organisation.\n"
                "- Bitte ergänzen.")
    if schluessel == "schlusswort":
        offen = [f"{i}: {u.get('einwand', {}).get('pruefung_vor_ort', '')}"
                 for i, u in urteile.items() if u.get("urteil") == "Offen: vor Ort prüfen"]
        text = [audit.get("fazit", "").strip(),
                f"Die Bewertung beruht auf {audit.get('norm', 'der Norm')} und auf den Leitlinien "
                "für Managementsystemaudits nach ISO 19011.",
                f"Das Audit wurde durch das dialektische KI-Prüfverfahren {APP_NAME} unterstützt. "
                "Ein Agententeam hat den Konformitätsnachweis konstruiert, ein zweites hat ihn "
                "angegriffen. Über jeden verbliebenen Zweifel hat der Auditor entschieden und "
                "diese Entscheidung begründet."]
        if offen:
            text.append("Folgende Punkte konnten anhand der Unterlagen nicht abschliessend "
                        "beurteilt werden und sind vor Ort zu prüfen:")
            text += [f"- {o}" for o in offen]
        return "\n".join(t for t in text if t)
    if schluessel == "veraenderungen_org":
        return ("Bitte ergänzen: organisatorische und personelle Veränderungen seit dem letzten "
                "Audit, etwa neue Funktionen, Wechsel in der Prozessverantwortung, "
                "Umstrukturierungen. Bei unveränderter Lage: Die Organisation ist unverändert.")
    if schluessel == "veraenderungen_prozess":
        return ("Bitte ergänzen: neue oder geänderte Prozesse, Produkte, Dienstleistungen, "
                "Anlagen oder IT-Systeme im geprüften Bereich. Bei unveränderter Lage: "
                "Der Prozess ist unverändert.")
    if schluessel == "kennzahlen_entwicklung":
        zahlen = [n for n in audit.get("notizen", []) if n.get("typ") == "Leistungsdaten"]
        if zahlen:
            return "\n".join(f"{n.get('quelle', '')} ({n.get('datum', '')}): {n.get('text', '')}"
                              for n in zahlen)
        return ("Bitte ergänzen: Entwicklung der relevanten Kennzahlen, etwa Mengen, "
                "Reklamationen, Termintreue, Ausschuss oder Durchlaufzeiten, mit Vergleich zur "
                "Vorperiode. Leistungsdaten können unter Plan bei den Nachweisen erfasst "
                "werden, dann "
                "erscheinen sie hier automatisch.")
    if schluessel == "erledigung_vorjahr":
        return ("Bitte ergänzen: Stand der Korrekturmassnahmen aus früheren Audits und Bewertung "
                "ihrer Wirksamkeit. Bei erstmaligem Audit oder ohne Vorbefunde: Im letzten Audit "
                "wurden keine Abweichungen festgestellt.")
    if schluessel == "umgang_hinweise":
        return ("Bitte ergänzen: welche Hinweise aus früheren Berichten aufgegriffen wurden und "
                "welche offen bleiben. Bei erstmaligem Audit: Nicht zutreffend.")
    if schluessel == "qualitaetsinitiativen":
        return ("Bitte ergänzen: Projekte, Evaluationen und Verbesserungen, die die Organisation "
                "aus eigener Initiative angestossen hat, einschliesslich der daraus abgeleiteten "
                "Massnahmen.")
    if schluessel == "selbstbewertung":
        return ("Bitte ergänzen: wie die Organisation ihren eigenen Erfüllungsgrad einschätzt und "
                "ob diese Selbsteinschätzung mit dem Auditbefund übereinstimmt. Eine Abweichung "
                "zwischen Selbstbild und Befund ist selbst ein Auditergebnis.")
    if schluessel == "naechste_pruefung":
        return ("Nächstes internes Audit dieses Prozesses: noch festzulegen.\n"
                "Schwerpunkte: Wirksamkeit der vereinbarten Korrekturmassnahmen"
                + (" sowie die oben offen gebliebenen Punkte." if k["offen"] else "."))
    if schluessel == "antrag":
        if k["abweichungen"] == 0:
            return ("Der Auditor empfiehlt der Leitung, die Konformität des geprüften Prozesses "
                    "zu bestätigen. Auflagen sind nicht erforderlich. Die Hinweise sollten im "
                    "Rahmen der fortlaufenden Verbesserung aufgegriffen werden.")
        return ("Der Auditor empfiehlt der Leitung, die festgestellten Abweichungen mit "
                "Korrekturmassnahmen zu belegen und deren Wirksamkeit innerhalb der "
                "vereinbarten Fristen nachzuweisen. Der Bericht ist in der nächsten "
                "Managementbewertung zu behandeln.")
    return ""


def felder_fuellen(audit: dict, nur_leere: bool = True):
    """Belegt die Berichtsfelder vor. Vom Auditor geaenderte Texte bleiben erhalten."""
    felder = audit.setdefault("bericht_felder", {})
    for schluessel, _, _, _ in BERICHT_FELDER:
        if nur_leere and felder.get(schluessel, "").strip():
            continue
        felder[schluessel] = feld_vorbelegung(audit, schluessel)
    return felder


def abweichungen_sortiert(audit: dict):
    """Teilt die bestaetigten Abweichungen in wesentliche und geringfuegige."""
    fest = {f.get("einwand"): f for f in
            audit.get("feststellungen", {}).get("feststellungen", [])}
    mass = {}
    for m in audit.get("massnahmen", []):
        mass.setdefault(m.get("feststellung"), m)
    major, minor = [], []
    for i, u in audit.get("urteile", {}).items():
        if u.get("urteil") != "Bestätigt: Abweichung":
            continue
        e = u.get("einwand", {})
        r = e.get("zertifizierungsrisiko") or {}
        klass = str(r.get("klassifizierung", "")).lower()
        schwer = ("haupt" in klass or "major" in klass or e.get("schwere") == "hoch")
        eintrag = {"id": i, "abschnitt": e.get("abschnitt", ""),
                   "feststellung": fest.get(i, {}).get("feststellung", e.get("begruendung", "")),
                   "anforderung": fest.get(i, {}).get("anforderung", ""),
                   "nachweis": fest.get(i, {}).get("objektiver_nachweis", ""),
                   "urteil_begruendung": u.get("begruendung", ""),
                   "massnahme": mass.get(i, {}).get("beschreibung", ""),
                   "sofort": mass.get(i, {}).get("sofortmassnahme", ""),
                   "verantwortlich": mass.get(i, {}).get("verantwortlich", ""),
                   "termin": mass.get(i, {}).get("termin", ""),
                   "wirksamkeit": mass.get(i, {}).get("wirksamkeitsnachweis", "")}
        (major if schwer else minor).append(eintrag)
    return major, minor


# ---------------------------------------------------------------------------
# Word-Bericht nach dem Schema eines Auditkurzberichts
# ---------------------------------------------------------------------------
def leitfaden_docx(audit: dict) -> bytes:
    """Gespraechsleitfaden fuer den Termin vor Ort als Word-Datei mit Platz fuer Notizen."""
    from docx import Document as NeuesDok
    from docx.shared import Pt

    d = NeuesDok()
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10.5)

    def absatz(text="", fett=False, kursiv=False, groesse=None):
        p = d.add_paragraph()
        lauf = p.add_run(str(text))
        lauf.bold = fett
        lauf.italic = kursiv
        if groesse:
            lauf.font.size = Pt(groesse)
        return p

    absatz("Gesprächsleitfaden für das Audit vor Ort", groesse=18, fett=True)
    absatz(f"{audit.get('titel', '')} · {audit.get('norm', '')}", groesse=12, fett=True)
    absatz(f"Prozess {audit.get('prozess', '')} · Termin {audit.get('termin', '')} · "
           f"Auditor {audit.get('auditor', '')}")
    absatz()
    absatz("Dieser Leitfaden ist ein Entwurf aus der Dokumentenprüfung. Die Fragen zeigen, wo "
           "die Unterlagen offen lassen, ob die Vorgabe auch gelebt wird. Was Sie vor Ort sehen "
           "und hören, entscheidet der Auditor, nicht das System.", kursiv=True)

    for nr in sorted(audit.get("leitfaden", {})):
        block = audit["leitfaden"][nr]
        d.add_page_break()
        absatz(f"Normabschnitt {nr} {block.get('titel', '')}", groesse=14, fett=True)
        for f in block.get("fragen", []):
            absatz()
            absatz(f"{f.get('id', '')} · {f.get('thema', '')}", fett=True)
            absatz(f.get("frage", ""))
            for n in f.get("nachfragen", []) or []:
                d.add_paragraph(str(n), style="List Bullet")
            tab = d.add_table(rows=0, cols=2)
            tab.style = "Table Grid"
            zeilen = [("Wen fragen", f.get("wen_fragen", "")),
                      ("Stichprobe", f.get("stichprobe", "")),
                      ("Worauf achten", f.get("beobachten", "")),
                      ("Gelebt, wenn", f.get("gelebt_wenn", "")),
                      ("Nur auf dem Papier, wenn", f.get("papier_wenn", "")),
                      ("Gefährdung, wenn nicht gelebt", f.get("gefaehrdung", "")),
                      ("Beobachtung vor Ort", beobachtung(audit, f.get("id", "")).get("text", "")),
                      ("Ergebnis", beobachtung(audit, f.get("id", "")).get("ergebnis", ""))]
            for name, wert in zeilen:
                r = tab.add_row().cells
                r[0].text = str(name)
                for p in r[0].paragraphs:
                    for lauf in p.runs:
                        lauf.bold = True
                r[1].text = str(wert or "")
            d.add_paragraph()

    puffer = io.BytesIO()
    d.save(puffer)
    return puffer.getvalue()


def bericht_docx(audit: dict) -> bytes:
    from docx import Document as NeuesDok
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    f = audit.get("bericht_felder", {})
    k = kennzahlen(audit)
    urteile = audit.get("urteile", {})
    major, minor = abweichungen_sortiert(audit)
    d = NeuesDok()

    stil = d.styles["Normal"]
    stil.font.name = "Calibri"
    stil.font.size = Pt(10.5)

    def absatz(text="", stilname=None, fett=False, kursiv=False, groesse=None):
        p = d.add_paragraph(style=stilname) if stilname else d.add_paragraph()
        for i, zeile in enumerate(str(text).split("\n")):
            lauf = p.add_run(("\n" if i else "") + zeile)
            lauf.bold = fett
            lauf.italic = kursiv
            if groesse:
                lauf.font.size = Pt(groesse)
        return p

    def feldtext(schluessel, ersatz="Nicht ausgefüllt."):
        wert = str(f.get(schluessel, "")).strip()
        if not wert:
            absatz(ersatz, kursiv=True)
            return
        for zeile in wert.split("\n"):
            z = zeile.strip()
            if not z:
                continue
            if z.startswith("- "):
                d.add_paragraph(z[2:], style="List Bullet")
            else:
                absatz(z)

    def tabelle(zeilen, breit=None):
        t = d.add_table(rows=0, cols=2)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        for name, wert in zeilen:
            r = t.add_row().cells
            r[0].text = str(name)
            for p in r[0].paragraphs:
                for lauf in p.runs:
                    lauf.bold = True
            r[1].text = str(wert)
        d.add_paragraph()
        return t

    # ---------------- Titelseite
    titel = absatz(f"Bericht zum internen Audit", groesse=20, fett=True)
    titel.alignment = WD_ALIGN_PARAGRAPH.LEFT
    absatz(f"{audit.get('norm', '')}", groesse=14, fett=True)
    absatz()
    absatz(audit.get("titel", ""), groesse=14, fett=True)
    absatz(ORGANISATION or "", groesse=12)
    absatz()
    tabelle([("Audit", audit.get("titel", "")),
             ("Auditierter Prozess oder Bereich", audit.get("prozess", "")),
             ("Normative Grundlage", audit.get("norm", "")),
             ("Geprüfte Normabschnitte", ", ".join(audit.get("kriterien", []))),
             ("Bedeutung und Risiko des Prozesses", audit.get("risiko", "")),
             ("Audittermin", audit.get("termin", "")),
             ("Auditor", audit.get("auditor", "")),
             ("Unparteilichkeit bestätigt",
              "ja, der Auditor prüft nicht die eigene Arbeit"
              if audit.get("unparteilich") else "nicht bestätigt"),
             ("Auditierte Funktion", audit.get("auditierte", "")),
             ("Berichtsdatum", f"{date.today():%d.%m.%Y}"),
             ("Auditkennung", audit.get("id", ""))])
    absatz("Verfasser", fett=True)
    absatz(audit.get("auditor", ""))
    d.add_page_break()

    # Inhaltsverzeichnis als Word-Feld. Word fuellt es beim Oeffnen mit F9.
    d.add_heading("Inhaltsverzeichnis", level=1)
    p = d.add_paragraph()
    lauf = p.add_run()
    anfang = OxmlElement("w:fldChar")
    anfang.set(qn("w:fldCharType"), "begin")
    anweisung = OxmlElement("w:instrText")
    anweisung.set(qn("xml:space"), "preserve")
    anweisung.text = r'TOC \o "1-3" \h \z \u'
    trenner = OxmlElement("w:fldChar")
    trenner.set(qn("w:fldCharType"), "separate")
    platzhalter = OxmlElement("w:t")
    platzhalter.text = ("Inhaltsverzeichnis. In Word mit der rechten Maustaste darauf klicken "
                        "und Felder aktualisieren wählen.")
    schluss = OxmlElement("w:fldChar")
    schluss.set(qn("w:fldCharType"), "end")
    for teil in (anfang, anweisung, trenner, platzhalter, schluss):
        lauf._r.append(teil)
    d.add_page_break()

    # ---------------- 1 Allgemeine Informationen
    d.add_heading("1  Allgemeine Informationen", level=1)
    d.add_heading("1.1  Zielsetzung des Audits", level=2)
    feldtext("ziel_audit")
    d.add_heading("1.2  Auditierter Prozess oder Bereich", level=2)
    feldtext("dienstleistung")
    d.add_heading("1.3  Organisation und Ansprechpartner", level=2)
    feldtext("institution")
    d.add_heading("1.4  Auditzeitplan", level=2)
    feldtext("zeitplan")
    d.add_heading("1.5  Geprüfter Standort", level=2)
    feldtext("standort")
    d.add_heading("1.6  Eingereichte Unterlagen", level=2)
    feldtext("unterlagen")
    d.add_heading("1.7  Haftungsausschluss", level=2)
    feldtext("haftung")
    d.add_heading("1.8  Verteiler", level=2)
    feldtext("verteiler")

    d.add_heading("2  Geltungsbereich", level=1)
    feldtext("geltungsbereich")
    d.add_heading("3  Begriffe", level=1)
    feldtext("begriffe")

    # ---------------- 4 Auswertung
    d.add_heading("4  Auswertung der Organisation und des Prozesses", level=1)
    for schluessel, nummer, name in [
            ("veraenderungen_org", "4.1", "Wichtige Veränderungen in der Organisation"),
            ("veraenderungen_prozess", "4.2", "Veränderungen am Prozess oder Angebot"),
            ("kennzahlen_entwicklung", "4.3", "Entwicklung der relevanten Kennzahlen"),
            ("erledigung_vorjahr", "4.4", "Erledigungsnachweise früherer Korrekturmassnahmen"),
            ("umgang_hinweise", "4.5", "Umgang mit Hinweisen aus früheren Berichten"),
            ("qualitaetsinitiativen", "4.6", "Eigene Qualitätsinitiativen"),
            ("selbstbewertung", "4.7", "Selbstbewertung der Organisation")]:
        d.add_heading(f"{nummer}  {name}", level=2)
        feldtext(schluessel)

    # ---------------- 5 Auditprozess
    d.add_heading("5  Zusammenfassung des Auditprozesses und der geprüften Inhalte", level=1)
    d.add_heading("5.1  Ausführungen", level=2)
    d.add_heading("5.1.1  Regelkreis Führung und Qualitätsmanagement", level=3)
    feldtext("regelkreis_fuehrung")
    d.add_heading("5.1.2  Regelkreis Leistungserbringung und Unterstützung", level=3)
    feldtext("regelkreis_leistung")
    d.add_heading("5.1.3  Angewandtes Prüfverfahren", level=3)
    angreifer = []
    for erg in audit.get("analysen", {}).values():
        for name in erg.get("teams", []):
            if name not in angreifer:
                angreifer.append(name)
    if not angreifer:
        angreifer = [agent_name(audit, "contra")]
    gegner = (angreifer[0] if len(angreifer) == 1
              else ", ".join(angreifer[:-1]) + " und " + angreifer[-1])
    absatz("Die Dokumentenanalyse wurde durch ein Mehragentensystem unterstützt. "
           f"{agent_name(audit, 'pro')} hat je Normabschnitt den Konformitätsnachweis "
           f"konstruiert. {gegner} "
           + ("hat" if len(angreifer) == 1 else "haben")
           + " diesen Nachweis angegriffen und nach widersprechenden Nachweisen und "
             "ungestützten Annahmen gesucht. Jedes Zitat wurde maschinell gegen die Quelldatei "
             "geprüft. Über jeden verbliebenen Zweifel hat der Auditor entschieden und die "
             "Entscheidung begründet.")
    tabelle([("Geprüfte Normabschnitte", k["abschnitte"]),
             ("Erzeugte Zweifel", k["einwaende"]),
             ("Davon durch den Auditor ausgeräumt", k["ausgeraeumt"]),
             ("Als Abweichung bestätigt", k["abweichungen"]),
             ("Als Verbesserungspotenzial bestätigt", k["potenziale"]),
             ("Zur Prüfung vor Ort offen", k["offen"]),
             ("Zitate in der Quelle wiedergefunden",
              f"{k['zitate_belegt']} von {k['zitate']}")])
    d.add_heading("5.1.4  Prüfung vor Ort", level=3)
    befunde = vorort_befunde(audit)
    if not befunde:
        absatz("Es wurden keine Beobachtungen vor Ort erfasst. Die Aussagen dieses Berichts "
               "stützen sich damit ausschliesslich auf die eingereichten Unterlagen.",
               kursiv=True)
    else:
        absatz(f"Aus der Dokumentenprüfung wurden {k['fragen']} Fragen für den Termin vor Ort "
               f"abgeleitet. Der Auditor hat {k['fragen_geprueft']} davon im Gespräch, durch "
               "Augenschein und anhand von Stichproben geprüft. Die folgenden Beobachtungen "
               "sind objektive Nachweise im Sinne von ISO 19011.")
        tv = d.add_table(rows=1, cols=4)
        tv.style = "Table Grid"
        kopfzeile = tv.rows[0].cells
        for i, name in enumerate(["Abschnitt", "Frage", "Beobachtung", "Ergebnis"]):
            kopfzeile[i].text = name
            for p in kopfzeile[i].paragraphs:
                for lauf in p.runs:
                    lauf.bold = True
        for b in befunde:
            r = tv.add_row().cells
            r[0].text = str(b["abschnitt"])
            r[1].text = str(b["frage"])
            r[2].text = str(b["beobachtung"]) or "keine Notiz"
            r[3].text = str(b["ergebnis"])
        d.add_paragraph()
    d.add_heading("5.2  Eröffnungs- und Abschlussgespräch", level=2)
    feldtext("gespraeche")

    # ---------------- 6 Auditergebnis
    d.add_heading("6  Auditergebnis", level=1)
    d.add_heading("6.1  Erfüllung der geprüften Normabschnitte", level=2)
    t = d.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    kopf = t.rows[0].cells
    for i, name in enumerate(["Normabschnitt", "Ergebnis", "Befund"]):
        kopf[i].text = name
        for p in kopf[i].paragraphs:
            for lauf in p.runs:
                lauf.bold = True
    for nr, erg in audit.get("analysen", {}).items():
        zeichen, lage = ampel_abschnitt(erg, urteile)
        wort = {"🟢": "erfüllt", "🟡": "mit Einschränkung erfüllt",
                "🔴": "nicht erfüllt"}[zeichen]
        r = t.add_row().cells
        r[0].text = f"{nr} {erg.get('titel', '')}"
        r[1].text = wort
        r[2].text = lage
    d.add_paragraph()
    feldtext("erfuellung", "")
    d.add_heading("6.2  Begründung", level=2)
    feldtext("begruendung")

    d.add_heading("6.3  Abweichungen und Korrekturmassnahmen", level=2)
    absatz("Eine Abweichung bezeichnet eine im Audit festgestellte Nichterfüllung einer "
           "Anforderung und ist mit einer Korrekturmassnahme zu belegen. Unterschieden werden "
           "wesentliche und geringfügige Abweichungen.")

    def abweichungsblock(titel, nummer, liste, frist_hinweis, leertext):
        d.add_heading(f"{nummer}  {titel}", level=3)
        if not liste:
            absatz(leertext)
            return
        for a in liste:
            absatz(f"{a['id']} · Normabschnitt {a['abschnitt']}", fett=True)
            tabelle([("Anforderung", a["anforderung"]),
                     ("Objektiver Nachweis", a["nachweis"]),
                     ("Feststellung", a["feststellung"]),
                     ("Begründung des Auditors", a["urteil_begruendung"]),
                     ("Sofortmassnahme", a["sofort"]),
                     ("Korrekturmassnahme", a["massnahme"]),
                     ("Verantwortlich", a["verantwortlich"]),
                     ("Termin", a["termin"]),
                     ("Nachweis der Wirksamkeit", a["wirksamkeit"])])
        absatz(frist_hinweis, kursiv=True)

    absatz("Eine wesentliche Abweichung liegt vor, wenn eine Anforderung systematisch oder "
           "vollständig nicht erfüllt ist. Eine geringfügige Abweichung ist ein Einzelfall oder "
           "eine formale Lücke ohne Systemversagen.", kursiv=True, groesse=9)
    abweichungsblock("Wesentliche Abweichungen mit Korrekturmassnahmen", "6.3.1", major,
                     "Nachweise zu den Korrekturmassnahmen sind termingerecht vorzulegen. "
                     "Die Wirksamkeit wird nachverfolgt.",
                     "Es wurden keine wesentlichen Abweichungen festgestellt.")
    abweichungsblock("Geringfügige Abweichungen mit Korrekturmassnahmen", "6.3.2", minor,
                     "Die Erledigung wird in der Regel beim nächsten Audit geprüft.",
                     "Es wurden keine geringfügigen Abweichungen festgestellt.")
    d.add_heading("6.3.3  Hinweise", level=3)
    feldtext("hinweise_text")
    d.add_heading("6.3.4  Empfehlungen", level=3)
    feldtext("empfehlungen_text")
    d.add_heading("6.4  Schlusswort", level=2)
    feldtext("schlusswort")

    # ---------------- 7 Planung
    d.add_heading("7  Planung der nächsten Überprüfung", level=1)
    feldtext("naechste_pruefung")

    # ---------------- 8 Antrag und Unterschriften
    d.add_heading("8  Antrag an die Leitung und Freigabe", level=1)
    feldtext("antrag")
    absatz()
    tabelle([("Auditor", f"{audit.get('auditor', '')}, Datum: ______________"),
             ("Unterschrift Auditor", "______________________________"),
             ("Zur Kenntnis genommen, Leitung des Bereichs", "______________________________"),
             ("Zur Kenntnis genommen, oberste Leitung", "______________________________")])
    absatz("Dieser Bericht ist als dokumentierte Information aufzubewahren "
           "(ISO 9001, Abschnitte 9.2 und 7.5).", kursiv=True, groesse=9)

    # ---------------- Anhang
    d.add_page_break()
    d.add_heading("Anhang A  Ausgeräumte Zweifel als Nachweis der Prüftiefe", level=1)
    absatz("Diese Zweifel wurden im Prüfverfahren erzeugt und vom Auditor mit Begründung "
           "ausgeräumt. Sie dokumentieren die Tiefe der Prüfung.")
    ausger = [(i, u) for i, u in urteile.items() if u.get("urteil") == "Ausgeräumt"]
    if not ausger:
        absatz("Keine.")
    for i, u in ausger:
        e = u.get("einwand", {})
        absatz(f"{i} · Normabschnitt {e.get('abschnitt', '')}", fett=True)
        absatz(f"Zweifel: {e.get('begruendung', '')}")
        absatz(f"Begründung der Ausräumung: {u.get('begruendung', '')}")
        absatz()

    d.add_heading("Anhang B  Zur Prüfung vor Ort offene Punkte", level=1)
    offen = [(i, u) for i, u in urteile.items() if u.get("urteil") == "Offen: vor Ort prüfen"]
    if not offen:
        absatz("Keine.")
    for i, u in offen:
        e = u.get("einwand", {})
        v = e.get("pruefung_vor_ort")
        v = v if isinstance(v, list) else ([v] if v else [])
        absatz(f"{i} · Normabschnitt {e.get('abschnitt', '')}", fett=True)
        absatz(f"Offener Zweifel: {e.get('begruendung', '')}")
        for schritt in v:
            d.add_paragraph(str(schritt), style="List Bullet")
        absatz()

    puffer = io.BytesIO()
    d.save(puffer)
    return puffer.getvalue()


# ---------------------------------------------------------------------------
# Oberflaeche
# ---------------------------------------------------------------------------
kataloge = normkataloge_laden()
audits = st.session_state.audits

if not st.session_state.get("begruesst"):
    # Der Hinweis gilt als gesehen, sobald er einmal angezeigt wurde. Sonst
    # erscheint er nach dem Schliessen über das Kreuz bei jeder Eingabe erneut.
    st.session_state.begruesst = True
    audix_begruessung()

kopf_text, kopf_hilfe = st.columns([5, 1])
with kopf_text:
    st.markdown(f"## {APP_NAME}")
    st.caption(((ORGANISATION + " · ") if ORGANISATION else "") + "Internes Auditwesen")
st.divider()

with st.sidebar:
    if LOGO.exists():
        st.image(str(LOGO), width=165)
    bereich = st.radio("Bereich", ["Auditprogramm", "Einzelaudit", "Massnahmen"],
                       label_visibility="collapsed")
    # Platzhalter. Wird im Einzelaudit mit Auditauswahl und Fortschritt gefüllt,
    # damit beides direkt unter der Navigation steht.
    audit_platz = st.container()
    st.divider()
    with st.expander("Arbeitsstand sichern und laden"):
        st.caption("Nichts wird dauerhaft gespeichert. Stand als Datei sichern und "
                   "später wieder laden.")
        st.download_button("Arbeitsstand sichern",
                           json.dumps(audits, ensure_ascii=False, indent=1),
                           file_name=f"auditstand_{date.today()}.json",
                           use_container_width=True)
        wieder = st.file_uploader("Arbeitsstand laden", type=["json"], key="restore")
        if wieder is not None and st.button("Übernehmen", use_container_width=True):
            try:
                st.session_state.audits = json.loads(wieder.getvalue().decode("utf-8"))
                st.rerun()
            except (json.JSONDecodeError, UnicodeDecodeError):
                st.error("Die Datei konnte nicht gelesen werden.")
    with st.expander("Was Audix tut"):
        st.markdown(HINWEISE)
        if st.button("Einführung erneut ansehen", use_container_width=True):
            st.session_state.begruesst = False
            st.rerun()
    with st.expander("Technik"):
        st.caption(f"Konstruktionsmodell {MODEL_PRO}\n\nFalsifikationsmodell {MODEL_CONTRA}")
        if not API_KEY:
            st.error("Kein API-Schlüssel in den Secrets hinterlegt.")

# ================================================================ Auditprogramm
if bereich == "Auditprogramm":
    st.session_state["aktuelles_audit"] = None
    with audit_platz:
        if audits:
            st.caption(f"{len(audits)} Audit(s) im Programm")
        else:
            st.caption("Noch kein Audit angelegt")
    with kopf_hilfe:
        audix_hilfe("programm")
    st.subheader("Auditprogramm des Jahres")
    if not kataloge:
        st.error("Kein Normkatalog gefunden. Legen Sie eine JSON-Datei im Ordner normen ab.")
        st.stop()
    with st.form("neu"):
        s1, s2, s3 = st.columns(3)
        titel = s1.text_input("Bezeichnung", "Prozessaudit Einkauf")
        prozess = s2.text_input("Prozess oder Bereich", "Einkauf")
        risiko = s3.selectbox("Bedeutung und Risiko", RISIKO)
        auditor = s1.text_input("Auditor")
        auditierte = s2.text_input("Auditierte Funktion")
        termin = s3.date_input("Geplanter Termin", date.today())
        norm = st.selectbox("Norm", list(kataloge))
        liste = {f"{a['nr']} {a['titel']}": a["nr"] for a in kataloge[norm]["abschnitte"]}
        wahl = st.multiselect("Zu prüfende Normabschnitte (leer lassen, wenn das System "
                              "sie anhand der Nachweise vorschlagen soll)", list(liste))
        unparteilich = st.checkbox("Der Auditor prüft nicht die eigene Arbeit "
                                   "(Unparteilichkeit nach ISO 9001 Abschnitt 9.2)")
        if st.form_submit_button("Audit ins Programm aufnehmen", type="primary"):
            if not unparteilich:
                st.error("Ohne bestätigte Unparteilichkeit darf das Audit nicht geplant werden.")
            else:
                audits.append({"id": date.today().strftime("%Y%m%d") + "-" + uuid.uuid4().hex[:4],
                               "titel": titel, "prozess": prozess, "risiko": risiko,
                               "termin": str(termin), "auditor": auditor,
                               "auditierte": auditierte, "norm": norm,
                               "kriterien": [liste[w] for w in wahl], "ziel": "", "umfang": "",
                               "unparteilich": True, "status": "geplant", "dokumente": {},
                               "notizen": [], "analysen": {}, "urteile": {},
                               "feststellungen": {}, "massnahmen": [], "protokoll": [],
                               "prompt_zusatz": {}, "fazit": ""})
                st.success("Audit angelegt. Weiter im Bereich Einzelaudit.")
                st.rerun()

    if audits:
        st.markdown("**Jahresübersicht**")
        st.dataframe([{"Bezeichnung": a["titel"], "Prozess": a.get("prozess", ""),
                       "Risiko": a.get("risiko", ""), "Termin": a.get("termin", ""),
                       "Auditor": a.get("auditor", ""),
                       "Abschnitte": ", ".join(a.get("kriterien", [])) or "noch offen",
                       "Abweichungen": kennzahlen(a)["abweichungen"],
                       "Offene Massnahmen": kennzahlen(a)["massnahmen_offen"],
                       "Status": a.get("status", "")} for a in audits],
                     use_container_width=True)
        st.markdown("**Abdeckung der Normabschnitte über das Programm**")
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

# ================================================================ Do und Check
elif bereich == "Einzelaudit":
    if not audits:
        with audit_platz:
            st.caption("Noch kein Audit vorhanden.")
        st.warning("Legen Sie zuerst im Auditprogramm ein Audit an.")
        st.stop()
    namen = {f"{a['titel']} · {a['termin']}": a for a in audits}
    with audit_platz:
        st.caption("Gewähltes Audit")
        audit = namen[st.selectbox("Audit", list(namen), label_visibility="collapsed")]
    st.session_state["aktuelles_audit"] = audit
    katalog = kataloge.get(audit.get("norm"))
    if not katalog:
        st.error("Der Normkatalog dieses Audits wurde nicht gefunden.")
        st.stop()
    abschnitte = {a["nr"]: a for a in katalog["abschnitte"]}
    k = kennzahlen(audit)
    audit.setdefault("prompt_zusatz", {})

    with audit_platz:
        erledigt = {1: bool(audit.get("kriterien")),
                    2: bool(quellen_sammeln(audit)),
                    3: bool(audit.get("analysen")),
                    4: k["fragen"] > 0 and k["fragen_geprueft"] == k["fragen"],
                    5: k["einwaende"] > 0 and k["beurteilt"] == k["einwaende"],
                    6: bool(audit.get("feststellungen")),
                    7: bool(audit.get("massnahmen"))}
        schritte = [(SCHRITT_NAME[nr], SCHRITT_PHASE[nr], erledigt[nr])
                    for nr in sorted(SCHRITT_NAME)]
        # Nur der Balken. Welche Aufgabe erledigt ist, steht in der Aufgabenübersicht
        # der gewählten Phase und muss hier nicht ein zweites Mal stehen.
        fertig_anzahl = sum(1 for _, _, f in schritte if f)
        st.progress(fertig_anzahl / len(schritte),
                    text=f"{fertig_anzahl} von {len(schritte)} Schritten erledigt")

    schritt = phasen_schaltflaeche(erledigt)

    # ---------------- 1 Vorbereitung (Plan)
    if schritt == 1:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("vorbereitung")
        with st.form("plan"):
            s1, s2 = st.columns(2)
            audit["titel"] = s1.text_input("Bezeichnung", audit.get("titel", ""))
            audit["prozess"] = s2.text_input("Prozess oder Bereich", audit.get("prozess", ""))
            audit["auditor"] = s1.text_input("Auditor", audit.get("auditor", ""))
            audit["auditierte"] = s2.text_input("Auditierte", audit.get("auditierte", ""))
            audit["termin"] = str(s1.date_input(
                "Termin", date.fromisoformat(audit.get("termin", str(date.today())))))
            audit["status"] = s2.selectbox("Status", STATUS,
                                           index=STATUS.index(audit.get("status", "geplant"))
                                           if audit.get("status") in STATUS else 0)
            audit["ziel"] = st.text_area("Auditziel", audit.get("ziel", ""),
                                         placeholder="Feststellung der Konformität und "
                                                     "Wirksamkeit des Prozesses")
            audit["umfang"] = st.text_area("Umfang und Grenzen", audit.get("umfang", ""),
                                           placeholder="Standort, Zeitraum, betrachtete "
                                                       "Tätigkeiten, Stichprobenumfang")
            liste = {f"{a['nr']} {a['titel']}": a["nr"] for a in katalog["abschnitte"]}
            vorauswahl = [t for t, v in liste.items() if v in audit.get("kriterien", [])]
            neu = st.multiselect("Zu prüfende Normabschnitte", list(liste), default=vorauswahl)
            audit["unparteilich"] = st.checkbox(
                "Unparteilichkeit bestätigt, der Auditor prüft nicht die eigene Arbeit",
                value=bool(audit.get("unparteilich")))
            if st.form_submit_button("Auditplan speichern", type="primary"):
                audit["kriterien"] = [liste[n] for n in neu]
                st.success("Gespeichert.")
        if not audit.get("unparteilich"):
            st.warning("Die Unparteilichkeit ist nicht bestätigt. ISO 9001 Abschnitt 9.2 "
                       "verlangt, dass Auditoren ihre eigene Arbeit nicht auditieren.")

        st.divider()
        st.markdown("**Was in den gewählten Abschnitten geprüft wird**")
        st.caption("Diese Übersicht ist zugleich Ihre Auditcheckliste und die Liste der "
                   "Unterlagen, die Sie beim Fachbereich anfordern sollten.")
        if not audit.get("kriterien"):
            st.info("Noch keine Normabschnitte gewählt. Entweder oben auswählen oder unter Do "
                    "bei der Prüfung vom System vorschlagen lassen.")
        for nr in audit.get("kriterien", []):
            a = abschnitte.get(nr)
            if not a:
                continue
            with st.expander(f"{nr} {a['titel']}", expanded=False):
                st.markdown(f"**Was die Norm verlangt**  \n{a['anforderung']}")
                st.markdown("**Worauf der Auditor achtet**")
                for p in a.get("pruefpunkte", []):
                    st.markdown(f"- {p}")
                if a.get("typische_nachweise"):
                    st.markdown("**Diese Unterlagen sollten Sie bereitstellen**")
                    for n in a["typische_nachweise"]:
                        st.markdown(f"- {n}")
        st.divider()
        st.markdown("**Auditteam zusammenstellen**")
        audit.setdefault("agenten", {})
        besetzung_feld(audit)
        st.markdown("**Das Team dieses Audits**")
        if not audit.get("kriterien"):
            team_zeile(audit, "vorbereitung", "wählt gleich die Normabschnitte aus")
        team_zeile(audit, "pro")
        for schluessel in contra_schluessel(audit):
            team_zeile(audit, schluessel)
        if besetzung(audit)["erwiderung"]:
            team_zeile(audit, "erwiderung")
        st.caption("Die übrigen Teams finden Sie dort, wo sie arbeiten, also bei Vor Ort, "
                   "beim Bericht und bei den Massnahmen.")

        if audit.get("kriterien"):
            zeilen = [f"# Auditplan {audit['titel']}", "",
                      f"Termin {audit['termin']}", f"Prozess {audit['prozess']}",
                      f"Auditor {audit['auditor']}", f"Auditierte {audit['auditierte']}",
                      "", f"## Auditziel\n{audit.get('ziel', '')}",
                      f"## Umfang\n{audit.get('umfang', '')}",
                      f"## Auditkriterien\n{audit['norm']}, Abschnitte "
                      + ", ".join(audit["kriterien"]), "", "## Checkliste"]
            for nr in audit["kriterien"]:
                a = abschnitte.get(nr, {})
                zeilen += [f"### {nr} {a.get('titel', '')}", a.get("anforderung", ""), "",
                           "Prüfpunkte"]
                zeilen += [f"- {p}" for p in a.get("pruefpunkte", [])]
                zeilen += ["", "Bereitzustellende Unterlagen"]
                zeilen += [f"- {n}" for n in a.get("typische_nachweise", [])] + [""]
            st.download_button("Auditplan mit Checkliste herunterladen", "\n".join(zeilen),
                               file_name=f"auditplan_{audit['id']}.md")

    # ---------------- 2 Nachweise (Plan)
    if schritt == 2:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("durchfuehrung")
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
        st.caption("Halten Sie Aussagen möglichst wörtlich fest. Das Falsifikationsteam "
                   "vergleicht sie mit den Vorgabedokumenten und findet so Widersprüche "
                   "zwischen Vorschrift und gelebter Praxis.")
        with st.form("notiz", clear_on_submit=True):
            s1, s2, s3 = st.columns(3)
            typ = s1.selectbox("Art", ["Interviewnotiz", "Beobachtung", "Leistungsdaten",
                                       "Eröffnungsgespräch", "Abschlussgespräch"])
            quelle = s2.text_input("Quelle (Rolle, Arbeitsplatz, Kennzahl)")
            datum = s3.date_input("Datum", date.today())
            text = st.text_area("Notiz", height=120)
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

    # ---------------- 3 Pruefung (Do)
    if schritt == 3:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("pruefung")

        besetzung_feld(audit)
        with st.expander("Wer in diesem Lauf arbeitet"):
            team_zeile(audit, "pro")
            for schluessel in contra_schluessel(audit):
                team_zeile(audit, schluessel)
            if besetzung(audit)["erwiderung"]:
                team_zeile(audit, "erwiderung")
            ohne_schwerpunkt = [s for s in contra_schluessel(audit)
                                if s != "contra" and not agent_rolle(audit, s).strip()]
            if ohne_schwerpunkt:
                st.warning("Ein zugeschaltetes Team hat noch keinen eigenen Schwerpunkt. Ohne "
                           "Schwerpunkt findet es weitgehend dasselbe wie das erste Team.")

        if not quellen_sammeln(audit):
            st.warning("Bitte zuerst unter Plan bei der Aufgabe Nachweise Unterlagen erfassen.")
        else:
            start = st.button("Prüfung starten", type="primary", use_container_width=True)
            if start:
                balken = st.progress(0.0, text="Vorbereitung")
                protokoll = st.empty()
                try:
                    je_abschnitt = 1 + len(contra_schluessel(audit)) + (
                        1 if besetzung(audit)["erwiderung"] else 0)
                    schritte_gesamt = max(
                        1, len(audit.get("kriterien") or [1]) * je_abschnitt + 1)
                    zaehler = {"n": 0}

                    def melden(text):
                        zaehler["n"] += 1
                        balken.progress(min(1.0, zaehler["n"] / schritte_gesamt), text=text)
                        protokoll.caption(text)

                    if not audit.get("kriterien"):
                        melden("Vorbereitungsteam wählt die Normabschnitte")
                        vor = plan_vorschlagen(audit, katalog)
                        gueltig = [n for n in vor.get("kriterien", []) if n in abschnitte]
                        audit["kriterien"] = gueltig or list(abschnitte)[:3]
                        audit["ziel"] = vor.get("ziel", audit.get("ziel", ""))
                        audit["umfang"] = vor.get("umfang", audit.get("umfang", ""))
                        audit["planbegruendung"] = vor.get("begruendungen", {})
                        audit["fehlende_nachweise"] = vor.get("fehlende_nachweise", [])
                        schritte_gesamt = len(audit["kriterien"]) * je_abschnitt + 1
                    for nr in audit["kriterien"]:
                        if nr in audit.get("analysen", {}):
                            continue
                        analyse_durchfuehren(audit, katalog, abschnitte[nr],
                                             besetzung(audit)["erwiderung"], melden)
                    audit["status"] = "durchgeführt"
                    balken.progress(1.0, text="Prüfung abgeschlossen")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))

            if audit.get("analysen"):
                s1, s2, s3, s4 = st.columns(4)
                s1.metric("Geprüfte Abschnitte", k["abschnitte"])
                s2.metric("Gefundene Zweifel", k["einwaende"])
                s3.metric("Zitate belegt", f"{k['zitate_belegt']} / {k['zitate']}")
                s4.metric("Noch zu beurteilen", k["einwaende"] - k["beurteilt"])
                st.caption("Diese Prüfung beruht allein auf Unterlagen. Bei der Aufgabe Vor Ort mache "
                           "ich aus jedem Zweifel eine Frage für das Gespräch, damit Sie prüfen "
                           "können, ob der Prozess auch gelebt wird.")
                if k["zitate"] and k["zitate_belegt"] < k["zitate"]:
                    st.warning(f"{k['zitate'] - k['zitate_belegt']} Zitate wurden in Ihren "
                               "Unterlagen nicht wiedergefunden. Diese Nachweise sind nicht "
                               "verwendbar, sie sind rot markiert.")
            if audit.get("fehlende_nachweise"):
                st.info("Das Vorbereitungsteam vermisst: " + "; ".join(audit["fehlende_nachweise"]))

            with st.expander("Wie Sie die Ergebnisse lesen"):
                st.markdown(LEGENDE)

            urteile = audit.get("urteile", {})
            risiko_zeichen = {"hauptabweichung": "🔴", "nebenabweichung": "🟡",
                              "kein risiko": "🟢"}
            for nr, erg in audit.get("analysen", {}).items():
                einwaende_abschnitt = erg.get("einwaende", {}).get("einwaende", [])
                erwid = {x.get("einwand"): x
                         for x in erg.get("erwiderungen", {}).get("erwiderungen", [])}
                zeichen, lage = ampel_abschnitt(erg, urteile)
                with st.container(border=True):
                    st.markdown(f"## {zeichen} Abschnitt {nr} {erg.get('titel', '')}")
                    st.markdown("**1 Normforderung (Soll-Zustand)**")
                    st.write(erg.get("anforderung", ""))
                    st.markdown(f"**Gesamtbefund des Systems** · {zeichen} {lage}")
                    st.caption("Befund des Systems auf Basis der Unterlagen. Die Feststellung "
                               "trifft der Auditor unter Check bei der Aufgabe Urteil.")
                    for t in erg.get("fall", {}).get("teilaussagen", []):
                        z, kurz = ampel_teilaussage(t, einwaende_abschnitt, urteile)
                        zweifel = [e for e in einwaende_abschnitt if e.get("ziel") == t.get("id")]
                        titel = t.get("pruefpunkt") or t.get("aussage") or t.get("id")
                        with st.expander(f"{z} Prüfpunkt · {titel}", expanded=False):

                            st.markdown("**2 Audit-Befund und Status**")
                            meine = [urteile.get(e["id"], {}).get("urteil") for e in zweifel]
                            meine = [m for m in meine if m]
                            belegte = sum(1 for n in t.get("nachweise", [])
                                          if n.get("verifiziert"))
                            st.markdown(
                                f"- System-Befund · {z} {kurz}\n"
                                f"- Datenbasis · {belegte} von {len(t.get('nachweise', []))} "
                                f"Zitaten in den Unterlagen gefunden, {len(zweifel)} begründete "
                                f"Zweifel, {len(t.get('annahmen', []))} Behauptungen ohne Nachweis\n"
                                f"- Ihr Urteil · "
                                + (", ".join(sorted(set(meine))) if meine
                                   else "noch nicht entschieden"))

                            st.markdown("**3 Analyse und Argumentation (Ist-Zustand)**")
                            st.markdown(f"*Sichtweise {agent_name(audit, 'pro')}*")
                            st.write(t.get("aussage", ""))
                            if t.get("argument"):
                                st.caption(t["argument"])
                            st.markdown(f"*Sichtweise {agent_name(audit, 'contra')}*")
                            if zweifel:
                                for e in zweifel:
                                    st.write(e.get("begruendung", ""))
                                    kurz_e = (erwid.get(e["id"].split("-", 1)[-1])
                                              or erwid.get(e["id"]))
                                    if kurz_e:
                                        st.caption(
                                            f"Erwiderung {agent_name(audit, 'erwiderung')} · "
                                            + ("zugestanden · " if kurz_e.get("zugestanden")
                                               else "") + str(kurz_e.get("erwiderung", "")))
                            else:
                                st.write("Kein Widerspruch gegen diesen Prüfpunkt.")

                            st.markdown("**4 Dokumentenprüfung (die Beweise)**")
                            if not t.get("nachweise"):
                                st.markdown("- Kein Dokument eingereicht, das diesen Prüfpunkt "
                                            "belegt. Das ist der Grund für die rote Ampel.")
                            for n in t.get("nachweise", []):
                                if n.get("verifiziert"):
                                    st.markdown(
                                        f"- Dokument · **{n.get('dokument', '')}**  \n"
                                        f"  Status · 🟢 verwendbar  \n"
                                        f"  Zitat · {n.get('zitat', '')}")
                                else:
                                    st.markdown(
                                        f"- Dokument · **{n.get('dokument', '')}**  \n"
                                        f"  Status · 🔴 nicht verwendbar  \n"
                                        f"  Grund · Das zitierte Textstück wurde im Dokument "
                                        f"nicht gefunden. Es gibt dafür keinen schriftlichen "
                                        f"Nachweis.")
                            for a in t.get("annahmen", []):
                                st.markdown(f"- Behauptung ohne Nachweis · {a}")

                            st.markdown("**5 Audit-Zweifel (Risikoanalyse der Unterlagen)**")
                            if not zweifel:
                                st.markdown("- Keine.")
                            for e in zweifel:
                                schwere = {"hoch": "🔴 hoch", "mittel": "🟡 mittel",
                                           "gering": "🟢 gering"}.get(e.get("schwere"),
                                                                     e.get("schwere", ""))
                                st.markdown(
                                    f"- ID · `{e.get('id')}`  \n"
                                    + (f"  Erhoben von · {e['team']}  \n"
                                       if e.get("team") else "")
                                    + f"  Art · {TYPEN.get(e.get('typ'), e.get('typ'))}  \n"
                                    f"  Schweregrad · {schwere}  \n"
                                    f"  Begründung · {e.get('begruendung', '')}")
                                nachweise_anzeigen(e.get("nachweise"),
                                                   "Kein Gegenbeleg. Der Zweifel stützt sich "
                                                   "auf das Fehlen eines Nachweises.")

                            st.markdown("**6 Empfehlung für die Prüfung vor Ort**")
                            schritte_vo = []
                            for e in zweifel:
                                v = e.get("pruefung_vor_ort")
                                schritte_vo += v if isinstance(v, list) else ([v] if v else [])
                            if schritte_vo:
                                for i, v in enumerate(schritte_vo, start=1):
                                    st.markdown(f"{i}. {v}")
                            else:
                                st.markdown("Keine besondere Prüfung vorgeschlagen.")

                            st.markdown("**7 Risiko im externen Zertifizierungsaudit**")
                            risiken = [e.get("zertifizierungsrisiko") for e in zweifel
                                       if isinstance(e.get("zertifizierungsrisiko"), dict)]
                            if not risiken:
                                st.caption("Das Falsifikationsteam hat dazu nichts geliefert. "
                                           "Schätzen Sie es im Urteil selbst ein.")
                            for r in risiken:
                                klass = str(r.get("klassifizierung", ""))
                                zr = risiko_zeichen.get(klass.lower(), "⚪")
                                st.markdown(
                                    f"- Klassifizierung · {zr} {klass}  \n"
                                    f"  Prognose · {r.get('prognose', '')}  \n"
                                    f"  Konsequenz · {r.get('konsequenz', '')}  \n"
                                    f"  Dringlichkeit · {r.get('dringlichkeit', '')}")
                            st.caption("Diese Einschätzung ist eine Prognose des Systems und "
                                       "bindet keine Zertifizierungsstelle.")
                    if st.button(f"Abschnitt {nr} erneut prüfen", key=f"rm_{nr}"):
                        del audit["analysen"][nr]
                        st.rerun()

            with st.expander("Agentenprotokoll, wer wann was geliefert hat"):
                if not audit.get("protokoll"):
                    st.caption("Noch kein Lauf.")
                for p in audit.get("protokoll", []):
                    st.markdown(f"**{p.get('rolle')}** · {p.get('modell')} · {p.get('zeit')} · "
                                f"{p.get('dauer_sekunden')} Sekunden · "
                                f"{p.get('zeichen_eingabe')} Zeichen Eingabe")
                    if p.get("zusatzanweisung"):
                        st.caption("Ergänzung des Auditors · " + p["zusatzanweisung"])
                    with st.expander("Rohausgabe"):
                        st.code(p.get("ausgabe", ""), language="json")

    # ---------------- 4 Vor Ort (Do)
    if schritt == 4:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("vorort")
        st.markdown("**Vom Papier zur gelebten Praxis**")
        st.caption("Ein Dokument beweist nur, dass etwas beschrieben ist. Hier übersetze ich die "
                   "Zweifel aus der Prüfung in Fragen für das Gespräch, in eine Stichprobe und in "
                   "Merkmale, an denen Sie erkennen, ob der Prozess tatsächlich gelebt wird. "
                   "Das Hinschauen bleibt Ihre Aufgabe, ich bereite es nur vor.")
        if not audit.get("analysen"):
            st.info("Bearbeiten Sie zuerst die Aufgabe Prüfung für mindestens einen Normabschnitt. "
                    "Die Fragen entstehen aus den Zweifeln dieser Prüfung.")
        else:
            audit.setdefault("beobachtungen", {})
            kz = kennzahlen(audit)
            fehlend = [nr for nr in sorted(audit["analysen"])
                       if nr not in audit.get("leitfaden", {})]
            s1, s2 = st.columns([3, 2])
            if s1.button(
                    "Fragen für das Gespräch erzeugen" if not audit.get("leitfaden")
                    else f"Fehlende Abschnitte ergänzen ({len(fehlend)})",
                    type="primary", disabled=not fehlend, use_container_width=True):
                balken = st.progress(0.0, text="Das Leitfadenteam arbeitet")
                try:
                    for i, nr in enumerate(fehlend, start=1):
                        balken.progress((i - 1) / len(fehlend),
                                        text=f"Fragen zu Abschnitt {nr}")
                        leitfaden_erzeugen(audit, nr)
                    balken.progress(1.0, text="Fertig")
                    st.rerun()
                except ModellFehler as fehler:
                    st.error(str(fehler))
            if audit.get("leitfaden"):
                with s2:
                    st.download_button(
                        "Leitfaden als Word herunterladen", leitfaden_docx(audit),
                        file_name=f"auditleitfaden_{audit['id']}.docx",
                        mime="application/vnd.openxmlformats-officedocument."
                             "wordprocessingml.document", use_container_width=True)
                st.progress(kz["fragen_geprueft"] / max(1, kz["fragen"]),
                            text=f"{kz['fragen_geprueft']} von {kz['fragen']} Fragen bearbeitet")
                st.caption("Nehmen Sie den Leitfaden mit in den Termin. Was Sie dort feststellen, "
                           "tragen Sie hier ein. Es erscheint beim passenden Zweifel unter "
                           "Check bei der Aufgabe Urteil und im Auditbericht als objektiver Nachweis.")

            for nr in sorted(audit.get("leitfaden", {})):
                block = audit["leitfaden"][nr]
                fragen = block.get("fragen", [])
                bearbeitet = sum(1 for f in fragen
                                 if beobachtung(audit, f["id"]).get("text", "").strip()
                                 or beobachtung(audit, f["id"]).get("ergebnis",
                                                                    VORORT_ERGEBNIS[0])
                                 != VORORT_ERGEBNIS[0])
                with st.expander(f"{nr} {block.get('titel', '')} · {len(fragen)} Fragen · "
                                 f"{bearbeitet} bearbeitet", expanded=False):
                    if st.button(f"Fragen zu {nr} neu erzeugen", key=f"lfneu_{nr}"):
                        del audit["leitfaden"][nr]
                        st.rerun()
                    if not fragen:
                        st.warning("Zu diesem Abschnitt kam keine verwertbare Antwort zurück. "
                                   "Erzeugen Sie die Fragen bitte erneut.")
                    for f in fragen:
                        fid = f["id"]
                        with st.container(border=True):
                            st.markdown(f"**{f.get('frage', '')}**")
                            kopf = [f"Thema · {f.get('thema', '')}" if f.get("thema") else "",
                                    f"Gesprächspartner · {f.get('wen_fragen', '')}"
                                    if f.get("wen_fragen") else "",
                                    f"geht auf Zweifel {f.get('einwand')} zurück"
                                    if f.get("einwand") else ""]
                            st.caption(" · ".join(x for x in kopf if x))
                            for n in f.get("nachfragen", []) or []:
                                st.markdown(f"- {n}")
                            if f.get("stichprobe"):
                                st.markdown(f"**Stichprobe** · {f['stichprobe']}")
                            if f.get("beobachten"):
                                st.markdown(f"**Worauf achten** · {f['beobachten']}")
                            links, rechts = st.columns(2)
                            links.success(f"Gelebt, wenn · {f.get('gelebt_wenn', '')}")
                            rechts.warning(f"Nur auf dem Papier, wenn · "
                                           f"{f.get('papier_wenn', '')}")
                            if f.get("gefaehrdung"):
                                st.error(f"Gefährdung, wenn das nicht gelebt wird · "
                                         f"{f['gefaehrdung']}")
                            alt = beobachtung(audit, fid)
                            text = st.text_area(
                                "Was Sie vor Ort festgestellt haben",
                                value=alt.get("text", ""), key=f"vo_t_{fid}", height=80,
                                placeholder="Wer hat was gesagt, was haben Sie gesehen, "
                                            "welche Stichprobe haben Sie gezogen")
                            ergebnis = st.radio(
                                "Ergebnis", VORORT_ERGEBNIS, horizontal=True, key=f"vo_e_{fid}",
                                index=VORORT_ERGEBNIS.index(alt["ergebnis"])
                                if alt.get("ergebnis") in VORORT_ERGEBNIS else 0)
                            audit["beobachtungen"][fid] = {"text": text, "ergebnis": ergebnis,
                                                           "abschnitt": nr,
                                                           "einwand": f.get("einwand", "")}
        st.divider()
        st.caption("Dieses Team arbeitet hier")
        team_zeile(audit, "vorort")

    # ---------------- 5 Urteil (Check)
    if schritt == 5:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("urteil")
        einwaende = alle_einwaende(audit)
        if not einwaende:
            st.info("Noch keine Zweifel. Führen Sie zuerst die Prüfung durch.")
        else:
            erw = {x.get("einwand"): x for erg in audit.get("analysen", {}).values()
                   for x in erg.get("erwiderungen", {}).get("erwiderungen", [])}
            st.progress(k["beurteilt"] / max(1, k["einwaende"]),
                        text=f"{k['beurteilt']} von {k['einwaende']} Zweifeln beurteilt")
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
                               f"betrifft Prüfpunkt {e.get('ziel')}"
                               + (f" · erhoben von {e['team']}" if e.get("team") else ""))
                    links, rechts = st.columns(2)
                    with links:
                        st.markdown("**Zweifel des Falsifikationsteams**")
                        st.write(e.get("begruendung"))
                        nachweise_anzeigen(e.get("nachweise"),
                                           "Kein Gegenbeleg, der Einwand stützt sich auf das "
                                           "Fehlen eines Nachweises.")
                    with rechts:
                        st.markdown("**Erwiderung des Konstruktionsteams**")
                        kurz = erw.get(eid.split("-", 1)[-1]) or erw.get(eid)
                        if kurz:
                            st.write(("zugestanden · " if kurz.get("zugestanden") else "")
                                     + str(kurz.get("erwiderung")))
                            nachweise_anzeigen(kurz.get("nachweise"), "Kein weiterer Nachweis.")
                        else:
                            st.write("keine Erwiderung")
                    v = e.get("pruefung_vor_ort")
                    v = v if isinstance(v, list) else ([v] if v else [])
                    if v:
                        st.caption("Vorschlag Prüfung vor Ort · " + " | ".join(str(x) for x in v))
                    r = e.get("zertifizierungsrisiko")
                    if isinstance(r, dict):
                        st.caption(f"Prognose externes Audit · {r.get('klassifizierung', '')} · "
                                   f"{r.get('konsequenz', '')}")
                    for frage, beob in beobachtung_zu_einwand(audit, eid):
                        if not (beob.get("text", "").strip()
                                or beob.get("ergebnis", VORORT_ERGEBNIS[0])
                                != VORORT_ERGEBNIS[0]):
                            continue
                        with st.container(border=True):
                            st.markdown(f"**Vor Ort geprüft · {beob.get('ergebnis', '')}**")
                            st.caption(frage.get("frage", ""))
                            st.write(beob.get("text", "") or "Keine Notiz erfasst.")
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
                balken = st.progress(0.0, text="Feststellungen werden formuliert")
                try:
                    audit["feststellungen"] = llm_json(
                        MODEL_PRO, BERICHT_SYSTEM,
                        f"Auditkriterien {audit['norm']}, Abschnitte "
                        f"{', '.join(audit['kriterien'])}\n\nEntscheidungen des Auditors\n"
                        f"{json.dumps(relevant, ensure_ascii=False)}\n\n"
                        "Beobachtungen des Auditors vor Ort. Sie sind objektive Nachweise aus "
                        "Gespräch und Augenschein und wiegen schwerer als jede Unterlage.\n"
                        f"{json.dumps(vorort_befunde(audit), ensure_ascii=False)}",
                        audit, "bericht")
                    fs = audit["feststellungen"].get("feststellungen", [])
                    balken.progress(0.5, text="Massnahmenteam leitet Verbesserungen ab")
                    if fs:
                        vor = llm_json(MODEL_CONTRA, MASSNAHMEN_SYSTEM,
                                       "Bestätigte Feststellungen\n"
                                       f"{json.dumps(fs, ensure_ascii=False)}", audit,
                                       "massnahmen")
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
                                "status": "offen", "wirksam_geprueft": "", "audit": audit["id"],
                                "audit_titel": audit.get("titel", "")})
                    audit["status"] = "berichtet"
                    balken.progress(1.0, text="Fertig")
                    st.rerun()
                except ModellFehler as f:
                    st.error(str(f))

    # ---------------- 6 Bericht (Check)
    if schritt == 6:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("bericht")
        team_zeile(audit, "bericht", "formuliert die Feststellungen, ohne sie einzustufen")
        if not audit.get("feststellungen"):
            st.info("Der Bericht entsteht, sobald Sie alle Zweifel beurteilt haben.")
        else:
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Abweichungen", k["abweichungen"])
            s2.metric("Potenziale", k["potenziale"])
            s3.metric("Vor Ort offen", k["offen"])
            s4.metric("Ausgeräumt", k["ausgeraeumt"])
            st.markdown("**Ergebnis je Normabschnitt**")
            st.dataframe([{"Abschnitt": f"{nr} {erg.get('titel', '')}",
                           "Ergebnis": " ".join(ampel_abschnitt(erg, audit.get("urteile", {})))}
                          for nr, erg in audit.get("analysen", {}).items()],
                         use_container_width=True, hide_index=True)
            audit["fazit"] = st.text_area(
                "Fazit des Auditors", value=audit.get("fazit", ""), height=100,
                help="Ihre Gesamteinschätzung. Sie fliesst in das Schlusswort des Berichts ein.")

            st.divider()
            st.markdown("### Berichtsformular")
            st.caption("Das Formular folgt dem Aufbau eines Auditberichts. Die Felder sind aus "
                       "den erfassten Daten vorbelegt und überschreibbar. Änderungen stehen "
                       "genau so im Word-Bericht.")
            audit.setdefault("bericht_felder", {})
            s1, s2 = st.columns(2)
            if s1.button("Leere Felder vorbelegen", use_container_width=True):
                felder_fuellen(audit, nur_leere=True)
                st.rerun()
            if s2.button("Alle Felder neu vorbelegen", use_container_width=True,
                         help="Überschreibt auch Ihre eigenen Texte."):
                felder_fuellen(audit, nur_leere=False)
                st.rerun()
            if not any(v.strip() for v in audit["bericht_felder"].values()):
                felder_fuellen(audit, nur_leere=True)

            gruppen = [("1  Allgemeine Informationen",
                        ["ziel_audit", "dienstleistung", "institution", "zeitplan", "standort",
                         "unterlagen", "haftung", "verteiler"]),
                       ("2 und 3  Geltungsbereich und Begriffe",
                        ["geltungsbereich", "begriffe"]),
                       ("4  Auswertung der Organisation und des Prozesses",
                        ["veraenderungen_org", "veraenderungen_prozess",
                         "kennzahlen_entwicklung", "erledigung_vorjahr", "umgang_hinweise",
                         "qualitaetsinitiativen", "selbstbewertung"]),
                       ("5  Auditprozess und geprüfte Inhalte",
                        ["regelkreis_fuehrung", "regelkreis_leistung", "gespraeche"]),
                       ("6  Auditergebnis",
                        ["erfuellung", "begruendung", "hinweise_text", "empfehlungen_text",
                         "schlusswort"]),
                       ("7 und 8  Planung und Antrag",
                        ["naechste_pruefung", "antrag"])]
            beschriftung = {sch: (lab, hilfe, hoehe)
                            for sch, lab, hilfe, hoehe in BERICHT_FELDER}
            for gruppentitel, schluessel_liste in gruppen:
                with st.expander(gruppentitel, expanded=False):
                    for sch in schluessel_liste:
                        lab, hilfe, hoehe = beschriftung[sch]
                        audit["bericht_felder"][sch] = st.text_area(
                            lab, value=audit["bericht_felder"].get(sch, ""), height=hoehe,
                            help=hilfe, key=f"bf_{audit['id']}_{sch}")

            major, minor = abweichungen_sortiert(audit)
            st.markdown("**Abweichungen im Bericht**")
            st.caption("Die Einteilung ergibt sich aus dem Urteil des Auditors und aus der "
                       "Prognose zum externen Audit. Wesentlich heisst, dass eine Anforderung "
                       "systematisch oder vollständig nicht erfüllt ist.")
            s1, s2 = st.columns(2)
            s1.metric("Wesentliche Abweichungen", len(major))
            s2.metric("Geringfügige Abweichungen", len(minor))
            if major or minor:
                st.dataframe(
                    [{"Art": art, "ID": a["id"], "Abschnitt": a["abschnitt"],
                      "Feststellung": a["feststellung"][:160],
                      "Korrekturmassnahme": a["massnahme"][:120],
                      "Verantwortlich": a["verantwortlich"], "Termin": a["termin"]}
                     for art, liste in [("wesentlich", major), ("geringfügig", minor)]
                     for a in liste],
                    use_container_width=True, hide_index=True)
            if any(not a["massnahme"] for a in major + minor):
                st.warning("Zu mindestens einer Abweichung fehlt noch eine Korrekturmassnahme. "
                           "Ergänzen Sie sie unter Act bei den Massnahmen, sonst bleibt die Tabelle "
                           "im Bericht leer.")

            st.divider()
            st.markdown("### Bericht herunterladen")
            s1, s2 = st.columns(2)
            s1.download_button(
                "Auditbericht als Word-Datei", bericht_docx(audit),
                file_name=f"Auditbericht_{audit.get('prozess', 'Prozess')}_{audit['id']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary", use_container_width=True)
            s2.download_button("Bericht als Webseite (HTML)", bericht_html(audit),
                               file_name=f"auditbericht_{audit['id']}.html", mime="text/html",
                               use_container_width=True)
            s1.caption("Vollständiger Bericht nach dem Schema, in Word weiter bearbeitbar.")
            s2.caption("Kurzfassung zum Ansehen, über Drucken als PDF speicherbar.")
            s3, s4 = st.columns(2)
            s3.download_button("Bericht als Text (Markdown)", bericht_markdown(audit),
                               file_name=f"auditbericht_{audit['id']}.md",
                               use_container_width=True)
            trail = {kk: vv for kk, vv in audit.items() if kk != "dokumente"}
            trail["nachweise"] = list(audit.get("dokumente", {}))
            s4.download_button("Audit Trail (JSON)",
                               json.dumps(trail, ensure_ascii=False, indent=2),
                               file_name=f"audit_trail_{audit['id']}.json",
                               use_container_width=True)
            with st.expander("Kurzfassung ansehen"):
                st.markdown(bericht_markdown(audit))

    # ---------------- 7 Massnahmen (Act)
    if schritt == 7:
        hilfe_l, hilfe_r = st.columns([5, 1])
        with hilfe_r:
            audix_hilfe("massnahmen")
        team_zeile(audit, "massnahmen", "schlägt Ursachen und Massnahmen vor")
        if not audit.get("massnahmen"):
            st.info("Noch keine Massnahmen. Sie entstehen nach dem Festschreiben der Urteile.")
        else:
            st.warning("Die Ursachen sind Hypothesen. Sie gehören vor der Freigabe mit den "
                       "Prozessverantwortlichen geprüft.")
            tabelle = st.data_editor(
                audit["massnahmen"], num_rows="dynamic", use_container_width=True,
                key=f"me_{audit['id']}",
                column_config={
                    "id": None, "audit": None, "audit_titel": None,
                    "feststellung": st.column_config.TextColumn("Feststellung", width="small"),
                    "beschreibung": st.column_config.TextColumn("Korrekturmassnahme",
                                                                width="large"),
                    "sofortmassnahme": st.column_config.TextColumn("Sofortmassnahme"),
                    "ursache_hypothese": st.column_config.TextColumn("Ursachenhypothese"),
                    "wirksamkeitsnachweis": st.column_config.TextColumn("Wirksamkeitsnachweis"),
                    "verantwortlich": st.column_config.TextColumn("Verantwortlich"),
                    "termin": st.column_config.TextColumn("Termin", width="small"),
                    "status": st.column_config.SelectboxColumn("Status", options=MSTATUS),
                    "wirksam_geprueft": st.column_config.TextColumn("Wirksamkeit geprüft am")})
            if st.button("Massnahmen übernehmen", type="primary"):
                for zeile in tabelle:
                    zeile.setdefault("id", uuid.uuid4().hex[:6])
                    zeile.setdefault("audit", audit["id"])
                    zeile.setdefault("audit_titel", audit.get("titel", ""))
                    zeile.setdefault("status", "offen")
                    zeile.setdefault("wirksam_geprueft", "")
                audit["massnahmen"] = list(tabelle)
                st.success("Gespeichert. Die Nachverfolgung finden Sie links im Bereich Act.")

# ================================================================ Act
else:
    st.session_state["aktuelles_audit"] = None
    with audit_platz:
        st.caption("Alle Massnahmen aus allen Audits")
    with kopf_hilfe:
        audix_hilfe("nachverfolgung")
    st.subheader("Massnahmen und Nachverfolgung der Wirksamkeit")
    zeilen = [dict(m, Audit=a.get("titel", ""),
                   ueberfaellig="ja" if (m.get("status") in ("offen", "in Umsetzung")
                                         and str(m.get("termin", "")) < str(date.today())) else "")
              for a in audits for m in a.get("massnahmen", [])]
    if not zeilen:
        st.info("Noch keine Massnahmen. Sie entstehen im Einzelaudit nach Ihrem Urteil.")
    else:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Massnahmen gesamt", len(zeilen))
        s2.metric("Offen", sum(1 for z in zeilen if z.get("status") in ("offen", "in Umsetzung")))
        s3.metric("Überfällig", sum(1 for z in zeilen if z["ueberfaellig"]))
        s4.metric("Wirksam bestätigt",
                  sum(1 for z in zeilen if z.get("status") == "wirksam bestätigt"))
        st.dataframe([{"Audit": z["Audit"], "Feststellung": z.get("feststellung", ""),
                       "Massnahme": z.get("beschreibung", ""),
                       "Verantwortlich": z.get("verantwortlich", ""),
                       "Termin": z.get("termin", ""), "Status": z.get("status", ""),
                       "Wirksamkeit geprüft": z.get("wirksam_geprueft", ""),
                       "überfällig": z["ueberfaellig"]} for z in zeilen],
                     use_container_width=True, hide_index=True)
        st.divider()
        st.markdown("**Wirksamkeit einer Massnahme prüfen**")
        auswahl = {f"{z['Audit']} · {z.get('feststellung', '')} · {z.get('beschreibung', '')[:50]}":
                   (z.get("audit"), z.get("id")) for z in zeilen}
        gewaehlt = st.selectbox("Massnahme", list(auswahl))
        aid, mid = auswahl[gewaehlt]
        ziel = next((m for a in audits if a["id"] == aid
                     for m in a.get("massnahmen", []) if m.get("id") == mid), None)
        if ziel:
            st.caption(f"Vereinbarter Wirksamkeitsnachweis · {ziel.get('wirksamkeitsnachweis', '')}")
            s1, s2 = st.columns(2)
            ergebnis = s1.selectbox("Ergebnis der Prüfung",
                                    ["wirksam bestätigt", "nicht wirksam", "umgesetzt"])
            pruefdatum = s2.date_input("Geprüft am", date.today())
            nachweis = st.text_area("Was wurde geprüft und mit welchem Ergebnis",
                                    value=ziel.get("wirksamkeit_notiz", ""), height=90)
            if st.button("Wirksamkeitsprüfung festhalten", type="primary"):
                ziel["status"] = ergebnis
                ziel["wirksam_geprueft"] = str(pruefdatum)
                ziel["wirksamkeit_notiz"] = nachweis
                st.success("Dokumentiert.")
                st.rerun()
