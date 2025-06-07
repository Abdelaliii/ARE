import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import defaultdict

st.set_page_config(page_title="MBTI-Test", layout="centered")

st.title("MBTI-basierter Persönlichkeitstest (erweitert)")
st.markdown("""
Beantworte die folgenden Aussagen ehrlich auf einer Skala von 1 bis 10:
- 1 = trifft gar nicht zu
- 10 = trifft voll zu

Ziel: Eine präzisere Einordnung deines Persönlichkeitstyps nach dem MBTI-Modell durch eine breitere Datenbasis.
""")

fragen = [
    # Extraversion vs Introversion (E/I) – 10 Fragen
    ("Ich erhalte Energie durch soziale Interaktion.", "E"),
    ("Ich spreche gern in größeren Gruppen.", "E"),
    ("Ich fühle mich nach längerer Zeit mit Menschen oft energetischer.", "E"),
    ("Ich knüpfe leicht neue Kontakte.", "E"),
    ("Ich genieße es, im Mittelpunkt zu stehen.", "E"),
    ("Ich bevorzuge tiefe Gespräche gegenüber Small Talk.", "I"),
    ("Ich lade meine Energie im Alleinsein auf.", "I"),
    ("Ich brauche Zeit für mich, um Erlebtes zu verarbeiten.", "I"),
    ("Ich beobachte lieber, bevor ich mich einbringe.", "I"),
    ("Ich wirke auf andere eher ruhig und zurückhaltend.", "I"),

    # Sensing vs Intuition (S/N) – 10 Fragen
    ("Ich achte stark auf Details in meiner Umgebung.", "S"),
    ("Ich ziehe bewährte Methoden neuen Ideen vor.", "S"),
    ("Ich verlasse mich gern auf meine Erfahrung.", "S"),
    ("Ich bin praktisch veranlagt.", "S"),
    ("Ich arbeite lieber Schritt für Schritt.", "S"),
    ("Ich denke oft über zukünftige Möglichkeiten nach.", "N"),
    ("Ich sehe schnell Zusammenhänge zwischen scheinbar unverbundenen Themen.", "N"),
    ("Ich lasse mich gern von meiner Intuition leiten.", "N"),
    ("Ich beschäftige mich gern mit Visionen und Ideen.", "N"),
    ("Ich bin kreativ und denke gern in Szenarien.", "N"),

    # Thinking vs Feeling (T/F) – 10 Fragen
    ("Ich treffe Entscheidungen primär logisch und sachlich.", "T"),
    ("Ich halte Emotionen selten für objektive Entscheidungsgrundlagen.", "T"),
    ("Ich diskutiere gern auf Basis von Fakten.", "T"),
    ("Ich bin direkt und ehrlich, auch wenn es unbequem ist.", "T"),
    ("Ich analysiere Probleme gern systematisch.", "T"),
    ("Ich achte darauf, niemanden mit meinen Entscheidungen zu verletzen.", "F"),
    ("Ich nehme Rücksicht auf die Emotionen anderer.", "F"),
    ("Ich entscheide oft mit dem Herzen.", "F"),
    ("Ich empfinde Harmonie im Team als sehr wichtig.", "F"),
    ("Ich spüre oft, was andere Menschen brauchen.", "F"),

    # Judging vs Perceiving (J/P) – 10 Fragen
    ("Ich arbeite gerne nach einem klaren Plan.", "J"),
    ("Ich erledige Aufgaben lieber früh als spät.", "J"),
    ("Ich fühle mich wohl mit festen Strukturen.", "J"),
    ("Ich plane gerne im Voraus.", "J"),
    ("Ich mag es, Dinge abzuschließen.", "J"),
    ("Ich bin spontan und offen für Änderungen.", "P"),
    ("Ich arbeite gerne unter flexiblen Bedingungen.", "P"),
    ("Ich fühle mich schnell eingeengt durch zu viele Vorgaben.", "P"),
    ("Ich entscheide gerne situativ.", "P"),
    ("Ich lasse Optionen gerne offen bis zum Schluss.", "P")
]

antworten = defaultdict(list)

for i, (frage, kategorie) in enumerate(fragen):
    st.write(f"**Frage {i+1}:** {frage}")
    antwort = st.slider("", min_value=1, max_value=10, step=1, key=f"frage_{i}")
    antworten[kategorie].append(antwort)

if st.button("Auswertung anzeigen"):
    st.subheader("Dein MBTI-Profil")
    dimensionen = [
        ("E", "I", "Extraversion", "Introversion"),
        ("S", "N", "Empfindung", "Intuition"),
        ("T", "F", "Denken", "Fühlen"),
        ("J", "P", "Urteilen", "Wahrnehmen")
    ]

    mbti_code = ""
    result_text = []
    scores = {}

    for a, b, a_name, b_name in dimensionen:
        a_score = np.mean(antworten[a]) if antworten[a] else 0
        b_score = np.mean(antworten[b]) if antworten[b] else 0
        code = a if a_score >= b_score else b
        mbti_code += code
        result_text.append(f"**{a_name} ({round(a_score,1)}) vs {b_name} ({round(b_score,1)}) → {code}**")
        scores[a_name] = a_score
        scores[b_name] = b_score

    tier_namen = {
        "ISTJ": "Eule – die strukturierte Planerin",
        "ISFJ": "Golden Retriever – der loyale Unterstützer",
        "INFJ": "Schwan – der visionäre Idealist",
        "INTJ": "Fuchs – der strategische Denker",
        "ISTP": "Panther – der flexible Problemlöser",
        "ISFP": "Reh – das ruhige Ästhetik-Wesen",
        "INFP": "Delfin – der kreative Träumer",
        "INTP": "Eule – der tiefgründige Analytiker",
        "ESTP": "Tiger – der impulsive Macher",
        "ESFP": "Papagei – der lebensfrohe Entertainer",
        "ENFP": "Kolibri – der enthusiastische Entdecker",
        "ENTP": "Affe – der ideenreiche Innovator",
        "ESTJ": "Löwe – der durchsetzungsstarke Umsetzer",
        "ESFJ": "Pferd – der hilfsbereite Teamplayer",
        "ENFJ": "Hund – der engagierte Mentor",
        "ENTJ": "Adler – der zielorientierte Anführer"
    }

    st.success(f"Dein Persönlichkeitstyp: **{tier_namen.get(mbti_code, mbti_code)}** ({mbti_code})")
    for zeile in result_text:
        st.write(zeile)

    st.markdown("### Grafische Auswertung")
    df = pd.DataFrame.from_dict(scores, orient='index', columns=['Score'])
    df = df.sort_values(by='Score')
    st.bar_chart(df)

    st.markdown("---")
    st.subheader("Psychologischer Bericht")

    lange_berichte = {
    "ISTJ": "ISTJs sind ruhig, ernsthaft, praktisch und verantwortungsvoll. Sie handeln durchdacht, organisiert und arbeiten zuverlässig. Sie bevorzugen klare Regeln und Strukturen und zeichnen sich durch eine starke Loyalität gegenüber Familie, Tradition und Institutionen aus. Entscheidungen treffen sie faktenbasiert und mit klarem Sinn für Pflicht. Im Beruf wie im Privatleben sind sie verlässlich, planvoll und gewissenhaft.",
    "ISFJ": "ISFJs sind fürsorglich, loyal und hilfsbereit. Sie übernehmen Verantwortung für andere mit großem Pflichtbewusstsein und Sinn für Harmonie. Ihre Stärke liegt in der praktischen Unterstützung und dem Aufbau langfristiger Beziehungen. Sie meiden Konflikte und arbeiten gewissenhaft im Hintergrund, ohne viel Aufhebens um sich zu machen.",
    "INFJ": "INFJs sind visionär, einfühlsam und werteorientiert. Sie haben ein tiefes Verständnis für zwischenmenschliche Beziehungen und streben danach, einen sinnvollen Beitrag zur Welt zu leisten. Sie handeln aus Überzeugung und verfügen über eine klare Vorstellung davon, was richtig ist. Mit Intuition erkennen sie Muster und leiten daraus langfristige Strategien ab.",
    "INTJ": "INTJs sind analytisch, strategisch und unabhängig. Sie setzen hohe Standards, sind organisiert und vorausschauend. Sie bevorzugen Effizienz und Logik und schätzen klare Strukturen. Mit einem starken Willen verfolgen sie ihre Ziele konsequent und arbeiten bevorzugt allein oder mit gleich starken Persönlichkeiten.",
    "ISTP": "ISTPs sind pragmatisch, ruhig und effizient. Sie analysieren Situationen objektiv und handeln lösungsorientiert. Flexibel und anpassungsfähig fühlen sie sich wohl, wenn sie mit ihren Händen oder mit Systemen arbeiten. Sie brauchen Freiheit und bevorzugen klare Fakten gegenüber Theorien.",
    "ISFP": "ISFPs sind zurückhaltend, freundlich und harmoniebedürftig. Sie handeln nach inneren Werten, ohne diese offen zur Schau zu stellen. Sie leben im Moment, achten auf Ästhetik und schätzen ein ruhiges, flexibles Umfeld. Ihre Fürsorglichkeit zeigen sie oft im Stillen.",
    "INFP": "INFPs sind idealistisch, kreativ und empathisch. Sie streben nach Authentizität und einem Leben im Einklang mit ihren Werten. Oft schreiben oder schaffen sie künstlerisch. INFPs haben ein tiefes Bedürfnis nach Bedeutung und setzen sich für Ideale und Gerechtigkeit ein.",
    "INTP": "INTPs sind analytisch, unabhängig und geistreich. Sie streben nach Wissen und logischer Erklärung. Sie lieben es, Probleme zu durchdenken und Theorien zu entwickeln. Praktische Anwendungen sind für sie zweitrangig, solange das Konzept stimmt.",
    "ESTP": "ESTPs sind energiegeladen, lösungsorientiert und spontan. Sie handeln direkt und pragmatisch, leben im Hier und Jetzt und sind durchsetzungsfähig. In Krisensituationen behalten sie die Nerven und fokussieren sich auf das Machbare.",
    "ESFP": "ESFPs sind lebensfroh, warmherzig und genießen das Zusammensein mit anderen. Sie sind spontan, handeln praktisch und schätzen direkte Erfahrungen. Sie begeistern sich für das Schöne und teilen dies gern mit anderen.",
    "ENFP": "ENFPs sind kreativ, kommunikativ und enthusiastisch. Sie lieben es, Ideen zu erkunden, neue Menschen kennenzulernen und andere zu inspirieren. Gleichzeitig sind sie emotional tiefgründig und motiviert von Werten.",
    "ENTP": "ENTPs sind erfinderisch, vielseitig und energiegeladen. Sie lieben Debatten, neue Herausforderungen und unkonventionelle Wege. Ihr Denken ist schnell, flexibel und oft visionär.",
    "ESTJ": "ESTJs sind entscheidungsfreudig, praktisch und durchsetzungsstark. Sie organisieren effizient, übernehmen Verantwortung und setzen klare Strukturen. Ihre Stärke liegt im Planen, Kontrollieren und Umsetzen.",
    "ESFJ": "ESFJs sind fürsorglich, herzlich und pflichtbewusst. Sie übernehmen Verantwortung für ihr Umfeld, schaffen Harmonie und fördern Gemeinschaft. Sie sind organisiert und handeln oft aus dem Wunsch heraus, anderen zu helfen.",
    "ENFJ": "ENFJs sind empathisch, führungsstark und inspirierend. Sie erkennen Potenziale in anderen und fördern diese. Sie sind kommunikativ, kooperativ und handeln im Sinne des Gemeinwohls.",
    "ENTJ": "ENTJs sind zielstrebig, strategisch und durchsetzungsfähig. Sie denken logisch, handeln entschlossen und übernehmen gern die Führung. Sie sind planvoll und streben nach Effizienz und Erfolg."
}


    st.info(lange_berichte.get(mbti_code, "Keine Beschreibung gefunden."))

    st.markdown("**Hinweis:** Dieses Ergebnis gibt eine Tendenz basierend auf den Antworten wieder. Es ersetzt keine professionelle psychologische Diagnose.")

    st.markdown("### Erläuterung zur Bewertungsmethodik")
    st.markdown("""
    Der MBTI-Test basiert auf vier Dichotomien: Extraversion–Introversion, Sensing–Intuition, Thinking–Feeling, und Judging–Perceiving.

    Für jede dieser vier Dimensionen wurden in diesem Test je zehn Aussagen gestellt, die auf einer Skala von 1 (trifft gar nicht zu) bis 10 (trifft voll zu) bewertet werden sollten.

    Die Punkte jeder Aussage werden jeweils der zugehörigen Dimension (z. B. E oder I) zugeschrieben. Am Ende wird der Durchschnittswert für jede Seite der Dimension berechnet. Die Seite mit dem höheren Durchschnitt dominiert die jeweilige Achse und fließt in den finalen vierstelligen MBTI-Typ (z. B. ENFP) ein.

    Das Ergebnis zeigt damit deine aktuellen Tendenzen in den vier Bereichen. Die Methode zielt nicht auf absolute Wahrheiten, sondern auf eine strukturierte Selbsteinschätzung zur persönlichen Reflexion.
    """)

    st.markdown("### Übersicht aller Typen mit Tierbezug und Originalbeschreibung")

    st.markdown("""
**ISTJ – Eule**: ruhig, ernst, erfolgreich durch Gründlichkeit und Zuverlässigkeit. Praktisch, sachlich, logisch, organisiert.  
**ISFJ – Golden Retriever**: freundlich, verantwortungsbewusst, loyal, sorgfältig, sensibel für Bedürfnisse anderer.  
**INFJ – Schwan**: sinn- und zusammenhangsuchend, idealistisch, visionär, hilfsbereit, tiefgründig.  
**INTJ – Fuchs**: unabhängig, strategisch, zukunftsorientiert, hat hohe Standards und klare Ziele.  
**ISTP – Panther**: pragmatisch, logisch, spontan, lösungsorientiert, ruhig und unabhängig.  
**ISFP – Reh**: freundlich, zurückhaltend, hilfsbereit, flexibel, schätzt Harmonie und Ästhetik.  
**INFP – Delfin**: idealistisch, loyal, empathisch, individualistisch, kreativ.  
**INTP – Eule**: theoretisch, logisch, neugierig, unabhängig denkend, analysierend.  
**ESTP – Tiger**: spontan, pragmatisch, energisch, lebt im Hier und Jetzt, entscheidungsfreudig.  
**ESFP – Papagei**: lebensfroh, kontaktfreudig, hilfsbereit, spontan, liebt das Leben und Menschen.  
**ENFP – Kolibri**: inspirierend, enthusiastisch, kreativ, flexibel, menschenorientiert.  
**ENTP – Affe**: neugierig, ideenreich, anpassungsfähig, energiegeladen, diskussionsfreudig.  
**ESTJ – Löwe**: realistisch, strukturiert, organisiert, entscheidungsfreudig, durchsetzungsstark.  
**ESFJ – Pferd**: harmoniebedürftig, hilfsbereit, fürsorglich, praktisch, organisiert.  
**ENFJ – Hund**: empathisch, charismatisch, motivierend, kommunikativ, verantwortlich.  
**ENTJ – Adler**: führungsstark, logisch, entschlossen, zielorientiert, strategisch denkend.
""")

