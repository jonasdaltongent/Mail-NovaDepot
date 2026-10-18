# Handleiding voor de leraar — Les 11: De zakelijke e-mail

<!-- Beginvragen (stap 1 van CONTEXT.md), 18-10-2026. Jonas vroeg "Maak de lessen voor de komende week" en koos in de
     chat de naam van de reeks; de rest volgt uit de bestanden:
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 11 van het jaarplan (module 3, eerste les), uur 2 van W08: 3ORLO wo 21 oktober
     (2de uur), 4ORLOb wo 21 oktober (3de uur), 4ORLOa do 22 oktober (7de uur) · vorige les: les 09 (Een veilig account),
     gegeven op di 20-10 (3ORLO, 4ORLOb) en wo 21-10 (4ORLOa), dus vóór les 11 in elke klas → lesstart = ophalen uit les 09 ·
     toestel lokaal 18, Windows 11 · werkdocument CO1_De-zakelijke-e-mail (Jonas, 18-10-2026: een nieuwe reeks
     "Communicatie <n>" voor module 3; opdracht *Communicatie 1 — De zakelijke e-mail*, onderwerp *Digitale competenties*),
     zonder zip of bronbestanden · de e-mail in een nagebootst mailvenster in het werkdocument, want de leerlingen hebben
     geen Gmail (Jonas, 09-10-2026; zoals MW les 3 en MWW les 3); er wordt niets verstuurd · context NovaDepot · na les 07
     (13-10-2026): een les die ze goed kunnen volgen, alles in de browser, één werkplek. -->

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, *Tom vraagt een leverancier een levering te verplaatsen* — module 3 van het jaarplan, les 11 (reeks *Communicatie 1*)
**Toestel:** **lokaal 18, Windows 11** — alles in Chrome, geen bestanden, geen Gmail
**Lesdag:** 3ORLO wo 21 oktober (2de uur) · 4ORLOb wo 21 oktober (3de uur) · 4ORLOa do 22 oktober (7de uur)
**Kernleerplandoel:** `BV2_04.01` (toepassen) · ondersteunend `BK2_02.01`, `BV2_02.08`, `BV2_02.12`
**Deadline:** vrijdag 23 oktober 2026, 20.00 uur · 20 punten
**Pages:** https://jonasdaltongent.github.io/Mail-NovaDepot/ · dia's: …/presentatie.html

Tom Peeters (magazijn) moet bij Zuivelhoeve De Melkkan de levering van bestelling 2026-0418 verplaatsen van donderdag 29 naar
vrijdag 30 oktober, want de koelcel krijgt onderhoud. Hij schreef al een mail, maar die is slordig. De leerling verbetert ze,
zet daarna de zes bouwstenen van een goede mail in de juiste volgorde en schrijft dan zelf de mail **voor Tom**, in het venster
*Nieuw bericht* van het werkdocument (alleen **Aan** en **Onderwerp**; Cc en Bcc volgen in les 12). Er wordt niets verstuurd.
Opbouw en bronnen: `lesvoorbereiding.md`.

---

## 1. Inhoud van het pakket

```text
W08 - Les 11 - ORLO - De zakelijke e-mail/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 10 dia's: 1 Zo start je · 2 lesstart (les 09) · 3 lesdoel · 4 Tom heeft een probleem · 5 het onderwerp (demo) · 6 de zes bouwstenen · 7 controle (aanspreking) · 8 stap 1 samen · 9 kloktijden · 10 exitvraag
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/                      # logo's, lettertypes, mascotte, het filmpje Inleveren, één schermafbeelding (stapjes/s5-codewoord.png)
├── werkdocument/
│   ├── CO1_De-zakelijke-e-mail.docx  # het werkdocument dat de leerling INLEVERT (kopie per leerling)
│   ├── maak_werkdocumenten.py        # maakt het opnieuw (python-docx)
│   ├── maak_stapjes.py               # maakt de uitsnede bij het stapje "Typ het codewoord" uit de pdf van Google Documenten
│   └── huisstijl.py, go-dalton-gent.png
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Geen zip en geen bronbestanden: er is niets te downloaden, uit te pakken of te uploaden.

### 1b. Hoe de lespagina werkt

Zoals les 09: links de **route** (*De mail van Tom*: stap 1–4 · *Klaar om te versturen*: stap 5), midden **één stap** met
*Toon stapjes* per handeling (de les begint in *Kort*, zonder knop, sinds 18-10-2026), rechts de **checklist** met 16 taken. Geen *Hulp nodig?*: grijze
regels onder de handeling en één tip per stap.

- **Startpagina** met drie beelden (de opdracht in Classroom → deze lespagina → je werkdocument CO1), daaronder *Dit maak je
  vandaag*: het skelet van de mail met de zes bouwstenen (balkjes, geen tekst).
- **Infokaders** (`.info`) bij het leesmateriaal, in de volgorde van de vragen, zonder de antwoorden aan te wijzen: *Zo schrijf
  je zakelijk* (stap 2), *De zes bouwstenen* (stap 3), *Wat moet de leverancier weten?* (stap 4).
- **Tekeningen in HTML en CSS** (geen schermafbeeldingen): de opdracht in Classroom, de mail van Tom (`.mail`), de tabellen van
  het werkdocument (`.teken-wd`), het kader met de feiten, het venster *Nieuw bericht* met lichtblauwe vakken (`.mailbouw.werk`,
  met het rode kader op het vak waar de leerling typt) en de mail met de zes bouwstenen (`.mailbouw`, ook op dia 6). In stap 4
  staat onder de handelingen het skelet van de ingevulde mail: *Ziet het bij jou er zo uit?* Sinds 19-10-2026 zeggen de
  handelingen van stap 4 zelf (zichtbaar in *Kort*): het adres uit het kader, het onderwerp van rij 1 mag, en *Je schrijft
  voor Tom: onderaan staat de handtekening van Tom*.
- **Stap 5: inleveren pas na de zelftest, met een codewoord** (zoals les 09). Eerst de zelftest van drie vragen (onderwerp,
  wat Kim Wouters moet weten, slotgroet); daarna het blok *Afwerken en inleveren* (codewoord, controletabel, het filmpje
  *Inleveren*) dat dicht blijft tot de drie vragen juist zijn. De groene regel toont dan het codewoord **ENVELOP**, dat de
  leerling bovenaan *Deel 4 — Nalezen* typt.
- **Extra** (telt niet mee): het antwoord van Kim Wouters (*Vrijdag 30 oktober past, maar pas om 11 uur*), in het werkdocument
  onder *Extra*. De pagina zegt hoe de leerling opnieuw inlevert: **Ongedaan maken**, dan weer **Inleveren**, met een rode
  waarschuwing dat het werk anders niet is ingeleverd.
- `localStorage` bewaart alleen vinkjes en de laatste stap (voorvoegsel `novadepot_mail_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository `jonasdaltongent/Mail-NovaDepot`, Pages op `main`, map `/ (root)`; adres hierboven. Dat adres staat op dia 9, in
`classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling

`classroom.json`: **Communicatie 1 — De zakelijke e-mail**, onderwerp *Digitale competenties*, 3ORLO, 4ORLOa en 4ORLOb,
20 punten, deadline vrijdag 23 oktober 2026, 20.00 uur. Bijlagen: de link naar de lespagina en `CO1_De-zakelijke-e-mail`
(**Een kopie maken voor elke leerling**). Geen zip. Als **concept** klaargezet op 18-10-2026 en teruggelezen (`DRAFT`, zonder tijdstip, deadline vr 23-10 20.00 uur, 20 punten, het
werkdocument als `STUDENT_COPY`). **Ingepland** door de hoofdsessie na het nakijken (18-10-2026, teruggelezen: `DRAFT` met een tijdstip): Classroom
publiceert zelf een half uur vóór elke les: wo 21-10 om 8.45 (3ORLO) en 9.55 (4ORLOb), do 22-10 om 13.15 (4ORLOa).

Instructietekst:

```text
1. Open de lespagina (link) en je werkdocument CO1_De-zakelijke-e-mail.
2. Je schrijft de mail in je werkdocument. Er wordt niets verstuurd.
3. Klik op Inleveren, ten laatste vrijdag 23 oktober 2026 om 20.00 uur.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/CO1_De-zakelijke-e-mail.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads
   converteren* aan) en voeg het toe met **Bijvoegen** › **Drive** › **Een kopie maken voor elke leerling**.
2. Voeg de lespagina toe met **Link**.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 9 klopt.
- [ ] De opdracht staat in de drie klassen, met twee bijlagen (de link en het werkdocument).
- [ ] Een testleerling krijgt een eigen kopie van het werkdocument, met de eigen naam in de titel.
- [ ] Les 09 is in die klas gegeven (dia 2 haalt er een vraag uit op); zo niet: sla dia 2 over.
- [ ] Kijk in de app *Na de les* of er over les 09 iets bij te sturen is.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Werkdocument openen | in de opdracht rechts bij **Jouw werk** op de naam klikken | `_afspraken/klikpaden.md` |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** › status **Ingeleverd** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl); het venster gezien met demoleerling3 (15-10-2026) |
| Opnieuw inleveren (Extra) | **Ongedaan maken**, daarna weer **Inleveren** | zelfde bron en dezelfde proef; zie *Wat nog moet blijken* |

Geen andere knoppen: de mail wordt niet verstuurd, dus er is geen knop *Sturen* of *Verzenden* in de les.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten (1'). Dia 2 lesstart (3'): wat doet Tom eerst bij een onbekende aanmelding (les 09), met vingers; de brug *"Toms account is weer veilig, vandaag schrijft hij er een belangrijke mail mee"*. Dia 3–6 (5'): lesdoel; de situatie en de slordige mail van Tom; **demo, hardop denkend** (het lege onderwerp wordt een precies onderwerp; *hey* wijs je aan zonder het op te lossen); de zes bouwstenen als tekening (lees ze de eerste keer hardop voor). Dia 7 controle (1'): *welke aanspreking?* |
| **Samen** | **3'** | Dia 8: iedereen doet stap 1; duim op. |
| **Keuzewerktijd** | **tot het einde** | Dia 9 blijft staan, met **kloktijden** (`_tools/zet_lesmomenten.py`); stap 2 tot 5, inleveren inbegrepen. Dia 10 in de laatste minuut, met *Vertrek je? Log uit, of Windows + L*. |

Stap 1–5: 3 + 8 + 8 + 13 + 8 = **40 minuten**.

**Eerste rondgang:** heeft iedereen het **eigen** werkdocument open (naam in de titel)? **Tweede rondgang, rond stap 4:** staan het
bestelnummer en de nieuwe datum in de mail? Staat overal *u*? Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert.
Minimumroute: deel 1, deel 3 en inleveren; deel 2 mag later, vóór de deadline.

---

## 4. Verbetersleutel

- **Codewoord** bovenaan *Deel 4 — Nalezen*: **ENVELOP**. Ontbreekt het, dan is de zelftest van stap 5 niet (juist) gemaakt.

De leerling levert alleen het werkdocument in (20 punten: 4 · 3 · 13).

### Deel 1 — Wat is er mis? (4)

Rij 1 staat er al (het lege onderwerp → *Levering bestelling 2026-0418 verplaatsen naar vrijdag 30 oktober*). De leerling
vindt vier fouten (1 punt per rij 2 tot 5, met een betere versie die klopt). Vier zijn genoeg; mogelijke fouten:

| Fout in de mail van Tom | Zo is het beter |
|---|---|
| *hey* | *Beste mevrouw Wouters,* |
| begint met *ik*, in kleine letters | *Mijn naam is Tom Peeters, magazijnverantwoordelijke bij NovaDepot.* |
| welke levering? geen bestelnummer of datum | *bestelling 2026-0418 van donderdag 29 oktober* |
| *later* | een nieuwe datum: *vrijdag 30 oktober, tussen 7 en 10 uur* |
| *kapot!!!* en 😅 | rustig en juist: *de koelcel krijgt onderhoud* |
| *jullie* | *u* |
| *asap* | *Kunt u mij tegen dinsdag 27 oktober laten weten of dat past?* |
| *thx*, *tom* | *Met vriendelijke groeten,* en een handtekening |

### Deel 2 — De bouwstenen (3)

| Bouwsteen | Letter |
|---|---|
| 1 Aanspreking | **B** (voorbeeld, al ingevuld) |
| 2 Inleiding | **E** |
| 3 De vraag | **F** |
| 4 Slot | **D** |
| 5 Slotgroet | **A** |
| 6 Handtekening | **C** |

Alle vijf juist = 3 punten, drie of vier juist = 2, één of twee = 1.

### Deel 3 — De mail van Tom (13)

| Criterium | Punten |
|---|---|
| Aan ingevuld (`verkoop@demelkkan.example`) en een precies onderwerp (welke bestelling, wat Tom vraagt) | 2 |
| *Beste mevrouw Wouters,* en een slotgroet zonder afkorting | 2 |
| de inleiding zegt wie schrijft en waarover | 1 |
| de vraag met het bestelnummer, de afgesproken datum, kort de reden en de nieuwe datum met uur | 4 |
| het slot zegt tegen wanneer Tom een antwoord wil, en bedankt | 2 |
| overal *u*; geen chattaal, hoofdletters om te roepen, *!!!* of emoji's | 1 |
| een handtekening met naam, functie, NovaDepot en toestel 214 | 1 |

Een goede mail (een van vele):

```text
Aan:        verkoop@demelkkan.example
Onderwerp:  Levering bestelling 2026-0418 verplaatsen naar vrijdag 30 oktober

Beste mevrouw Wouters,

Mijn naam is Tom Peeters, magazijnverantwoordelijke bij NovaDepot. Ik mail u over onze bestelling 2026-0418.

U levert die bestelling op donderdag 29 oktober om 7.30 uur. Die dag krijgt onze koelcel onderhoud, en dan kunnen
we geen zuivel ontvangen. Kunt u de levering verplaatsen naar vrijdag 30 oktober, tussen 7 en 10 uur?

Kunt u mij tegen dinsdag 27 oktober laten weten of dat past? Alvast bedankt.

Met vriendelijke groeten,

Tom Peeters
Magazijnverantwoordelijke · NovaDepot · toestel 214
```

De controletabel onderaan het werkdocument heeft dezelfde criteria (samen 20); de leerling typt een x, jij zet de punten in
dezelfde rij.

### Essentiële fouten — geef hier altijd feedback op

- Geen **bestelnummer** of geen **nieuwe datum** (een mail met alleen *later*).
- *Hey* of *Hallo*, of *u* en *je* door elkaar.
- Chattaal, hoofdletters om te roepen of emoji's.
- Geen naam onder de mail.

Verbeteren en opnieuw inleveren mag. De feedback op deel 3 is de voorbereiding op E1 (les 15).

---

## 5. Schermafbeeldingen (optioneel)

Eén: `assets/stapjes/s5-codewoord.png`, een uitsnede van *Deel 4 — Nalezen* uit de pdf van het Google-document. Alle andere
beelden zijn tekeningen in HTML en CSS. Zie `assets/screenshots/LEESMIJ.md`.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt `CO1_De-zakelijke-e-mail.docx` opnieuw. Vereist `python-docx` en `Pillow`. De uitsnede bij het codewoord:
`python3 werkdocument/maak_stapjes.py <pdf van het Google-document>` (vereist `PyMuPDF`).

`werkdocument/huisstijl.py` kreeg op 18-10-2026 een parameter `velden` bij `mailvenster` (standaard de vier velden). Die
verbetering moet ook in `_sjablonen/lespakket/werkdocument/huisstijl.py`; de lesbouwer wijzigt het sjabloon niet.

---

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json`: `BV2_04.01` (kern), `BK2_02.01`, `BV2_02.08` en `BV2_02.12` (ondersteunend). De repository krijgt de `pre-push`
hook uit `_sjablonen/pre-push`.

---

## 8. Wat nog moet blijken in de klas

- **Inleveren pas na de zelftest** (zoals les 09): maken de leerlingen de drie vragen, en typen ze *ENVELOP* bij *Nalezen*?
  Vinden ze het tweede blok *Afwerken en inleveren* onder de zelftest, en leveren ze in zonder te vragen?
- **Haalbaarheid in 40 minuten.** Stap 4 (13 minuten) vraagt zes stukken tekst; niemand was snel klaar in les 07 en les 09 is nog
  niet geëvalueerd. Zakt het tempo in stap 4, dan is de minimumroute (deel 1, deel 3, inleveren) de uitweg.
- **Schrijven in het venster van het werkdocument** in Google Documenten: de vakken groeien mee met een langer antwoord, maar dat is
  alleen in de pdf nagekeken, niet met een leerling aan het toetsenbord. De emoji in de mail van Tom (deel 1) staat in de pdf juist;
  hoe ze op Windows 11 in Chrome verschijnt, is niet gezien.
- **De Extra en *Ongedaan maken*:** de knop komt uit de proef met demoleerling3 (15-10-2026), niet uit deze opdracht. Klikken
  leerlingen op *Ongedaan maken* en vergeten ze weer in te leveren? Kijk in Classroom of er na de deadline opdrachten op
  *Toegewezen* staan. Is het te risicovol, haal dan de laatste regel van de Extra weg (de Extra telt niet mee).
- **Aanspreking en *u*:** schrijven ze *Beste mevrouw Wouters,* of blijft het bij *Beste Kim*? Meng ze *u* en *je*? (dia 10 vraagt het.)
- **Dia 6** is een nieuw soort tekening (een mail met genummerde labels). Hoeft ze meer uitleg dan de notities geven?
- **Dia 2** haalt een vraag uit les 09 op. Over les 09 is er nog geen verslag: de vraag kan te makkelijk of te moeilijk blijken.
- De lespagina en dia's zijn gemeten met scripts en in de ingebouwde browser bekeken (geen horizontaal scrollen op 683, 1024,
  1280 en 1440 px; elke dia binnen de rand), maar niet door een mens op een Windows-toestel: kijk de tekeningen na.
