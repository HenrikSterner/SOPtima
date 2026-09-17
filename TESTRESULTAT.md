# SOPtima – testresultat

## 1. Samlet resultat

Testen er gennemført den 14. september 2026 mod den aktuelle arbejdsmappe efter rettelserne i `app.py`.

| Kontrol | Resultat |
|---|---|
| Python-kompilering | Bestået |
| Automatiske tests | 63 bestået, 1 miljøbetinget skip |
| Streamlit AppTest | Alle seks trin renderer uden exceptions |
| Visuel Edge-rendering | Bestået ved 1440×1100; farver, hierarki, spacing og navigation gennemgået |
| Kritisk tilstandstest | Bestået: ændring af K fjerner eksisterende løsning |
| Demogendannelse | Bestået: kræver eksplicit bekræftelse |
| 350 elever / 85 lærere | Bestået med komplet fordeling og konfliktfri tidsplan |
| Testtid, fuld suite | 20,39 sekunder |
| Streamlit-server | Bestået: health endpoint returnerede HTTP 200 / `ok` |
| `git diff --check` | Bestået |

Den automatiske suite ligger i `tests/test_app.py`. Den kan køres i et normalt udviklingsmiljø med:

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

## 2. Dækkede testområder

De 63 beståede tests dækker følgende dele af `TESTBESKRIVELSE.md`:

- UTF-8, cp1252, danske tegn og automatisk CSV-separator
- alternative kolonnenavne, tomme/ugyldige rækker og Excel-ark
- fag-, niveau-, initial- og ønskernormalisering
- tvetydige lærernavne og dublerede elev-ID'er
- stram og forståelig validering af maksimumstal
- kandidatmatch og faglig gyldighed
- K, individuelt max, tilladt overskridelse og låst max
- I-grænse og korrekt belastning ved dobbeltvejledning
- forskellige lærerønsker og retvisende ønskestatistik
- deterministisk fordeling og ikke-forringelse ved højere algoritmedybde
- uafhængig kontrol af assignments, loads og double loads
- parallelle sessioner, konfliktkæde og konflikttrekant
- ingen lærerdobbeltbookinger i tidsplanen
- ugyldige tider, utilstrækkelig dag, fast og flydende frokost
- tidsplanimport fra senere Excel-ark og atomar afvisning af mangelfulde rækker
- Excel-arkvalg, roundtrip, frosne overskrifter og formleneutralisering
- HTML-escaping og afstemning af tidsplaneksport
- entydige, sikre Word-filnavne i ZIP
- opstart, status, seks-trins-navigation og bekræftet demogendannelse
- standardflowet med 350 fiktive elever og 85 fiktive lærere
- ti genererede fordelingsscenarier med forskellige fag, kapaciteter og ønsker
- belastningsscenariet med 1.000 elever og 200 lærere

## 3. Ydelsesresultat

På det aktuelle testmiljø tog den kombinerede 350/85-test cirka 6,0 sekunder. Den omfatter:

- indlæsning og normalisering af demodata
- 120 fordelingsforsøg
- afsluttende invariantkontrol
- tidsplan med standardtiderne kl. 09.00–15.00, 20 minutter pr. elev, to pauser, fem minutters skift, fast frokost og dybde 80
- kontrol af samtlige lærerbookinger

Den browserløse UI-rejse, der beregner en løsning og derefter ændrer K, tog cirka 7,3 sekunder. Begge ligger under de foreslåede mål i testbeskrivelsen.

Belastningstesten med 1.000 elever, 200 lærere og 20 fordelingsforsøg tog cirka 2,6 sekunder og gav en komplet, invariantkontrolleret fordeling.

## 4. Fejl og risici, der blev rettet

| Område | Rettelse |
|---|---|
| Forældede resultater | Fordelingen signeres nu med data, kapaciteter og samtlige regler. Enhver ændring fjerner fordeling, tidsplan og forberedt Word-eksport og viser årsagen. |
| Forældet tidsplan | Tidsplanen signeres med kilde og alle tidsindstillinger og fjernes straks, når en relevant værdi ændres. |
| Manuel redigering | Gemning kontrollerer nu lærer-ID, fagmatch, K, individuelt max og I. Hårde brud afvises atomart; tilladte overskridelser vises som advarsel. |
| Ønskestatistik | Opfyldelse tæller forskellige afgivne lærerønsker. Ét ønske tæller ikke længere som to, hvis læreren har begge fag. |
| Lærere uden fag | Rækken bevares og vises som en konkret valideringsfejl i stedet for at forsvinde under import. |
| Ugyldige maxværdier | Negative, decimale, tekstlige og for store værdier afvises med en handlingsanvisende fejl. Klare heltal og fx `12 elever` accepteres. |
| Manglende max | Lærere med max 0 vises med navne og forklaring samt henvisning til max-forslaget. |
| Tvetydige ønsker | To lærere med samme navn vælges ikke længere vilkårligt; brugeren skal vælge med initialer. |
| Excel-sikkerhed | Tekst, der starter med en regnearksoperator, neutraliseres før eksport. |
| Word/ZIP | Filnavne indeholder elev-ID, renses og gøres entydige, så elever med samme navn ikke overskriver hinanden. |
| Tidsplanimport | `Navn (initialer)` normaliseres til samme læreridentitet som initialerne alene. |
| Flydende frokost | Frokosten optræder nu også i meget korte planer med kun én vejledningsrunde. |
| Datakilder | Upload af egne elev- eller lærerdata fjerner det modsvarende demodatasæt, så virkelige data ikke blandes skjult med demoen. |
| Demogendannelse | Brugeren skal bekræfte, før aktuelle data erstattes. |
| UI-design | Den aktive `main_v2()` indlæser nu hele designsystemet med ensartede kort, knapper, tabeller, navigation, fokusmarkering, responsivitet og reduceret bevægelse. |
| Navigation | Siderne har ens hero-område, datakildebadge, status, forrige/næste-navigation og et tydeligt næste trin. |
| Resultatvisning | Komplet resultat får eksplicit godkendelsesstatus; kapacitetsbrud har høj visuel prioritet; max og hård grænse vises hver for sig. |
| Tidsplanflow | Datakilde, dagsindstillinger og generering er opdelt i tre tydelige sektioner. En ufuldstændig fordeling deaktiverer generering. |
| Browseridentitet | Sidetitlen er ændret fra demotitel til `SOPtima · vejlederfordeling og tidsplan`. |

## 5. UI-vurdering

Følgende kriterier er verificeret gennem kodegennemgang og Streamlit AppTest:

- præcis én aktiv seks-trins-navigation
- entydigt sidenavn og formål på hvert trin
- synlig datakilde og workflowstatus
- primære handlinger er visuelt og sprogligt tydelige
- eksisterende data bevares ved en fejlet upload
- kritiske handlinger og ugyldige tilstande er blokeret
- alle data-/regelændringer kan spores til en ugyldiggjort løsning
- fejl kan foldes ud som en komplet liste
- tomt resultat og tom liste har en forklarende tomtilstand
- terminologien `Ingen ønsker` og `Ingen vejleder` er adskilt
- fokusmarkering, reduceret bevægelse og responsive layoutregler er tilføjet
- stylingklasserne på den aktive kodevej har tilhørende CSS
- native Streamlit-komponenter og specialdesign bruger samme grønne temafarve

## 6. Ikke gennemførte manuelle accepttests

Følgende dele af testbeskrivelsen kræver mennesker, fysiske browsere eller programmer, som ikke kan simuleres pålideligt i den automatiske terminaltest. De er derfor ikke markeret som bestået:

1. modereret brugertest med mindst fem SOP-koordinatorer/lærere
2. visuel pixel- og layoutkontrol i fysisk Chrome, Edge og eventuelt Firefox/Safari ved alle beskrevne skærmstørrelser
3. fuld tastatur- og skærmlæsertest samt manuel WCAG-kontrastvurdering
4. åbning og udskrivning i den installerede desktopversion af Microsoft Excel og Word
5. længerevarende hukommelsesprofilering og udholdenhedstest over flere timer
6. real DOCX-indholdstest i dette miljø

Den sidste test blev forsøgt, men Windows’ programkontrol blokerede den downloadede `lxml`-DLL, som `python-docx` kræver. ZIP-navngivning og antal filer er testet med en isoleret dokumentstub, men egentlig åbning af de genererede DOCX-filer skal gentages på en maskine, hvor projektets deklarerede `python-docx`-afhængighed kan indlæses.

## 7. Konklusion

Den automatiserbare kerne er bestået efter rettelserne. Programlogikken har nu særskilt slutkontrol, og UI'en beskytter brugeren mod de væsentligste fejltilstande, som den oprindelige testbeskrivelse identificerede.

Programmet er teknisk klar til den sidste faglige og visuelle accept. En endelig produktionsgodkendelse bør først gives efter de seks manuelle punkter ovenfor, især modereret brugertest, fysisk browserkontrol og åbning af Word-eksporten.
