# SOPtima – detaljeret testbeskrivelse

## 1. Dokumentets formål

Dette dokument beskriver, hvordan SOPtima skal testes både teknisk og fra brugerens perspektiv. Testen skal dokumentere, at programmet:

- læser og normaliserer elev- og lærerdata korrekt
- afviser eller tydeligt markerer data, der ikke kan bruges
- fordeler elever i overensstemmelse med fag, ønsker, kapacitet, K-grænse og I-grænse
- viser et retvisende resultat, også efter manuelle ændringer
- opretter en konfliktfri og tidsmæssigt gyldig vejledningsplan
- producerer korrekte og anvendelige eksportfiler
- opleves professionelt, troværdigt, konsekvent og effektivt gennem hele arbejdsgangen

Dokumentet er udarbejdet ud fra den aktive programkode i `app.py` og kravene i `PROGRAMBESKRIVELSE.md`. Den aktive brugerflade er `main_v2()` med seks trin i venstremenuen.

## 2. Kvalitetsmål

SOPtima kan godkendes, når følgende overordnede mål er opfyldt:

1. **Korrekthed:** Ingen elev tildeles en lærer, som ikke underviser i det relevante fag.
2. **Regeloverholdelse:** Globale og individuelle maksimumstal samt I-grænsen håndteres præcist og forklares korrekt.
3. **Dataintegritet:** Skærmbilleder, beregningsresultat, statistik, tidsplan og eksporterede filer viser samme data.
4. **Konfliktfri plan:** En lærer kan aldrig være knyttet til to samtidige vejledninger.
5. **Sporbar tilstand:** Brugeren kan altid se, om data er indlæst, godkendt, ændret eller beregnet, og gamle resultater må ikke fremstå som aktuelle.
6. **Fejltolerance:** Fejl skal forklares i almindeligt dansk, tæt på det sted hvor de kan rettes, uden at programmet går ned.
7. **Professionel brugeroplevelse:** Den primære arbejdsgang skal være tydelig uden forkundskab, og gentagne handlinger skal kunne udføres hurtigt og sikkert.
8. **Tilgængelighed:** Centrale funktioner skal kunne bruges med tastatur, tydelig fokusmarkering, tilstrækkelig kontrast og forståelige labels.

## 3. Omfang

### 3.1 Funktioner, der er omfattet

- indlæsning af CSV, XLSX og XLSM for elev- og lærerdata
- direkte upload af en godkendt XLSX/XLSM-liste til tidsplanen
- redigering af elev- og lærerdata i tabeller
- normalisering af tekst, fag, niveauer, initialer og ønsker
- validering, fagdækning og datagodkendelse
- redigering og forslag til lærernes maksimumstal
- fordelingsalgoritmen og dens scoring
- manuel ændring af tildelte vejledere
- resultatvisning og statistik
- Excel-, Word/ZIP- og HTML-eksport
- tidsplan, pauser, frokost, parallelle runder og lærerfiltrering
- navigation, status, beskeder, visuel kvalitet og tilgængelighed
- sessionsadfærd og ugyldiggørelse af resultater

### 3.2 Afgrænsning

Følgende er ikke selvstændige funktioner i den nuværende løsning og testes derfor kun indirekte:

- login, roller og flerbrugerrettigheder
- permanent database eller lagring mellem sessioner
- automatisk afsendelse af filer til elever eller lærere
- garanti for en matematisk globalt optimal fordeling eller tidsplan
- indholdsmæssig kvalitet af en ekstern Word-skabelon, som ikke følger med programmet

## 4. Testprincipper

Fordelings- og tidsplanalgoritmerne er heuristiske. Testen må derfor ikke låse sig til én bestemt lærerfordeling, medmindre datasættet kun har én lovlig løsning. I stedet kontrolleres faste egenskaber, også kaldet invariants:

- alle tildelte lærere findes i lærerlisten
- hver tildelt lærer dækker det konkrete fag
- belastning tælles én gang pr. elev pr. lærer, også ved dobbeltvejledning
- antallet af dobbeltvejledninger beregnes korrekt
- en løsning må kun overskride individuelt maksimum, når indstillingerne tillader det
- global K må aldrig overskrides af den automatiske fordeling
- højere algoritmedybde må ikke give en dårligere score, når de tidligere forsøg er en delmængde af de nye
- to sessioner med en fælles lærer må ikke ligge samtidig
- alle komplette elever optræder præcis én gang i elevplanen
- resultater og eksporter skal kunne afstemmes række for række

Ved visuelle og brugerorienterede tests er målet ikke blot, at en funktion kan findes. Den skal være let at forstå, give sikker feedback og reducere risikoen for fejl.

## 5. Prioritet og alvor

| Prioritet | Betydning | Eksempel |
|---|---|---|
| P0 | Kritisk for korrekt drift eller data | Forkert fagvejleder, lærerkonflikt i tidsplan, korrupt eksport |
| P1 | Væsentlig for opgaven eller tilliden til resultatet | Gammelt resultat vises efter ændrede regler, misvisende max-status |
| P2 | Mærkbar kvalitets- eller effektivitetsfejl | Uklart næste trin, dårlig tabeloversigt, mangelfuld feedback |
| P3 | Kosmetisk eller mindre forbedring | Uens afstande, mindre tekstvariationer |

En fejl klassificeres som kritisk, hvis den kan føre til en fagligt forkert fordeling, en dobbeltbooket lærer, tab af brugerdata eller et resultat, som brugeren med rimelighed kan opfatte som godkendt, selv om det er forældet.

## 6. Testmiljøer

Mindst følgende kombinationer skal dækkes før frigivelse:

| Område | Minimumsdækning |
|---|---|
| Python | En understøttet Python 3-version i driftsmiljøet |
| Afhængigheder | Laveste og seneste kompatible version inden for intervallerne i `requirements.txt` |
| Desktopbrowser | Seneste stabile Chrome og Edge |
| Ekstra browser | Seneste stabile Firefox eller Safari, hvis denne browser anvendes på skolen |
| Skærme | 1920×1080, 1366×768 og 1280×720 ved 100 % zoom |
| Smal visning | 768 px og ca. 390 px bredde; mindst læsning, navigation og fejlhåndtering skal fungere |
| Zoom | 100 %, 200 % og browserens tekstforstørrelse |
| Regneark | Microsoft Excel samt genåbning med `openpyxl` som automatisk integritetskontrol |
| Sprog/data | Dansk tekst med æ, ø, å samt ældre cp1252- og fejlkodede data |

Testsessionen skal startes fra en ren Streamlit-session. Browseropdatering, ny fane og ny session testes særskilt, fordi programmet gemmer arbejdsdata i `st.session_state` og ikke i en permanent database.

## 7. Testdata

Alle persondata i test skal være fiktive. De eksisterende filer `testdata/sop_test_elever.csv` og `testdata/sop_test_lærere.csv` bruges som realistisk volumenprøve med henholdsvis 350 elever og 85 lærere.

Derudover oprettes små, kontrollerede datasæt:

| Datasæt | Indhold | Formål |
|---|---|---|
| D01 Minimal | 1 elev, 2 fag, 2 lærere med ét fag hver, rigelig kapacitet | Grundforløb med entydigt facit |
| D02 Dobbeltvejledning | 1 elev, samme lærer kan begge fag | Test af I = 0 og I = 1 |
| D03 Parallel | 4 elever og 8 forskellige lærere | Alle fire kan ligge i samme runde |
| D04 Konfliktkæde | 3 elever: AB+CD, CD+EF, EF+GH | Første og tredje kan være parallelle |
| D05 Konflikttrekant | 3 elever: AB+CD, CD+EF, EF+AB | Kræver tre runder |
| D06 Kapacitet | Flere elever end den individuelle og globale kapacitet | Test af max, K, låsning og manglende tildelinger |
| D07 Ønsker | Ingen, ét, to, ukendte, dublerede og delvist kompatible ønsker | Validering og ønskestatistik |
| D08 Normalisering | `Dansk A`, ` dansk `, `DANSK`, æ/ø/å, cp1252 og dobbeltkodet UTF-8 | Matching og læsbarhed |
| D09 Ugyldige rækker | Manglende navn, klasse, fag, lærer-ID, lærerfag og max | Fejlbeskeder og blokering |
| D10 Dubletter | Dublerede initialer, elev-ID'er og elevnavne | Sammenlægning, entydighed og filnavne |
| D11 Tidsgrænser | Kort dag, start = slut, frokost uden for dagen og mange pauser | Tidsvalidering og forklaring |
| D12 Sikker tekst | `<script>`, HTML-tegn, linjeskift og værdier der starter med `=`, `+`, `-` eller `@` | HTML-escaping og regnearksinjektion |
| D13 Volumen | 350/85, 1.000/200 og et kapacitetsmæssigt presset sæt | Ydelse og stabilitet |

For hvert datasæt gemmes input, indstillinger, forventede invariants og faktisk resultat. Til algoritmetests gemmes også score og kørselstid.

## 8. Testforløb og detaljerede testtilfælde

### 8.1 Opstart, navigation og session

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| NAV-01 | P1 | Start programmet i en ren session. | Programmet åbner uden exception. Demodata er tydeligt markeret som fiktive, status stemmer med de faktisk indlæste data, og anbefalet næste trin er entydigt. |
| NAV-02 | P1 | Gå gennem trin 1–6 via venstremenuen i vilkårlig rækkefølge. | Det valgte trin, sidens overskrift og indhold stemmer overens. Ingen indtastede data mistes alene ved navigation. |
| NAV-03 | P1 | Forsøg at åbne Beregn før data er godkendt. | Beregn-knappen er deaktiveret, og beskeden angiver præcist, hvor data skal rettes eller godkendes. |
| NAV-04 | P1 | Redigér elevdata efter en beregning. | Godkendelsesstatus, fordeling og afledt tidsplan ugyldiggøres med det samme. Brugeren kan ikke forveksle det gamle resultat med et aktuelt. |
| NAV-05 | P1 | Redigér lærerdata efter en beregning. | Samme krav som NAV-04. |
| NAV-06 | P0 | Ændr K, I, global K, overskridelsesregel, låsning, parprioritet, klasseprioritet eller algoritmedybde efter en beregning. | Resultatet markeres tydeligt som forældet eller fjernes, og der kræves en ny beregning før resultat og tidsplan kan betragtes som gyldige. |
| NAV-07 | P0 | Ændr et individuelt max efter en beregning. | Fordelingens statistik må ikke fortsat fremstå som beregnet efter det gamle max. Resultatet ugyldiggøres eller genberegnes kontrolleret. |
| NAV-08 | P1 | Klik **Gendan demodata** efter egne data og en tidsplan. | Der vises en klar bekræftelse. Egne data erstattes kun som følge af klikket, og gammel fordeling/tidsplan fjernes. |
| NAV-09 | P1 | Genindlæs browsersiden og åbn programmet i en ny fane/session. | Adfærden svarer til den dokumenterede sessionsmodel. Brugeren advares om, at ikke-downloadede data ikke er permanent gemt. |
| NAV-10 | P2 | Fremprovokér en uventet filfejl og fortsæt navigationen. | Fejlen vises kontrolleret, og resten af sessionen er fortsat anvendelig. Ingen teknisk traceback vises til slutbrugeren. |

### 8.2 Filindlæsning og parsing

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| IMP-01 | P0 | Upload gyldig elev-CSV med semikolon og UTF-8 BOM. | Alle rækker og danske tegn læses korrekt. Kolonner matches, og tomme slutrækker ignoreres. |
| IMP-02 | P0 | Gentag IMP-01 med komma, tabulator og cp1252. | Separator og understøttet encoding genkendes uden datatab. |
| IMP-03 | P0 | Upload gyldig XLSX og XLSM. | Første ark læses som elevdata, og data svarer til regnearket. |
| IMP-04 | P1 | Upload elevark med alternative kolonnenavne som `Navn`, `Hold`, `Subject 1` og `Wish 1`. | De dokumenterede alternativer matches. Ikke-understøttede navne giver en konkret liste over krævede kolonner. |
| IMP-05 | P0 | Upload elevrækker med manglende navn, kun ét fag eller begge fag tomme. | Rækker må ikke forsvinde tavst på en måde, der skjuler en datakvalitetsfejl. Mangler vises med række/elev og kan rettes før godkendelse. |
| IMP-06 | P1 | Upload elev uden klasse/hold. | Værdien håndteres konsekvent som ukendt, vises tydeligt og bevares i resultat/eksport. Det afklares i UI, om klasse er påkrævet eller valgfri. |
| IMP-07 | P0 | Upload lærer-CSV/XLSX med initialer, navn, flere fag, hold og max. | Dublerede rækker for samme initial samles uden dublerede fag/hold; laveste max fra ekstra max-ark anvendes som beskrevet. |
| IMP-08 | P0 | Upload lærer med ID men uden fag. | Læreren og fejlen må ikke forsvinde tavst. Brugeren får en konkret besked og kan rette rækken. Godkendelse blokeres. |
| IMP-09 | P1 | Upload lærerark uden max-kolonne eller med blanke/ugyldige maxværdier. | Programmet forklarer, at max bliver 0 eller kræver rettelse. En komplet fordeling må ikke forventes uden synlig advarsel. |
| IMP-10 | P1 | Test max som `12`, `12 elever`, `12,5`, `-1`, `abc` og `999`. | Kun den besluttede heltalsregel accepteres. Negative, tvetydige og for store værdier afvises eller korrigeres synligt; ingen skjult delstrengstolkning må overraske brugeren. |
| IMP-11 | P1 | Upload fil med forkerte eller manglende kolonner. | Beskeden nævner filtypen, de manglende felter og et konkret næste skridt. Sessionens hidtidige gyldige data bevares. |
| IMP-12 | P1 | Upload tom fil, beskadiget Excel-fil og fil med korrekt filendelse men forkert indhold. | Programmet går ikke ned. Der vises en forståelig fejl uden intern staksporing. |
| IMP-13 | P1 | Upload tidsplanfil med overskrift på første ark. Gentag med titelrækker før overskriften og med tabellen på et senere ark. | Tabellen findes inden for de første 25 rækker på et relevant ark, og elev-, klasse-, lærer- og fagfelter læses korrekt. |
| IMP-14 | P0 | Upload tidsplanfil hvor en udfyldt elevrække mangler en lærer eller et fag. | Hele importen afvises med de relevante rækkenumre; en delvist importeret plan må ikke bruges. |
| IMP-15 | P1 | Upload tidsplanfil med tomme rækker og valgfri bemærkningskolonne. | Tomme rækker ignoreres. Bemærkninger bevares i elevplan, lærerplan, tidslinje og eksport. |
| IMP-16 | P1 | Upload en fil større end den aftalte driftsgrænse. | Systemet reagerer kontrolleret med størrelse/ventetid eller en tydelig grænse; browseren må ikke fryse uden feedback. |

### 8.3 Normalisering og matchning

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| NOR-01 | P0 | Match `Dansk A`, ` dansk ` og `DANSK` mod lærerfaget `Dansk`. | Alle varianter matcher samme kanoniske fag. Den oprindelige læsbare fagtekst bevares i visning og eksport. |
| NOR-02 | P0 | Test niveauerne A, B, C og `A*` på flere fag. | Niveau fjernes kun ved fagmatchning, ikke utilsigtet inde i fagnavne. |
| NOR-03 | P0 | Test `Kommunikation og IT`, `komm/it` og `komm/info`. | Varianterne matcher samme fag. |
| NOR-04 | P0 | Test idéhistorie, idræt og erhvervsøkonomi med korrekte og ældre fejlencodinger. | De kendte varianter matches, og teksten vises uden `Ã`, `Â` eller erstatningstegn. |
| NOR-05 | P1 | Test flere mellemrum, foranstillede/efterstillede mellemrum og store/små bogstaver i lærerinitialer og ønsker. | Initialer og ønsker matches entydigt uden at ændre lærerens læsevenlige navn. |
| NOR-06 | P1 | Angiv ønske som initial, fuldt navn og `Navn (initialer)`. | Alle tre former opløses til samme lærer-ID. |
| NOR-07 | P0 | Opret to lærere med samme normaliserede navn, men forskellige initialer, og brug navnet som ønske. | Systemet må ikke vælge vilkårligt. Brugeren skal vælge en entydig lærer eller få en tvetydighedsfejl. |
| NOR-08 | P1 | Test teknikfag med konkrete retninger og nummererede kursusnavne. | Den konkrete retning bevares og matches efter den beskrevne regel; generisk `Teknikfag` må ikke give falske match. |

### 8.4 Validering og godkendelse

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| VAL-01 | P0 | Ingen elever og/eller ingen lærere. | Datagodkendelse og beregning er blokeret med særskilte, konkrete fejl. |
| VAL-02 | P0 | En elev mangler Fag 1 eller Fag 2. | Eleven identificeres, og beregning er blokeret. |
| VAL-03 | P0 | Et elevfag har ingen mulig lærer. | Elev og fag vises i fejl-/dækningstabellen med status `MANGLER LÆRER`; beregning er blokeret. |
| VAL-04 | P1 | Et ønske peger på et ukendt ID. | Ukendte ønsker samles forståeligt pr. elev, og brugeren kan vælge en gyldig lærer eller `Ingen ønsker`. |
| VAL-05 | P1 | En kendt ønskelærer dækker ingen af elevens fag. | Ønsket markeres som inkompatibelt, og godkendelse blokeres, indtil det rettes eller fjernes. |
| VAL-06 | P2 | En ønskelærer dækker kun ét af de to fag. | Det vises som information, ikke som en falsk fuld dækning. Det fremgår hvilket fag der ikke matches. |
| VAL-07 | P1 | Ret samtlige fejl via tabeller og forslag. | Fejl forsvinder umiddelbart, data kan godkendes, og status i sidepanelet opdateres konsekvent. |
| VAL-08 | P0 | Ændr en godkendt celle til ugyldig og tilbage igen. | Godkendelse ophæves ved ændringen. Den gendannes ikke automatisk alene ved at rette tilbage; brugeren foretager en tydelig ny godkendelse. |
| VAL-09 | P1 | Dubler samme elev-ID eller indsæt to identiske elever. | Systemet opdager eller håndterer dubletten efter en dokumenteret regel. Eksport og Word-filer må ikke skjule eller overskrive en elev. |
| VAL-10 | P1 | Sæt alle relevante læreres max til 0 og godkend data. | Systemet viser før beregning, at fagene har 0 anvendelig kapacitet, eller forklarer efter beregning præcist hvorfor ingen kan tildeles. |

### 8.5 Fordelingsalgoritme og regler

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| FOR-01 | P0 | Beregn D01. | Eleven får de to eneste fagligt gyldige lærere. Begge slots er udfyldt, loads er 1 for hver lærer, og score/statistik stemmer. |
| FOR-02 | P0 | Beregn D02 med I = 0. | Samme lærer må ikke tildeles begge fag. Hvis ingen anden lærer findes, mangler mindst én tildeling. |
| FOR-03 | P0 | Beregn D02 med I = 1. | Samme lærer kan få begge fag; lærerens belastning er 1 og dobbeltbelastning er 1. |
| FOR-04 | P0 | Flere dobbeltvejledninger end I hos samme lærer. | Automatisk fordeling overstiger aldrig I. Manglende tildeling foretrækkes frem for et skjult I-brud. |
| FOR-05 | P0 | Global K til, individuelle max over og under K. | Effektivt normalmaksimum er `min(K, individuelt max)`. Ingen automatisk load overstiger K. |
| FOR-06 | P0 | Tillad overskridelse af individuelt max op til K. | Overskridelse kan kun ske, når global K er aktiv, overskridelse er tilladt og max ikke er låst. Den vises tydeligt i resultatet. |
| FOR-07 | P0 | Slå `Lås lærernes max-tal fast` til. | Individuelle maksimumstal overskrides ikke, heller ikke når overskridelse ellers er tilladt. |
| FOR-08 | P0 | Slå global K fra. | Individuelle max anvendes direkte. Indstillingen om overskridelse op til K har ingen virkning. |
| FOR-09 | P0 | Kapaciteten er utilstrækkelig. | Programmet returnerer kontrolleret den bedst fundne løsning, viser manglende elever/fagpladser og angiver kapacitetsblokering; det må ikke opfinde en fagligt ugyldig lærer. |
| FOR-10 | P1 | To ønsker er begge fagligt og kapacitetsmæssigt mulige. | Begge ønsker opfyldes, når dette ikke bryder en højere prioriteret regel. |
| FOR-11 | P1 | Ønske konkurrerer med komplet dækning eller K. | Komplet og regelgyldig fordeling prioriteres over ønsket i overensstemmelse med scoringsrækkefølgen. |
| FOR-12 | P1 | Samme lærer er det eneste, dublerede ønske og tildeles begge fag. | Statistikken skelner mellem to opfyldte fagslots og to forskellige ønskelærere. Labelen `Begge ønsker` må ikke være misvisende, hvis eleven kun afgav ét ønske. |
| FOR-13 | P1 | Aktivér/deaktivér prioritering af samme lærerpar på et kontrolleret datasæt. | Hårde regler er uændrede. Når score ellers er lige, må aktiv prioritet ikke give dårligere parsamling end inaktiv. |
| FOR-14 | P1 | Aktivér/deaktivér klassesamling. | Hårde regler er uændrede. Når score ellers er lige, må aktiv prioritet ikke give dårligere samling pr. klasse/lærer. |
| FOR-15 | P1 | Beregn samme input og indstillinger to gange. | Resultat og score er reproducerbare med den deterministiske seed-strategi, eller variationen er tydeligt dokumenteret. |
| FOR-16 | P1 | Kør med dybde 20, 40, 120 og 180. | Score kan være den samme eller blive lavere; den må ikke blive højere, fordi højere dybde genbruger de tidligere seed-forsøg. |
| FOR-17 | P0 | Genberegn loads direkte fra assignments. | For hver lærer er vist load lig antallet af unikke elever med læreren, ikke antallet af fagslots. Summen af loads svarer til de unikke lærer-elev-relationer. |
| FOR-18 | P0 | Genberegn scorekomponenter uafhængigt. | K-overlast, ufuldstændige elever, individuel overlast, 0/1/2 opfyldte slots, maksimal load, load-kvadrater og bonusser stemmer med `stats`. |
| FOR-19 | P1 | Test en elev hvor begge fagtekster er samme fag. | Reglen for én eller to vejledere er tydelig og konsekvent; I og belastning håndteres korrekt. |
| FOR-20 | P2 | Kontrollér progressbar ved kort og lang beregning. | Den starter ved 0 %, stiger uden tilbagespring, ender ved 100 % og ledsages af en beskrivelse af den igangværende handling. |

### 8.6 Resultat, manuel redigering og statistik

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| RES-01 | P0 | Afstem resultatets elevtabel mod assignments. | Alle elever optræder én gang, fag og vejledere står i korrekt slot, og manglende værdier vises entydigt. |
| RES-02 | P0 | Afstem lærerbelastning og fagstatistik manuelt. | Antal elever, max, over-max-status, fagkapacitet og ikke-tildelte rækker er matematisk korrekte. |
| RES-03 | P1 | Kontrollér de fem resultatkort for elever med 0, 1 og 2 afgivne ønsker. | Begreberne er entydige. Det fremgår, om tallet handler om afgivne ønsker, opfyldte ønsker eller opfyldte fagslots. |
| RES-04 | P0 | Vælg manuelt en lærer, som ikke underviser i faget. | Ændringen afvises med elev, lærer og fag. Den eksisterende gyldige fordeling bevares samlet. |
| RES-05 | P0 | Lav manuelt en ændring, der overskrider K, individuelt max eller I. | Programmet advarer før gemning eller kræver en eksplicit, dokumenteret tilsidesættelse. Resultatstatus og statistik viser alle brud; tidsplanen må ikke fremstille resultatet som fuldt regelgyldigt. |
| RES-06 | P1 | Fjern en tildeling med `Ingen ønsker`. | Slot bliver reelt tomt, loads/dobbeltloads opdateres, eleven vises under Ikke tildelte, og eksisterende tidsplan ugyldiggøres. |
| RES-07 | P1 | Skift en dobbeltvejledning til to lærere og tilbage igen. | Loads og dobbeltloads opdateres korrekt i begge retninger. |
| RES-08 | P1 | Lav flere manuelle ændringer, hvor én er ugyldig. | Gemning er atomar: enten gemmes alle gyldige ændringer efter tydelig bekræftelse, eller ingen ændringer gemmes. |
| RES-09 | P1 | Søg, sortér og gennemgå tabeller med 350 elever. | Brugeren kan finde en elev/lærer hurtigt, og sortering eller filtrering ændrer ikke koblingen mellem viste rækker og de underliggende assignments. |
| RES-10 | P1 | Sammenlign advarselsbanner, resultatkort, fanen Ikke tildelte og Excel-arket. | Alle steder rapporterer samme antal og samme berørte elever/fag. |

### 8.7 Eksport

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| EXP-01 | P0 | Download standardfordelingen til Excel. | Filen kan åbnes uden reparationsadvarsel. De valgte ark findes én gang, har korrekt rækkefølge, filter og frosset top-række. |
| EXP-02 | P1 | Vælg hver kombination af eksportark, inklusive ingen ark. | Kun valgte ark eksporteres. Ingen valgte ark giver en klar besked og ingen tom/defekt fil. |
| EXP-03 | P0 | Afstem Excel-fordeling, lærerbelastning, fagstatistik og ikke-tildelte med UI. | Data er identiske med den aktuelle, ikke-forældede løsning. IDs og labels er entydige. |
| EXP-04 | P1 | Eksportér aktuelle inputark, genimportér dem og genberegn. | Elev-, lærer- og kapacitetsdata bevares semantisk i en roundtrip. |
| EXP-05 | P1 | Download lærerdata, genimportér og sammenlign fag/hold/max. | Op til 20 fag og 4 hold bevares uden kolonneforskydning eller datatab. |
| EXP-06 | P0 | Generér Word/ZIP for elever med unikke navne, dublerede navne og filnavnstegn. | ZIP kan åbnes, der er præcis én fil pr. elev, filnavne er gyldige og entydige, og hver fil indeholder korrekt elev, klasse, fag og vejledere. |
| EXP-07 | P1 | Generér Word-filer med og uden ekstern skabelon. | Begge kodeveje giver læsbare DOCX-filer. Manglende `python-docx` giver en kontrolleret fejl. |
| EXP-08 | P0 | Eksportér værdier fra D12. | HTML er escaped. Regnearksfelter må ikke kunne udløse uønskede formler ved åbning; hvis formler bevares bevidst, skal risikoen være dokumenteret og accepteret. |
| EXP-09 | P1 | Eksportér tidsplan til Excel og HTML. | Begge formater indeholder elevplan, lærerplan, tidslinje, lærerpar og indstillinger med samme tider og data som UI. |
| EXP-10 | P2 | Udskriv HTML-tidsplan til PDF/print. | Indhold afskæres ikke, sideskift er fornuftige, kontrast bevares, og brede tabeller kan læses eller håndteres tydeligt. |

### 8.8 Tidsplanlogik

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| TID-01 | P0 | Forsøg med en elev, der mangler en af sine to vejledere. | Tidsplanen afvises med besked om, at fordelingen skal rettes først. |
| TID-02 | P0 | Generér D03. | Alle fire elever kan ligge i samme runde, fordi ingen deler lærer. Planen bruger én vejledningsrunde. |
| TID-03 | P0 | Generér D04. | Elever med fælles lærer ligger ikke samtidig; første og tredje elev kan ligge parallelt. |
| TID-04 | P0 | Generér D05. | Planen bruger mindst tre runder, og ingen lærer er dobbeltbooket. |
| TID-05 | P0 | Test dobbeltvejledning, hvor samme lærer står i begge fag. | Læreren optræder kun én gang i konfliktberegning og lærerplan for sessionen, men begge relevante fag fremgår. |
| TID-06 | P0 | Kontrollér alle par af samtidige elevsessioner i et stort resultat. | Snittet mellem deres sæt af lærer-ID'er er tomt. Dette køres som automatisk invariant efter hver plantest. |
| TID-07 | P1 | Kør samme data med tidsplandybde 10, 80 og 300. | Antallet af runder må ikke blive højere ved større dybde. Kørsel stopper korrekt, når den nedre grænse nås. |
| TID-08 | P1 | Slå samling af samme lærerpar til og fra. | Konflikter og antal runder ændres ikke negativt. Med funktionen slået til ligger gentagne lærerpar mindst lige så sammenhængende målt ved den definerede affinitet. |
| TID-09 | P0 | Sæt starttid lig eller senere end sluttid. | Planen afvises med en konkret tidsfejl. Plan over midnat understøttes ikke uden særskilt krav. |
| TID-10 | P0 | Brug en dag, der er for kort. | Fejlen angiver valgt sluttid, tidligst mulige sluttid, manglende minutter, antal runder samt tid til vejledning, skift, pauser og frokost. |
| TID-11 | P0 | Fast frokost helt eller delvist uden for tidsrummet. | Planen afvises. Beskeden forklarer, at hele frokosten skal ligge inden for dagen. |
| TID-12 | P0 | Fast frokost rammer en planlagt session. | Sessionen flyttes, så ingen vejledning overlapper frokosten. Den ekstra ventetid indgår i samlet tidsforbrug og eventuel fejlberegning. |
| TID-13 | P1 | Flydende frokost ved lige og ulige antal runder. | Frokost ligger omtrent midt i planen, har korrekt varighed og overlapper ingen session. |
| TID-14 | P1 | 0 pauser, flere pauser end mellemrum og pausevarighed 0. | Faktisk pauseantal begrænses til mulige placeringer, pauser fordeles uden overlap, og indstillinger viser det faktiske antal. |
| TID-15 | P0 | Kontrollér formel for varighed uden fast frokostventetid. | Samlet tid svarer til runder × elevtid + skift + faktiske pauser + frokost. Sidste elevslut og den viste planlagte tid stemmer. |
| TID-16 | P1 | Filtrér lærerplanen på hver lærer. | Kun den valgte lærers sessioner og frokost vises. Elev, fag, medvejleder og tider stemmer med elevplanen. |
| TID-17 | P1 | Sammenlign tabel og visuel Gantt for én og alle lærere. | Hver blok har korrekt start, slut, varighed og label. Blokke overlapper ikke på samme lærerrække. |
| TID-18 | P1 | Brug godkendt upload som direkte kilde, mens en tidligere beregnet løsning findes. | Det fremgår utvetydigt, hvilken kilde der bruges. Uploaden har forrang som beskrevet, og gammel tidsplan blandes ikke ind. |
| TID-19 | P1 | Brug samme lærer skrevet med forskellig case, mellemrum eller label i uploadfilen. | Samme person normaliseres til én lærer. Der opstår ikke skjult parallel dobbeltbooking på grund af tekstvariation. |
| TID-20 | P1 | Kontrollér tidsplan med 350 elever. | Alle elever optræder præcis én gang, lærerrækker kan afstemmes, og UI/eksport forbliver anvendelig. |

## 9. Professionel og strømlinet brugergrænseflade

### 9.1 UX-acceptkriterier

Brugerfladen godkendes ikke alene ud fra teknisk funktion. Følgende kvaliteter skal være synlige i alle seks trin:

| Område | Acceptkriterium |
|---|---|
| Visuelt hierarki | Hver side har én tydelig hovedoverskrift, en kort formålsforklaring og én primær næste handling. Sekundære handlinger konkurrerer ikke visuelt med den primære. |
| Procesforståelse | Trinnummer, aktuel status, afsluttede trin og næste anbefalede handling er synlige og indbyrdes konsistente. Brugeren kan besøge tidligere trin uden at miste arbejde. |
| Terminologi | `Lærer`, `vejleder`, `max`, `kapacitet`, `K`, `I`, `ønske` og `dobbeltvejledning` bruges konsekvent eller forklares første gang. |
| Datakilde | Det er altid tydeligt, om skærmen viser demodata, egne inputdata, en beregnet løsning eller en uploadet godkendt tidsplanliste. |
| Tilstand | `Ikke indlæst`, `ikke godkendt`, `godkendt`, `ændret`, `beregner`, `beregnet` og `forældet` har tydeligt forskellige tekster og må ikke kun skelnes med farve. |
| Formularer | Labels angiver enhed og gyldigt interval. Standardværdier er fagligt rimelige. Hjælpetekst forklarer konsekvensen frem for blot at gentage labelen. |
| Fejl | Fejl står tæt på problemet, nævner berørt elev/række/fag, forklarer hvorfor og angiver præcis rettelse. Flere fejl kan gennemgås uden en afkortning, der skjuler kritiske problemer. |
| Bekræftelse | Upload, godkendelse, beregning, manuelle ændringer og eksport giver tydelig, kort feedback. En succesbesked må først vises, når handlingen reelt er afsluttet. |
| Tabeller | Kolonneoverskrifter er entydige, vigtigste kolonner er synlige først, lange tekster kan læses, og tomme tabeller har en forklarende tomtilstand. |
| Redigering | Redigerbare og låste celler kan skelnes. Brugeren kan se, om en ændring er gemt, afvist eller endnu ikke anvendt. |
| Resultattillid | Hårde regelbrud står over ønskestatistik. Grøn succes må ikke vises samtidig med uadresserede kritiske brud. |
| Eksport | Downloadknapper fortæller format og indhold. Filnavne er stabile og forståelige. Kun aktuelle resultater kan eksporteres. |
| Responsivt layout | Ved lille bredde stables felter logisk, tekst overlapper ikke, og primære knapper forbliver synlige. Brede tabeller får kontrolleret vandret rulning. |
| Visuel konsistens | Farver, typografi, kort, afrunding, mellemrum, knaptyper, ikoner og feedbackkomponenter følger ét samlet designsystem. Special-CSS er faktisk indlæst på den aktive kodevej. |
| Tilgængelighed | Alle funktioner kan nås med tastatur; fokus er synligt; labels er maskinlæsbare; kontrast sigter mod WCAG 2.2 AA; 200 % zoom giver ikke tab af funktion. |
| Sprog | Dansk er grammatisk, kort og handlingsorienteret. Tekster bruger samme tiltaleform og samme stavning, fx `lærerpar` kontra `vejlederpar`. |
| Oplevet ydelse | Almindelig redigering reagerer straks. Længere beregninger viser fremdrift og låser ikke brugeren i en uklar tilstand. |

### 9.2 UX-testcases

| ID | Pri. | Test og handling | Forventet resultat |
|---|---:|---|---|
| UX-01 | P1 | Bed en ny bruger forklare programmets formål og næste skridt efter 10 sekunder på startsiden. | Brugeren kan korrekt forklare både formål, datakilde og næste handling uden hjælp. |
| UX-02 | P1 | Lad brugeren udføre hele forløbet med egne filer uden instruktion ud over opgaven. | Mindst 90 % af testbrugerne gennemfører uden hjælp og uden kritiske fejl. |
| UX-03 | P1 | Indsæt tre forskellige inputfejl. | Brugeren kan lokalisere og rette alle tre direkte fra beskederne uden at åbne programbeskrivelsen. |
| UX-04 | P1 | Ændr en regel efter beregning og spørg, om resultatet stadig gælder. | 100 % svarer korrekt ud fra UI'ets synlige status; ingen er i tvivl om behovet for genberegning. |
| UX-05 | P1 | Bed brugeren forklare K, individuelt max, tilladt overskridelse og låst max. | Brugeren kan forudsige den effektive grænse i tre konkrete eksempler. Hjælpetekst er tilstrækkelig. |
| UX-06 | P1 | Bed brugeren finde alle ufordelte elever og årsagen. | Opgaven løses på højst to minutter uden manuel sammenligning mellem flere uklare skærme. |
| UX-07 | P1 | Bed brugeren skifte en vejleder manuelt og kontrollere konsekvensen. | Brugeren opdager fag-, kapacitets- og I-konsekvens før eller umiddelbart efter gemning. |
| UX-08 | P2 | Bed brugeren eksportere kun Fordeling og Lærerbelastning. | De korrekte ark vælges og downloades i første forsøg. |
| UX-09 | P1 | Bed en lærer finde sin dagsplan og en elev finde sit tidspunkt. | Begge opgaver løses på højst 30 sekunder i UI eller den eksporterede plan. |
| UX-10 | P1 | Vis en umulig tidsopsætning. | Brugeren kan på baggrund af fejlen vælge en passende ny sluttid eller reducere de rigtige varigheder i første eller andet forsøg. |
| UX-11 | P2 | Gennemgå alle trin ved 1366×768. | Ingen vigtig handling ligger permanent uden for synsfeltet, og lange sider har naturlige sektioner og overskrifter. |
| UX-12 | P1 | Gennemfør kerneflowet udelukkende med tastatur. | Fokusorden følger læseretningen; dropdowns, editorer, faner og knapper kan betjenes; fokusmarkering forsvinder ikke. |
| UX-13 | P1 | Kør automatisk tilgængelighedsscanning og manuel kontrastkontrol. | Ingen kritiske fejl; normal tekst og væsentlige kontroller opfylder det valgte AA-kontrastmål. |
| UX-14 | P2 | Zoom til 200 % og test 390 px bredde. | Ingen tekst eller funktion skjules, overlapper eller kræver todimensional rulning uden for egentlige datatabeller. |
| UX-15 | P2 | Sammenlign hero, sidepanel, statuskort, metrickort, knapper og tabeller på alle trin. | Komponenterne har ensartet design. Ingen HTML-klasse står ustylet på den aktive `main_v2()`-kodevej. |
| UX-16 | P2 | Kontrollér browserfane før og efter skift fra demo til egne data. | Titel og badge beskriver produktet og den aktuelle datakilde professionelt; siden må ikke fortsat hedde `demo med fiktive data` ved arbejde med egne data. |
| UX-17 | P2 | Test succes-, info-, advarsels- og fejlbeskeder sammen. | Farve, ikon og tekstniveau bruges konsekvent; status kan forstås uden farvesyn. |
| UX-18 | P2 | Observer en erfaren bruger ved anden gennemførsel. | Gentaget flow kræver færre handlinger, og brugeren kan springe direkte til relevante trin uden at blive vildledt af procesindikatoren. |

### 9.3 Modereret brugertest

Der gennemføres mindst fem sessioner med personer, som ligner de faktiske brugere: SOP-koordinatorer, studieadministrative medarbejdere eller lærere med planlægningsansvar. Mindst to deltagere må ikke have set programmet før.

Hver deltager løser disse opgaver:

1. identificér om data er demo eller egne data
2. upload elevdata og ret et ukendt ønske
3. upload lærerdata og ret manglende fagdækning
4. forklar og indstil K, I og individuelle maksimumstal
5. beregn og vurder om resultatet kan godkendes
6. lav en manuel ændring og kontrollér konsekvenserne
7. eksportér et specificeret sæt Excel-ark
8. opret en tidsplan med fast frokost
9. find en bestemt elevs tid og en bestemt lærers dagsplan
10. håndtér en plan, som ikke kan nås inden sluttidspunktet

Der registreres gennemførsel, tidsforbrug, fejlklik, behov for hjælp, tvivlspunkter og deltagernes tillid til resultatet på en skala fra 1–5. Målet er mindst 90 % opgavegennemførsel uden hjælp, ingen kritiske brugerfejl og en median på mindst 4 for forståelighed og tillid.

## 10. Ikke-funktionelle tests

### 10.1 Ydelse

Følgende er foreslåede frigivelsesmål og bør justeres efter den faktiske driftsserver:

| ID | Scenarie | Foreslået mål |
|---|---|---|
| YD-01 | Første visning med demodata | Indhold og navigation anvendelig inden for 3 sekunder |
| YD-02 | Validering/redigering af 350 elever og 85 lærere | Synlig opdatering inden for 1 sekund efter Streamlit-rerun |
| YD-03 | Fordeling af 350/85 ved standarddybde 120 | Højst 30 sekunder på referencehardware; fremdrift vises løbende |
| YD-04 | Tidsplan af 350 elever ved dybde 80 | Højst 15 sekunder på referencehardware |
| YD-05 | Excel-eksport af standarddatasættet | Højst 5 sekunder og gyldig fil |
| YD-06 | Gentagne beregninger i samme session | Ingen vedvarende vækst i hukommelse eller gradvis væsentlig forringelse |

Kørselstid rapporteres som median og 95-percentil over mindst fem kørsler. Hardware, Python-version, Streamlit-version og datasæt registreres.

### 10.2 Robusthed og gendannelse

- Afbryd en upload og en beregning; verificér at seneste gyldige data ikke bliver erstattet af en halv tilstand.
- Fremprovokér manglende `openpyxl` og `python-docx`; de relevante eksportfunktioner skal fejle lokalt uden at blokere resten af programmet.
- Gentag upload, beregning, manuel redigering og tidsplan mindst 20 gange i samme session.
- Kontrollér at progressbar og spinner afsluttes ved både succes og fejl.
- Kontrollér at en fejl i en ny fil ikke ødelægger den tidligere gyldige filtilstand.

### 10.3 Sikkerhed og privatliv

- Uploadede filer må kun behandles i den aktuelle session som oplyst, og produktionsopsætningens logning/cache skal verificeres særskilt.
- Tekst fra filer skal escapes i HTML og må ikke kunne indsætte script eller markup.
- CSV/Excel-formelinjektion testes på alle eksporterede tekstkolonner.
- ZIP-filnavne skal renses for ugyldige tegn, stiforløb som `../` og dubletter.
- Fejlbeskeder må ikke vise serverstier, kildekode eller fulde stack traces.
- Store eller komprimerede filer skal have en aftalt størrelsesgrænse og må ikke kunne udtømme serverens hukommelse ukontrolleret.

## 11. Automatisering

### 11.1 Enhedstests

Følgende funktioner bør dækkes direkte med `pytest`:

- `repair_text`, `normal_key`, `canonical_subject` og `subject_from_course`
- `parse_students`, `parse_teachers`, `parse_teacher_upload` og `parse_schedule_upload`
- `resolve_wishes`, `input_validation`, `data_readiness` og `subject_coverage`
- `build_candidates`, `wished_count`, `score_solution` og `subject_stats`
- `optimize` med små datasæt og invariantkontrol
- `make_schedule` med kendte konfliktgrafer og tidsgrænser
- samtlige eksportfunktioner med genåbning og dataafstemning

### 11.2 Property-baserede tests

Generér mange små tilfældige, men gyldige elev-/lærersæt og kontrollér efter hver beregning:

```text
assignment -> eksisterende lærer
assignment -> lærer dækker elevens fag
automatisk load <= K, når global K bruges
double_load <= I
load == antal unikke elever pr. lærer
samme input + samme indstillinger -> samme resultat
højere dybde -> score ikke højere
fælles lærer -> forskellige tidsrunder
antal elevplansrækker == antal elever
```

### 11.3 Integration og UI-automatisering

Brug Streamlits testfaciliteter til komponent- og tilstandstest og et browserbaseret værktøj til komplette brugerrejser. Automatiseringen skal især kontrollere:

- aktivering/deaktivering af knapper
- navigation og status i sidepanelet
- at dataændringer ugyldiggør godkendelse, løsning og tidsplan
- upload og download i reelle browserforløb
- fejlbeskeder og tomtilstande
- tastaturnavigation og tilgængelighedsregression
- skærmbilleder ved faste viewport-størrelser for visuel regression

Heuristiske outputtabeller sammenlignes ikke ukritisk som fulde snapshots. Der sammenlignes primært invariants, scorekomponenter og semantiske data. Visuelle snapshots bruges kun til layout og styling.

## 12. Sporbarhed til programbeskrivelsen

| Kravområde i `PROGRAMBESKRIVELSE.md` | Primære tests |
|---|---|
| Indlæsning og klargøring | IMP-01–16, NOR-01–08 |
| Validering | VAL-01–10 |
| Kandidatlister og fagmatch | NOR-01–08, VAL-03, FOR-01, FOR-09 |
| K, individuelle max og overskridelse | FOR-05–09, RES-02, RES-05 |
| Dobbeltvejledning I | FOR-02–04, FOR-17, RES-07 |
| Elevønsker | VAL-04–06, FOR-10–12, RES-03 |
| Par- og klassesamling | FOR-13–14, TID-08 |
| Heuristik og scoring | FOR-15–18 |
| Resultat og manuel redigering | RES-01–10 |
| Tidsplan og konflikter | TID-01–20 |
| Pauser, frokost og tidsgrænse | TID-09–15 |
| Eksport | EXP-01–10 |
| Trinvis brugerrejse | NAV-01–10, UX-01–18 |

## 13. Særlige risikobaserede kontrolpunkter i den nuværende kode

Følgende punkter skal testes først, fordi en statisk gennemgang viser en særlig risiko for uoverensstemmelse mellem programbeskrivelse, tilstand og brugerens forventning:

1. **Forældet resultat efter regelændring:** Elev- og lærerændringer nulstiller løsningen, men regel- og maxkontroller skal verificeres særskilt. Et gammelt resultat må ikke vises som aktuelt efter ændring af K, I, kapacitet eller prioriteringer.
2. **Manuelle regelbrud:** Den manuelle vejledereditor kontrollerer fagmatch. Det skal verificeres, at K, individuelt max og I også kontrolleres eller mindst markeres meget tydeligt efter ændringen.
3. **Lærere uden fag:** Ved filparsing kan rækker uden fag blive frasorteret, før brugeren ser valideringen. Testen skal sikre, at fejlen ikke skjules.
4. **Manglende max:** En lærer uden læsbart max får i praksis kapacitet 0. Datagodkendelsen skal gøre konsekvensen tydelig.
5. **Ønskestatistik:** `wished_count` tæller tildelte fagslots, der matcher ønskelisten. En dobbeltvejledning kan derfor se ud som to opfyldte ønsker, selv om eleven kun afgav ét lærerønske. Terminologi og beregning skal verificeres.
6. **Visuel kodevej:** `main_v2()` anvender CSS-klasser til brand, hero, sidepanel og status, mens dens lokale stylesheet kun indeholder en del af de tilhørende regler. Visuel regression skal bekræfte, at den aktive kodevej faktisk viser det tilsigtede samlede design.
7. **Browserens sidetitel:** Den aktive sidekonfiguration omtaler en demo med fiktive data, også når brugeren arbejder med egne data. Det testes som et professionalitets- og tillidsproblem.
8. **Dublerede elevnavne i ZIP:** Word-filnavnet dannes af elevnavnet. To elever med samme navn skal stadig give to entydige filer.
9. **Regnearksinjektion:** Tekst skrives direkte til Excel-celler. Data der begynder med regnearksoperatorer skal behandles sikkert.
10. **Kilde til tidsplan:** En uploadet godkendt liste har forrang for en eksisterende beregning. Kilden skal være vedvarende og tydeligt markeret ved generering og eksport.

Disse punkter er ikke en erstatning for en egentlig testkørsel. De er prioriterede hypoteser, som skal bekræftes eller afkræftes med de beskrevne tests.

## 14. Testregistrering

Hver gennemført test registreres med:

- test-ID og programversion/commit
- dato, tester og miljø
- anvendt datasæt og indstillinger
- forventet resultat
- faktisk resultat
- status: bestået, fejlet, blokeret eller ikke kørt
- skærmbillede eller eksportfil ved UI- og outputfejl
- fejl-ID, alvor, reproduktionstrin og berørte data

Ved algoritmefejl vedlægges et minimalt datasæt, som reproducerer fejlen. Ved visuelle fejl vedlægges viewport, browser, zoom og både forventet og faktisk skærmbillede.

## 15. Godkendelseskriterier før frigivelse

En version er klar til faglig accept, når:

- alle P0-tests er bestået
- mindst 95 % af P1-tests er bestået, og ingen åben P1-fejl kan give et misvisende resultat
- alle algoritmiske invariants består på de kontrollerede og property-genererede datasæt
- UI, Excel, Word/ZIP og HTML er afstemt på mindst ét lille og ét realistisk datasæt
- den komplette 350/85-brugerrejse er gennemført i to understøttede desktopbrowsere
- der ikke er kritiske tilgængelighedsfejl i kerneflowet
- ydelsesmål er målt og enten opfyldt eller eksplicit accepteret
- den modererede brugertest opfylder målene for gennemførsel og tillid
- kendte begrænsninger er synlige for brugeren på det relevante tidspunkt, ikke kun i teknisk dokumentation

Den endelige faglige godkendelse bør foretages af en SOP-ansvarlig, mens teknisk godkendelse foretages af udvikler/testansvarlig. Begge skal godkende, at et resultat med advarsler præsenteres så tydeligt, at det ikke kan forveksles med en fuldt regelgyldig fordeling.
