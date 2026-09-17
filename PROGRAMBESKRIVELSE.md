# SOPtima – program- og algoritmebeskrivelse

## 1. Formål

SOPtima er et Streamlit-program, der hjælper med at fordele elever mellem vejledere og efterfølgende lave en tidsplan for vejledningen.

Programmet forsøger at finde en fordeling, som:

- giver hver elev en vejleder til hvert af elevens to fag
- tager hensyn til elevens ønsker
- sikrer, at vejlederen underviser i det relevante fag
- overholder vejledernes kapacitet og maksimumstal
- begrænser antallet af dobbeltvejledninger
- fordeler eleverne så jævnt som muligt mellem vejlederne
- kan samle elever med samme vejlederpar eller fra samme klasse

## 2. Programmets overordnede arbejdsgang

Programmet er opdelt i seks trin:

1. Elevdata
2. Lærerdata
3. Regler og maksimumstal
4. Beregn fordeling
5. Resultat og eksport
6. Tidsplan

Programmet arbejder trinvis. Hvis elevdata, lærerdata, maksimumstal eller fordelingsregler ændres, ugyldiggøres den eksisterende fordeling og tidsplan. Fordelingen skal derefter beregnes igen, så resultatet altid bygger på de aktuelle oplysninger.

## 3. Indlæsning og klargøring af data

### Elevdata

Elevdata kan indlæses fra CSV- eller Excel-filer. Programmet forventer som minimum:

- elevnavn
- fag 1
- fag 2

Klasse eller hold anbefales og bruges blandt andet ved klassesamling. Hvis feltet mangler, vises eleven med ukendt klasse/hold. Der kan desuden angives op til to ønskede vejledere, projekttitel og projektbeskrivelse.

For hver elev registreres navn, klasse, fag, eventuelle vejlederønsker samt projektoplysninger.

### Lærerdata

Lærerdata kan også indlæses fra CSV- eller Excel-filer. Programmet bruger blandt andet:

- lærerens initialer eller ID
- lærerens navn
- de fag, læreren underviser i
- eventuelle hold
- lærerens maksimumstal

Hvis maksimumstal findes på et ekstra Excel-ark, læses de også ind. Lærerens fag og ID normaliseres, så små forskelle i mellemrum, store/små bogstaver og tegnsætning ikke skaber unødige fagkonflikter.

### Normalisering

Før data bruges, renses tekst for blandt andet:

- overflødige mellemrum
- forskelle på store og små bogstaver
- ældre eller fejlkodet UTF-8-tekst
- forskellige skrivemåder for fag
- niveauangivelser som A, B eller C

Formålet er, at fag og personer kan matches korrekt, selv om de er skrevet lidt forskelligt i inputfilerne.

## 4. Validering

Før fordelingen kan beregnes, kontrollerer programmet blandt andet:

- om der findes elever og lærere
- om alle elever har to fag
- om alle lærere har et gyldigt ID
- om elev-ID'er er entydige
- om alle elevfag har mindst én mulig vejleder
- om ønskede vejledere findes
- om et ønske angivet som lærernavn er entydigt
- om en ønsket vejleder faktisk underviser i elevens fag
- om lærerdata er tilstrækkelige til en fordeling

Ugyldige eller ufuldstændige data vises som advarsler eller fejl i brugergrænsefladen. Data skal godkendes, før beregningen kan startes.

## 5. Kandidatlister

Før selve fordelingen bygges en kandidatliste med alle lærere, der underviser i hvert fag.

Eksempel:

```text
Matematik  -> AB, CD, EF
Dansk      -> GH, IJ
Fysik      -> AB, KL
```

Ved en elevs fag bruges kun lærere fra den relevante kandidatliste.

## 6. Fordelingsregler

Brugeren kan vælge følgende regler:

### Global K-grænse

`K` er det maksimale antal elever, en lærer som udgangspunkt må have.

### Individuelt maksimum

Hver lærer kan have sit eget maksimumstal. Hvis global K anvendes, bruges den laveste værdi af den globale grænse og lærerens individuelle maksimum:

```text
kapacitet(lærer) = min(K, individuelt maksimum)
```

### Overskridelse af individuelt maksimum

Det kan tillades, at en lærer overskrider sit individuelle maksimum, så længe den globale K-grænse ikke overskrides. Denne mulighed kan slås fra.

### Dobbeltvejledning

`I` angiver, hvor mange elever en lærer højst må vejlede i begge elevens fag. Hvis samme lærer tildeles begge fag, tæller det som én elev i lærerens belastning, men som én dobbeltvejledning.

### Samme vejlederpar

Hvis dette prioriteres, forsøger algoritmen at genbruge de samme vejlederpar til flere elever med samme fagkombination.

### Samme klasse eller hold

Hvis dette prioriteres, forsøger algoritmen at samle elever fra samme klasse eller hold hos de samme lærere.

## 7. Fordelingsalgoritmen

Det er en heuristisk algoritme med flere beregningsforsøg. Den søger altså efter en god løsning gennem en række prioriterede valg og forbedringer.

### 7.1 Initialisering

For hvert beregningsforsøg starter programmet med en tom fordeling og holder styr på lærernes belastning og antallet af dobbeltvejledninger. Lige gode valg kan afgøres forskelligt fra forsøg til forsøg.

### 7.2 Beregning af fagpres

For hvert fag beregnes et pres:

```text
pres = antal elever med faget / samlet tilgængelig kapacitet
```

Et fag med mange elever og få ledige vejlederpladser får derfor et højt pres.

### 7.3 Prioritering af vanskelige elever og fag

Eleverne sorteres efter:

1. højeste fagpres
2. færrest gyldige ønskede vejledere
3. færrest mulige vejledere til fagene
4. tilfældig tie-breaker

Det betyder, at de mest begrænsede fag og elever behandles først. På den måde reduceres risikoen for, at de sidste elever står uden en mulig vejleder.

### 7.4 Første tildeling: elevønsker

Algoritmen forsøger først at tildele elevens ønskede vejledere, hvis:

- vejlederen underviser i det pågældende fag
- vejlederen har ledig kapacitet
- tildelingen ikke overskrider grænsen for dobbeltvejledning

Ved flere mulige ønsker vælges normalt den lærer, der har lavest belastning.

Algoritmen forsøger samtidig at undgå et valg, som gør elevens andet fag umuligt at bemande.

### 7.5 Anden tildeling: resterende fag

Fag, der endnu ikke har fået en vejleder, behandles efter fagpres og antal mulige vejledere.

For hver kandidat vurderes blandt andet:

- om kandidaten vil overskride kapaciteten
- om elevens andet fag fortsat kan dækkes
- om kandidaten er blandt elevens ønsker
- hvor stor lærerens aktuelle belastning er

Den bedste lovlige kandidat tildeles.

### 7.6 Flytning af elever ved overbelastning

Efter den første fordeling undersøges lærere, der ligger over deres kapacitet.

Algoritmen forsøger op til otte gange at flytte elever til andre mulige lærere. Elever, hvor den overbelastede lærer ikke var et ønske, flyttes typisk først, så elevønsker bevares bedst muligt.

### 7.7 Forbedring af elevønsker

Derefter forsøger algoritmen op til fire gange at erstatte en ikke-ønsket vejleder med en ønsket vejleder, hvis det kan ske uden at bryde kapacitets- eller dobbeltvejledningsreglerne.

### 7.8 Forbedring af begge ønsker

Til sidst forsøges kombinationer af elevens to ønsker. Hvis begge ønsker kan opfyldes, og resultatet er bedre end den nuværende løsning, ændres elevens to tildelinger samlet.

### 7.9 Gentagelse og valg af bedste løsning

Hele processen gentages det antal gange, som brugeren vælger under “Algoritmedybde for fordeling”. Hvert forsøg bruger en anden tilfældig rækkefølge ved lige gode valg.

Den løsning, der får den laveste score, gemmes som resultat.

## 8. Scoring af en løsning

Løsninger vurderes efter en samlet kvalitetsvurdering. En lav score er bedst.

Scoren prioriterer i denne rækkefølge:

1. overskridelse af global K-grænse
2. elever uden komplet vejlederfordeling
3. overskridelse af individuelle maksimumstal
4. elever uden opfyldte ønsker
5. manglende opfyldelse af elevønsker
6. genbrug af vejlederpar og lærere fra samme klasse
7. lav maksimal belastning
8. jævn belastning målt med summen af belastningernes kvadrater

En forenklet beskrivelse af målet er:

```text
minimér:
    alvorlige kapacitetsbrud
  + manglende tildelinger
  + ikke-opfyldte ønsker
  + skæv belastning
  - bonus for genbrug af vejlederpar og klasse-samling
```

Kapacitetsbrud og manglende tildelinger vægtes meget højere end ønsker og samling. Derfor vil en komplet fordeling normalt prioriteres over en fordeling, der opfylder flere ønsker, men efterlader elever uden vejleder.

## 9. Oprettelse af tidsplan

Tidsplanen kan baseres på den beregnede fordeling eller på en godkendt Excel-liste med elev, fag og vejledere.

En tidsplan kan kun oprettes, hvis alle elever har to vejledere.

### Er tidsplanen også heuristisk?

Ja. Tidsplanen konstrueres også heuristisk, men på en lidt anden måde end selve elevfordelingen.

Den vigtigste opgave er at undgå konflikter: En lærer må ikke være sat til at vejlede to elever på samme tid. Programmet finder derfor en gyldig placering af eleverne i parallelle runder ved hjælp af en grådig grafbaseret metode. Den placerer altid den aktuelt mest begrænsede elev i den tidligste runde, hvor der ikke opstår en konflikt.

Metoden har to egenskaber:

- Den producerer en plan, hvor elever med fælles vejledere ikke ligger samtidig.
- Den garanterer ikke, at planen har det absolut færrest mulige antal runder.

For at forbedre resultatet gentages konstruktionen flere gange med forskellige valg, når flere elever er lige vanskelige at placere. Den bedste af de fundne planer vælges. Derfor betyder en højere algoritmedybde normalt større chance for en kort plan, men også længere beregningstid.

Det er altså mere præcist at sige, at tidsplanen er en heuristisk optimering under hårde konfliktregler: Konflikterne skal overholdes, mens antallet af runder og samlingen af vejlederpar optimeres så godt som muligt.

### 9.1 Opret vejlederpar

For hver elev gemmes de to tildelte vejledere som et vejlederpar. Hvis samme lærer har begge fag, behandles læreren kun én gang i konfliktberegningen.

### 9.2 Gruppér vejlederpar

Hvis “Saml samme lærerpar mest muligt” er valgt, placeres elever med samme vejlederpar i sammenhængende blokke.

### 9.3 Opret konflikter

To elever er i konflikt, hvis de deler mindst én vejleder. De kan derfor ikke have vejledning samtidig.

Konflikterne kan ses som en graf:

```text
Elev A ── Elev B
  │         │
Elev C ── Elev D
```

En kant betyder, at de to elever ikke må placeres i samme runde.

### 9.4 Fordel elever i runder

Programmet bruger en DSATUR-lignende graf-farvning:

1. Vælg den elev, der har flest forskellige allerede brugte runder blandt sine konflikter.
2. Hvis flere elever står lige, vælges den med flest konflikter.
3. Giv eleven den første runde, der ikke bruges af en konflikt.
4. Gentag, indtil alle elever er placeret.

Elever i samme runde kan vejledes parallelt, fordi de ikke deler vejleder.

Der beregnes flere mulige farvninger med forskellige tilfældige tie-breakers. Den farvning med færrest runder vælges. Algoritmen stopper tidligt, hvis den når den teoretiske nedre grænse:

```text
mindste mulige antal runder = det største antal elever hos én vejleder
```

### 9.5 Rækkefølge af runder

Hvis vejlederpar skal samles, ændres rækkefølgen af hele runderne, så runder med fælles vejlederpar så vidt muligt ligger ved siden af hinanden. Dette ændrer ikke konflikterne eller antallet af runder.

### 9.6 Tidspunkter, pauser og frokost

For hver runde bruges:

- starttidspunkt
- antal minutter pr. elev
- skiftetid mellem runder
- almindelige pauser
- frokostpause

En forsimplet beregning af den samlede tid er:

```text
samlet tid =
    antal runder × elevtid
  + (antal runder - 1) × skiftetid
  + almindelige pauser
  + frokostpause
```

Almindelige pauser fordeles omtrent jævnt mellem runderne. Frokosten kan enten placeres:

- på et fast tidspunkt for alle lærere
- flydende omtrent midt i vejledningsplanen

Hvis planen ikke kan nås inden sluttidspunktet, stopper programmet og viser en forklaring med den nødvendige sluttid samt hvilke dele af planen, der bruger tiden.

## 10. Brugerens trin-for-trin-brug

### Trin 1 – Elevdata

1. Upload elevarket.
2. Kontrollér elevnavn, klasse og de to fag.
3. Angiv eventuelle ønskede vejledere.
4. Ret fejl eller ukendte ønsker.

### Trin 2 – Lærerdata

1. Upload lærerarket.
2. Kontrollér initialer, navne og fag.
3. Kontrollér eventuelle hold og maksimumstal.
4. Ret lærere uden fag eller ID.

### Trin 3 – Regler og maksimumstal

1. Vælg, om global K skal bruges.
2. Angiv K.
3. Angiv maksimum for dobbeltvejledning `I`.
4. Vælg, om individuelle maksimumstal må overskrides op til K.
5. Vælg, om maksimumstal skal låses.
6. Vælg prioritering af vejlederpar og klasser.
7. Vælg algoritmedybde.

### Trin 4 – Beregn

1. Tryk på “Beregn fordeling”.
2. Programmet gennemfører flere fordelingsforsøg.
3. Den bedste løsning gemmes.
4. En statuslinje viser beregningens fremdrift.

### Trin 5 – Resultat og eksport

Her vises blandt andet:

- antal elever med komplet fordeling
- antal opfyldte ønsker
- lærernes belastning
- fagstatistik
- elever eller fag uden vejleder

Fordelingen kan redigeres manuelt. Ved gemning kontrolleres fagmatch, global K, lærernes maksimumstal og grænsen for dobbeltvejledning. Tilladte overskridelser af et individuelt maksimum vises som advarsler. En manuel ændring ugyldiggør en allerede oprettet tidsplan. Resultatet kan eksporteres til Excel, og der kan genereres Word-filer til eleverne.

Ønskestatistikken tæller forskellige afgivne lærerønsker. Hvis samme ønskede lærer tildeles begge fag, tæller det derfor som ét opfyldt lærerønske og ikke som to forskellige ønsker.

### Trin 6 – Tidsplan

1. Brug den beregnede fordeling eller upload en godkendt fordelingsliste.
2. Angiv start- og sluttidspunkt.
3. Angiv minutter pr. elev.
4. Angiv skiftetid.
5. Angiv antal og længde på pauser.
6. Vælg frokosttype og frokostlængde.
7. Vælg algoritmedybde for tidsplanen.
8. Tryk på “Generér tidsplan”.
9. Kontrollér elevplan, lærerplan og samlet tidslinje.
10. Eksportér tidsplanen som Excel eller HTML.

## 11. Output

Fordelingsresultatet kan indeholde følgende Excel-ark:

- Fordeling
- Lærerbelastning
- Fagstatistik
- Ikke tildelte
- Elevdata
- Lærerdata

Tidsplanen indeholder:

- Elevplan
- Lærerplan
- Tidslinje
- Lærerpar
- Indstillinger

Der kan også genereres individuelle Word-filer samlet i en ZIP-fil.

## 12. Beregningstid

Beregningstiden vokser især med:

- antallet af elever
- antallet af lærere
- antallet af mulige vejledere pr. fag
- antallet af alternative fordelinger, programmet afprøver

En højere algoritmedybde kan give en bedre fordeling eller tidsplan, men øger beregningstiden.

For tidsplanen afhænger beregningstiden også af, hvor mange elever der deler vejledere. Jo flere fælles vejledere, desto flere tidsmæssige konflikter skal programmet tage hensyn til.

## 13. Begrænsninger

- Algoritmen er heuristisk og garanterer ikke den globalt optimale løsning.
- Hvis der ikke er tilstrækkelig fagdækning eller kapacitet, kan nogle fag eller elever stå uden vejleder.
- Tidsplanen kræver, at alle elever har to vejledere.
- Manuel redigering bør efterfølges af en kontrol af belastning og maksimumstal.
- Resultatet afhænger af valgte K-, I- og kapacitetsværdier samt algoritmedybden.
- Flere beregningsforsøg forbedrer normalt robustheden, men gør beregningen langsommere.
