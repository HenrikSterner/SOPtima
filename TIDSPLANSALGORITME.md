# Tidsplansfordelingsalgoritme

Dette dokument beskriver den nuværende tidsplansfunktion i SOPtima og et forslag til at udvide den med interaktiv træk-og-slip-planlægning.

## Formål

Tidsplanen omsætter den beregnede SOP-fordeling til én fælles vejledningstid pr. elev. En aktivitet kan derfor bestå af en elev, to fag, en hovedvejleder og eventuelt en medvejleder. De to vejledere skal være til rådighed samtidig.

Målet er at gennemføre flest mulige vejledninger parallelt uden dobbeltbookinger og inden for den dag, som brugeren har valgt.

## Nuværende funktionalitet

### Forudsætninger og inddata

Tidsplanen bygger på den allerede godkendte SOP-fordeling. Hver elev skal have to vejledere; mangler en af dem, kan tidsplanen ikke genereres.

Brugeren angiver følgende rammer:

- Start- og sluttid for dagen.
- Varighed pr. elev.
- Minutter mellem vejledninger.
- Antal og varighed af almindelige pauser.
- Frokostpause: fast klokkeslæt for alle eller flydende midt i planen.
- Om pauser skal fordeles før og efter frokost.
- Om elever med samme lærerpar skal samles mest muligt.
- Om algoritmen skal forsøge at mindske huller i den enkelte lærers plan.
- Eventuelle lærerspærringer. En spærring vælges som et start- og sluttidspunkt fra en liste med femminuttersintervaller inden for den valgte dag.
- Algoritmedybde, som bestemmer hvor mange alternative rundeplaceringer der afprøves.

Standardindstillingerne er kl. 08:15–16:15, 20 minutter pr. elev, 0 minutters skiftetid, to pauser á 10 minutter, pauser fordelt omkring frokost, samling af lærerpar og forsøg på at undgå lærerhuller.

### Vejledninger som konfliktgraf

Hver elevvejledning bliver behandlet som en aktivitet. To aktiviteter er i konflikt, hvis de har mindst én lærer til fælles. Aktiviteter uden fælles lærer kan ligge i samme runde og foregå parallelt.

Algoritmen opbygger derfor en konfliktgraf:

- En knude repræsenterer en elevvejledning.
- En forbindelse mellem to knuder betyder, at de deler en lærer.
- En farve repræsenterer en vejledningsrunde.
- Aktiviteter med samme farve kan afholdes samtidig.

Dermed må en lærer aldrig blive tildelt to elever i samme runde, mens andre lærerpar kan vejlede parallelt.

### Beregning af runder

Runderne beregnes med en DSATUR-inspireret kantfarvningsstrategi. Den vælger løbende den endnu ikke placerede aktivitet med flest konflikter i allerede anvendte runder og placerer den i den første konfliktfrie runde.

SOPtima prøver flere variationer af denne proces. Resultatet vælges primært ud fra færrest mulige runder. Den teoretiske nedre grænse er belastningen på den mest bookede lærer: Har en lærer 17 vejledninger, kan planen ikke have færre end 17 runder.

Når indstillingen *Forsøg at undgå huller* er slået til, bliver rækkefølgen af de fundne runder derefter optimeret. Det ændrer ikke, hvilke aktiviteter der kan være parallelle, men forsøger at lægge den enkelte lærers runder tættere sammen. Hvis *Saml samme lærerpar mest muligt* er valgt, bruges det også som et sekundært mål ved rundernes rækkefølge.

### Lærerspærringer

En lærerspærring er en hård regel: Ingen vejledning med den pågældende lærer må overlappe det spærrede tidsrum.

Overlappende spærringer for samme lærer samles automatisk. Når der er spærringer, prioriteres runder, som kan starte først uden at ramme de deltagende læreres spærringer. Ved den konkrete placering flyttes en runde frem til efter spærringen, hvis den ellers ville overlappe den.

En spærring kan derfor skabe et hul i den fælles tidslinje. Det er bevidst, fordi læreren aldrig må bookes i sin spærring.

### Pauser og frokost

Alle pauser er fælles og lægges kun mellem vejledningsrunder:

- Almindelige pauser placeres efter en afsluttet runde og før en efterfølgende runde; aldrig før første eller efter sidste vejledning.
- Ved flydende almindelige pauser fordeles de så vidt muligt før og efter frokosten. Med to pauser bliver det normalt én på hver side af frokost.
- En fast frokostpause respekterer det valgte klokkeslæt for alle lærere. Hvis en runde ellers ville krydse frokosten, indsættes frokosten først.
- En flydende frokostpause lægges omkring midten af den samlede rundeplan.

Pauser og frokost indgår i tidsforbruget og i lærerplanens tabel. De vises ikke som blokke i den visuelle læreroverigt.

### Kapacitetskontrol

Efter alle runder er placeret, sammenlignes planens reelle sluttid med den valgte sluttid. Beregningen medregner vejledningsrunder, skiftetid, almindelige pauser, frokost og nødvendig ventetid som følge af spærringer eller fast frokost.

Hvis hele planen kan være inden for dagen, oprettes:

- Elevplan med tidspunkt, elev, klasse, fag og vejledere.
- Lærerplan med vejledninger, medvejleder, pauser, frokost og spærringer.
- Fælles tidslinje og oversigt over lærerpar.
- Visuel lærerplan, Excel-eksport med samlet lærerplan og et ark pr. lærer samt elev- og lærer-HTML.

Hvis planen ikke kan nås, afbrydes genereringen i den nuværende version. Brugeren får en forklaring med antal runder, vejledningsminutter, pauser, frokost, de mest belastede lærere, nødvendig ventetid og den tidligst mulige sluttid. Ved lærerspærringer nævnes også de berørte elever og at de resterende vejledninger skal lægges en anden dag, hvis rammerne fastholdes.

Den nuværende version laver altså **ikke** en gemt delplan med uplacerede aktiviteter. Den viser i stedet en fejlbesked, når blot én del af planen falder efter dagens sluttid.

## Foreslået udvidelse: delvis plan og blok for uplacerede aktiviteter

I stedet for at afvise hele planen bør algoritmen altid beholde de aktiviteter, der lovligt kan placeres inden for dagen. Aktiviteter, som ikke kan placeres, skal flyttes til en selvstændig blok efter dagens tidsrum:

> Kan ikke placeres inden for det valgte tidsrum – kræver anden tid eller dag

Blokken er en planlægningsmarkering og ikke et reelt mødetidspunkt. Den må derfor ikke eksporteres som en falsk kalenderaftale. Hver række bør vise elev, klasse, fag, hovedvejleder, medvejleder, varighed og en forklaring, for eksempel *sluttid overskredet*, *lærerspærring* eller *konflikt med allerede låst aktivitet*.

Den samlede optimeringsrækkefølge ændres dermed til:

1. Maksimér antallet af aktiviteter placeret inden for dagens tidsrum.
2. Overhold alle hårde regler: lærerkonflikter, lærerspærringer, aktivitetens varighed, frokost og pauser.
3. Minimer antallet af runder og planens samlede længde.
4. Minimer lærerhuller og saml samme lærerpar, når det ikke forringer de tre første mål.
5. Placér resten i blokken *Kræver anden tid eller dag*.

Det gør resultatet brugbart, også på dage hvor tidsrammen er for kort.

## Foreslået udvidelse: træk og slip

### Brugeroplevelse

Den visuelle lærerplan udvides med redigerbare vejledningsblokke. En bruger kan trække en blok til et andet tidspunkt eller til en anden runde. Under flytningen vises et klart signal:

- Grøn: placeringen er lovlig.
- Gul: flytningen er mulig, men andre aktiviteter skal flyttes.
- Rød: placeringen bryder en hård regel og kan ikke accepteres.

Når blokken slippes, vælges en af to tilstande:

- **Lås kun den flyttede aktivitet:** Flytningen accepteres kun, hvis tidspunktet allerede er konfliktfrit.
- **Juster planen dynamisk:** Den flyttede aktivitet låses til det ønskede tidspunkt, og SOPtima genberegner kun de konfliktramte aktiviteter.

Brugeren bør kunne fortryde og gentage ændringer samt se en ændringslog med tidspunkt, bruger, aktivitet, gammel placering og ny placering.

### Validering ved et slip

Før en flytning accepteres, skal systemet kontrollere:

1. At start og slut ligger på den tilladte tidsopløsning, eksempelvis fem minutter.
2. At aktivitetens fulde varighed er bevaret.
3. At eleven ikke samtidig har en anden aktivitet.
4. At ingen af aktivitetens vejledere har en anden vejledning samtidig.
5. At ingen vejleder rammer en personlig spærring.
6. At aktiviteten ikke overlapper frokost eller en fælles pause.
7. Om aktiviteten ligger inden for dagens tidsrum eller skal markeres som uplaceret.

### Dynamisk genberegning

En flytning bør ikke udløse en helt ny plan, hvis det kan undgås. I stedet udføres en lokal reparation:

1. Den flyttede aktivitet bliver låst på brugerens valgte tidspunkt.
2. Aktiviteter, der nu overlapper den på elev- eller lærerniveau, frigøres midlertidigt.
3. Kun de frigjorte aktiviteter og de berørte læreres tidsrum genberegnes.
4. Uændrede aktiviteter bevares så vidt muligt på deres oprindelige plads.
5. Aktiviteter, som stadig ikke kan placeres, flyttes til blokken *Kræver anden tid eller dag*.

Hvis en lokal reparation ikke kan finde en løsning, kan brugeren vælge *Optimér hele planen*. Her bevares alle manuelt låste aktiviteter, mens resten beregnes på ny.

### Tekniske ændringer

Hver planlagt aktivitet skal have et stabilt aktivitets-id og mindst disse felter:

```text
activity_id, elev_id, lærer_ids, varighed, start, slut, status, låst, årsag
```

`status` bør mindst kunne være `planlagt`, `låst`, `uplaceret` eller `konflikt`. En uplaceret aktivitet har ingen reel start- og sluttid, men en forklarende årsag.

Planlægningsfunktionen skal returnere både `planned` og `unplaced` i stedet for at kaste en fejl, når dagen er for kort. Fejl skal fortsat bruges ved ugyldige indstillinger, eksempelvis sluttid før starttid eller en spærring med omvendt rækkefølge.

Webgrænsefladen skal sende en flytning til serveren med planens versionsnummer. Serveren validerer ændringen, beregner den opdaterede plan og returnerer en ny version. Versionsnummeret forhindrer, at to samtidige redigeringer overskriver hinanden uden varsel.

## Udvidelse af lærerens HTML: Word og Teams

Den genererede lærer-HTML skal ved hver planlagt elevvejledning indeholde en handlingskolonne. Herfra skal læreren kunne hente elevens SOP-dokument og åbne den fælles Teams-samtale for den konkrete vejledning.

Uplacerede aktiviteter skal fortsat stå i blokken *Kræver anden tid eller dag*, men må ikke få mødelink eller Word-download, før de har et bekræftet tidspunkt og vejlederpar.

### 1. Word-dokument pr. elev

For hver planlagt vejledning vises handlingen **Download Word**. Den henter et Word-dokument med filnavnet:

```text
SOP - Elevens navn.docx
```

Filnavnet skal renses for ugyldige tegn. Hvis to elever har samme navn, tilføjes elev-id eller klasse for at undgå, at en fil overskriver en anden, for eksempel `SOP - Alex Jensen - S 2024q.docx`.

Dokumentet udfyldes automatisk med mindst følgende oplysninger:

| Felt i dokumentet | Datakilde |
| --- | --- |
| Elevnavn | Elevens navn |
| Klasse | Elevens klasse |
| Fag 1 | Elevens første SOP-fag |
| Vejleder fag 1 | Navn på vejleder for fag 1 |
| Fag 2 | Elevens andet SOP-fag |
| Vejleder fag 2 | Navn på vejleder eller medvejleder for fag 2 |

Den eksisterende Word-skabelon kan genbruges med pladsholderne `Elevnavn`, `Klasse`, `Fag1 og niveau`, `Vejleder fag 1`, `Fag2 og niveau` og `Vejleder fag 2`. Hvis skabelonen mangler, skal SOPtima generere et enkelt dokument med de samme oplysninger.

En selvstændig HTML-fil kan ikke pålideligt indeholde mange Word-filer. Derfor skal lærer-eksporten leveres som enten:

1. En ZIP-pakke med lærer-HTML og en Word-fil pr. elev, hvor HTML-linkene peger på de medfølgende filer; eller
2. En webudgave af lærerplanen, hvor en download-knap kalder SOPtimas server og får dannet dokumentet ved klik.

ZIP-løsningen er bedst til lokal deling. Serverløsningen er bedst, hvis der senere skal logges, hvem der henter dokumenterne, eller hvis dokumentet skal kunne regenereres efter ændringer i planen.

### 2. Teams-samtale pr. elev

Ved siden af Word-handlingen vises **Åbn Teams-chat**. Linket åbner en fælles Teams-chat med eleven og de to vejledere og kan have et emne som:

```text
SOP – Elevens navn – Klasse
```

Et eksempel med Henrik Sterner og eleven `next27888@edu.nextkbh.dk` er:

```text
https://teams.microsoft.com/l/chat/0/0?users=hst%40nextkbh.dk,next27888%40edu.nextkbh.dk&topicName=SOP%20-%20Elevens%20navn%20-%20Klasse
```

For en vejledning med både hovedvejleder og medvejleder skal elevens og begge læreres Microsoft 365-adresser indgå som kommaseparerede værdier i `users`-parameteren. Alle dynamiske værdier skal URL-kodes.

Linket åbner en Teams-chat i kladdeform; Teams-brugeren skal selv sende den første besked. Det kan ikke med et almindeligt link oprettes og sendes en chat lydløst.

Begrebet *Teams-kanal* skal bruges præcist: Linket ovenfor opretter eller åbner en **gruppechat**, ikke en kanal. En kanal kræver et eksisterende Team og en kanal-id. Hvis der skal oprettes en særskilt kanal pr. elev automatisk, kræver det Microsoft Graph, en godkendt Entra-app, de nødvendige Team-tilladelser og skolens IT-godkendelse. Gruppechatten er derfor den anbefalede løsning for en enkelt SOP-vejledning.

### Nødvendige data og kontroller

For at handlingerne kan genereres korrekt, skal datamodellen udvides med:

```text
elev.teams_email
lærer.teams_email
```

Lærerens Teams-adresse skal knyttes til lærerens initialer. For eksempel er `hst` knyttet til `hst@nextkbh.dk` og `matr` til `matr@nextkbh.dk`.

Før Teams-linket vises, kontrollerer SOPtima, at eleven og alle vejledere har en gyldig Microsoft 365-adresse. Mangler blot én adresse, vises en deaktiveret handling med forklaringen *Teams-link kan ikke oprettes: mangler Teams-adresse for …*.

## Anbefalet implementeringsrækkefølge for eksporthandlinger

1. Tilføj Teams-mail til elev- og lærerdata samt validering af adresser.
2. Tilføj en handlingskolonne i lærer-HTML med Teams-chatlink pr. planlagt elev.
3. Genbrug Word-skabelonen og generér dokumenterne med det rensede filnavn `SOP - Elevens navn.docx`.
4. Pak HTML og Word-filer sammen som en ZIP-eksport, eller tilføj en serverbaseret download-rute.
5. Tilføj eventuelt Microsoft Graph senere, hvis Teams-møder eller kanaler skal oprettes og udsendes helt automatisk.

## Anbefalet implementeringsrækkefølge

1. Ændr planlægningsresultatet, så det returnerer planlagte og uplacerede aktiviteter i stedet for at afvise hele planen ved kapacitetsmangel.
2. Vis den separate blok for uplacerede aktiviteter i Streamlit, Excel og de genererede HTML-filer.
3. Tilføj stabile aktivitets-id'er, låsemarkeringer og ændringslog.
4. Tilføj træk-og-slip med streng validering uden automatisk flytning af andre aktiviteter.
5. Tilføj lokal dynamisk genberegning og til sidst funktionen *Optimér hele planen*.

Denne rækkefølge giver først et robust og ærligt resultat ved kapacitetsmangel og bygger derefter den interaktive redigering ovenpå.
