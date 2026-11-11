# Validering – 29. oktober 2026

## Rettelse af blandede versioner – 11. november 2026

Brugerens Outlook-kontoknap gav HTTP 404, fordi en Python-server startet før opdateringen stadig kørte. Backend-koden lå i hukommelsen, mens statiske filer blev læst fra den opdaterede projektmappe. Den gamle server blev lukket via API efter kontrol af, at ingen analyse eller udsendelse var aktiv. Genstart med samme datamappe bevarede alle 13 kørsler, og Outlook-kontokaldet returnerede HTTP 200 med én konto.

Version 1.1.1 indlæser brugerfladens filer ved serverstart, så de følger backend-versionen. Session-endpointet oplyser versionen. Genvejen advarer, hvis en anden version allerede kører, og installationen kontrollerer datamappens app-lås, også når appen kører som python.exe. En manglende Outlook-route giver en konkret genstartbesked.

10 målrettede API- og installationstests bestod, inklusive ændrede statiske filer under kørsel og en rigtig Windows-fillås holdt af Python. Edge-browsertesten bestod uden JavaScript-fejl. Den genbyggede pakke blev startet med isolerede data og port 19002; version 1.1.1, statiske filer og Outlook-kontokald blev verificeret. Brugerens server kører nu også 1.1.1 med alle 13 kørsler bevaret. Outlooks afsendelseskode er uændret. Ingen mails blev sendt.


## Outlook Classic-udvidelse – 11. november 2026

Implementeret på `feature/outlook-classic` fra `v1.0-before-outlook`. Afsendelsesmetode og konto vælges under Indstillinger. Klassisk Outlook bruger det eksisterende lokale Outlook-login. Graph er fortsat tilgængelig, og gamle kladder uden metode behandles som Graph.

- 88 automatiske tests bestod i Windows-brugersessionen, inklusive DPAPI og 16 nye Outlook-tests. Outlook-tests simulerer afsendelse, kontofejl, offline tilstand, timeout, midlertidige PDF-filer, godkendelser og gamle kladder.
- Edge-browsertesten bestod uden JavaScript-fejl. Den kontrollerede både Graph og Outlook, kontoindlæsning, gemte indstillinger efter genindlæsning, afsender i godkendelsen og status efter simuleret Outlook-afsendelse. Desktop og mobil blev kontrolleret.
- `scripts/validate_outlook.py` bestod med rigtig klassisk Outlook: kontoidentifikation, profil/lagerbinding, afsender, modtager og PDF-vedhæftning blev kontrolleret på en usendt syntetisk mail, som blev kasseret. Scriptet kalder aldrig Send.
- Den genbyggede Windows-pakke bestod `scripts/validate_package.py` med rigtig LibreOffice og Excel: 13 sider for fastansatte og 2 sider for vikarer. Den pakkede Outlook-worker fandt og verificerede en afsenderkonto; ingen mails blev sendt.
- Testene brugte separate data og port 19001/19002. Standardporten er stadig 9000. Brugerens kørende app blev ikke stoppet.
- Installationsscriptets PowerShell-syntaks og ZIP-integritet er kontrolleret. Installation på en ren Windows-pc er ikke afprøvet.

Rigtig mailafsendelse og levering er endnu ikke afprøvet med den ønskede afsenderkonto. Før medarbejderudsendelse: vælg konto, importér en testfil, brug **Send test til mig selv**, og kontrollér modtagelse og Sendt post. Outlooks lokale sikkerhedsregler kan vise dialoger eller blokere automatisering. Ved ukendt resultat skal både Udbakke og Sendt post kontrolleres før et nyt forsøg.


## Installationspakke – 11. november 2026

ZIP-filen er genbygget og ZIP-integriteten kontrolleret. Den indeholder appmappen, install.cmd, install.ps1 og START-HER.txt. Python og Python-afhængigheder er indbygget. Installationen forsøger automatisk at installere manglende LibreOffice via WinGet og kontrollerer derefter runtime igen. Excel og Microsoft-konto kræver separat opsætning.

66 eksisterende tests bestod i sandboxen. Windows DPAPI-testen bestod særskilt i den normale Windows-session med syntetiske data i en separat midlertidig mappe. Fem nye installationstests bestod: eksisterende programmer, vellykket download, manglende WinGet, mislykket download og manglende Excel. Installationstests simulerer systemkommandoer; de installerer ikke software. PowerShell-syntaks er kontrolleret.

Den fulde pakketest med PDF-eksport kunne ikke starte, fordi port 9000 allerede var optaget; den kørende app blev ikke stoppet. Frisk installation og faktisk LibreOffice-download på en ren Windows-pc er endnu ikke afprøvet. De tidligere PDF-resultater nedenfor er historiske, ikke gentaget ved denne ændring. Ingen mails blev sendt.

Genskabelige build- og Python-cachemapper samt den gamle udpakkede dist/package-71c04568-pakke er fjernet. Kildekode, tests, udviklingsmiljø og tidligere testartefakter er bevaret.

## Opdatering: delvis import og Cc

67 automatiske tests bestod, inklusive Windows DPAPI kørt i brugerens Windows-session. Testene dækker nu også eksterne navngivne områder, lokale navne, afhængigheder mellem ark, dynamiske og 3D-referencer, manglende/ødelagte ark, gemte Excel-fejl, fælles formler, PDF-gennemgang samt Cc i API, kladder, godkendelse og Graph-payload. Eksisterende kladder uden Cc understøttes. Mailtransport er simuleret.

Den opdaterede Edge-browsertest bestod uden JavaScript-fejl. Den kontrollerede Cc før import, efter genåbning, automatisk lagring før godkendelse, afsendelse med Cc og testmail uden Cc. En blandet fil med en ekstern navngiven reference og en fejlramt medarbejder fortsatte med den gyldige medarbejder. Desktop- og mobilvisning er kontrolleret.

Den genbyggede Windows-app bestod også pakketesten med rigtig LibreOffice og Excel: 13 sider for fastansatte og 2 sider for vikarer. Begge importfiler indeholdt en ekstern navngiven reference og et fejlramt ekstra medarbejderark; den gyldige medarbejder blev behandlet videre. Det pakkede API gemte og ændrede Cc, og den pakkede GUI indeholdt det nye felt.

Den indberettede XLSX-fil blev undersøgt i en midlertidig kopi. Seks eksterne navngivne referencer udløste den tidligere globale fejl. Den nye import fortsatte med 96 ark, hvor 27 bestod modtager- og timekontrollen og kunne gå videre til PDF. Gemte Excel-fejl i A1 blev registreret som advarsler med krav om PDF-gennemgang. En repræsentativ PDF blev eksporteret med rigtig Excel og indeholdt de forventede to sider. Originalfilen var uændret, midlertidige arbejdsfiler blev slettet, og der blev ikke sendt mails.

`scripts/validate_import.py <fil.xlsx> <fastansat|vikar> --render-first` kan gentage en lokal kontrol uden at ændre live-kørsler eller sende mails. Scriptet udskriver kun optællinger.

## Tidligere grundvalidering – 22. oktober 2026

## Automatiske tests

38 tests dækker Excel-match, synlige fejl, dubletter, modstridende referencer, vikartimer, API-beskyttelse, individuelle beskeder, periodeskift, genstart, afsendelsesstatus, PDF-integritet, opbevaring og Windows DPAPI. Graph-svar og mailafsendelse simuleres.

Testkommando: `.venv\Scripts\python.exe -m pytest -q`.

Windows DPAPI-testen skal køre som en normal Windows-bruger med en aktiv logonsession. En isoleret sandbox uden denne session kan ikke benytte brugerens DPAPI eller Excel COM.

## Browser

`scripts/validate_browser.py` er gennemført i Microsoft Edge med syntetiske Excel-filer og simuleret mailtransport. Testen kontrollerer:

- Måned/år og prioritet over filnavnet.
- Fejlramte ark med præcise cellehenvisninger og deaktiveret modtagervalg.
- Individuel besked, autosave, periodeskift og genåbning af kladde.
- Godkendelse med både inkluderede og udeladte personer.
- Udsendelse til korrekt modtager og testmail til afsenderen selv.
- Desktopvisning, mobilvisning uden vandret overflow og indstillinger.

Ingen JavaScript-fejl blev registreret.

## Rigtig PDF-eksport

`scripts/validate_windows.py` er gennemført med installeret Microsoft Excel og LibreOffice/UNO. Originalkoden fra begge ZIP-filer og den nye kode behandlede samme syntetiske skabeloner:

| Type | Sidetal | Sammenligning af rasteriserede sider |
|---|---:|---|
| Fastansat | 13 | Alle sider identiske |
| Vikar | 2 | Begge sider identiske |

Der kontrolleres også læsbar PDF-tekst, korrekt medarbejderidentitet i fastansattes dokument og uændret kildefil.

## Windows-pakke

`scripts/validate_package.py` er gennemført på den pakkede `FastansatVikar.exe`, startet fra en anden arbejdsmappe. Den lokale server, indbyggede Python-runtime, statiske filer, PDF-programkontrol, LibreOffice-worker, separate Excel-worker og kontrolleret nedlukning fungerer. Den pakkede app genererede 13 og 2 sider som forventet.

PowerShell-installations- og buildscripts er syntakskontrolleret. Installationsscriptets oprettelse af brugerens genveje og Windows-oprydningstask er ikke udført som del af testen; testkørslerne ændrer ikke en eksisterende installation.

## Microsoft-konti og udsendelse

Der er ikke gennemført Microsoft-login eller sendt rigtige mails som del af denne validering. Organisationens Entra-appregistrering og samtykke skal konfigureres ved første brug. PDF-sammenligningen mellem gammel og ny eksport anvender syntetiske ark; en rigtig XLSX-fil og én PDF blev også kontrolleret den 29. oktober som beskrevet ovenfor. Gennemgå de konkrete PDF'er og en testmail til dig selv før medarbejderudsendelse.
