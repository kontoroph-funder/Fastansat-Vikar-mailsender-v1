# Fastansatte & Vikarer

Lokal Windows-app til at analysere Excel, gennemgå personlige PDF'er, redigere individuelle mails og sende via Microsoft Graph eller en konto i klassisk Outlook til Windows.

## Installation

1. Pak **FastansatVikar-Windows.zip** helt ud.
2. Dobbeltklik på **install.cmd**. Appen installeres for den aktuelle Windows-bruger, og der oprettes en genvej på skrivebordet og i Startmenuen. Python er inkluderet i appen; det skal ikke installeres separat. Manglende LibreOffice hentes automatisk via WinGet; Windows kan bede om administratorgodkendelse. Excel skal installeres og aktiveres separat med brugerens Office-licens. Se `START-HER.txt` i ZIP-filen.
3. Åbn **Fastansatte og Vikarer** fra genvejen. Browseren åbner `http://127.0.0.1:9000`.
4. Under **Indstillinger** vælges afsendelsesmetode. Til klassisk Outlook: åbn Outlook, log ind på afsenderkontoen, vælg **Klassisk Outlook på denne pc**, tryk **Hent konti fra Outlook**, vælg kontoen og **Gem indstillinger**. Til Graph: angiv Microsoft client-id og tenant-id og vælg **Gem & forbind Microsoft 365**.
5. Brug **Kontrollér PDF-programmer**. Vikarer kræver installeret Microsoft Excel. Fastansatte kræver LibreOffice med fungerende Python/UNO; programmappe og eventuel Python-sti kan indstilles her.

Selve appinstallationen kræver normalt ikke administratorrettigheder; installation af LibreOffice kan kræve dem. Installationen opretter også en oprydningsopgave for den aktuelle bruger. ZIP-filen indeholder kun appmappen, `install.cmd`, `install.ps1` og `START-HER.txt`. Udviklingsfiler og testdata distribueres ikke. Hvis WinGet mangler, installeres Microsoft App Installer fra Microsoft Store, hvorefter installationen køres igen. Returkode 2 betyder, at appen er installeret, men PDF-programmer kræver opsætning.

Microsoft-appregistreringen skal være en offentlig desktop-app med delegerede rettigheder `User.Read` og `Mail.Send` samt **Allow public client flows** slået til. Der anvendes ingen client secret. Organisationens Microsoft-politikker kan kræve administratorens samtykke. Appen sender fra brugerens egen postkasse; fælles postkasser er ikke implementeret.

## Klassisk Outlook

Outlook-metoden bruger det installerede **klassiske Outlook til Windows** i den aktuelle Windows-brugers profil. Appen kræver ikke client-id, tenant-id eller Graph-samtykke ved denne metode. Den valgte konto skal allerede kunne sende almindelige mails i Outlook. Det nye Outlook understøtter ikke denne COM-integration. Office-licens og Outlook-installation følger ikke med pakken.

Vælg og gem kontoen før import af en ny Excel-fil. Kørsler gemmer deres metode; gamle kørsler og kvitteringer uden metode behandles som Graph. Skift til den passende metode under Indstillinger for at fortsætte en gammel kladde. Konto- og metodeskift ugyldiggør tidligere godkendelser og er blokeret under udsendelse. Outlook-kontoens profil, postkasselager og SMTP-adresse kontrolleres ved godkendelse og igen inden afsendelse. Manglende eller tvetydige konti stopper udsendelsen; standardkontoen bruges aldrig som reserve. Fælles postkasser og aliaser er ikke særskilt implementeret.

Appen sender automatisk efter **Godkend & send nu**; dette er ikke en kladdeeksport. Status **Afleveret til Outlook** betyder, at Outlook har accepteret Send-kaldet, ikke at mailen er leveret. Mails kan vente i Udbakke. Ved afbrudt eller usikkert resultat pauses kørslen; kontrollér både Udbakke og Sendt post, før du klargør et nyt forsøg. En mail, der allerede venter i Udbakke, må ikke gensendes. En ukendt testmail bør også kontrolleres, før der sendes en ny test.

Outlook kan vise sikkerhedsforespørgsler eller blokere automatisering efter organisationens regler. Appen ændrer ikke sikkerhedsindstillinger. Outlook-kald kører i en separat proces med tidsgrænse (45 sekunder for kontokontrol, 180 sekunder for afsendelse). Kun appens worker stoppes ved timeout; brugerens Outlook lukkes aldrig. Midlertidige maildata og PDF-kopier slettes efter kaldet. Rester efter et programnedbrud ryddes efter 48 timer ved appens almindelige oprydning. Outlooks egne kopier i Udbakke/Sendt post omfattes ikke af appens oprydning.

## Arbejdsgang

1. Vælg **Fastansatte** eller **Vikarer**, måned og år. Dropdowns starter på den aktuelle dato i Danmark. Filnavnet ændrer aldrig valget.
2. Vælg en `.xlsx`-fil på højst 25 MB og klik **Analysér fil**.
3. Gennemgå fejloversigten og medarbejderne. Klik på en medarbejder, åbn PDF'en, og rediger eventuelt emne og tekst. Beskeder gemmes automatisk efter en kort pause; **Gem besked** gemmer straks.
   **Cc** er en valgfri mailadresse for hele kørslen. Den får en kopi af hver medarbejdermail med PDF. Adressen kan angives før import eller gemmes med **Gem Cc** bagefter; den gemmes også ved gennemgang af udsendelsen. Cc vises i godkendelsen, og ændringer kræver ny godkendelse. Testmail sendes kun til dig selv, uden Cc.
4. Brug eventuelt **Send test til mig selv**. Testen bruger den valgte persons PDF, men sendes kun til den indloggede afsender.
5. Vælg **Gennemgå & send** eller send kun til den valgte person. Godkendelsesdialogen viser afsender, lønperiode, modtagere og udeladte poster.
6. Åbn gemte kladder under **Historik & kladder**. **Afslut appen** lukker den lokale server, når der ikke foregår analyse eller afsendelse. At lukke browserfanen alene stopper ikke serveren.

**Måned/år:** Lønperioden løber fra den 16. i foregående måned til den 15. i den valgte måned. Januar viser begge år. Ved periodeskift opdateres skabelonbaserede beskeder. Individuelle beskeder bevares og skal markeres som gennemgået. Efter et afsendelsesforsøg låses perioden. Periodevalget filtrerer ikke PDF-indholdet: fastansattes PDF indeholder stadig oversigten og alle årets 12 månedsafsnit fra den uploadede projektmappe.

## Excel-regler og fejl

**Fastansatte:** Navn læses fra `Timeregistrering …` i F1. Registreringsarkets R2 skal matches entydigt med et oversigtsarks S2. Oversigtens formler må ikke referere til en anden medarbejder. Månedsoverskrifter i kolonne A skal dække januar–december i korrekt rækkefølge. PDF'en består af oversigten A1:G16 og alle 12 månedsafsnit; et afsnit kan fylde flere sider.

Fejl viser medarbejder, arknavn, celle og den fundne adresse. Manglende/ugyldige adresser, dubletter, modstridende referencer og PDF-fejl blokerer kun berørte personer. Korrekte personer kan stadig sendes. Oversigtsark uden en entydig registrering vises særskilt. Der gættes aldrig på modtageren. Ret adresser i Excel og opret en ny kørsel med den rettede fil.

**Vikarer:** Navnet kommer fra fanen. E-mail læses fra B2, eller B1 hvis B2 er tom. Mindst én positiv værdi i B5:B25 kræves. Numeriske decimaltal og danske teksttal understøttes. PDF'en bruger A1:K75, liggende A4 og sideskift før række 39. Output skal være to sider; forkert layout vises som en fejl.

Filkontrollen behandler fejl pr. ark. Eksterne formler med brugbare gemte værdier fryses i en midlertidig kopi. Eksterne eller ugyldige navngivne områder fjernes med en advarsel; medarbejdere, hvis formler bruger de pågældende områder, blokeres. Manglende eksterne værdier, ulæselige ark og afhængige ark vises med fejl, mens andre medarbejdere behandles videre. Ukendte dynamiske referencer vurderes konservativt. Manglende værdier erstattes aldrig med nul.

Gemte Excel-fejl og fjernet betinget formatering vises som advarsler. PDF-generering fortsætter, men du skal åbne PDF'en, kontrollere de nævnte celler og vælge **PDF og besked er gennemgået** før afsendelse. Fejlbeskeder viser ark, celle eller navngivet område. Originalfilen ændres ikke. Regn eventuelt filen igennem og gem den i Excel før import, hvis eksterne værdier mangler.

Makroindhold, aktive forbindelser, ødelagt fælles projektmappestruktur og andre problemer, der ikke kan afgrænses sikkert til et ark, afvises fortsat for hele filen. En tidligere import med global fejl kan forsøges igen via **Ny kørsel**, hvis der ikke blev sendt mails fra den.

## Mailstatus og opbevaring

- **Accepteret af Outlook** betyder, at Microsoft Graph har accepteret afsendelsen. Det er ikke en leveringskvittering.
- Ved ukendt resultat efter timeout eller programafbrydelse gensendes mailen ikke automatisk. Kontrollér **Sendt post** i Outlook og vælg **Mailen er sendt** eller **Klargør nyt forsøg**. Et nyt forsøg kræver derefter godkendelse.
- Den samme fil/type/periode kan ikke genimporteres, mens der allerede findes en kørsel eller kvittering, medmindre den tidligere kørsel blev annulleret eller udløb uden accepterede/usikre afsendelser. Ændrede filer vurderes som nye kørsler: sørg for at fravælge personer, der allerede er sendt fra en tidligere kørsel.
- Kladder har 48 timers levetid. Arbejdsfiler, modtagerdata og beskeder slettes ved afslutning, annullering eller udløb. Når alle sendbare personer er accepteret, afsluttes kørslen automatisk. Udeladte sendbare personer holder kørslen åben, indtil den afsluttes manuelt eller udløber.
- Oprydning sker ved start, hvert minut mens appen kører og via installationens oprydningsopgave hver time, når brugeren er logget ind. Når pc'en er slukket, sker oprydningen ved næste mulighed. Hvis opgaven ikke kan registreres, meddeler installationen det; start appen for at rydde op.
- Tekniske kvitteringer uden navne, mailadresser eller beskeder beholdes i 365 dage. De indeholder kørsels-ID, filhash, type, periode, tidspunkter og optællinger. Den afsendte mail og dens PDF kan stadig findes i Outlook.
- Data ligger i `%LOCALAPPDATA%\FastansatVikar\data`. Microsoft-tokens er krypteret med Windows DPAPI for den aktuelle bruger. Arbejdsdata er lokale filer under brugerens profil; der anvendes ikke yderligere kryptering af Excel, PDF og SQLite.

Afslut gamle kørsler og stop den tidligere mailautomatisering før første produktionskørsel. Gamle kladder og loginlagre overføres ikke automatisk. Appen er lokal på én pc og én Windows-bruger; den er ikke en fælles webtjeneste.

## Udvikling og test

Python 3.13, Windows, Excel og LibreOffice/UNO er anvendt ved udviklingen.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m app
```

`requirements-lock.txt` fastholder versionerne brugt i Windows-pakken (inklusive test- og buildværktøjer). `requirements.txt` er runtime-afhængigheder. `PAYROLL_DATA_DIR` kan pege på en separat datamappe til udvikling. `PAYROLL_NO_BROWSER=1` undlader automatisk browseråbning. `PAYROLL_PORT` kan vælge en separat lokal port til isolerede tests; standarden er 9000. Brug altid også en separat `PAYROLL_DATA_DIR` ved parallel test af den pakkede app.

`scripts/validate_browser.py` tester den rigtige brugerflade i installeret Edge med syntetiske filer og en simuleret mailtransport. Der sendes ingen mails. Screenshots gemmes i `.test-data/browser`.

`scripts/validate_windows.py <LoenAutomatisering-windows.zip> <Mail sender.zip>` sammenligner PDF'er fra original og ny kode på syntetiske Excel-filer med rigtig Excel og LibreOffice. Sidetal og rasteriserede sider sammenlignes. De originale kilder læses kun; originalens Excel-layout testes i en separat Excel-proces.

Byg installationspakken:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1
```

## Kodegrundlag

Mail senders CSS/layout og Excel-udskriftsindstillinger er videreført. LoenAutomatiserings XLSX-sanitizer og LibreOffice/UNO-worker er genbrugt. Medarbejderanalyse, fælles API, godkendelse, beskedredigering og frontend-adfærd er samlet i denne app. Det tidligere Gmail-opstartsflow og PySide-udgaven er ikke inkluderet.

`scripts/validate_outlook.py` kontrollerer en rigtig Outlook-konto og klargøring af en syntetisk mail, som kasseres uden at blive sendt. `scripts/validate_browser.py` tester begge afsendelsesmetoder med simuleret mailtransport.
