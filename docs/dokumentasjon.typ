
#set document(
  title: "Skriv ut dokumenter automatisk",
  author: "Ståle Gjelsten",
)

#set page(
  paper: "a4",
  margin: (x: 2.5cm, y: 2.5cm),
  numbering: "1",
)

#set text(
  lang: "nb",
  size: 11pt,
)

#let kbd(body) = text(font: "Libertinus Keyboard")[#body]

#show raw.where(block: false): it => box(
  it,
  fill: black.lighten(95%),
  inset: (x: 2pt, y: 0pt),
  outset: (y: 3pt),
  radius: 2pt,
)

#show raw.where(block: true): it => block(
  it,
  fill: black.lighten(95%),
  inset: 8pt,
  radius: 2pt,
  width: 100%,
)


#set par(justify: true)

//#set heading(numbering: "1.")

#show link: it => {
  underline(stroke: (paint: red.darken(80%), thickness: 1pt, dash: "dotted"), offset: 0.15em)[#text(fill: red.darken(80%))[#it]]
  if type(it.dest) == type("hello") {
    sym.wj
    h(1.0pt)
    sym.wj
    super(box(height: 0.05em, text(size: 0.7em, fill: red.darken(80%))[#sym.arrow.tr]))
  }
}




#align(center)[
  #text(size: 20pt, weight: "bold")[
    Skriv ut alle elevbesvarelser fra itslearning
  ]

  #text(size: 12pt)[
    Ståle Gjelsten \
    Dahlske videregående skole
  ]
  #v(1cm)
]

= Hensikt

Dette programmet er laget for å gjøre det enkelt for lærere å skrive ut alle elevbesvarelsene fra itslearning på én gang. I stedet for å åpne og skrive ut hver besvarelse manuelt, kan du bruke dette programmet til å skrive ut alt automatisk. Programmet kan legge til elevens navn i toppteksten og sidetall i bunnteksten på dokumentene hvis du ønsker det.

= Hva programmet gjør

Dette programmet skriver ut alle dokumenter og bilder i en mappe (eller zip-fil) og alle undermapper automatisk. Programmet har tre forskjellige moduser avhengig av filtype:

== 1. Word-filer (.docx)
Word-dokumenter skrives ut som de er skrevet. Hvis du velger det, legges elevens navn til i toppteksten og sidenummer i bunnteksten (format: "Side 1 av 5").

== 2. HTML-filer, bilder og tekstfiler
For hver elev kombineres alle HTML-filer (.html, .htm), bilder (.jpg, .jpeg, .png, .gif, .bmp) og tekstfiler (.txt) til én midlertidig HTML-fil. Denne skrives ut med elevens navn og sidenummer hvis du velger det. Den midlertidige filen slettes automatisk etter utskrift.

== 3. PDF-filer (.pdf)
PDF-filer skrives ut som de er, uten topptekst eller bunntekst (dette er en teknisk begrensning).

#block(
  fill: blue.lighten(79%),
  inset: 1em,
  radius: 0.3em,
)[
  *Viktig:* Originaldokumentene endres IKKE. Før utskriften starter vises en interaktiv meny hvor du kan velge innstillinger:

  - *Topptekst og bunntekst*: elevens navn øverst (mappenavnet) og sidenummer nederst (format: "Side 1 av 5")
  - *Kommentarer*: om kommentarer i Word-dokumenter skal skrives ut

  Standardinnstillingen legger til topptekst og bunntekst. Bruk piltastene og #kbd[Space] for å endre innstillingene, og tast #kbd[Enter] for å starte utskriften.
]

= Hvordan bruke programmet

== Steg 0: Last ned programmet

Last ned programmet #link("https://raw.githubusercontent.com/stalegjelsten/print-word-files/main/print.ps1")[*print.ps1*] (høyreklikk og velg "Lagre lenke som..." eller "Save link as...").

#block(
  fill: blue.lighten(79%),
  inset: 1em,
  radius: 0.3em,
)[
*Viktig etter nedlasting:* Windows blokkerer automatisk filer som er lastet ned fra internett. 

  Du må fjerne denne blokkeringen før du kan kjøre programmet: høyreklikk på `print.ps1` → *Egenskaper* → huk av *«Fjern blokkering»* (eller *«Unblock»*) nederst → klikk *OK*.
]



== Steg 1: Last ned besvarelser fra itslearning

Logg inn på itslearning og gå til oppgaven du vil skrive ut besvarelser fra. 

#figure(
  image("assets/itslearning-download-answers.png", width: 60%),
  caption: [Nedlasting av besvarelser fra itslearning]
)<itslearning>

+ Vis kun elevene som har levert oppgaven ved å velge *Vis:* *Levert*. Se @itslearning.
+ *Merk alle elevene* du vil skrive ut besvarelser for (huk av øverst for å velge alle)
+ Klikk på *Handlinger*
+ *"Last ned besvarelser"*
+ Itslearning bruker litt tid på å samle alle besvarelsene til én fil. Trykk på *klikk her for å laste ned* når filene er klare.
+ En zip-fil#footnote[En zip-fil er mappe som pakket sammen til en fil slik at det er enkelt å laste den ned og flytte den] lastes ned til datamaskinen din (vanligvis i Nedlastinger-mappen)

== Steg 2: Kjør utskriftsprogrammet

+ Åpne mappen hvor du lagret `print.ps1`.
+ Høyreklikk på `print.ps1` og velg *"Kjør med PowerShell"* eller *"Run with PowerShell"*
  - Et vindu åpnes med informasjon om hvilken printer som er valgt
+ Det åpnes automatisk et vindu hvor du kan velge enten:
  - *Zip-filen* du lastet ned fra itslearning
  - *En mappe* som inneholder dokumenter
+ Velg filen eller mappen og klikk OK
+ Programmet skanner filene og viser en interaktiv meny
+ I menyen kan du:
  - Avhuke enkeltfiler du *ikke* vil skrive ut med #kbd[Space]
  - Slå av/på *topptekst og bunntekst* (mappenavn + sidenummer)
  - Slå av/på utskrift av *kommentarer* i Word-dokumenter
  - Navigere med #kbd[↑] og #kbd[↓]
+ Tast #kbd[Enter] for å starte utskriften, eller #kbd[Esc] for å avbryte
+ Programmet skriver ut filene og viser fremdrift i terminalen
+ Når det er ferdig vises en oppsummering -- tast #kbd[Enter] for å lukke vinduet
  - Første gang programmet kjøres får du spørsmål om å opprette en snarvei på skrivebordet. Snarveien gjør at du kan dobbeltklikke for å kjøre programmet direkte.

= Krav

Programmet fungerer kun på Windows, og det krever at følgende programmer er installert:


- *Microsoft Word* -- For å skrive ut Word-dokumenter og HTML-filer
- *Adobe Acrobat Reader* -- For å skrive ut PDF-filer


= Endre innstillinger

Du kan tilpasse programmet ved å åpne `print.ps1` i Notisblokk og endre disse linjene øverst i filen:

- *Linje 4* (`$CONFIG_MARGIN_CM`): Endre sidemarger i centimeter (standard: 2.0 cm)
- *Linje 5* (`$CONFIG_IMAGE_WIDTH_CM`): Endre maksimal bildebredde i centimeter (standard: 17.0 cm)
- *Linje 6* (`$CONFIG_PRINTER`): Endre hvilken printer som skal brukes

*Standard printer er:* `\\TDCSPRN30\Sikker_UtskriftCS`

= Feilsøking

== Hvis du får feilmelding om at skriptet "ikke er signert" eller "execution policy"

Dette skyldes at Windows blokkerer filer lastet ned fra internett, og/eller at datamaskinen din blokkerer PowerShell-skript av sikkerhetsgrunner. Du må gjøre begge stegene under (ingen av dem krever administrator):

*Steg 1 -- Fjern blokkeringen av filen:*

Høyreklikk på `print.ps1` → *Egenskaper* → huk av *«Fjern blokkering»* (eller *«Unblock»*) nederst → klikk *OK*.

Alternativt kan du åpne PowerShell i samme mappe og kjøre:

```powershell
Unblock-File -Path .\print.ps1
```

_Dette må gjøres én gang per datamaskin._

*Steg 2 -- Tillat kjøring av lokale skript:*

+ Trykk på Start-knappen og søk etter "PowerShell"
+ Klikk på "Windows PowerShell"
+ Når det åpner seg et vindu med blå bakgrunn, skriv inn følgende og tast deretter #kbd[Enter]:

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

+ Skriv #kbd[J] og tast #kbd[Enter] når du får spørsmål
+ Du kan nå lukke dette vinduet og prøve å kjøre `print.ps1` på nytt

_Dette trenger du bare å gjøre én gang på datamaskinen din._

== Hvis PDF-filer ikke skrives ut

Installer #link("https://get.adobe.com/no/reader/")[Adobe Acrobat Reader]. 

== Hvis Word- eller HTML-filer ikke skrives ut

Kontroller at Microsoft Word er installert på datamaskinen.

== Hva gjør filen .no-shortcut-prompt?

Hvis du takker nei til å opprette en snarvei til programmet på skrivebordet så opprettes det en spesiell fil på maskinen din som heter `.no-shortcut-prompt`. Hvis denne filen eksisterer så vil ikke programmet spørre deg om du vil opprette snarveien på skrivebordet. Det er ufarlig å slette filen.

= Kontakt og lisens
Programmet er utviklet av #link("https://github.com/stalegjelsten")[Ståle Gjelsten]. Funksjonene for å skrive ut PDF-filer, å kombinere HTML-filer , å legge til topp- og bunntekst og menysystemet er utviklet sammen med språkmodellen #link("https://www.anthropic.com/claude/sonnet")[Claude Sonnet 4.6 fra Anthropic].

Programmet er lisensiert med #link("https://opensource.org/license/MIT")[MIT-lisens], som betyr at det kan fritt brukes, endres og deles videre. Programmet leveres "som det er", og uten noen form for garantier.
