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

#set par(justify: true)

//#set heading(numbering: "1.")

#show link: it => {underline(text(maroon.darken(40%), it), stroke: (dash: "dotted"))}


#align(center)[
  #text(size: 20pt, weight: "bold")[
    Skriv ut elevbesvarelser automatisk
  ]

  #v(1cm)
]

= Hensikt

Dette programmet er laget for å gjøre det enkelt for lærere å skrive ut alle elevbesvarelser fra itslearning på én gang. I stedet for å åpne og skrive ut hver besvarelse manuelt, kan du bruke dette programmet til å skrive ut alt automatisk. Programmet kan legge til elevens navn i toppteksten og sidetall i bunnteksten på dokumentene hvis du ønsker det.

= Hva programmet gjør

Dette programmet skriver ut alle dokumenter og bilder i en mappe (eller zip-fil) og alle undermapper automatisk. Programmet har tre forskjellige moduser avhengig av filtype:

== 1. Word-filer (.docx)
Word-dokumenter skrives ut som de er skrevet. Hvis du velger det, legges elevens navn til i toppteksten og sidenummer i bunnteksten (format: "Side 1 av 5").

== 2. HTML-filer, bilder og tekstfiler
For hver mappe kombineres alle HTML-filer (.html, .htm), bilder (.jpg, .jpeg, .png, .gif, .bmp) og tekstfiler (.txt) til én midlertidig HTML-fil. Denne skrives ut med elevens navn og sidenummer hvis du velger det. Den midlertidige filen slettes automatisk etter utskrift.

== 3. PDF-filer (.pdf)
PDF-filer skrives ut som de er, uten topptekst eller bunntekst (dette er en teknisk begrensning).

#block(
  fill: blue.lighten(79%),
  inset: 1em,
  radius: 0.3em,
)[
  *Viktig:* Originaldokumentene endres IKKE. Før utskriften starter får du velge om du vil legge til topptekst og bunntekst på Word-filer og HTML-filer:

  - *Elevens navn* øverst (mappenavnet fra zip-filen)
  - *Sidenummer* nederst (format: "Side 1 av 5")

  Standard er å legge til topptekst og bunntekst (trykk bare #kbd[Enter]), men du kan velge å hoppe over dette ved å trykke #kbd[n]-tasten.
]

= Hvordan bruke programmet

== Steg 0: Last ned programmet

#link("https://raw.githubusercontent.com/stalegjelsten/print-word-files/main/print.ps1")[Last ned print.ps1] (høyreklikk og velg "Lagre lenke som..." eller "Save link as...") fra GitHub.

Lagre filen et sted på datamaskinen din (for eksempel på Skrivebordet).

== Steg 1: Last ned besvarelser fra itslearning

Logg inn på itslearning og gå til oppgaven du vil skrive ut besvarelser fra.

#figure(
  image("assets/itslearning-download-answers.png", width: 100%),
  caption: [Nedlasting av besvarelser fra itslearning]
)

+ Vis kun elevene som har levert oppgaven ved å velge *Vis:* *Levert*
+ *Merk alle elevene* du vil skrive ut besvarelser for (huk av øverst for å velge alle)
+ Klikk på *Handlinger*
+ *"Last ned besvarelser"*
+ En zip-fil lastes ned til datamaskinen din (vanligvis i Nedlastinger-mappen)

== Steg 2: Kjør utskriftsprogrammet

+ Høyreklikk på `print.ps1` og velg *"Kjør med PowerShell"* eller *"Run with PowerShell"*
+ Et vindu åpnes. Les informasjonen som vises
+ Det åpnes automatisk et vindu hvor du kan velge enten:
  - *Zip-filen* du lastet ned fra itslearning (anbefalt)
  - *En mappe* som inneholder dokumenter
+ Velg filen/mappen og klikk OK
+ Programmet viser en oversikt over alle filer som skal skrives ut
+ Du blir spurt om du vil legge til mappenavn og sidenummer på utskriftene (trykk Enter for ja, eller skriv 'n' for nei)
+ Programmet skriver nå ut alle dokumenter og bilder
+ Når det er ferdig, trykk Enter for å lukke vinduet

= Krav

Programmet fungerer kun på Windows, og det krever at følgende programmer er installert:


- *Microsoft Word* -- For å skrive ut Word-dokumenter og HTML-filer
- *Adobe Acrobat Reader* -- For å skrive ut PDF-filer


= Endre innstillinger

Du kan tilpasse programmet ved å åpne `print.ps1` i Notisblokk og endre disse linjene øverst i filen:

- *Linje 4* (`$CONFIG_MARGIN_CM`): Endre sidemarger i centimeter (standard: 2.0 cm)
- *Linje 5* (`$CONFIG_IMAGE_WIDTH_CM`): Endre maksimal bildebredde i centimeter (standard: 17.0 cm)
- *Linje 6* (`$CONFIG_PRINTER`): Endre hvilken printer som skal brukes

*Standard printer er:* `\\TDCSOM30\Sikker_UtskriftCS`

= Feilsøking

== Hvis du får feilmelding om "execution policy"

Dette betyr at datamaskinen din blokkerer PowerShell-skript av sikkerhetsgrunner. Slik fikser du det:

+ Høyreklikk på Start-knappen (nederst til venstre på skjermen)
+ Velg "Windows PowerShell (administrator)" eller "Terminal (administrator)"
+ Når det åpner seg et vindu med hvit eller blå tekst, skriv inn følgende og trykk Enter:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
+ Skriv `J` og trykk Enter når du får spørsmål
+ Du kan nå lukke dette vinduet og prøve å kjøre `print.ps1` på nytt

_Dette trenger du bare å gjøre én gang på datamaskinen din._

== Hvis PDF-filer ikke skrives ut

Installer #link("https://get.adobe.com/no/reader/")[Adobe Acrobat Reader]. 

== Hvis Word- eller HTML-filer ikke skrives ut

Kontroller at Microsoft Word er installert på datamaskinen.
