# SignMatch – TES & MCSS Analys

**SignMatch** är ett avancerat verktyg för kommunal vård/hemtjänst, som samkör planeringssystemet **TES** och signeringssystemet **MCSS** för att skapa en heltäckande, automatiserad uppföljning. Programmet ger statistik, avvikelseanalys och ansvarsfördelning per medarbetare, visualiserat i en snygg Excel-rapport.

---

## Funktioner

**Matchar insatser** mellan TES (planering) och MCSS (signering)
**Statistik per medarbetare**: signerat, ej signerat, sent signerat
**Identifierar**: Ej i TES, Ej utförda besök, Fel matchning
**TES-registrering**: Andel reg. på plats, mobil, dator, web, mm.
**Excel-rapport**: Tydliga flikar, tabeller och summeringar
**Autoformat**: Alla tabeller formateras automatiskt med Office-blå stil

---

## Exempelflöde

1. **MCSS.xlsx** (signeringslista) och **TES.xlsx** (planeringslista) exporteras ur era system och placeras i samma mapp som programmet.
2. Programmet körs, matchar, summerar och kategoriserar insatser.
3. Resultatet sparas i en Excel-fil, där ansvar och eventuella avvikelser visas tydligt per medarbetare.

---

## Installation
**Kräver Python 3.10+**

Installera beroenden:
pip install pandas openpyxl xlsxwriter rapidfuzz

---

## Användning
1. Lägg MCSS.xlsx och TES.xlsx i samma mapp.
2. Kör programmet: python SignMatch_v4_24_8.py
3. Färdig fil skapas, t.ex. SignMatch_v4_24_8.xlsx.

---

## Output: Excel-flikar
Flik                      Innehåll
Sammanställning	          Statistik per medarbetare (signerat, ej signerat, andelar)
Ej signerade insatser	    Brukare, insats, ansvarig TES-person
Ej i TES	                Insatser där brukaren inte fanns i TES inom tidsfönstret
Ej utförda besök	        Brukaren fanns men status var t.ex. "ej utfört" eller "bomkörning"
Fel matchning	            Matchingslogiken kan inte avgöra namn på brukare eller medarbetare
TES-insatser	            Registreringsstatistik per utförare (på plats/mobil/dator/webb)

----

## Avancerad matchning
- Fuzzy matching och normalisering av namn gör att stavfel, tecken, eller olika namnordning ändå matchas rätt.
- Ansvarig TES-medarbetare identifieras även om brukarnamnet har mindre skillnader.

----

## Anpassning & support
- Ändra gärna filnamnet på Excel-rapporten i koden (variabeln excel_namn i main()).
- Vill ni lägga till fler kolumner eller flikar? Kontakta utvecklare eller skapa ett issue här på Github!

----

### Kontakt
Frågor eller vidareutveckling: Marcus Kihl
marcus.kihl@orebro.se  

----

### Licens
MIT-licens, använd fritt men ange källa vid vidareutveckling.
