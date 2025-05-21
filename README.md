# SignMatch

Instruktion – SignMatch v4.23

Förberedelser
1. Säkerställ att du har Python 3.10 eller senare installerat
2. Installera nödvändiga bibliotek (i kommandotolk / terminal): pip install pandas openpyxl xlsxwriter rapidfuzz

Filförberedelser
1. Exportera data till Excel från MCSS och TES för samma tidsspann.
2. Placera de två filer i samma mapp som SignMatch_v4_23.py:
- MCSS.xlsx – utdrag ur MCSS
- TES.xlsx – export från TES

Kör du programmet
1. Öppna terminalen i mappen där filerna ligger
2. Kör scripet: python SignMatch_v4_23.py

Vad programmet gör
1. Läser MCSS och TES och matchar utifrån brukarnamn och tid
2. Identifierar:
- Ej signerade insatser, Ej utförda besök (ej utfört, bomkörning etc.)
- Felmatchningar (brukarnamn går inte att koppla)
- Ansvarig TES-utförare där möjligt

Resultat
Programmet skapar en Excel-fil:
- SignMatch_v4_23.xlsx

Med följande flikar:
- Sammanställning:
Statistik per medarbetare (signerat, ej signerat, andelar)
- Ej signerade insatser:
Brukare + insats + ansvarig TES-personal
- Ej i TES:
Insatser där brukaren inte fanns i TES inom tidsfönstret
- Ej utförda besök:
Brukaren fanns men status i TES var t.ex. "ej utfört"
- Fel matchning:
TES-besök fanns men namn kunde inte matchas
