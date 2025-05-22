# ===========================================
# Program: SignMatch
# Version: 4.24.8
# Datum: 2024-05-22
# ===========================================

import pandas as pd
import re
import unicodedata
from datetime import timedelta
from rapidfuzz import fuzz

MINUTER_EFTER = 240
MINUTER_INNAN = 240
BASE_FUZZY_SCORE = 60
NAMNKEY_MERGE_THRESHOLD = 95

def strip_accents(text):
    """Tar bort accenter (ex: é -> e) men behåller åäö."""
    if not isinstance(text, str): return text
    text = unicodedata.normalize('NFKD', text)
    return ''.join([c for c in text if not unicodedata.combining(c) or c in 'åäö'])

def namnkey(namn):
    """Standardiserar namn för bättre matchning oavsett stavning/ordning."""
    namn = str(namn).strip().lower()
    namn = strip_accents(namn)
    namn = namn.replace(',', ' ').replace('-', ' ')
    namn = re.sub(r'[^a-zåäö\\s]', '', namn)
    namn_delar = sorted(set(namn.split()))
    return ' '.join(namn_delar)

def advanced_name_matching(n1, n2):
    """Använder fuzzy matching för att hitta liknande namn."""
    return max([
        fuzz.ratio(n1, n2),
        fuzz.partial_ratio(n1, n2),
        fuzz.token_sort_ratio(n1, n2),
        fuzz.token_set_ratio(n1, n2)
    ])

def merge_namnkeys(namnkeys):
    """Samlar ihop namn-keys som egentligen är samma person (med olika stavning/ordning)."""
    grupper = {}
    mapping = {}
    for nk in namnkeys:
        found = False
        for gk in grupper:
            if advanced_name_matching(nk, gk) >= NAMNKEY_MERGE_THRESHOLD:
                mapping[nk] = gk
                found = True
                break
        if not found:
            grupper[nk] = True
            mapping[nk] = nk
    return mapping

def parse_datetime_any(date, tid):
    """Gör om datum/tid-strängar till ett pd.Timestamp."""
    return pd.to_datetime(f"{str(date)} {str(tid).replace('kl ', '')}", errors='coerce')

def prepare_mcss_data(filepath):
    """Läser in MCSS-filen, skapar nyckelkolumner för matchning."""
    df = pd.read_excel(filepath)
    df['planerad_tid'] = df.apply(lambda row: parse_datetime_any(row['Skulle utföras (Datum)'], row['Skulle utföras (Tid)']), axis=1)
    df['boende_key'] = df['Boende'].apply(namnkey)
    df['signerat_key'] = df['Signerat av'].apply(namnkey)
    return df

def prepare_tes_data(filepath):
    """Läser in TES-filen, skapar nyckelkolumner för matchning."""
    df = pd.read_excel(filepath, header=2)
    df = df[df['Datum'].notna()]
    df['besok_tid'] = pd.to_datetime(df['Tid'], errors='coerce')
    df = df[df['besok_tid'].notna()]
    df['boende_key'] = df['Brukare'].apply(namnkey)
    df['utförare_key'] = df['Utförare'].apply(namnkey)
    return df

def extract_visningsnamn_map(mcss, tes):
    """Bygger en map mellan namnkey och 'visningsnamn' (för att alltid visa rätt namn)."""
    map_mcss = mcss[['signerat_key', 'Signerat av']].dropna().drop_duplicates().set_index('signerat_key')['Signerat av'].to_dict()
    map_tes = tes[['utförare_key', 'Utförare']].dropna().drop_duplicates().set_index('utförare_key')['Utförare'].to_dict()
    return {**map_mcss, **map_tes}

def match_visits(mcss, tes):
    """
    Matchar varje rad i MCSS (som är 'Ej signerad') mot TES för att hitta ansvariga.
    Kategoriserar varje rad i:
      - Ej i TES (planerat besök saknas helt)
      - Ej utförda besök (besök finns, men är 'Ej utfört', 'Bomkörning', eller 'Delvis utfört')
      - Fel matchning (besök hittas, men namn matchar inte)
      - Annars: kopplar TES-utförare till ansvarig
    """
    ej_signerade = []
    ansvariga = {}
    kategoriserade = []

    for _, row in mcss[mcss['Utfördes (Tid)'] == 'Ej signerad'].iterrows():
        plan_tid = row['planerad_tid']
        tidsstart = plan_tid - timedelta(minutes=MINUTER_INNAN)
        tidsslut = plan_tid + timedelta(minutes=MINUTER_EFTER)
        tes_fonster = tes[(tes['besok_tid'] >= tidsstart) & (tes['besok_tid'] <= tidsslut)]
        tes_boende = tes_fonster[tes_fonster['boende_key'] == row['boende_key']]

        if not tes_boende.empty:
            status = str(tes_boende.iloc[0]['Status']).strip().lower()
            if status in ['ej utfört', 'bomkörning', 'delvis utfört']:
                kategoriserade.append((row, 'Ej utförda besök', 'Status: ' + status))
            else:
                ansvarig = tes_boende.iloc[0]['Utförare']
                key = namnkey(ansvarig)
                ansvariga[key] = ansvariga.get(key, 0) + 1
                ej_signerade.append({**row.to_dict(), 'TES-utförare': ansvarig, 'Utförare_key': key})
            continue

        bästa, bästa_score = None, 0
        for _, kandidat in tes_fonster.iterrows():
            score = advanced_name_matching(row['Boende'], kandidat['Brukare'])
            if score > bästa_score:
                bästa, bästa_score = kandidat, score

        if bästa is not None and bästa_score >= BASE_FUZZY_SCORE:
            status = str(bästa['Status']).strip().lower()
            if status in ['ej utfört', 'bomkörning', 'delvis utfört']:
                kategoriserade.append((row, 'Ej utförda besök', 'Status: ' + status))
            else:
                ansvarig = bästa['Utförare']
                key = namnkey(ansvarig)
                ansvariga[key] = ansvariga.get(key, 0) + 1
                ej_signerade.append({**row.to_dict(), 'TES-utförare': ansvarig, 'Utförare_key': key})
        elif tes_fonster.empty:
            kategoriserade.append((row, 'Ej i TES', 'Inget besök i tidsfönster'))
        else:
            kategoriserade.append((row, 'Fel matchning', 'Namn matchar inte'))

    return ansvariga, pd.DataFrame(ej_signerade), pd.DataFrame([{**r.to_dict(), 'Kategori': k, 'Kommentar': c} for r, k, c in kategoriserade])

def generate_statistics(mcss, ansvariga, ej_signerade_df, kategorier_df, visningsnamn_map):
    """
    Bygger sammanställningen över antal signerade/ej signerade per medarbetare,
    och summerar ihop ALLA rader med samma visningsnamn, oavsett ursprunglig nyckel.
    Tar även med specialkategorier som 'Ej i TES', 'Fel matchning' etc.
    """
    # Samla ihop alla namn-keys från MCSS och TES
    alla_keys = set(mcss['signerat_key'].unique()).union(set(ej_signerade_df['Utförare_key'].unique()))
    key_merge_map = merge_namnkeys(alla_keys)

    # Lägg till group_key för MCSS och ej_signerade_df (för att kunna summera rätt)
    mcss['group_key'] = mcss['signerat_key'].map(key_merge_map)
    ej_signerade_df['group_key'] = ej_signerade_df['Utförare_key'].map(key_merge_map)

    # Hämta visningsnamn för varje group_key (t.ex. Efternamn, Förnamn)
    visningsnamn_grupp = pd.Series(visningsnamn_map).groupby(key_merge_map.get).first()
    visningsnamn_grupp = visningsnamn_grupp.apply(lambda x: ' '.join(sorted(re.sub(r'[,-]', ' ', str(x)).split())))

    # Summera antal signerade per medarbetare (group_key)
    stat = (
        mcss[mcss['Utfördes (Tid)'] != 'Ej signerad']
        .groupby('group_key').agg(Signerade=('Insats', 'count')).reset_index()
    )
    # Summera antal ej signerade insatser per medarbetare (group_key)
    ansvariga_df = ej_signerade_df.groupby('group_key').size().reset_index(name='Ej signerade')

    # Slå ihop
    stat_all = pd.merge(stat, ansvariga_df, on='group_key', how='outer').fillna(0)
    stat_all['Signerade'] = stat_all['Signerade'].astype(int)
    stat_all['Ej signerade'] = stat_all['Ej signerade'].astype(int)
    stat_all['Medarbetare'] = stat_all['group_key'].map(visningsnamn_grupp)
    stat_all['Totalt'] = stat_all['Signerade'] + stat_all['Ej signerade']
    stat_all['Andel ej signerade (%)'] = (stat_all['Ej signerade'] / stat_all['Totalt'].replace(0,1) * 100).round(1)

    # Lägg till specialkategorier
    for kategori in kategorier_df['Kategori'].unique():
        antal = kategorier_df[kategorier_df['Kategori'] == kategori].shape[0]
        stat_all = pd.concat([stat_all, pd.DataFrame([{
            'Medarbetare': kategori,
            'Signerade': 0,
            'Ej signerade': antal,
            'Totalt': antal,
            'Andel ej signerade (%)': 100.0
        }])], ignore_index=True)

    # ----- NYTT: slå ihop ALLA rader med samma visningsnamn -----
    # (oavsett om de kom från olika group_key/namnkey/stavning)
    stat_all['Medarbetare'] = stat_all['Medarbetare'].str.strip()
    stat_all = (
        stat_all
        .groupby('Medarbetare', as_index=False)
        .agg({
            'Signerade': 'sum',
            'Ej signerade': 'sum',
            'Totalt': 'sum',
            # Räkna om andelen på riktigt
            # (om det är en kategori, blir det 100%)
            'Andel ej signerade (%)': 'mean'
        })
    )
    stat_all['Andel ej signerade (%)'] = (
        stat_all['Ej signerade'] / stat_all['Totalt'].replace(0,1) * 100
    ).round(1)

    # -----

    # Returnera bara önskade kolumner i rätt ordning
    return stat_all[['Medarbetare', 'Signerade', 'Ej signerade', 'Totalt', 'Andel ej signerade (%)']]

def autofit_and_table(writer, df, sheet_name):
    """Autoanpassar kolumnbredd, och gör tabellen blå/likadan på alla flikar."""
    worksheet = writer.sheets[sheet_name]
    for i, column in enumerate(df.columns):
        max_len = max(df[column].astype(str).map(len).max(), len(str(column)))
        worksheet.set_column(i, i, max_len + 2)
    table_name = f"Tab_{re.sub(r'[^A-Za-z0-9_]', '_', sheet_name)[:31]}"
    worksheet.add_table(
        0, 0, len(df), len(df.columns) - 1,
        {
            'columns': [{'header': col} for col in df.columns],
            'name': table_name,
            'style': 'Table Style Medium 9'
        }
    )

def generate_tes_statistics(tes):
    """
    Summerar antal Ja i TES-kolumner per utförare, filtrerar bort irrelevanta besök.
    Rensar 'Planerad för...' från namn, och konverterar Ja/tomt till 1/0.
    """
    df = tes.copy()
    # Filtrera bort inköp, ledsagning och grupptid
    df = df[
        ~df['Besök'].str.lower().str.contains('inköp|ledsagning', na=False) &
        ~df['Brukare'].str.lower().str.contains('grupptid', na=False)
    ]
    # Omvandla "Ja" till 1, allt annat till 0 (för summering)
    for col in [
        'Reg. på plats',
        'Manuellt reg. i mobil',
        'Manuell reg. dator',
        'Uppdaterad från webb',
        'Uppdaterad från mobil',
        'Mobiltid ändrad'
    ]:
        df[col] = df[col].apply(lambda x: 1 if str(x).strip().lower() == 'ja' else 0)

    stat = df.groupby('Utförare').agg(
        Reg_pa_plats=('Reg. på plats', 'sum'),
        Manuellt_mobil=('Manuellt reg. i mobil', 'sum'),
        Manuellt_dator=('Manuell reg. dator', 'sum'),
        Uppdaterad_webb=('Uppdaterad från webb', 'sum'),
        Uppdaterad_mobil=('Uppdaterad från mobil', 'sum'),
        Mobiltid_andrad=('Mobiltid ändrad', 'sum'),
        Total=('Utförare', 'count')
    ).reset_index()

    stat['Andel felaktigt registrerade (%)'] = (
        (stat['Manuellt_mobil'] + stat['Manuellt_dator']) / stat['Total'] * 100
    ).round(1)
    stat = stat.rename(columns={
        'Utförare': 'Medarbetare',
        'Reg_pa_plats': 'Reg. på plats',
        'Manuellt_mobil': 'Manuellt reg. i mobil',
        'Manuellt_dator': 'Manuellt reg. dator',
        'Uppdaterad_webb': 'Uppdaterad från webb',
        'Uppdaterad_mobil': 'Uppdaterad från mobil',
        'Mobiltid_andrad': 'Mobiltid ändrad'
    })
    # Ta bort "(planerad för ...)" från medarbetarnamn
    stat['Medarbetare'] = stat['Medarbetare'].astype(str).str.replace(r'\s*\([Pp]lanerad för.*?\)', '', regex=True).str.strip()
    stat['Medarbetare'] = stat['Medarbetare'].str.replace(r'\s+', ' ', regex=True)
    return stat

def export_to_excel(statistik, ej_signerade_df, kategorier_df, visningsnamn_map, filepath, tes_statistik=None):
    """
    Skriver statistik, avvikelsekategorier och TES-insatsdata till Excel-fil,
    med autoanpassade kolumner och samma tabellformat för alla flikar.
    """
    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
        statistik.to_excel(writer, sheet_name='Sammanställning', index=False)
        autofit_and_table(writer, statistik, 'Sammanställning')

        ej_df = ej_signerade_df.rename(columns={
            'Boende': 'Kund',
            'TES-utförare': 'Ansvariga för ej signering'
        })[[
            'Kund', 'Insats', 'Skulle utföras (Datum)', 'Skulle utföras (Tid)', 'Ansvariga för ej signering'
        ]]
        ej_df.to_excel(writer, sheet_name='Ej signerade insatser', index=False)
        autofit_and_table(writer, ej_df, 'Ej signerade insatser')

        for kategori in ['Ej i TES', 'Ej utförda besök', 'Fel matchning']:
            df = kategorier_df[kategorier_df['Kategori'] == kategori]
            if not df.empty:
                if kategori in ['Ej utförda besök', 'Fel matchning']:
                    df = df.rename(columns={
                        'Boende': 'Kund',
                        'Kommentar': 'Orsak'
                    })[[
                        'Kund', 'Insats', 'Skulle utföras (Datum)', 'Skulle utföras (Tid)', 'Orsak'
                    ]]
                df.to_excel(writer, sheet_name=kategori, index=False)
                autofit_and_table(writer, df, kategori)

        if tes_statistik is not None:
            tes_statistik.to_excel(writer, sheet_name='TES-insatser', index=False)
            autofit_and_table(writer, tes_statistik, 'TES-insatser')

def main():
    """Huvudfunktion. Byt namn på Excelfilen här om du vill."""
    excel_namn = 'SignMatch_v4_24_8.xlsx'  # <-- Ändra gärna till valfritt filnamn!
    mcss = prepare_mcss_data('MCSS.xlsx')
    tes = prepare_tes_data('TES.xlsx')
    visningsnamn_map = extract_visningsnamn_map(mcss, tes)
    ansvariga, ej_signerade_df, kategorier_df = match_visits(mcss, tes)
    statistik = generate_statistics(mcss, ansvariga, ej_signerade_df, kategorier_df, visningsnamn_map)
    tes_statistik = generate_tes_statistics(tes)
    export_to_excel(statistik, ej_signerade_df, kategorier_df, visningsnamn_map, excel_namn, tes_statistik)
    print(f"KLART! Allt sparat i {excel_namn}")

if __name__ == "__main__":
    main()
