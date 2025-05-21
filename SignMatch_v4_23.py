
# ===========================================
# Program: SignMatch
# Version: 4.23
# Datum: 2024-05-21
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
    if not isinstance(text, str): return text
    text = unicodedata.normalize('NFKD', text)
    return ''.join([c for c in text if not unicodedata.combining(c) or c in 'åäö'])

def namnkey(namn):
    namn = str(namn).strip().lower()
    namn = strip_accents(namn)
    namn = namn.replace(',', ' ').replace('-', ' ')
    namn = re.sub(r'[^a-zåäö\s]', '', namn)
    namn_delar = sorted(set(namn.split()))
    return ' '.join(namn_delar)

def advanced_name_matching(n1, n2):
    return max([
        fuzz.ratio(n1, n2),
        fuzz.partial_ratio(n1, n2),
        fuzz.token_sort_ratio(n1, n2),
        fuzz.token_set_ratio(n1, n2)
    ])

def merge_namnkeys(namnkeys):
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
    return pd.to_datetime(f"{str(date)} {str(tid).replace('kl ', '')}", errors='coerce')

def prepare_mcss_data(filepath):
    df = pd.read_excel(filepath)
    df['planerad_tid'] = df.apply(lambda row: parse_datetime_any(row['Skulle utföras (Datum)'], row['Skulle utföras (Tid)']), axis=1)
    df['boende_key'] = df['Boende'].apply(namnkey)
    df['signerat_key'] = df['Signerat av'].apply(namnkey)
    return df

def prepare_tes_data(filepath):
    df = pd.read_excel(filepath, header=2)
    df = df[df['Datum'].notna()]
    df['besok_tid'] = pd.to_datetime(df['Tid'], errors='coerce')
    df = df[df['besok_tid'].notna()]
    df['boende_key'] = df['Brukare'].apply(namnkey)
    df['utförare_key'] = df['Utförare'].apply(namnkey)
    return df

def extract_visningsnamn_map(mcss, tes):
    map_mcss = mcss[['signerat_key', 'Signerat av']].dropna().drop_duplicates().set_index('signerat_key')['Signerat av'].to_dict()
    map_tes = tes[['utförare_key', 'Utförare']].dropna().drop_duplicates().set_index('utförare_key')['Utförare'].to_dict()
    return {**map_mcss, **map_tes}

def match_visits(mcss, tes):
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
    alla_keys = set(mcss['signerat_key'].unique()).union(set(ej_signerade_df['Utförare_key'].unique()))
    key_merge_map = merge_namnkeys(alla_keys)

    mcss['group_key'] = mcss['signerat_key'].map(key_merge_map)
    ej_signerade_df['group_key'] = ej_signerade_df['Utförare_key'].map(key_merge_map)

    visningsnamn_grupp = pd.Series(visningsnamn_map).groupby(key_merge_map.get).first()
    visningsnamn_grupp = visningsnamn_grupp.apply(lambda x: ' '.join(sorted(re.sub(r'[,-]', ' ', str(x)).split())))

    stat = mcss[mcss['Utfördes (Tid)'] != 'Ej signerad'].groupby('group_key').agg(Signerade=('Insats', 'count')).reset_index()
    ansvariga_df = ej_signerade_df.groupby('group_key').size().reset_index(name='Ej signerade')

    stat_all = pd.merge(stat, ansvariga_df, on='group_key', how='outer').fillna(0)
    stat_all['Signerade'] = stat_all['Signerade'].astype(int)
    stat_all['Ej signerade'] = stat_all['Ej signerade'].astype(int)
    stat_all['Medarbetare'] = stat_all['group_key'].map(visningsnamn_grupp)
    stat_all['Totalt'] = stat_all['Signerade'] + stat_all['Ej signerade']
    stat_all['Andel ej signerade (%)'] = (stat_all['Ej signerade'] / stat_all['Totalt'].replace(0,1) * 100).round(1)

    for kategori in kategorier_df['Kategori'].unique():
        antal = kategorier_df[kategorier_df['Kategori'] == kategori].shape[0]
        stat_all = pd.concat([stat_all, pd.DataFrame([{
            'Medarbetare': kategori,
            'Signerade': 0,
            'Ej signerade': antal,
            'Totalt': antal,
            'Andel ej signerade (%)': 100.0
        }])], ignore_index=True)

    return stat_all[['Medarbetare', 'Signerade', 'Ej signerade', 'Totalt', 'Andel ej signerade (%)']]

def autofit_and_table(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for i, column in enumerate(df.columns):
        max_len = max(df[column].astype(str).map(len).max(), len(str(column)))
        worksheet.set_column(i, i, max_len + 2)
    worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
        'columns': [{'header': col} for col in df.columns],
        'name': f'Tab_{sheet_name.replace(" ", "_")[:31]}'
    })

def export_to_excel(statistik, ej_signerade_df, kategorier_df, visningsnamn_map, filepath):
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

def main():
    mcss = prepare_mcss_data('MCSS.xlsx')
    tes = prepare_tes_data('TES.xlsx')
    visningsnamn_map = extract_visningsnamn_map(mcss, tes)
    ansvariga, ej_signerade_df, kategorier_df = match_visits(mcss, tes)
    statistik = generate_statistics(mcss, ansvariga, ej_signerade_df, kategorier_df, visningsnamn_map)
    export_to_excel(statistik, ej_signerade_df, kategorier_df, visningsnamn_map, 'SignMatch_v4_23.xlsx')
    print("KLART! Allt sparat i SignMatch_v4_23.xlsx")

if __name__ == "__main__":
    main()
