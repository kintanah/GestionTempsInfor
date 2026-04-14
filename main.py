import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import traceback
import re

# --- CONFIGURATION ---
st.set_page_config(page_title="Outil Infor_Spoon - Réconciliation", layout="wide")
st.title("📊 Validation Signature : Date-Nom : Description (Heures)")


def fix_encoding(text):
    if isinstance(text, str):
        try:
            return text.encode('latin-1').decode('utf-8')
        except:
            return text
    return text


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ecarts')
    return output.getvalue()


# --- CHARGEMENT ---
st.sidebar.header("📁 Chargement")
file_beeline = st.sidebar.file_uploader("Beeline (CSV/XLSX)", type=['xlsx', 'csv'])
file_ts = st.sidebar.file_uploader("Timesheet (CSV/XLSX)", type=['xlsx', 'csv'])
file_mapping = st.sidebar.file_uploader("Mapping (CSV/XLSX)", type=['xlsx', 'csv'])

# ✅ MULTI CP (séparé uniquement par ;)
cp_filter = st.sidebar.text_input("👤 Chefs de Projet", "")
st.sidebar.caption("Séparer les noms uniquement avec ';' (ex: Wolff;Dupont;Martin)")

cp_list = [cp.strip().lower() for cp in cp_filter.split(";") if cp.strip()]

if file_beeline and file_ts and file_mapping and cp_list:
    try:
        # 1. Lecture Mapping
        df_map = pd.read_csv(file_mapping) if file_mapping.name.endswith('.csv') else pd.read_excel(file_mapping)

        ts_to_aligned = pd.Series(
            df_map.iloc[:, 2].values,
            index=df_map.iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()

        bee_to_aligned = pd.Series(
            df_map.iloc[:, 2].values,
            index=df_map.iloc[:, 1].astype(str).str.lower().str.strip()
        ).dropna().to_dict()

        # 2. Lecture Sources
        df_bee = pd.read_csv(file_beeline) if file_beeline.name.endswith('.csv') else pd.read_excel(file_beeline)
        df_ts = pd.read_csv(file_ts) if file_ts.name.endswith('.csv') else pd.read_excel(file_ts)

        # 3. Traitement Beeline
        df_bee_filt = df_bee[
            df_bee.iloc[:, 12].astype(str).str.lower().apply(
                lambda x: any(cp in x for cp in cp_list)
            )
        ].copy()

        df_bee_filt['comment_raw'] = df_bee_filt.iloc[:, 9].astype(str).apply(fix_encoding)

        df_bee_filt['bee_resp_aligned'] = df_bee_filt.iloc[:, 11].astype(str).str.lower().str.strip().map(
            bee_to_aligned
        ).fillna(df_bee_filt.iloc[:, 11])

        df_bee_filt['bee_date_saisie'] = df_bee_filt.iloc[:, 0]

        # 4. Traitement Timesheet
        df_ts['user_clean'] = df_ts.iloc[:, 1].astype(str).str.lower().str.strip()
        df_ts['resp_aligned'] = df_ts['user_clean'].map(ts_to_aligned).fillna(df_ts.iloc[:, 1])
        df_ts['dt_ref'] = pd.to_datetime(df_ts.iloc[:, 2], errors='coerce')

        df_ts_bill = df_ts[
            df_ts.iloc[:, 5].astype(str).str.strip().str.lower() == 'no'
        ].copy()

        # Groupement
        ts_final = df_ts_bill.groupby(['resp_aligned', 'dt_ref'], as_index=False).agg({
            df_ts.columns[7]: 'sum',
            df_ts.columns[8]: lambda x: " | ".join(x.astype(str).unique())
        })

        ts_final.columns = ['responsable', 'date_travail', 'Heures_TS', 'Desc_TS']

        # 5. Fonction Match Signature
        def match_signature(row):
            if pd.isnull(row['date_travail']):
                return pd.Series(["❌ Date vide", "N/A", "N/A"])

            target_date_str = row['date_travail'].strftime('%d/%m/%Y')
            target_resp = str(row['responsable']).lower()
            target_firstname = target_resp.split()[0]
            h_val = float(row['Heures_TS'])

            hour_pattern = rf"\({h_val:g}\s*(h|hr|hrs)\)"

            for _, bee_row in df_bee_filt.iterrows():
                lines = [l.strip() for l in str(bee_row['comment_raw']).split('\n') if l.strip()]

                for line in lines:
                    line_low = line.lower()

                    if target_date_str in line:
                        if target_resp in line_low or target_firstname in line_low:
                            if re.search(hour_pattern, line_low):
                                return pd.Series([
                                    "✅ Billé",
                                    bee_row['bee_resp_aligned'],
                                    bee_row['bee_date_saisie']
                                ])

            return pd.Series(["❌ Pas biller", "🚫 INCONNU", "N/A"])

        # Application
        ts_final[['Statut_Beeline', 'Responsable_Beeline', 'Date_Saisie_Beeline']] = ts_final.apply(
            match_signature, axis=1
        )

        # 6. Affichage
        st.subheader(f"Réconciliation pour les CP : {'; '.join(cp_list)}")

        dev_list = sorted(ts_final['responsable'].unique())
        sel_devs = st.sidebar.multiselect("Filtrer par Responsable", dev_list, default=dev_list)

        df_viz = ts_final[ts_final['responsable'].isin(sel_devs)].copy()

        if not df_viz.empty:
            st.plotly_chart(
                px.pie(
                    df_viz,
                    names='Statut_Beeline',
                    color='Statut_Beeline',
                    color_discrete_map={
                        '✅ Billé': '#27ae60',
                        '❌ Pas biller': '#e74c3c'
                    }
                ),
                use_container_width=True
            )

            df_display = df_viz.copy()
            df_display['date_travail'] = df_display['date_travail'].dt.strftime('%Y-%m-%d')

            st.dataframe(
                df_display.sort_values(by='date_travail', ascending=False),
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                "📥 Télécharger Rapport Excel",
                to_excel(df_viz),
                "reconciliation_Infor_Spoon.xlsx"
            )
        else:
            st.warning("Aucune donnée à afficher.")

    except Exception as e:
        st.error("Erreur technique lors de l'analyse")
        st.code(traceback.format_exc())

else:
    st.info("👋 Bonjour ! Charge les 3 fichiers et saisis au moins un CP (séparés par ';').")