import streamlit as st
from streamlit_gsheets import GSheetsConnection
import gspread # On importe l'outil de précision
import pandas as pd

# 1. TOUJOURS en premier
st.set_page_config(page_title="Calculateur Liasses fiscales 2033", layout="wide")

# 2. Création de la connexion
conn = st.connection("gsheets", type=GSheetsConnection)

st.title("🏠 Calculateur des liasses fiscales 2033")
st.write("Remplissez les champs ci-dessous pour calculer vos liasses fiscales 2033.")

# Organisation par onglets
tab1, tab2, tab3 = st.tabs(["🏗️ Acquisition & Travaux", "💰 Exploitation (Annuelle)", "📉 Antériorité & Report"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        prix_achat = st.number_input("Prix d'achat (€)", min_value=0, value=200000, step=1000)
        frais_notaire = st.number_input("Frais de notaire (€)", min_value=0, value=15000)
        frais_agence = st.number_input("Frais d'agence (€)", min_value=0, value=0)
    with col2:
        mobilier = st.number_input("Mobilier (€)", min_value=0, value=5000)
        travaux = st.number_input("Travaux à amortir (€)", min_value=0, value=0)
        debut_activite = st.date_input("Date de début d'activité", value=None)

with tab2:
    col_a, col_b = st.columns(2)
    with col_a:
        loyer = st.number_input("Loyer annuel CC (€)", min_value=0, value=12000)
        charges = st.number_input("Charges de copropriété (€/an)", min_value=0, value=1200)
        taxe_fonciere = st.number_input("Taxe foncière (€/an)", min_value=0, value=800)
    with col_b:
        interets = st.number_input("Intérêts d'emprunt (€/an)", min_value=0, value=3000)

with tab3:
    st.info("Ces champs concernent vos reports des années précédentes.")
    amort_excedentaires = st.number_input("Amortissements excédentaires ANTÉRIEURS (€)", min_value=0, value=0)
    deficits_anterieurs = st.number_input("Déficits ANTÉRIEURS (€)", min_value=0, value=0)
    dispo_anterieurs = st.number_input("Disponibilité Antérieurs (€)", min_value=0, value=0)

from fpdf import FPDF

def create_pdf(p1, p2, p3, p4, p5):
    pdf = FPDF(orientation='L', unit='mm', format='A4') # Mode Paysage pour plus de place
    pdf.set_auto_page_break(auto=True, margin=15)
    
    def ajouter_page_tableau(titre, data, largeurs_manuelles=None):
        pdf.add_page()
        pdf.set_font("helvetica", 'B', 14)
        pdf.cell(0, 10, txt=titre, ln=True, align='C')
        pdf.ln(5)
        
        if not data: return

        nb_cols = len(data[0])
        
        # Si on n'a pas de largeurs manuelles, on divise par le nombre de colonnes
        if largeurs_manuelles is None:
            col_widths = [270 / nb_cols] * nb_cols
        else:
            col_widths = largeurs_manuelles

        pdf.set_font("helvetica", size=7)
        
        for row in data:
            start_y = pdf.get_y()
            clean_row = [str(item).replace('"', '') if item is not None else "" for item in row]
            
            # Calcul de la hauteur de ligne
            h_list = [len(pdf.multi_cell(col_widths[i], 5, txt=clean_row[i], split_only=True)) * 5 
                      for i in range(min(len(clean_row), len(col_widths)))]
            line_h = max(h_list) if h_list else 5
            
            curr_x = pdf.get_x()
            for i in range(min(len(clean_row), len(col_widths))):
                pdf.rect(curr_x, start_y, col_widths[i], line_h)
                pdf.multi_cell(col_widths[i], 5, txt=clean_row[i], align='L')
                curr_x += col_widths[i]
                pdf.set_xy(curr_x, start_y)
            
            pdf.ln(line_h)

    # Appels simplifiés (le code calcule les largeurs tout seul)
    ajouter_page_tableau("PAGE 1 : BILAN", p1)
    ajouter_page_tableau("PAGE 2 : COMPTE DE RÉSULTAT", p2, largeurs_manuelles=[90, 20, 40, 40, 40, 40])
    ajouter_page_tableau("PAGE 3 : IMMOBILISATIONS", p3)
    ajouter_page_tableau("PAGE 3 : AMORTISSEMENTS", p4)
    ajouter_page_tableau("PAGE 4 : SUIVI DES DÉFICITS", p5)
    
    return pdf.output()



    # Définition des largeurs de colonnes pour chaque page (total ~190)
    # Page 1 : Libellé (90), Code (20), Brut (40), Net (40)
    ajouter_page_tableau("PAGE 1 : BILAN", p1, [90, 20, 40, 40])
    
    # Page 2 : Libellé (110), Code (20), Montant (60)
    # Exemple pour 5 colonnes lues dans ton Sheet (Libellé, Code, vide, vide, Montant)
    # On donne 150mm au texte, 20mm au code, et on répartit le reste
    ajouter_page_tableau("PAGE 2 : COMPTE DE RÉSULTAT", p2, largeurs_manuelles=[150, 20, 33, 33, 34])
    
    # Page 3 : Libellé (70) + 4 colonnes de chiffres (30 chacune)
    ajouter_page_tableau("PAGE 3 : IMMOBILISATIONS", p3, [70, 30, 30, 30, 30])
    ajouter_page_tableau("PAGE 3 : AMORTISSEMENTS", p4, [70, 30, 30, 30, 30])
    
    # Page 4 : Libellé (130), Montant (60)
    ajouter_page_tableau("PAGE 4 : DÉFICITS", p5, [130, 60])
    
    return pdf.output()


if st.button("Enregistrer les données"):
    try:
        # 1. On récupère les identifiants de ton fichier secrets.toml
        creds = st.secrets["connections"]["gsheets"]
        
        # 2. On établit une connexion directe "Pro"
        gc = gspread.service_account_from_dict(creds)
        
        # 3. On ouvre le fichier par son URL
        sh = gc.open_by_url(creds["spreadsheet"])
        
        # 4. On sélectionne la feuille (Nom de l'onglet, ex: "Feuille1")
        
        # On transforme la date en texte format "français"
        date_str = debut_activite.strftime("%d/%m/%Y") if debut_activite else ""

        # 5. On écrit précisément dans les cases cibles
        # update_acell est la commande magique pour la précision
        worksheet = sh.worksheet("test python")
        worksheet.update_acell('B4', prix_achat)
        worksheet.update_acell('B5', frais_notaire)
        worksheet.update_acell('B6', frais_agence)
        worksheet.update_acell('B8', travaux)
        worksheet.update_acell('B7', mobilier)
        worksheet.update_acell('B12', date_str)

        worksheet.update_acell('B84', loyer)
        worksheet.update_acell('B85', charges)
        worksheet.update_acell('B86', taxe_fonciere)
        worksheet.update_acell('B87', interets)

        worksheet.update_acell('B93', amort_excedentaires)
        worksheet.update_acell('B98', deficits_anterieurs)
        worksheet.update_acell('B103', dispo_anterieurs)
        
        
        

        st.success(f"✅ Liasse mise à jour !")
        
        st.divider() # Petite ligne de séparation visuelle

        # --- 2. RÉCUPÉRATION ET AFFICHAGE DES 4 PAGES ---
        with st.spinner("Extraction des résultats..."):
                        
            # On cible l'onglet où se trouvent les résultats (ex: index 1)
            ws_res = sh.get_worksheet(1) 
            
            # --- NETTOYAGE ET AFFICHAGE ESTHÉTIQUE ---
            def afficher_tableau_pro(data, titre):
                if not data:
                    return st.warning(f"Aucune donnée trouvée pour {titre}")
                
                df = pd.DataFrame(data)
                
                # 1. On remplace les None par du vide avant de nettoyer
                df = df.fillna("")
                
                # 2. Nettoyage : suppression des lignes/colonnes entièrement vides
                # On utilise "" maintenant car les None ont été remplacés
                df = df.mask(df == "").dropna(how='all').dropna(axis=1, how='all').fillna("")
                
                st.subheader(titre)
                return st.dataframe(df, use_container_width=True, hide_index=True)

            # Création des onglets
            tab_p1, tab_p2, tab_p3, tab_p4 = st.tabs([
                "📄 Page 1 : Bilan", 
                "📄 Page 2 : Résultat", 
                "📄 Page 3 : Immo & Amort.", 
                "📄 Page 4 : Déficits"
            ])

            with tab_p1:
                p1_data = ws_res.get('A109:G124')
                afficher_tableau_pro(p1_data, "État de l'Actif / Passif")

            with tab_p2:
                p2_data = ws_res.get('A129:E144')
                # Petit plus : Affichage d'un chiffre clé en haut
                if p2_data:
                    res_fiscal = p2_data[-1][-3]
                    st.metric("RÉSULTAT FISCAL FINAL", f"{res_fiscal} €")
                afficher_tableau_pro(p2_data, "Détail du Résultat Fiscal")

            with tab_p3:
                p3_data = ws_res.get('A151:I158')
                afficher_tableau_pro(p3_data, "Tableau des Immobilisations")
                
                p4_data = ws_res.get('A162:I169')
                afficher_tableau_pro(p4_data, "Tableau des Amortissements")

            with tab_p4:
                p5_data = ws_res.get('A178:C182')
                afficher_tableau_pro(p5_data, "Suivi des Déficits")


            # --- BOUTON DE TÉLÉCHARGEMENT PDF ---
            # On génère le bytearray
            pdf_raw = create_pdf(p1_data, p2_data, p3_data, p4_data, p5_data)
            
            # On convertit impérativement en bytes pour Streamlit
            pdf_final = bytes(pdf_raw) 
            
            st.download_button(
                label="📥 Télécharger la Liasse en PDF",
                data=pdf_final,
                file_name=f"liasse_fiscale_{date_str.replace('/', '-')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )


    except Exception as e:
        st.error(f"Erreur de liaison : {e}")