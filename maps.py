import re
import time
import xml.etree.ElementTree as ET
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Import Excel → Google My Maps",
    page_icon="🗺️",
    layout="wide",
)

st.title("🗺️ Convertisseur Excel vers Google My Maps")
st.markdown(
    "Importe un fichier Excel de prospection (comme `maps .xlsx`), filtre tes cibles, "
    "nomme chaque point avec la **Raison Sociale** et génère un fichier prêt pour **Google My Maps**."
)


def geocoder_adresse(requete: str):
    """Récupère (lat, lon) via l'API Adresse officielle (data.gouv.fr)."""
    if not requete or not str(requete).strip():
        return None, None
    try:
        resp = requests.get(
            "https://api-adresse.data.gouv.fr/search/",
            params={"q": str(requete).strip(), "limit": 1},
            timeout=10,
        )
        if resp.status_code == 200:
            features = resp.json().get("features", [])
            if features:
                lon, lat = features[0]["geometry"]["coordinates"]
                return float(lat), float(lon)
    except requests.exceptions.RequestException:
        pass
    return None, None


def nettoyer_texte(val) -> str:
    if pd.isna(val):
        return ""
    texte = str(val).strip()
    return "" if texte.lower() in ("nan", "none", "nat") else texte


def est_adresse(texte: str) -> bool:
    if not texte:
        return False
    motif = (
        r"\b\d{5}\b|\b(RUE|ROUTE|RTE|PLACE|CHEMIN|IMPASSE|AVENUE|AVE|BD|BOULEVARD|"
        r"ALLEE|ALL|SQUARE|QUAI|LIEU DIT|LIEU-DIT|ZAC|ZA)\b"
    )
    return bool(re.search(motif, texte.upper()))


def extraire_effectif_num(val):
    """Extrait un entier depuis la tranche d'effectif (ex: '10 à 19 salariés' -> 19)."""
    texte = nettoyer_texte(val)
    if not texte or est_adresse(texte):
        return None
    texte_clean = texte.replace("\u202f", "").replace(" ", "")
    nombres = [int(n) for n in re.findall(r"\d+", texte_clean)]
    if not nombres:
        return None
    return max(nombres)


def normaliser_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normalise les colonnes selon le format standard ou le format décalé de maps.xlsx."""
    df = df_raw.copy()

    # Supprime les colonnes totalement vides
    df = df.dropna(axis=1, how="all")

    noms_entreprises = []
    adresses_completes = []
    effectifs = []
    activites = []
    categories_entreprise = []
    cas = []
    statuts = []
    formes_juridiques = []
    sirens = []

    for _, row in df.iterrows():
        # 1. Nom de l'entreprise (priorité à Raison Sociale / Raison sociale / Dénomination)
        nom = ""
        for col_nom in ("Raison Sociale", "Raison sociale", "Dénomination", "Nom"):
            if col_nom in df.columns:
                nom = nettoyer_texte(row.get(col_nom))
                if nom:
                    break
        if not nom:
            nom = "Entreprise sans nom"

        # 2. Détection d'une ligne au format décalé (ex. lignes 20+ de maps.xlsx)
        val_eff_col = nettoyer_texte(row.get("Tranche Effectif")) if "Tranche Effectif" in df.columns else ""
        val_cat_col = nettoyer_texte(row.get("Catégorie")) if "Catégorie" in df.columns else ""
        val_res_col = nettoyer_texte(row.get("Résultat Net (€)")) if "Résultat Net (€)" in df.columns else ""
        val_ca_col = nettoyer_texte(row.get("Chiffre d'Affaires (€)")) if "Chiffre d'Affaires (€)" in df.columns else ""

        ligne_decalee = est_adresse(val_eff_col) or ("salariés" in val_res_col.lower())

        # 3. Adresse
        if "Adresse_Complete" in df.columns and nettoyer_texte(row.get("Adresse_Complete")):
            adr = nettoyer_texte(row.get("Adresse_Complete"))
        elif ligne_decalee and val_eff_col:
            adr = val_eff_col
        else:
            adr_rue = nettoyer_texte(row.get("Adresse")) if "Adresse" in df.columns else ""
            cp = nettoyer_texte(row.get("Code Postal")).replace(".0", "") if "Code Postal" in df.columns else ""
            ville = nettoyer_texte(row.get("Ville")) if "Ville" in df.columns else ""
            parties = [p for p in (adr_rue, cp, ville) if p and p.lower() not in adr_rue.lower()]
            adr = " ".join(parties).strip()

        # 4. Effectif, Activité, Catégorie entreprise, CA, Statut, Forme juridique
        if ligne_decalee:
            eff = val_res_col if val_res_col else None
            act = val_cat_col
            cat_ent = val_ca_col
            ca_val = nettoyer_texte(row.get("Unnamed: 9")) if "Unnamed: 9" in df.columns else ""
            statut_val = nettoyer_texte(row.get("Unnamed: 10")) if "Unnamed: 10" in df.columns else ""
            forme_val = nettoyer_texte(row.get("Unnamed: 11")) if "Unnamed: 11" in df.columns else ""
        else:
            eff = val_eff_col if val_eff_col else None
            act = ""
            cat_ent = val_cat_col
            ca_val = val_ca_col
            statut_val = nettoyer_texte(row.get("Statut")) if "Statut" in df.columns else ""
            forme_val = nettoyer_texte(row.get("Forme Juridique")) if "Forme Juridique" in df.columns else ""

        siren_val = nettoyer_texte(row.get("SIREN")).replace(".0", "") if "SIREN" in df.columns else ""

        noms_entreprises.append(nom)
        adresses_completes.append(adr)
        effectifs.append(eff)
        activites.append(act)
        categories_entreprise.append(cat_ent)
        cas.append(ca_val)
        statuts.append(statut_val)
        formes_juridiques.append(forme_val)
        sirens.append(siren_val)

    return pd.DataFrame(
        {
            "Nom Entreprise": noms_entreprises,
            "SIREN": sirens,
            "Adresse": adresses_completes,
            "Tranche Effectif": effectifs,
            "Activité": activites,
            "Catégorie": categories_entreprise,
            "Chiffre d'Affaires (€)": cas,
            "Statut": statuts,
            "Forme Juridique": formes_juridiques,
        }
    )


def generer_kml(df_points: pd.DataFrame, nom_calque: str = "Prospection Entreprises") -> bytes:
    """Génère un fichier KML où chaque point est nommé par la Raison Sociale."""
    kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
    doc = ET.SubElement(kml, "Document")
    nom_doc = ET.SubElement(doc, "name")
    nom_doc.text = nom_calque or "Prospection Entreprises"

    for _, row in df_points.iterrows():
        nom_entreprise = nettoyer_texte(row.get("Nom Entreprise")) or "Entreprise sans nom"

        placemark = ET.SubElement(doc, "Placemark")
        titre = ET.SubElement(placemark, "name")
        titre.text = nom_entreprise

        desc = ET.SubElement(placemark, "description")
        champs_desc = [
            ("Entreprise", nom_entreprise),
            ("Adresse", nettoyer_texte(row.get("Adresse"))),
            ("Effectif", nettoyer_texte(row.get("Tranche Effectif"))),
            ("Activité", nettoyer_texte(row.get("Activité"))),
            ("Catégorie", nettoyer_texte(row.get("Catégorie"))),
            ("Chiffre d'Affaires (€)", nettoyer_texte(row.get("Chiffre d'Affaires (€)"))),
            ("SIREN", nettoyer_texte(row.get("SIREN"))),
            ("Statut", nettoyer_texte(row.get("Statut"))),
            ("Forme juridique", nettoyer_texte(row.get("Forme Juridique"))),
        ]
        lignes_desc = [f"{libelle} : {val}" for libelle, val in champs_desc if val]
        desc.text = "\n".join(lignes_desc)

        lat = row.get("lat")
        lon = row.get("lon")
        if pd.notna(lat) and pd.notna(lon):
            point = ET.SubElement(placemark, "Point")
            coords = ET.SubElement(point, "coordinates")
            coords.text = f"{lon},{lat},0"

    return ET.tostring(kml, encoding="utf-8", xml_declaration=True)


fichier_upload = st.file_uploader("📂 Importer ton fichier Excel (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    try:
        xls = pd.ExcelFile(fichier_upload)
        feuille = st.selectbox("Feuille Excel", options=xls.sheet_names, index=0)
        df_raw = pd.read_excel(xls, sheet_name=feuille)
    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Excel : {e}")
        st.stop()

    df = normaliser_dataframe(df_raw)
    st.success(f"✅ {len(df)} entreprises chargées depuis **{fichier_upload.name}**.")

    # --- FILTRES ---
    col_f1, col_f2, col_f3 = st.columns(3)
    df_filtre = df.copy()

    with col_f1:
        eff_nums = df_filtre["Tranche Effectif"].apply(extraire_effectif_num)
        max_eff = int(eff_nums.max()) if eff_nums.notna().any() else 0

        if max_eff > 0:
            plage_eff = st.slider(
                "Effectif max estimé (0 au max)",
                min_value=0,
                max_value=max_eff,
                value=(0, max_eff),
                step=1,
            )
            inclure_sans_eff = st.checkbox(
                "Inclure les lignes sans effectif renseigné", value=True
            )
            masque_eff = eff_nums.between(plage_eff[0], plage_eff[1], inclusive="both")
            if inclure_sans_eff:
                masque_eff = masque_eff | eff_nums.isna()
            else:
                masque_eff = masque_eff.fillna(False)
            df_filtre = df_filtre[masque_eff]
        else:
            st.caption("Aucune valeur numérique détectée dans 'Tranche Effectif'.")

    with col_f2:
        options_filtre_texte = sorted(
            set(
                df_filtre["Activité"].dropna().astype(str).tolist()
                + df_filtre["Catégorie"].dropna().astype(str).tolist()
            )
        )
        options_filtre_texte = [x for x in options_filtre_texte if x.strip()]
        sel_cat = st.multiselect(
            "Filtrer par activité / catégorie",
            options=options_filtre_texte,
            default=[],
        )
        if sel_cat:
            masque_cat = df_filtre["Activité"].isin(sel_cat) | df_filtre["Catégorie"].isin(sel_cat)
            df_filtre = df_filtre[masque_cat]

    with col_f3:
        recherche_mot = st.text_input(
            "Recherche mot-clé (raison sociale, adresse, SIREN...)",
            placeholder="Ex: MAZIERES, Boulangerie, PME...",
        )
        if recherche_mot.strip():
            motif = recherche_mot.strip().lower()
            masque_mot = df_filtre.apply(
                lambda r: motif in " ".join(r.astype(str)).lower(), axis=1
            )
            df_filtre = df_filtre[masque_mot]

    st.info(f"🎯 **{len(df_filtre)}** entreprise(s) sélectionnée(s) pour l'export.")
    st.dataframe(df_filtre, use_container_width=True)

    # --- GÉOCODAGE & EXPORT MY MAPS ---
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        geocoder = st.checkbox(
            "📍 Géocoder les adresses (coordonnées GPS pour carte & KML)", value=True
        )
    with col_opt2:
        nom_export = st.text_input(
            "Nom du calque / fichier My Maps (sans extension)",
            value="Prospection_Entreprises_MyMaps",
        )

    if st.button("🚀 Préparer les fichiers My Maps"):
        df_export = df_filtre.copy()

        if geocoder:
            lats, lons = [], []
            barre = st.progress(0)
            statut = st.empty()
            total = len(df_export)

            for idx, (_, row) in enumerate(df_export.iterrows()):
                adr = nettoyer_texte(row.get("Adresse"))
                nom_ent = nettoyer_texte(row.get("Nom Entreprise"))

                lat, lon = geocoder_adresse(adr)
                if (lat is None or lon is None) and nom_ent:
                    lat, lon = geocoder_adresse(f"{nom_ent} 37")

                lats.append(lat)
                lons.append(lon)

                if (idx + 1) % 10 == 0 or (idx + 1) == total:
                    barre.progress((idx + 1) / max(total, 1))
                    statut.text(f"Géocodage en cours... {idx + 1}/{total}")
                time.sleep(0.05)

            df_export["lat"] = lats
            df_export["lon"] = lons
            barre.empty()
            statut.empty()

            df_carte = df_export.dropna(subset=["lat", "lon"])
            st.success(f"🗺️ {len(df_carte)}/{len(df_export)} entreprises géocodées avec succès.")
            if not df_carte.empty:
                st.map(df_carte[["lat", "lon"]])

        # Colonne 'Name' explicitement nommée par la Raison Sociale pour l'import CSV My Maps
        df_mymaps = df_export.copy()
        df_mymaps.insert(0, "Name", df_mymaps["Nom Entreprise"])

        csv_bytes = df_mymaps.to_csv(index=False).encode("utf-8-sig")
        kml_bytes = generer_kml(df_export, nom_calque=nom_export)

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label=f"📥 Télécharger {nom_export}.csv (Import My Maps)",
                data=csv_bytes,
                file_name=f"{nom_export}.csv",
                mime="text/csv",
            )
        with dl2:
            st.download_button(
                label=f"📥 Télécharger {nom_export}.kml (Calque Google Earth / My Maps)",
                data=kml_bytes,
                file_name=f"{nom_export}.kml",
                mime="application/vnd.google-earth.kml+xml",
            )
