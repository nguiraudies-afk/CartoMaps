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
    "Importe un fichier Excel de prospection (comme `Prospection La Riche.xlsx`), filtre tes cibles "
    "et génère un fichier prêt à importer dans **Google My Maps**."
)


def geocoder_adresse(adresse: str):
  """Récupère (lat, lon) via l'API Adresse officielle (data.gouv.fr)."""
  if not adresse or not str(adresse).strip():
    return None, None
  try:
    resp = requests.get(
        "https://api-adresse.data.gouv.fr/search/",
        params={"q": str(adresse).strip(), "limit": 1},
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


def obtenir_nom_entreprise(row) -> str:
  """Récupère le nom de l'entreprise en priorisant 'Raison sociale'."""
  for col in ("Raison sociale", "Dénomination"):
    val = row.get(col)
    if pd.notna(val):
      texte = str(val).strip()
      if texte and texte.lower() != "nan":
        return texte
  return "Entreprise"


def extraire_effectif_num(val):
  """Extrait un entier depuis la tranche d'effectif (gère nombres ou chaînes type '10 à 19')."""
  if pd.isna(val):
    return None
  if isinstance(val, (int, float)):
    return int(val)
  texte = str(val).replace("\u202f", "").replace(" ", "")
  # Cherche tous les nombres (ex: '10 à 19' -> prend le plus grand ou le premier)
  nombres = [int(n) for n in re.findall(r"\d+", texte)]
  if not nombres:
    return None
  return max(nombres)


def generer_kml(df_points: pd.DataFrame) -> bytes:
  """Génère un fichier KML compatible Google My Maps / Google Earth."""
  kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
  doc = ET.SubElement(kml, "Document")
  nom_doc = ET.SubElement(doc, "name")
  nom_doc.text = "Prospection Entreprises"

  for _, row in df_points.iterrows():
    nom_entreprise = obtenir_nom_entreprise(row)

    placemark = ET.SubElement(doc, "Placemark")
    titre = ET.SubElement(placemark, "name")
    titre.text = nom_entreprise

    desc = ET.SubElement(placemark, "description")
    details = [
        f"Raison sociale : {nom_entreprise}",
        f"Adresse : {row.get('Adresse_Complete', '')}",
        f"Effectif : {row.get('Tranche Effectif', 'Non renseigné')}",
        f"Catégorie : {row.get('Catégorie', '')}",
        f"SIREN : {row.get('SIREN', '')}",
        f"CA (€) : {row.get(\"Chiffre d'Affaires (€)\", '')}",
    ]
    desc.text = "\n".join([
        d
        for d in details
        if not d.endswith(": ") and not d.lower().endswith(": nan")
    ])

    lat = row.get("lat")
    lon = row.get("lon")
    if pd.notna(lat) and pd.notna(lon):
      point = ET.SubElement(placemark, "Point")
      coords = ET.SubElement(point, "coordinates")
      coords.text = f"{lon},{lat},0"

  return ET.tostring(kml, encoding="utf-8", xml_declaration=True)


fichier_upload = st.file_uploader(
    "📂 Importer ton fichier Excel (.xlsx)", type=["xlsx"]
)

if fichier_upload is not None:
  try:
    df = pd.read_excel(fichier_upload)
  except Exception as e:
    st.error(f"❌ Impossible de lire le fichier Excel : {e}")
    st.stop()

  # Construction d'une adresse complète propre si code postal / ville sont séparés
  def construire_adresse(row):
    adr = str(row.get("Adresse") or "").strip()
    cp = str(row.get("Code Postal") or "").replace(".0", "").strip()
    ville = str(row.get("Ville") or "").strip()
    if cp and cp not in adr:
      adr = f"{adr} {cp}".strip()
    if ville and ville.lower() not in adr.lower():
      adr = f"{adr} {ville}".strip()
    return adr

  df["Adresse_Complete"] = df.apply(construire_adresse, axis=1)

  st.success(
      f"✅ {len(df)} lignes chargées depuis **{fichier_upload.name}**."
  )

  # --- FILTRES ---
  col_f1, col_f2, col_f3 = st.columns(3)
  df_filtre = df.copy()

  with col_f1:
    if "Tranche Effectif" in df_filtre.columns:
      eff_nums = df_filtre["Tranche Effectif"].apply(extraire_effectif_num)
      max_eff = int(eff_nums.max()) if eff_nums.notna().any() else 0

      if max_eff > 0:
        plage_eff = st.slider(
            "Effectif (de 0 au max)",
            min_value=0,
            max_value=max_eff,
            value=(0, max_eff),
            step=1,
        )
        inclure_sans_eff = st.checkbox(
            "Inclure les lignes sans effectif renseigné", value=True
        )

        masque_eff = eff_nums.between(
            plage_eff[0], plage_eff[1], inclusive="both"
        )
        if inclure_sans_eff:
          masque_eff = masque_eff | eff_nums.isna()
        else:
          masque_eff = masque_eff.fillna(False)

        df_filtre = df_filtre[masque_eff]
      else:
        st.caption("Aucune valeur numérique détectée dans 'Tranche Effectif'.")

  with col_f2:
    if "Catégorie" in df_filtre.columns:
      categories = sorted(
          df_filtre["Catégorie"].dropna().astype(str).unique().tolist()
      )
      sel_cat = st.multiselect(
          "Filtrer par catégorie", options=categories, default=[]
      )
      if sel_cat:
        df_filtre = df_filtre[df_filtre["Catégorie"].astype(str).isin(sel_cat)]

  with col_f3:
    recherche_mot = st.text_input(
        "Recherche mot-clé (nom, rue, SIREN...)",
        placeholder="Ex: ZAC, Rue, PME...",
    )
    if recherche_mot.strip():
      motif = recherche_mot.strip().lower()
      masque = df_filtre.apply(
          lambda r: motif in " ".join(r.astype(str)).lower(), axis=1
      )
      df_filtre = df_filtre[masque]

  st.info(
      f"🎯 **{len(df_filtre)}** entreprise(s) sélectionnée(s) pour l'export."
  )
  st.dataframe(df_filtre, use_container_width=True)

  # --- GÉOCODAGE & EXPORT MY MAPS ---
  col_opt1, col_opt2 = st.columns(2)
  with col_opt1:
    geocoder = st.checkbox(
        "📍 Géocoder les adresses (coordonnées GPS pour carte & KML)",
        value=True,
    )
  with col_opt2:
    nom_export = st.text_input(
        "Nom du dossier / fichier My Maps (sans extension)",
        value="Prospection_MyMaps",
    )

  if st.button("🚀 Préparer les fichiers My Maps"):
    df_export = df_filtre.copy()

    if geocoder:
      lats, lons = [], []
      barre = st.progress(0)
      statut = st.empty()
      total = len(df_export)

      for idx, adr in enumerate(df_export["Adresse_Complete"]):
        lat, lon = geocoder_adresse(adr)
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
      st.success(
          f"🗺️ {len(df_carte)}/{len(df_export)} adresses géocodées avec"
          " succès."
      )
      if not df_carte.empty:
        st.map(df_carte[["lat", "lon"]])

    # Colonnes propres pour Google My Maps
    colonnes_mymaps = [
        c
        for c in [
            "Raison sociale",
            "Dénomination",
            "Adresse_Complete",
            "Tranche Effectif",
            "Catégorie",
            "SIREN",
            "Chiffre d'Affaires (€)",
            "lat",
            "lon",
        ]
        if c in df_export.columns
    ]
    df_mymaps = df_export[colonnes_mymaps]

    csv_bytes = df_mymaps.to_csv(index=False).encode("utf-8-sig")
    kml_bytes = generer_kml(df_export)

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
