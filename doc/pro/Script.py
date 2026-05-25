import pandas as pd
import os
import re
from datetime import datetime
import time

start_time = time.time()
print("🚀 Début du script")

# 📁 Détection du chemin OneDrive
base_path = os.path.expanduser("~")  # C:\Users\<Nom>
onedrive_path = os.path.join(base_path, "OneDrive - Alliance")

# 📁 Chemins dynamiques
dossier = os.path.join(onedrive_path, "700 - EQUIPE METHODES ET OUTILS", "750 - CONTRÔLES", "WORKDAY")
fichier_source = os.path.join(onedrive_path, "700 - EQUIPE METHODES ET OUTILS", "750 - CONTRÔLES", "753 - DATA & ANALYTICS", "Macro", "DOADemandes-202505141523.xlsx")

# ✅ Charger le fichier source
df_source = pd.read_excel(fichier_source)
print(f"✅ Fichier source chargé ({len(df_source)} lignes)")

# ✅ Indexation des fichiers Admin
admin_files = {}
for f in os.listdir(dossier):
    if "Admin-Renault Delegation Level" in f and f.endswith(".xlsx") and "Copie" not in f:
        match = re.search(r"\d{4}-\d{2}-\d{2}", f)
        if match:
            file_date = datetime.strptime(match.group(), "%Y-%m-%d")
            admin_files[file_date.strftime("%Y-%m")] = f

print(f"✅ {len(admin_files)} fichiers Admin indexés")

# ✅ Reprise automatique désactivée
resume_file = "resume.txt"
start_index = 0
print(f"🔄 Reprise à partir de la ligne {start_index + 1}")

# ✅ Traitement
results = []
last_month_year = None
df_admin_cache = None
ok_count, ko_count = 0, 0

for index, ligne in df_source.iloc[start_index:].iterrows():
    ipn = ligne.get("IPN Emetteur")
    date_emetteur = ligne.get("Date de validation Emetteur")
    date_redacteur = ligne.get("Date de validation Rédacteur")

    date_val = date_emetteur if pd.notna(date_emetteur) else date_redacteur
    if pd.isna(ipn) or pd.isna(date_val):
        results.append(["", ""])
        ko_count += 1
        continue

    date_val = pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    if pd.isna(date_val):
        results.append(["", ""])
        ko_count += 1
        continue

    mois_annee_val = date_val.strftime("%Y-%m")
    print(f"➡️ Ligne {index + 1}/{len(df_source)} | IPN: {ipn} | Mois: {mois_annee_val}")

    # ✅ Chargement Admin si changement de mois
    if mois_annee_val != last_month_year:
        admin_file = admin_files.get(mois_annee_val)
        if not admin_file and mois_annee_val == "2024-08":
            admin_file = admin_files.get("2024-07")

        if not admin_file:
            print(f"❌ Aucun fichier Admin pour {mois_annee_val}")
            results.append(["", ""])
            ko_count += 1
            continue

        admin_path = os.path.join(dossier, admin_file)
        try:
            df_admin_cache = pd.read_excel(admin_path, sheet_name="Sheet1", header=13)

            required_cols = ["Worker User Name ( IPN )", "Company", "Location Hierarchy"]
            if not all(col in df_admin_cache.columns for col in required_cols):
                print(f"⚠️ Colonnes manquantes dans {admin_file}, IPN ignoré")
                results.append(["", ""])
                ko_count += 1
                continue

            last_month_year = mois_annee_val
            print(f"✅ Fichier Admin chargé : {admin_file} ({len(df_admin_cache)} lignes)")
        except Exception as e:
            print(f"❌ Erreur chargement {admin_file}: {e}")
            results.append(["", ""])
            ko_count += 1
            continue

    # ✅ Recherche IPN
    match = df_admin_cache[df_admin_cache["Worker User Name ( IPN )"] == ipn]
    if match.empty:
        print("❌ IPN non trouvé dans Admin")
        results.append(["", ""])
        ko_count += 1
    else:
        company = str(match.iloc[0]["Company"]).replace(".", "").strip().upper()
        location = match.iloc[0]["Location Hierarchy"]
        print(f"✅ Company: {company} | Location: {location}")
        results.append([company, location])
        ok_count += 1

    # ✅ Sauvegarde de la position (optionnel)
    with open(resume_file, "w") as f:
        f.write(str(index))

# Supprimer les colonnes si elles existent déjà
for col in ["Company", "Location Hierarchy"]:
    if col in df_source.columns:
        df_source.drop(columns=[col], inplace=True)

# Insérer en K et L
df_source.insert(10, "Company", [r[0] for r in results])
df_source.insert(11, "Location Hierarchy", [r[1] for r in results])

# ✅ Sauvegarde finale
df_source.to_excel(fichier_source, index=False)
elapsed = round(time.time() - start_time, 2)
print(f"✅ Terminé : {ok_count} OK, {ko_count} KO | Temps : {elapsed} sec")