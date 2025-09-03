# Gestion des cellules fusionnées Excel dans Power BI

## 📌 Contexte  
Lorsqu’un fichier Excel contient des **cellules fusionnées** (souvent dans des plannings ou calendriers), il est lisible pour un humain, mais problématique pour Power BI :  
- Les valeurs fusionnées apparaissent une seule fois.  
- Après import, toutes les lignes en dessous se retrouvent avec des **valeurs nulles**.  
- Résultat : une perte d’information qui empêche toute analyse fiable.  

Exemple utilisé : calendrier d’activités d’une station de montagne (exploitation hivernale, fermeture, travaux lourds, maintenance annuelle, ouverture).  

---

## 🎯 Objectif  
Automatiser dans Power Query l’équivalent du **“tirer vers le bas” d’Excel** afin de répliquer correctement les valeurs manquantes.  

---

## 🛠️ Solution : `Table.FillDown`  
Le script Power Query applique la fonction **FillDown** sur la colonne concernée afin de remplir automatiquement les cellules vides avec la dernière valeur connue.  

### Extrait du script  
```m
  (...)
    AddColumn_CopieActivite = Table.AddColumn(#"En-têtes promus", "Copie de Activité", each [Activité]),
    TableFillDown = Table.FillDown(AddColumn_CopieActivite, {"Copie de Activité"}),
    #"Colonnes renommées" = Table.RenameColumns(TableFillDown,{{"Copie de Activité", "Activité après FillDown"}})
  (...)
```

---

## 📂 Contenu du dépôt
- **DATA_FILLDOWN_XLSX.xlsx** : fichier source avec cellules fusionnées (exemple calendrier).  
- **FillDown.pq** : script Power Query pour appliquer la transformation.  
- **README.md** : documentation et guide d’utilisation.  

---

## 🚀 Résultat attendu  
- Chaque ligne du tableau retrouve son activité.  
- Les données sont propres, structurées et directement exploitables dans Power BI.  

❌ Import brut : valeurs nulles (perte d’information)  
✅ Avec *FillDown* : réplication automatique des valeurs manquantes  

---

## 💡 Astuce  
Cette approche n’est pas limitée aux calendriers : elle s’applique à tout fichier Excel comportant des cellules fusionnées (rapports RH, plannings, reporting budgétaire…).  
