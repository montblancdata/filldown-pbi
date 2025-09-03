let
    Source = Excel.Workbook(File.Contents("C:\MontBlancData\DATA_FILLDOWN_XLSX.xlsx"), null, true),
    Feuil1_Sheet = Source{[Item="Feuil1",Kind="Sheet"]}[Data],
    #"Type modifié" = Table.TransformColumnTypes(Feuil1_Sheet,{{"Column1", type text}, {"Column2", type text}}),
    #"En-têtes promus" = Table.PromoteHeaders(#"Type modifié", [PromoteAllScalars=true]),
    AddColumn_CopieActivite = Table.AddColumn(#"En-têtes promus", "Copie de Activité", each [Activité]),
    TableFillDown = Table.FillDown(AddColumn_CopieActivite, {"Copie de Activité"}),
    #"Colonnes renommées" = Table.RenameColumns(TableFillDown,{{"Copie de Activité", "Activité après FillDown"}})
in
    #"Colonnes renommées"