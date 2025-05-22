let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela3"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Codigo", type text}, {"Tipo", type text}, {"Municipio", type text}, {"Res. Cons.", Int64.Type}, {"km inicial", type number}, {"km final", type number}, {"Extensão", type number}, {"Jurisdição", type text}, {"Administração", type text}, {"Conservação", Int64.Type}, {"Superficie", type text}, {"Observação", type text}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "KMS", each 
    if [km final] >= [km inicial] then
        List.Numbers([km inicial], [km final] - [km inicial] + 1, 1)
    else
        List.Numbers([km inicial], [km inicial] - [km final] + 1, -1)),
    #"KMS Expandido" = Table.ExpandListColumn(#"Personalização Adicionada", "KMS"),
    #"Texto Extraído Antes do Delimitador" = Table.TransformColumns(#"KMS Expandido", {{"KMS", each Text.BeforeDelimiter(Text.From(_, "pt-BR"), ","), type text}}),
    #"Coluna Duplicada" = Table.DuplicateColumn(#"Texto Extraído Antes do Delimitador", "KMS", "KMS - Copiar"),
    #"Coluna Duplicada1" = Table.DuplicateColumn(#"Coluna Duplicada", "Codigo", "Codigo - Copiar"),
    #"Colunas Mescladas" = Table.CombineColumns(#"Coluna Duplicada1",{"KMS - Copiar", "Codigo - Copiar"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Mesclado"),
    #"Duplicatas Removidas" = Table.Distinct(#"Colunas Mescladas", {"Mesclado"}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Duplicatas Removidas",{"Codigo", "Tipo", "Municipio", "Res. Cons.", "km inicial", "km final", "Extensão", "Jurisdição", "Administração", "Conservação", "KMS", "Superficie", "Observação", "Mesclado"}),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Colunas Reordenadas",{"KMS", "Codigo"}),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Outras Colunas Removidas",{"Codigo", "KMS"})
in
    #"Colunas Reordenadas1"
