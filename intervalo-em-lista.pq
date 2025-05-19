let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela1"]}[Content],
    #"Personalização Adicionada" = Table.AddColumn(Fonte, "KMS", each if [km final] >= [km inicial] then
List.Numbers([km inicial], [km final] - [km inicial] + 1, 1)
else
List.Numbers([km inicial], [km inicial] - [km final] + 1, -1)),
    #"KMS Expandido" = Table.ExpandListColumn(#"Personalização Adicionada", "KMS"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"KMS Expandido",{"Codigo", "Tipo", "Municipio", "Res. Cons.", "km inicial", "km final", "KMS", "Extensão", "Jurisdição", "Administração", "Conservação", "Superficie", "Observação"}),
    #"Texto Extraído Antes do Delimitador" = Table.TransformColumns(#"Colunas Reordenadas", {{"KMS", each Text.BeforeDelimiter(Text.From(_, "pt-BR"), ","), type text}}),
    #"Coluna Duplicada" = Table.DuplicateColumn(#"Texto Extraído Antes do Delimitador", "KMS", "KMS - Copiar"),
    #"Coluna Duplicada1" = Table.DuplicateColumn(#"Coluna Duplicada", "Codigo", "Codigo - Copiar"),
    #"Colunas Mescladas" = Table.CombineColumns(#"Coluna Duplicada1",{"Codigo - Copiar", "KMS - Copiar"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Mesclado"),
    #"Duplicatas Removidas" = Table.Distinct(#"Colunas Mescladas", {"Mesclado"}),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Duplicatas Removidas",{"KMS", "Codigo"}),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Outras Colunas Removidas",{"Codigo", "KMS"})
in
    #"Colunas Reordenadas1"
