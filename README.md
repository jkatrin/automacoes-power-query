# Roteiro: Transformação de Dados de Intervalo (1 ao 10) em uma lista (1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

## Passo a Passo no Microsoft Excel - Malha Rodoviaria Atualizada

**Certificar que não tem filtro, selecionar a planilha inteira, todas as linhas e colunas (de Código a Lote).
![selecao-planilha](https://github.com/user-attachments/assets/a07a1577-8ce3-4345-abc6-53b65dade09b)


## Passo a Passo no Microsoft Excel - Verificador de Malha

**Abrir planilha de "Verificador_Malha_Rev--"
(A planilha se mantém protegida para não haver alterações indevidas de malha rodoviária e da fórmula)
**Clicar na aba "Revisão" e dentro do campo de "Alterações" clicar em "Desproteger Planilha".
**Deixarei a senha como 1234 (Após isso, conseguirá visualizar a fórmula, e editá-la juntamente com a malha rodoviaria em caso de atualizações)

**
** 


Passo a Passo no Power Query
**Verificar se na coluna Km inicial e Km final: estão em números decimais
**Verificar se a coluna Rodovia está com todas as informações padronizadas em apenas um tipo (para poder ordenar): SP 050, SP_050, SP-050 ou SP050
**Adicionar coluna personalizada com nome KMS: **[intervalo-em-lista.pq]
**Expandir para nova lista
**Extrair antes do Delimitador:","
**Duplicar a coluna Rodovia e KMS
**Mesclar as duas colunas e remover duplicados
**Excluir a coluna mesclada
**Subir para o excel
Código M (Power Query)
Veja o arquivo correspondente: [intervalo-em-lista.pq]
