# Roteiro: Transformação de Dados de Intervalo (1 ao 10) em uma lista (1, 2, 3, 4, 5, 6, 7, 8, 9, 10) 

## Objetivo
Tenho uma malha rodoviaria, e nela tenho os dados: | Rodovia | Km Inicial | Km Final |
Quero transformar essa única linha com o intervalo, em uma lista seperando cada Km por linha para fazer uma planilha de automização no excel.

## Passo a Passo no Power Query
1. **Verificar se na coluna Km inicial e Km final: estão em números decimais
2. **Verificar se a coluna Rodovia está com todas as informações padronizadas em apenas um tipo (para poder ordenar): SP 050, SP_050, SP-050 ou SP050
3. **Adicionar coluna personalizada com nome KMS: **[`intervalo-em-lista.pq`]
4. **Expandir para nova lista
5. **Extrair antes do Delimitador:","
6. **Duplicar a coluna Rodovia e KMS
7. **Mesclar as duas colunas e remover duplicados
8. **Excluir a coluna mesclada
9. **Subir para o excel

## Código M (Power Query)
Veja o arquivo correspondente: [`intervalo-em-lista.pq`]
