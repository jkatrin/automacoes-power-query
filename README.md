# Roteiro: Transformação de Dados de Intervalo (1 ao 10) em uma lista (1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

## Passo a Passo no Microsoft Excel - Malha Rodoviaria Atualizada

Certificar que não tem filtro, selecionar a planilha inteira, todas as linhas e colunas (de Código a Lote).

![selecao-planilha](https://github.com/user-attachments/assets/a07a1577-8ce3-4345-abc6-53b65dade09b)

Na aba "Inserir" dentro do campo "Tabelas" clique em "Tabela" e "Ok".  
Isso vai permitir você usar o POWER QUERY, sua função principal é extrair, transformar e carregar dados (ETL) de maneira automatizada e sem a necessidade de programação, facilitando a análise e o uso dos dados no Excel, Power BI e outros produtos Microsoft.  
Na aba "Dados", no campo "Obter e Transformar Dados" clique em "Da Tabela/Intervalo"  
Irá abrir o Power Query  

## Passo a Passo no Editor Power Query

Verificar se na coluna Km inicial e Km final: estão em números decimais

![ndecimal](https://github.com/user-attachments/assets/6837b740-f536-46ba-9a85-1bced2c17265)

Com o Crtl pressionado clique nas colunas "Código", "Km inicial" e "Km final"  
Clique com o botão direito do mouse, e em seguida "Remover outras colunas"  
Verificar se a coluna Rodovia está com todas as informações padronizadas em apenas um tipo (para poder ordenar): SP 050, SP_050, SP-050 ou SP050  
Adicionar coluna personalizada com nome KMS:  
Colar a fórmula do Power Query (Código M):  

![coluna-personalizada](https://github.com/user-attachments/assets/ab6c71ea-eaed-4daf-a9bf-f28db8643a23)

(Confere se o nome Km Inicial, Final, e Código estão escritos igual dentro da fórmula, se atentar a letras maiúsculas ou minúsculas também)  
Expandir para nova lista  

![lista](https://github.com/user-attachments/assets/bb313d74-9fcf-4c96-9915-56d8194a4dbb)

Extrair antes do Delimitador: Coloque no campo "," e clique em "Ok"  

![delimit](https://github.com/user-attachments/assets/a39cb475-b63e-471f-8596-8b05a3dd58f7)

Duplicar a coluna Rodovia e KMS: clique com o botão direito na coluna "Código" e em seguida em "Duplicar Coluna", faça o mesmo para a coluna "Kms"  
Com duas colunas com "Copiar" no final (a de código e a de kms) com o Crtl pressionado clique nessas duas colunas e depois com o botão direito, em "Mesclar colunas" e "Ok"  

![mescla](https://github.com/user-attachments/assets/3fd6e438-c950-4bea-a75c-bd50a08826d5)

Colunas mescladas, clique com o botão direito e em "Remover Duplicadas", isso vai garantir que a rodovia não tenha mais de uma linha com o mesmo km  
Excluir a coluna mesclada, a km inicial e km final, vai ficar apenas colunas de "Código" e "Kms" (Por questão da máquina suportar o processo de carregamento)  
Subir para o excel ("Fechar e Carregar")  

Código M (Power Query)  
Veja o arquivo correspondente: [intervalo-em-lista.pq]  


## Passo a Passo no Microsoft Excel - Verificador de Malha

Abrir planilha de "Verificador_Malha_Rev--"  
(A planilha se mantém protegida para não haver alterações indevidas de malha rodoviária e da fórmula)  
Clicar na aba "Revisão" e dentro do campo de "Alterações" clicar em "Desproteger Planilha".  
Deixarei a senha como 1234 (Após isso, conseguirá visualizar a fórmula, e editá-la juntamente com a malha rodoviaria em caso de atualizações)  
Clique com o botão direito na aba "VF_DR01" e reexiba a aba "banco-de-dados"  
Pegamos nossa ETL da malha nova (é a edição avançada que fizemos da malha, depois da edição subimos para ao Excel de volta, copiar e cola valores nesse banco de dados)  
Aparecerá assim, copiamos  

![etl](https://github.com/user-attachments/assets/ff23266b-3c09-451a-8cfc-a8ecdefa051a)

E colamos aqui,  

![db](https://github.com/user-attachments/assets/5a9afa5c-d8bf-453e-a075-5743737b1ec2)

Prontinho, quase lá! Agora vamos apenas ajustar a seleção da fórmula, porque ser que tenha aumentado o Km e consequentemente as linhas  

![ajuste](https://github.com/user-attachments/assets/e775cf0a-5536-4989-925a-cc9c9ca1d914)

A coluna1 com o 1 é apenas a somatória da formula, vá até o final dela e a deixe ao lado da ultima linha da rodovia e km (se a malha aumentou, a rodovia e km vai estar bem abaixo desta coluna 1) você arrasta para preencher até o final  

![ajuste1](https://github.com/user-attachments/assets/e2ebea25-f0e2-481a-b876-a5df0762d0bd)

Agora, na aba "VF_DR01" vamos ajustar a seleção da fórmula  
No campo em vermelho, coloque o número da última linha do banco-de-dados para que a fórmula puxe até o final  
Antes, ia até essa linha  

![ajuste4](https://github.com/user-attachments/assets/487a582b-6f0e-4792-8e32-62836c5a0c38)

Agora vai até essa linha  

![ajuste5](https://github.com/user-attachments/assets/d151080b-a025-42a0-a0c1-b2697534683a)

## Passo a Passo Proteger Planilha de Edições  
Coloque o filtro na planilha, porque após travar não será possível  
Selecione as células da fórmula (abaixo da coluna "Verificador)  
Na aba "Revisão", no campo de "Alterações" clique em "Proteger Planilha", certifique-se que esteja no check: Selecionar células bloqueadas, selecionar células desbloqueadas, classificar, usar autofiltro.  
**Oculte a aba "banco-de-dados"  
**Faça o mesmo processo para as demais regionais e MALHA ATUALIZADA!!  
