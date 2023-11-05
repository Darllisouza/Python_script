## Arquivo Retorno 

Este código Python usa a biblioteca Boto3 para se conectar ao serviço Amazon S3 da AWS. Ele busca arquivos em um bucket S3 dentro de pastas específicas e verifica a data de modificação dos arquivos. Se a data for posterior a uma data definida, o código faz o download do arquivo, renomeando-o para evitar conflitos de nome. Ele lida com a possibilidade de mais objetos no bucket usando um token de continuação e exibe uma mensagem quando o download é concluído. 


## Ficha de Compensação Bancária

Conecta-se a um banco de dados PostgreSQL, executa consultas SQL para extrair dados de transações de duas tabelas e os armazena em DataFrames do Pandas. Em seguida, exporta esses DataFrames para arquivos Excel. O código configura o envio de e-mails via Microsoft Outlook e cria e envia e-mails com as planilhas anexadas, incluindo mensagens personalizadas.
