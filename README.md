# vba
Projeto em VBA criado para automatizar o envio de e-mails a partir de uma base de dados dividida em 4 abas do Excel [Exemplo com informações Vendas para as áreas]

# Estrutura e objetivo principal das rotinas:
     1. Adicionar Referência Outlook
             - Prepara/hablita o excel do usuário (máquina) para envio de e-mails pelo Microsoft Outlook
     2. Preparação do Arquivo
             - Identifica e armazena em variáveis as informações da pasta e do arquivo atual
     4. Início da rotina
             - Inicia a contagem do tempo para execução e identifica a região da base com os dados que serão utilizados
     5. Execução da rotina
             - Inicia a repetição do código (loop) para cada área/endereço de e-mail que receberá o e-mail
     6. Cria o arquivo temporário
             - Cria um arquivo excel temporário que será anexado ao e-mail, e parte da base que será incluída no corpo do e-mail
     7.  Cria um novo e-mail (outlook)
             - Abre uma instância do Microsoft Outlook e preenche com todas as informações dos passos anteriores
     9. Tratamentos de erro
             - Em caso de erro, finaliza a macro e exibe a mensagem do que foi parcialmente realizado (e-mails já enviados, por exemplo)
     10. Volta/finaliza os ajustes do excel
             - Antes de encerrar a macro, volta os padrões de comportamento padrão do Microsoft Excel
     11. Exibe a mensagem final sobre a execução de toda a rotina
             - Prepara e exibe uma mensagem amigável com a contagem de e-mails enviados e o tempo total para execução da macro 
