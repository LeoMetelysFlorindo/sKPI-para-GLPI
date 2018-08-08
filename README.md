# sKPI-para-GLPI
Um SKPI para o sistema GLPI
Aplicativo Web para apoio ao Sistema GLPI 08/08/2018 11:26h

FERRAMENTA DE DESENVOLVIMENTO: VISUAL STUDIO 2017 COMMUNITY LINGUAGEM DE PROGRAMAÇÃO: ASP.NET

DETALHES DESTE APLICATIVO: Este aplicativo foi desenvolvido para o GLPI da empresa Manuaus Ambiental em 26 de Abril de 2012. 
A ideia básica era ter uma ferramanta de apoio que permitisse o cálculo de Ociosidade e tempo de espera dos ticktes que
foram fechados pelas diversas áreas envolvidas. Os nomes internos de tecnicos foram fornecidos na própria aplicação e 
as colunas de AREAS foram preenchidas conforme a identificação de cada técnico. 
A aplicação realiza a leitura interna de dados do sistema GLPI realizando um simples select de dados e apresentando os mesmos
em uma tabela simples permitindo transferir estas informações da tabela gerada para o Excel, possibilitando assim um 
trabalho externo nos dados coletados. 

Para uma funcionalidade específica é possível acrescentar novos dados para cálculo às colunas aqui apresentadas. 

É importante lembrar que os dados da string de conexão devem ser adaptados para o GLPI que você utiliza e que a conexão remota
do MYSQL do GLPI deve estar ativa para permitir o uso desta aplicação.  Para isto você deve acessar o servidor MYSQL do GLPI e
conceder as permissões necessárias com o comando GRANT ON para permitir o acesso remoto ao banco de dados do GLPI.


Atenciosamente, 
Manoel Leonardo Metelis Florindo 
Especialista em Gestão de TI / Tecnólogo em Processamento de Dados 
