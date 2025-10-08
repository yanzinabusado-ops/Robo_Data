# SAP Robot

Automação para atualização de datas em pedidos via SAP GUI Scripting, com interface gráfica moderna em PyQt6.

## Visão Geral
- Leitura de dados a partir de uma planilha Excel (pedido, linha e nova data)
- Conexão com SAP GUI e navegação automática pela transação ME22N
- Atualização da data por item e salvamento com verificação de erros/avisos
- Interface gráfica para acompanhar progresso, status e logs em tempo real
- Geração de arquivo de log (CSV) com o resultado detalhado de cada item

## Pré-requisitos
- Windows (SAP GUI Scripting é suportado apenas no Windows)
- SAP GUI para Windows instalado e aberto, com Scripting habilitado
  - No cliente: SAP Logon > Options > Accessibility & Scripting > Scripting > Enable scripting
  - No servidor: o administrador SAP deve permitir scripting no perfil
- Python 3.10+ (64 bits recomendado)
- Acesso à planilha Excel com colunas obrigatórias: Pedido, Linha, NovaData

## Estrutura do Projeto
- RoboSAP_GUI.py: Interface gráfica (PyQt6)
- Sap.py: Lógica de automação do SAP (pandas + pywin32)
- Robo.ico/Robo.png: Ícones/imagens
- RoboSAP_GUI.spec: Especificação do PyInstaller (opcional para empacotamento)

## Instalação
1) Crie e ative um ambiente virtual (opcional, recomendado)
- CMD/PowerShell:
  - py -3 -m venv .venv
  - .venv\Scripts\activate

2) Instale as dependências
- pip install -r requirements.txt

## Execução
- Interface gráfica:
  - python RoboSAP_GUI.py
- Execução direta (modo console):
  - python Sap.py

Durante a execução pela interface, você pode selecionar um arquivo Excel personalizado ou usar o caminho padrão definido no código.

## Planilha Excel (formato esperado)
- Colunas obrigatórias: Pedido, Linha, NovaData
- NovaData aceita formatos comuns de data; o sistema converte para dd.mm.aaaa
- Exemplo (conceitual):
  - Pedido: 4500000001
  - Linha: 10
  - NovaData: 25/12/2025

## Caminhos e Logs
- Caminho padrão do Excel (em Sap.py):
  - Variável ARQUIVO_PADRAO aponta para o arquivo de trabalho padrão
- Logs de execução detalhados (CSV):
  - Gerados na pasta LOG_PASTA (definida em Sap.py)
- Logs da interface (arquivo .log por usuário/executável):
  - %USERPROFILE%\SAP_Robo_Logs

## Dicas de Uso
- Deixe o SAP GUI aberto e logado antes de iniciar o robô
- Evite usar o computador durante a automação para não interferir na sessão SAP
- Se um item já estiver com a data correta, ele será marcado como PULADO

## Empacotamento (opcional)
- Com PyInstaller instalado:
  - pyinstaller RoboSAP_GUI.spec
- Alternativa (sem .spec):
  - pyinstaller --noconsole --icon Robo.ico --name "SAP Robot" RoboSAP_GUI.py
- O executável ficará em dist/.

## Solução de Problemas
- Erro ao conectar ao SAP:
  - Verifique se o SAP GUI está aberto e logado
  - Confirme que o Scripting está habilitado (cliente e servidor)
  - Em algumas máquinas, é necessário que Python/robô e SAP GUI sejam da mesma arquitetura (64 bits)
- Erro ao ler Excel:
  - Garanta que o arquivo existe e não está protegido por senha
  - Instale o driver/engine do Excel (openpyxl já está no requirements)
- Campos não encontrados na tela:
  - Telas do SAP podem variar; garanta que a transação ME22N está acessível e a visualização é a padrão
- Bloqueios de TI/Antivírus:
  - Alguns agentes podem bloquear automações; peça exceção se necessário

## Segurança e Conformidade
- Não compartilhe credenciais SAP
- Execute o robô apenas em ambientes autorizados e com permissões adequadas
- Respeite as políticas internas de Scripting SAP e auditoria

## Manutenção
- Ajuste os seletores/IDs do SAP em Sap.py caso a tela mude
- Mantenha bibliotecas atualizadas com cuidado (testar antes em ambiente de homologação)

## Licença
- Uso interno. Ajuste conforme a política da sua organização.
