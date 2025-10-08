"""
Módulo de automação SAP responsável por atualizar datas de itens de pedidos
na transação ME22N utilizando SAP GUI Scripting.

Principais responsabilidades:
- Ler dados de um arquivo Excel (pedido, linha e nova data)
- Conectar à sessão do SAP GUI já aberta e logada
- Navegar para os itens informados e aplicar as alterações de data
- Registrar logs de execução e fornecer feedback para a interface gráfica
"""

import win32com.client
import pandas as pd
import time
import pythoncom
import os
from datetime import datetime

# Caminhos padrão para entrada e logs
ARQUIVO_PADRAO = r"\\br03file\pcoudir\Operacoes\10. Planning Raw Material\Gerenciamento de materiais\Atividades diarias\Robo Atualizacao de Datas Fornecedores\Alterar_pedidos.xlsx"
LOG_PASTA = r"\\br03file\pcoudir\Operacoes\10. Planning Raw Material\Gerenciamento de materiais\Atividades diarias\Robo Atualizacao de Datas Fornecedores\Log"

# Variáveis globais para comunicação com a interface gráfica
progress_callback = None
status_callback = None
log_callback = None
arquivo_excel_customizado = None

# Reservado para futura implementação de cancelamento cooperativo
_cancelamento_solicitado = False

def set_callbacks(progress_cb=None, status_cb=None, log_cb=None):
    """Registra callbacks para comunicação com a interface gráfica.

    Parâmetros:
        progress_cb (Callable[[int], None] | None): Callback para progresso (0-100).
        status_cb (Callable[[str, str], None] | None): Callback para status (texto, tipo).
        log_cb (Callable[[str], None] | None): Callback para mensagens de log.
    """
    global progress_callback, status_callback, log_callback
    progress_callback = progress_cb
    status_callback = status_cb
    log_callback = log_cb

def set_arquivo_excel(caminho_arquivo):
    """Define um caminho de arquivo Excel alternativo ao padrão."""
    global arquivo_excel_customizado
    arquivo_excel_customizado = caminho_arquivo

def get_arquivo_excel():
    """Retorna o caminho do arquivo Excel a ser utilizado na execução."""
    return arquivo_excel_customizado if arquivo_excel_customizado else ARQUIVO_PADRAO

def emit_progress(value):
    """Emite atualização de progresso para a interface (0 a 100)."""
    if progress_callback:
        progress_callback(value)

def emit_status(message, status_type):
    """Emite atualização de status para a interface.

    Parâmetros:
        message (str): Texto descritivo do status.
        status_type (str): Tipo do status (ex.: running, success, warning, error).
    """
    if status_callback:
        status_callback(message, status_type)

def emit_log(message):
    """Emite mensagem de log para a interface e console padrão."""
    if log_callback:
        log_callback(message)
    print(message)

def conectar_sap():
    """Estabelece conexão com a sessão ativa do SAP GUI.

    Retorna:
        session (obj) | None: Objeto de sessão SAP em caso de sucesso; caso contrário, None.
    """
    emit_log("🔄 Inicializando conexão com SAP...")
    emit_status("Inicializando SAP", "running")

    pythoncom.CoInitialize()
    try:
        emit_log("🔄 Obtendo SAP GUI...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")

        emit_log("🔄 Conectando ao engine...")
        application = SapGuiAuto.GetScriptingEngine

        emit_log("🔄 Estabelecendo conexão...")
        connection = application.Children(0)

        emit_log("🔄 Inicializando sessão...")
        session = connection.Children(0)

        emit_log("✅ Conexão SAP estabelecida com sucesso!")
        emit_status("Conectado ao SAP", "success")
        return session
    except Exception as e:
        error_msg = f"❌ Não foi possível conectar ao SAP: {e}"
        emit_log(error_msg)
        emit_status("Erro na conexão SAP", "error")
        return None

def esperar_objeto(session, objeto_id, tentativas=10, intervalo=0.5):
    """Aguarda um objeto da interface do SAP ficar disponível.

    Parâmetros:
        session: Sessão SAP ativa.
        objeto_id (str): Caminho/ID do objeto no SAP GUI.
        tentativas (int): Número máximo de tentativas.
        intervalo (float): Intervalo, em segundos, entre tentativas.

    Retorna:
        O objeto encontrado.

    Lança:
        Exception: Se o objeto não for encontrado no tempo limite.
    """
    for tentativa in range(tentativas):
        try:
            return session.findById(objeto_id)
        except:
            if tentativa < tentativas - 1:
                emit_log(f"🔄 Aguardando objeto {objeto_id}... (tentativa {tentativa + 1}/{tentativas})")
            time.sleep(intervalo)
    raise Exception(f"Objeto {objeto_id} não encontrado após {tentativas*intervalo}s.")

def limpar_tela_sap(session):
    """Limpa diálogos/resíduos e retorna à tela principal da sessão SAP."""
    try:
        for i in range(5):
            try:
                session.findById("wnd[1]/tbar[0]/btn[12]").press()
            except:
                break

        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)

        try:
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
            time.sleep(0.5)
        except:
            pass

    except Exception as e:
        emit_log(f"⚠️ Erro ao limpar tela: {e}")

def verificar_erro_sap(session):
    """Verifica mensagens da barra de status do SAP e retorna um erro informativo, se houver.

    Retorna:
        str | None: Mensagem de erro/aviso de interesse; None se não houver.
    """
    try:
        status_bar = session.findById("wnd[0]/sbar")
        if status_bar:
            message_type = status_bar.MessageType
            message_text = status_bar.Text.strip()

            if message_type in ['E', 'A']:
                return f"Erro SAP: {message_text}"
            elif message_type == 'W':
                emit_log(f"⚠️ Aviso SAP: {message_text}")
            elif message_type == 'I':
                if "sem alteração" in message_text.lower() or "não foi feita" in message_text.lower():
                    return f"Informação SAP: {message_text}"
        return None
    except:
        return None

def formatar_data(valor):
    """Formata valores de data para o padrão dd.mm.aaaa esperado pelo SAP."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%d.%m.%Y")

    valor_str = str(valor).strip()
    if "." in valor_str:
        return valor_str
    try:
        data = pd.to_datetime(valor_str, dayfirst=True, errors="coerce")
        if pd.notna(data):
            return data.strftime("%d.%m.%Y")
    except:
        pass
    return valor_str

def alterar_data(session, pedido, linha, nova_data, max_tentativas=2):
    """Altera a data de entrega de um item de pedido na ME22N.

    Parâmetros:
        session: Sessão SAP ativa.
        pedido (str | int): Número do pedido.
        linha (str | int): Número da linha (item) do pedido.
        nova_data (str): Data no formato dd.mm.aaaa.
        max_tentativas (int): Tentativas de repetição em caso de falha.

    Retorna:
        Tuple[str, str]: (status, mensagem) onde status ∈ {"SUCESSO", "PULADO", "ERRO"}.
    """
    for tentativa in range(1, max_tentativas + 1):
        try:
            emit_log(f"🔄 Processando pedido {pedido}, linha {linha} (tentativa {tentativa}/{max_tentativas})")
            limpar_tela_sap(session)

            # Abrir transação ME22N
            emit_log(f"📋 Abrindo transação ME22N para pedido {pedido}")
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "me22n"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)

            erro = verificar_erro_sap(session)
            if erro:
                raise Exception(f"Erro ao abrir ME22N: {erro}")

            # Inserir número do pedido
            emit_log(f"🔍 Buscando pedido {pedido}")
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
            session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = str(pedido)
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1.5)

            erro = verificar_erro_sap(session)
            if erro:
                raise Exception(f"Pedido não encontrado: {erro}")

            # Navegar para a linha específica
            emit_log(f"📝 Navegando para linha {linha}")
            combo_id = ("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
                        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/"
                        "subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST")
            linha_int = int(float(linha))
            key_str = f"{linha_int // 10:4d}"

            combo = esperar_objeto(session, combo_id, tentativas=5)
            combo.setFocus()
            combo.key = key_str
            time.sleep(1)

            # Alterar a data
            emit_log(f"📅 Alterando data para {nova_data}")
            campo_data = (
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/"
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5/"
                "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1320/"
                "tblSAPLMEGUITC_1320/ctxtMEPO1320-EEIND[2,0]"
            )
            celula = esperar_objeto(session, campo_data, tentativas=5)
            valor_atual = celula.text.strip()

            if valor_atual == nova_data:
                msg = f"⚠️ Pedido {pedido}, linha {linha} já estava com a data {nova_data}"
                emit_log(msg)
                return "PULADO", msg

            celula.text = nova_data
            celula.caretPosition = 2
            time.sleep(0.5)

            if celula.text.strip() != nova_data:
                raise Exception(f"O campo de data não foi atualizado. Esperado: {nova_data}, Atual: {celula.text.strip()}")

            # Salvar alterações
            emit_log(f"💾 Salvando alterações do pedido {pedido}")
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1.5)

            # Confirmar salvamento (quando aplicável)
            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.5)
            except:
                pass
            try:
                session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                time.sleep(0.5)
            except:
                pass

            erro = verificar_erro_sap(session)
            if erro:
                raise Exception(f"Erro após salvar: {erro}")

            msg = f"✅ Pedido {pedido}, linha {linha} atualizado para {nova_data}"
            emit_log(msg)
            return "SUCESSO", msg

        except Exception as e:
            error_msg = f"❌ Erro na tentativa {tentativa}: {e}"
            emit_log(error_msg)
            limpar_tela_sap(session)

            if tentativa < max_tentativas:
                emit_log(f"🔄 Tentando novamente em 2 segundos...")
                time.sleep(2)
                continue
            else:
                final_msg = f"❌ Pedido {pedido}, linha {linha} falhou após {max_tentativas} tentativas: {e}"
                emit_log(final_msg)
                return "ERRO", final_msg

    return "ERRO", f"❌ Falha desconhecida no pedido {pedido}, linha {linha}"

def salvar_logs_csv(resultados, pasta_base):
    """Gera arquivo CSV consolidando os resultados do processamento.

    Parâmetros:
        resultados (list[dict]): Lista com o resultado por item.
        pasta_base (str): Diretório onde o CSV será salvo.

    Retorna:
        str: Caminho completo do arquivo CSV gerado.
    """
    emit_log("💾 Salvando logs...")
    log_df = pd.DataFrame(resultados)
    if not os.path.exists(pasta_base):
        os.makedirs(pasta_base)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_base = os.path.join(pasta_base, f"log_alteracoes_{timestamp}.csv")
    log_df.to_csv(arquivo_base, index=False, sep=";", encoding="utf-8-sig")
    emit_log(f"📝 Log salvo em: {arquivo_base}")
    return arquivo_base

def main():
    """Ponto de entrada do processo em modo console."""
    # Obter o arquivo Excel a ser usado
    arquivo_atual = get_arquivo_excel()

    emit_log(f"🤖 SAP Robot iniciado")
    emit_log(f"📁 Arquivo: {arquivo_atual}")
    emit_status("Carregando dados", "running")
    emit_progress(5)

    # Verificar se arquivo existe
    if not os.path.exists(arquivo_atual):
        error_msg = f"❌ Arquivo {arquivo_atual} não encontrado."
        emit_log(error_msg)
        emit_status("Arquivo não encontrado", "error")
        return

    # Carregar Excel
    try:
        emit_log("📊 Carregando dados do Excel...")
        df = pd.read_excel(arquivo_atual, sheet_name=0)
        emit_log(f"✅ {len(df)} registros carregados do Excel")
        emit_progress(10)
    except Exception as e:
        error_msg = f"❌ Erro ao ler Excel: {e}"
        emit_log(error_msg)
        emit_status("Erro no Excel", "error")
        return

    # Conectar SAP
    emit_progress(15)
    session = conectar_sap()
    if not session:
        emit_log("❌ Não foi possível conectar ao SAP. Encerrando.")
        emit_status("Falha na conexão", "error")
        return

    emit_progress(25)
    emit_status("Processando pedidos", "running")

    resultados = []
    total_registros = len(df)

    for idx, row in df.iterrows():
        # Progresso proporcional (25% a 90% durante o processamento)
        progress_atual = 25 + int((idx / total_registros) * 65)
        emit_progress(progress_atual)

        pedido = row["Pedido"]
        linha = int(row["Linha"])
        nova_data_raw = row["NovaData"]
        nova_data = formatar_data(nova_data_raw)

        emit_log(f"\n{'='*50}")
        emit_log(f"📋 Processando {idx+1}/{total_registros}: Pedido {pedido}, Linha {linha}")
        emit_status(f"Processando pedido {pedido}", "running")

        status, mensagem = alterar_data(session, pedido, linha, nova_data)

        resultados.append({
            "Pedido": pedido,
            "Linha": linha,
            "Nova Data": nova_data,
            "Status": status,
            "Mensagem": mensagem,
            "Data Execução": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    # Finalizar processamento
    emit_progress(95)
    emit_status("Salvando logs", "running")

    # Salvar logs
    log_file = salvar_logs_csv(resultados, LOG_PASTA)

    # Contagem de resultados
    sucessos = sum(1 for r in resultados if r["Status"] == "SUCESSO")
    pulos = sum(1 for r in resultados if r["Status"] == "PULADO")
    erros = sum(1 for r in resultados if r["Status"] == "ERRO")

    # Emitir resumo
    emit_progress(100)
    emit_log(f"\n{'='*50}")
    emit_log("📋 Resumo da execução:")
    emit_log(f"✅ Sucessos: {sucessos}")
    emit_log(f"⚠️ Pulados: {pulos}")
    emit_log(f"❌ Erros: {erros}")
    emit_log(f"📝 Log salvo em: {log_file}")
    emit_log("🤖 Robô finalizado com sucesso!")

    if erros == 0:
        emit_status("Concluído com sucesso", "success")
    elif sucessos > 0:
        emit_status(f"Concluído com {erros} erros", "warning")
    else:
        emit_status("Concluído com erros", "error")

if __name__ == "__main__":
    main()
