import sys
import os
import time
import getpass
import requests
import json
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QProgressBar, QLabel, QPushButton, QMessageBox, 
    QSizePolicy, QFrame, QGraphicsDropShadowEffect, QSpacerItem,
    QFileDialog, QLineEdit
)
from PyQt6.QtGui import QFont, QIcon, QCursor, QPixmap, QPainter, QColor, QPalette
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QRect

# Módulo de automação SAP (operações de backend)
import Sap

USUARIO = getpass.getuser().upper()

# Paleta base de cores da interface (aderente à identidade visual)
BG_COLOR = '#82298c'
FG_COLOR = "#ffffff"
ACCENT_COLOR = '#a94bb9'
LOG_BG_COLOR = '#682077'
BUTTON_COLOR = '#9436a6'
BUTTON_HOVER = '#b44fc6'

# Mapa de cores para uso consistente na UI
COLORS = {
    'primary': BUTTON_COLOR,
    'primary_hover': BUTTON_HOVER,
    'accent': ACCENT_COLOR,
    'background': BG_COLOR,
    'surface': '#732082',
    'surface_elevated': '#7d2b8a',
    'text': FG_COLOR,
    'text_secondary': '#e8d5ed',
    'text_muted': '#c9b3d1',
    'border': 'rgba(255, 255, 255, 0.1)',
    'success': '#4ade80',
    'warning': '#fb923c',
    'error': '#f87171'
}

# Diretório de logs da aplicação (contexto do usuário)
LOGS_DIR = os.path.join(os.path.expanduser("~"), "SAP_Robo_Logs")
if not os.path.exists(LOGS_DIR):
    os.makedirs(LOGS_DIR)

# Caminho do arquivo de configuração persistente
CONFIG_FILE = os.path.join(LOGS_DIR, "config.json")

def load_config():
    """Carrega configurações salvas"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_config(config):
    """Salva configurações"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar configuração: {e}")

# --- Componentes de UI ---
class CleanCard(QFrame):
    def __init__(self, elevated=False):
        super().__init__()
        self.setFrameStyle(QFrame.Shape.NoFrame)
        
        bg_color = COLORS['surface_elevated'] if elevated else COLORS['surface']
        self.setStyleSheet(f"""
            CleanCard {{
                background-color: {bg_color};
                border: 1px solid {COLORS['border']};
                border-radius: 16px;
                padding: 24px;
            }}
        """)
        
        if elevated:
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(25)
            shadow.setColor(QColor(0, 0, 0, 30))
            shadow.setOffset(0, 8)
            self.setGraphicsEffect(shadow)

class ModernButton(QPushButton):
    def __init__(self, text, style='primary', icon_only=False):
        super().__init__(text)
        self.button_style = style
        self.icon_only = icon_only
        self.setMinimumHeight(48)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._setup_style()
        
    def _setup_style(self):
        if self.button_style == 'primary':
            self.setStyleSheet(f"""
                ModernButton {{
                    background-color: {COLORS['primary']};
                    color: {COLORS['text']};
                    border: none;
                    border-radius: 12px;
                    font-weight: 600;
                    font-size: 15px;
                    padding: 14px 28px;
                }}
                ModernButton:hover {{
                    background-color: {COLORS['primary_hover']};
                    transform: translateY(-1px);
                }}
                ModernButton:pressed {{
                    transform: translateY(1px);
                }}
                ModernButton:disabled {{
                    background-color: rgba(255, 255, 255, 0.1);
                    color: {COLORS['text_muted']};
                }}
            """)
        elif self.button_style == 'secondary':
            self.setStyleSheet(f"""
                ModernButton {{
                    background-color: transparent;
                    color: {COLORS['text']};
                    border: 2px solid {COLORS['border']};
                    border-radius: 12px;
                    font-weight: 500;
                    font-size: 15px;
                    padding: 12px 26px;
                }}
                ModernButton:hover {{
                    background-color: rgba(255, 255, 255, 0.05);
                    border-color: {COLORS['accent']};
                }}
                ModernButton:disabled {{
                    border-color: rgba(255, 255, 255, 0.05);
                    color: {COLORS['text_muted']};
                }}
            """)
        elif self.button_style == 'ghost':
            self.setStyleSheet(f"""
                ModernButton {{
                    background-color: transparent;
                    color: {COLORS['text_secondary']};
                    border: none;
                    border-radius: 8px;
                    font-weight: 500;
                    font-size: 14px;
                    padding: 8px 16px;
                }}
                ModernButton:hover {{
                    background-color: rgba(255, 255, 255, 0.08);
                    color: {COLORS['text']};
                }}
            """)

class StatusDot(QLabel):
    def __init__(self):
        super().__init__()
        self.setFixedSize(12, 12)
        self.status = 'idle'
        self._update_style()
        
    def set_status(self, status):
        self.status = status
        self._update_style()
        
    def _update_style(self):
        colors = {
            'idle': '#6b7280',
            'running': COLORS['accent'],
            'success': COLORS['success'],
            'error': COLORS['error'],
            'warning': COLORS['warning']
        }
        
        color = colors.get(self.status, colors['idle'])
        self.setStyleSheet(f"""
            StatusDot {{
                background-color: {color};
                border-radius: 6px;
                border: 2px solid rgba(255, 255, 255, 0.2);
            }}
        """)

class CleanProgressBar(QProgressBar):
    def __init__(self):
        super().__init__()
        self.setTextVisible(False)
        self.setFixedHeight(6)
        self.setStyleSheet(f"""
            CleanProgressBar {{
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 3px;
                border: none;
            }}
            CleanProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 {COLORS['primary']}, stop:1 {COLORS['accent']});
                border-radius: 3px;
            }}
        """)

# --- WorkerThread Integrado ---
class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal(float)

    def __init__(self, arquivo_excel=None):
        super().__init__()
        self._running = True
        self.arquivo_excel = arquivo_excel
        timestamp = time.strftime("%d%m%Y_%H%M%S")
        self.log_path = os.path.join(LOGS_DIR, f"{USUARIO}_{timestamp}.log")

    def stop(self):
        self._running = False
        # Ponto de extensão: sinalizar parada no módulo SAP (se/quando implementado)
        Sap.sinalizar_parada()

    def run(self):
        start_time = time.time()
        self._write_log(f"=== Execução iniciada em {time.strftime('%d/%m/%Y %H:%M:%S')} ===")

        # Configurar callbacks para o módulo SAP
        Sap.set_callbacks(
            progress_cb=self.progress_signal.emit,
            status_cb=self.status_signal.emit,
            log_cb=self.log_signal.emit
        )

        # Definir o arquivo Excel a ser usado (se fornecido)
        if self.arquivo_excel:
            Sap.set_arquivo_excel(self.arquivo_excel)
            self._write_log(f"Arquivo selecionado: {self.arquivo_excel}")

        try:
            # Executar o processo SAP integrado
            Sap.main()
            
            if self._running:
                self._write_log("✓ Execução concluída com sucesso")
                self.status_signal.emit("Concluído com sucesso", "success")
                
        except Exception as e:
            error_msg = f"✗ Erro crítico: {e}"
            self._write_log(error_msg)
            self.log_signal.emit(error_msg)
            self.status_signal.emit("Erro crítico", "error")

        end_time = time.time()
        self.finished_signal.emit(end_time - start_time)

    def _write_log(self, texto):
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                timestamp = datetime.now().strftime("%H:%M:%S")
                f.write(f"{timestamp} {texto}\n")
        except Exception as e:
            print(f"Erro ao escrever log: {e}")

# --- Janela principal ---
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAP Robot")
        self.setMinimumSize(900, 700)
        self.resize(1000, 800)
        
        # Carregar configurações persistidas
        self.config = load_config()
        self.arquivo_selecionado = self.config.get('ultimo_arquivo', None)
        
        self._setup_window()
        self._setup_ui()
        
        # Restaura o último arquivo selecionado, quando aplicável
        if self.arquivo_selecionado and os.path.exists(self.arquivo_selecionado):
            self.file_path_input.setText(self.arquivo_selecionado)
            self.clear_file_btn.setVisible(True)
            self.adicionar_log(f"📂 Arquivo carregado das configurações: {os.path.basename(self.arquivo_selecionado)}")
        
        self.worker = None

    def _setup_window(self):
        # Ícone da aplicação
        url_icon = "https://i.ibb.co/m5LgjRfL/Robo.png"
        try:
            resposta_icon = requests.get(url_icon, timeout=3)
            if resposta_icon.status_code == 200:
                pixmap_icon = QPixmap()
                pixmap_icon.loadFromData(resposta_icon.content)
                self.setWindowIcon(QIcon(pixmap_icon))
        except:
            pass
        
        # Estilos globais da aplicação (tema)
        style_sheet = f"""
            * {{
                font-family: 'Inter', 'SF Pro Display', 'Segoe UI', system-ui, sans-serif;
            }}
            
            QWidget {{
                background-color: {COLORS['background']};
                color: {COLORS['text']};
            }}
            
            QTextEdit {{
                background-color: rgba(0, 0, 0, 0.2);
                color: {COLORS['text']};
                font-family: 'JetBrains Mono', 'Fira Code', 'Consolas', monospace;
                font-size: 13px;
                border: none;
                border-radius: 12px;
                padding: 16px;
                line-height: 1.6;
                selection-background-color: {COLORS['accent']};
            }}
            
            QLineEdit {{
                background-color: rgba(0, 0, 0, 0.2);
                color: {COLORS['text']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                padding: 12px;
                font-size: 13px;
            }}
            
            QLineEdit:focus {{
                border: 1px solid {COLORS['accent']};
            }}
            
            QLineEdit:disabled {{
                background-color: rgba(0, 0, 0, 0.1);
                color: {COLORS['text_muted']};
            }}
            
            QScrollBar:vertical {{
                background-color: transparent;
                width: 6px;
                border-radius: 3px;
            }}
            
            QScrollBar::handle:vertical {{
                background-color: rgba(255, 255, 255, 0.2);
                border-radius: 3px;
                min-height: 20px;
            }}
            
            QScrollBar::handle:vertical:hover {{
                background-color: rgba(255, 255, 255, 0.3);
            }}
            
            QScrollBar:horizontal {{
                background-color: transparent;
                height: 6px;
                border-radius: 3px;
            }}
            
            QScrollBar::handle:horizontal {{
                background-color: rgba(255, 255, 255, 0.2);
                border-radius: 3px;
                min-width: 20px;
            }}
            
            QScrollBar::handle:horizontal:hover {{
                background-color: rgba(255, 255, 255, 0.3);
            }}
        """
        self.setStyleSheet(style_sheet)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(32, 32, 32, 32)
        main_layout.setSpacing(24)

        # Cabeçalho
        header_layout = QHBoxLayout()
        
        # Logo
        logo_label = QLabel()
        self._setup_logo(logo_label)
        header_layout.addWidget(logo_label)
        
        header_layout.addSpacing(16)
        
        # Título e subtítulo
        title_container = QVBoxLayout()
        title = QLabel("SAP Robot")
        title.setFont(QFont("Inter", 28, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {COLORS['text']}; margin: 0;")
        
        subtitle = QLabel("Automação de alteração de Datas")
        subtitle.setFont(QFont("Inter", 14, QFont.Weight.Normal))
        subtitle.setStyleSheet(f"color: {COLORS['text_secondary']}; margin: 0;")
        
        title_container.addWidget(title)
        title_container.addWidget(subtitle)
        title_container.setSpacing(4)
        header_layout.addLayout(title_container)
        
        header_layout.addStretch()
        
        # Indicador de status atual
        status_layout = QHBoxLayout()
        self.status_dot = StatusDot()
        self.status_text = QLabel("Pronto")
        self.status_text.setFont(QFont("Inter", 13, QFont.Weight.Medium))
        self.status_text.setStyleSheet(f"color: {COLORS['text_secondary']};")
        
        status_layout.addWidget(self.status_dot)
        status_layout.addSpacing(8)
        status_layout.addWidget(self.status_text)
        status_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        header_layout.addLayout(status_layout)
        main_layout.addLayout(header_layout)

        # Espaçamento entre seções
        main_layout.addSpacing(8)

        # Seletor de arquivo de entrada (Excel)
        file_selector_layout = QHBoxLayout()
        file_selector_layout.setSpacing(12)
        
        file_icon = QLabel("📂")
        file_icon.setFont(QFont("Inter", 16))
        file_selector_layout.addWidget(file_icon)
        
        self.file_path_input = QLineEdit()
        self.file_path_input.setPlaceholderText("Arquivo padrão (clique em Procurar para alterar)")
        self.file_path_input.setReadOnly(True)
        self.file_path_input.setMaximumHeight(44)
        file_selector_layout.addWidget(self.file_path_input, stretch=1)
        
        self.browse_btn = ModernButton("Procurar", 'ghost')
        self.browse_btn.setMinimumWidth(100)
        self.browse_btn.setMaximumHeight(44)
        self.browse_btn.clicked.connect(self.selecionar_arquivo)
        file_selector_layout.addWidget(self.browse_btn)
        
        self.clear_file_btn = ModernButton("✕", 'ghost')
        self.clear_file_btn.setFixedSize(44, 44)
        self.clear_file_btn.setToolTip("Limpar seleção")
        self.clear_file_btn.clicked.connect(self.limpar_arquivo)
        self.clear_file_btn.setVisible(False)
        file_selector_layout.addWidget(self.clear_file_btn)
        
        main_layout.addLayout(file_selector_layout)
        main_layout.addSpacing(8)

        # Área de logs
        log_card = CleanCard(elevated=True)
        log_layout = QVBoxLayout(log_card)
        log_layout.setSpacing(16)
        
        # Cabeçalho do console de logs
        log_header = QHBoxLayout()
        log_title = QLabel("Console")
        log_title.setFont(QFont("Inter", 16, QFont.Weight.Medium))
        log_header.addWidget(log_title)
        log_header.addStretch()
        
        # Identificação do usuário
        user_label = QLabel(f"@{USUARIO.lower()}")
        user_label.setFont(QFont("Inter", 12))
        user_label.setStyleSheet(f"color: {COLORS['text_muted']};")
        log_header.addWidget(user_label)
        
        log_layout.addLayout(log_header)
        
        # Área do console de logs
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMinimumHeight(250)
        self.log_area.setPlainText("🤖 SAP Robot carregado e pronto para execução...")
        log_layout.addWidget(self.log_area)
        
        main_layout.addWidget(log_card, stretch=1)

        # Controles inferiores (progresso e ações)
        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(16)
        
        # Indicador de progresso
        progress_container = QVBoxLayout()
        progress_container.setSpacing(8)
        
        progress_info = QHBoxLayout()
        progress_label = QLabel("Progresso")
        progress_label.setFont(QFont("Inter", 12, QFont.Weight.Medium))
        progress_label.setStyleSheet(f"color: {COLORS['text_secondary']};")
        
        self.progress_value = QLabel("0%")
        self.progress_value.setFont(QFont("Inter", 12, QFont.Weight.Medium))
        self.progress_value.setStyleSheet(f"color: {COLORS['text_muted']};")
        
        progress_info.addWidget(progress_label)
        progress_info.addStretch()
        progress_info.addWidget(self.progress_value)
        
        self.progress_bar = CleanProgressBar()
        
        progress_container.addLayout(progress_info)
        progress_container.addWidget(self.progress_bar)
        
        controls_layout.addLayout(progress_container, stretch=1)
        
        # Botões de ação
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)
        
        self.cancel_btn = ModernButton("Cancelar", 'secondary')
        self.cancel_btn.clicked.connect(self.cancelar_execucao)
        self.cancel_btn.setEnabled(False)
        buttons_layout.addWidget(self.cancel_btn)
        
        self.start_btn = ModernButton("Executar", 'primary')
        self.start_btn.clicked.connect(self.iniciar_execucao)
        buttons_layout.addWidget(self.start_btn)
        
        controls_layout.addLayout(buttons_layout)
        main_layout.addLayout(controls_layout)

    def _setup_logo(self, logo_label):
        url_logo = "https://i.ibb.co/Zp4D8B90/neodent-logo.png"
        try:
            resp_logo = requests.get(url_logo, timeout=3)
            if resp_logo.status_code == 200:
                pix = QPixmap()
                pix.loadFromData(resp_logo.content)
                logo_label.setPixmap(pix.scaledToHeight(44, Qt.TransformationMode.SmoothTransformation))
            else:
                self._set_fallback_logo(logo_label)
        except:
            self._set_fallback_logo(logo_label)
    
    def _set_fallback_logo(self, logo_label):
        logo_label.setText("●")
        logo_label.setFont(QFont("Inter", 24))
        logo_label.setStyleSheet(f"color: {COLORS['accent']};")

    def selecionar_arquivo(self):
        """Abre o diálogo do sistema para selecionar o arquivo Excel de entrada."""
        diretorio_inicial = os.path.expanduser("~")
        if self.arquivo_selecionado and os.path.exists(self.arquivo_selecionado):
            diretorio_inicial = os.path.dirname(self.arquivo_selecionado)
        
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Arquivo Excel",
            diretorio_inicial,
            "Arquivos Excel (*.xlsx *.xls);;Todos os arquivos (*.*)"
        )
        
        if arquivo:
            self.arquivo_selecionado = arquivo
            self.file_path_input.setText(arquivo)
            self.clear_file_btn.setVisible(True)
            self.adicionar_log(f"📂 Arquivo selecionado: {os.path.basename(arquivo)}")
            
            self.config['ultimo_arquivo'] = arquivo
            save_config(self.config)
            
    def limpar_arquivo(self):
        """Limpa a seleção atual do arquivo, voltando ao arquivo padrão."""
        self.arquivo_selecionado = None
        self.file_path_input.clear()
        self.file_path_input.setPlaceholderText("Arquivo padrão (clique em Procurar para alterar)")
        self.clear_file_btn.setVisible(False)
        self.adicionar_log("🗑️ Seleção de arquivo limpa - usando arquivo padrão")
        
        if 'ultimo_arquivo' in self.config:
            del self.config['ultimo_arquivo']
            save_config(self.config)

    def iniciar_execucao(self):
        if self.worker and self.worker.isRunning():
            return
        
        if self.arquivo_selecionado:
            if not os.path.exists(self.arquivo_selecionado):
                QMessageBox.warning(
                    self,
                    "Arquivo não encontrado",
                    f"O arquivo selecionado não existe:\n{self.arquivo_selecionado}"
                )
                return
        
        self.log_area.clear()
        self.progress_bar.setValue(0)
        self.progress_value.setText("0%")
        self.status_dot.set_status('running')
        self.status_text.setText("Iniciando")
        
        # Inicializa e executa a thread de trabalho (processo SAP)
        self.worker = WorkerThread(arquivo_excel=self.arquivo_selecionado)
        self.worker.log_signal.connect(self.adicionar_log)
        self.worker.progress_signal.connect(self.atualizar_progresso)
        self.worker.status_signal.connect(self.atualizar_status)
        self.worker.finished_signal.connect(self.execucao_finalizada)
        self.worker.start()

        self.start_btn.setEnabled(False)
        self.start_btn.setText("Executando...")
        self.cancel_btn.setEnabled(True)
        self.browse_btn.setEnabled(False)
        self.clear_file_btn.setEnabled(False)

    def adicionar_log(self, texto):
        if texto.startswith("[") and "]" in texto[:10]:
            self.log_area.append(texto)
        else:
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_area.append(f"[{timestamp}] {texto}")
        
        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def atualizar_progresso(self, valor):
        self.progress_bar.setValue(valor)
        self.progress_value.setText(f"{valor}%")

    def atualizar_status(self, texto, tipo):
        self.status_dot.set_status(tipo)
        self.status_text.setText(texto)

    def cancelar_execucao(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.quit()
            self.worker.wait()
            
            self.status_dot.set_status('warning')
            self.status_text.setText("Cancelado")
            self.adicionar_log("⚠️ Execução cancelada pelo usuário")
            
            self.cancel_btn.setEnabled(False)
            self.start_btn.setEnabled(True)
            self.start_btn.setText("Executar")
            self.browse_btn.setEnabled(True)
            if self.arquivo_selecionado:
                self.clear_file_btn.setEnabled(True)
            
            self.progress_bar.setValue(0)
            self.progress_value.setText("0%")

    def execucao_finalizada(self, tempo_total):
        self.cancel_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.start_btn.setText("Executar")
        self.browse_btn.setEnabled(True)
        if self.arquivo_selecionado:
            self.clear_file_btn.setEnabled(True)
        
        tempo_min = tempo_total / 60
        status_atual = self.status_text.text().lower()
        
        msg = QMessageBox(self)
        msg.setWindowTitle("Execução Finalizada")
        
        if "sucesso" in status_atual or "concluído" in status_atual:
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText("🎉 SAP Robot finalizado com sucesso!")
        elif "erro" in status_atual:
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("⚠️ SAP Robot finalizado com erros")
        else:
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText("📋 SAP Robot finalizado")
        
        msg.setInformativeText(f"⏱️ Tempo total: {tempo_min:.1f} minutos\n📝 Logs salvos automaticamente")
        
        msg_style = f"""
            QMessageBox {{
                background-color: {COLORS['surface']};
                color: {COLORS['text']};
                font-family: 'Inter', sans-serif;
                border-radius: 12px;
            }}
            QMessageBox QLabel {{
                color: {COLORS['text']};
                font-size: 14px;
                padding: 8px;
            }}
            QMessageBox QPushButton {{
                background-color: {COLORS['primary']};
                color: {COLORS['text']};
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: 600;
                font-size: 14px;
                min-width: 80px;
            }}
            QMessageBox QPushButton:hover {{
                background-color: {COLORS['primary_hover']};
            }}
            QMessageBox QPushButton:pressed {{
                background-color: {COLORS['primary']};
            }}
        """
        msg.setStyleSheet(msg_style)
        msg.exec()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, 
                'Confirmar Fechamento',
                'SAP Robot ainda está executando.\n\nDeseja realmente fechar?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.worker.stop()
                self.worker.quit()
                self.worker.wait(3000)
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # Metadados da aplicação
    app.setApplicationName("SAP Robot")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Neodent")
    
    janela = MainWindow()
    janela.show()
    
    sys.exit(app.exec())