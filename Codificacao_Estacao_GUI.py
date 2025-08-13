import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QHBoxLayout, QLabel, QLineEdit, QPushButton,
                           QFileDialog, QTabWidget, QTableWidget, QTableWidgetItem,
                           QMessageBox, QProgressBar, QHeaderView, QComboBox, QFrame)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QIcon, QPainter, QColor, QBrush
from Codificacao_Estacao_Core import EstacaoManager

def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, funciona para dev e para PyInstaller"""
    try:
        # PyInstaller cria um diretório temporário e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Codificação de Estações")
        self.setMinimumSize(1200, 800)
        # Definir ícone personalizado
        icon_path = resource_path("assets/pluviometer.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        # Definir o estilo global incluindo o fundo cinza
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #919191;
            }
            QLabel {
                color: #333333;
                background-color: transparent;
            }
            QPushButton {
                background-color: #005a9e;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0078d4;
            }
            QComboBox, QComboBox QAbstractItemView, QComboBox::item {
                color: #222 !important;
                background: white;
            }
            QFrame#Card {
                background: white;
                border-radius: 14px;
                border: 1px solid #333232;
            }
        """)
        # Inicializa o gerenciador de estações
        self.estacao_manager = EstacaoManager()
        # Configurar a interface
        self.setup_ui()
        
    def setup_ui(self):
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        layout = QVBoxLayout(central_widget)
        
        # Cabeçalho com sombra
        header_widget = QFrame()
        header_widget.setObjectName("Card")
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 20, 20, 20)
        header_layout.setSpacing(20)
        
        # Logo RHN (esquerda)
        rhn_label = QLabel()
        rhn_path = resource_path("assets/rhn_logo.png")
        rhn_pixmap = QPixmap(rhn_path)
        if not rhn_pixmap.isNull():
            rhn_pixmap = rhn_pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            rhn_label.setPixmap(rhn_pixmap)
        rhn_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_layout.addWidget(rhn_label, 1)
        
        # Título (centro)
        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("Codificação de Estações Pluviométricas")
        title_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #003366; margin: 20px; background-color: transparent;")
        title_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(title_label)
        header_layout.addWidget(title_widget, 4)
        
        # Logo ANA (direita)
        ana_label = QLabel()
        ana_path = resource_path("assets/ana_logo.png")
        ana_pixmap = QPixmap(ana_path)
        if not ana_pixmap.isNull():
            ana_pixmap = ana_pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            ana_label.setPixmap(ana_pixmap)
        ana_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(ana_label, 1)
        
        header_widget.setFixedHeight(170)
        layout.addWidget(header_widget)
        
        # Área principal em cartão
        main_card = QFrame()
        main_card.setObjectName("Card")
        main_layout = QVBoxLayout(main_card)
        main_layout.setSpacing(18)
        
        # Seleção de arquivo
        file_widget = QWidget()
        file_layout = QHBoxLayout(file_widget)
        
        self.file_label = QLabel("Nenhum arquivo selecionado")
        file_layout.addWidget(self.file_label)
        
        btn_arquivo = QPushButton("Selecionar Arquivo")
        btn_arquivo.setIcon(QIcon.fromTheme("document-open"))
        btn_arquivo.clicked.connect(self.selecionar_arquivo)
        file_layout.addWidget(btn_arquivo)
        
        main_layout.addWidget(file_widget)
        
        # Área de seleção de formato
        format_widget = QWidget()
        format_layout = QHBoxLayout(format_widget)
        
        format_label = QLabel("Formato de Saída:")
        format_layout.addWidget(format_label)
        
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Excel (.xlsx)", "Access (.mdb)"])
        format_layout.addWidget(self.format_combo)
        
        main_layout.addWidget(format_widget)
        
        # Tabela de resultados com zebra striping
        self.tabela = QTableWidget()
        self.setup_tabela()
        self.tabela.setAlternatingRowColors(True)
        self.tabela.setStyleSheet(self.tabela.styleSheet() + "\nQTableWidget { alternate-background-color: #f5faff; }")
        main_layout.addWidget(self.tabela)
        
        # Barra de progresso
        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #005a9e;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
            }
        """)
        main_layout.addWidget(self.progress)
        
        # Botão processar com ícone
        btn_processar = QPushButton("Processar Arquivo")
        btn_processar.setIcon(QIcon.fromTheme("media-playback-start"))
        btn_processar.clicked.connect(self.processar_arquivo)
        main_layout.addWidget(btn_processar)
        
        layout.addWidget(main_card)
        
        # Rodapé em cartão
        footer_widget = QFrame()
        footer_widget.setObjectName("Card")
        footer_layout = QVBoxLayout(footer_widget)
        
        # Informações
        author_label = QLabel("Desenvolvido por: Matheus da Silva Castro")
        author_label.setStyleSheet("font-weight: bold; color: #333333;")
        author_label.setAlignment(Qt.AlignCenter)
        footer_layout.addWidget(author_label)
        
        ref_label = QLabel("Referência da codificação:")
        ref_label.setAlignment(Qt.AlignCenter)
        ref_label.setStyleSheet("color: #333333;")
        footer_layout.addWidget(ref_label)
        
        link_label = QLabel('<a href="https://ecivilufes.wordpress.com/wp-content/uploads/2011/04/inventc3a1rio-estac3a7c3b5es-pluviomc3a9tricas.pdf">Manual de Codificação</a>')
        link_label.setOpenExternalLinks(True)
        link_label.setAlignment(Qt.AlignCenter)
        footer_layout.addWidget(link_label)
        
        layout.addWidget(footer_widget)
        
    def setup_tabela(self):
        """Configura a tabela de resultados com todas as colunas necessárias."""
        colunas = [
            "Nome", "Latitude", "Longitude", "Código Gerado", 
            "Altitude", "Área Drenagem", "Bacia", "Sub-Bacia", "Rio",
            "Estado", "Município", "Responsável",
            "Escala", "Descarga Líquida", "Sedimentos",
            "Qualidade Água", "Pluviômetro", "Telemétrica",
            "Status"
        ]
        
        self.tabela.setColumnCount(len(colunas))
        self.tabela.setHorizontalHeaderLabels(colunas)
        
        # Configurar o cabeçalho da tabela
        header = self.tabela.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        
        # Estilo da tabela
        self.tabela.setStyleSheet("""
            QTableWidget {
                gridline-color: #dcdcdc;
                border: 1px solid #dcdcdc;
                background-color: white;
                color: #333333;
            }
            QHeaderView::section {
                background-color: #005a9e;
                color: white;
                padding: 4px;
                border: 1px solid #004080;
            }
        """)
        
    def selecionar_arquivo(self):
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Arquivo",
            "",
            "Access (*.mdb);;Todos os Arquivos (*.*)"
        )
        
        if arquivo:
            self.file_label.setText(arquivo)
            
    def processar_arquivo(self):
        arquivo = self.file_label.text()
        if arquivo == "Nenhum arquivo selecionado":
            QMessageBox.warning(self, "Erro", "Selecione um arquivo primeiro!")
            return
            
        try:
            # Processar arquivo
            dados = self.estacao_manager.processar_arquivo(arquivo)
            
            # Configurar tabela
            self.tabela.setRowCount(len(dados))
            self.progress.setMaximum(len(dados))
            
            # Processar cada registro
            for i, registro in enumerate(dados):
                try:
                    # Atualizar tabela
                    self.atualizar_linha_tabela(i, registro)
                    
                except Exception as e:
                    self.tabela.setItem(i, self.tabela.columnCount() - 1, 
                                      QTableWidgetItem(f"Erro: {str(e)}"))
                    
                self.progress.setValue(i + 1)
                QApplication.processEvents()
            
            # Exportar resultados
            formato = self.format_combo.currentText()
            extensao = ".xlsx" if "Excel" in formato else ".mdb"
            
            caminho_saida, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Arquivo",
                "",
                f"{'Excel' if extensao == '.xlsx' else 'Access'} (*{extensao})"
            )
            
            if caminho_saida:
                if not caminho_saida.endswith(extensao):
                    caminho_saida += extensao
                self.estacao_manager.exportar_resultados(dados, caminho_saida)
                QMessageBox.information(self, "Sucesso", 
                    f"Processamento concluído! Arquivo salvo em:\n{caminho_saida}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar arquivo: {str(e)}")
            
    def atualizar_linha_tabela(self, linha, registro):
        """Atualiza uma linha da tabela com os dados do registro."""
        # Dados principais
        self.tabela.setItem(linha, 0, QTableWidgetItem(registro.get('Nome', '')))
        self.tabela.setItem(linha, 1, QTableWidgetItem(str(registro.get('Latitude', ''))))
        self.tabela.setItem(linha, 2, QTableWidgetItem(str(registro.get('Longitude', ''))))
        self.tabela.setItem(linha, 3, QTableWidgetItem(str(registro.get('Codigo', ''))))
        
        # Dados adicionais
        self.tabela.setItem(linha, 4, QTableWidgetItem(str(registro.get('Altitude', ''))))
        self.tabela.setItem(linha, 5, QTableWidgetItem(str(registro.get('AreaDrenagem', ''))))
        self.tabela.setItem(linha, 6, QTableWidgetItem(registro.get('BaciaNome', '')))
        self.tabela.setItem(linha, 7, QTableWidgetItem(registro.get('SubBaciaNome', '')))
        self.tabela.setItem(linha, 8, QTableWidgetItem(registro.get('RioNome', '')))
        self.tabela.setItem(linha, 9, QTableWidgetItem(registro.get('EstadoSigla', '')))
        self.tabela.setItem(linha, 10, QTableWidgetItem(registro.get('MunicipioNome', '')))
        self.tabela.setItem(linha, 11, QTableWidgetItem(registro.get('ResponsavelNome', '')))
        
        # Campos booleanos
        self.tabela.setItem(linha, 12, QTableWidgetItem(registro.get('Escala', '')))
        self.tabela.setItem(linha, 13, QTableWidgetItem(registro.get('DescargaLiquida', '')))
        self.tabela.setItem(linha, 14, QTableWidgetItem(registro.get('Sedimentos', '')))
        self.tabela.setItem(linha, 15, QTableWidgetItem(registro.get('QualidadeAgua', '')))
        self.tabela.setItem(linha, 16, QTableWidgetItem(registro.get('Pluviometro', '')))
        self.tabela.setItem(linha, 17, QTableWidgetItem(registro.get('Telemetrica', '')))
        
        # Status
        self.tabela.setItem(linha, 18, QTableWidgetItem("Sucesso"))
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 