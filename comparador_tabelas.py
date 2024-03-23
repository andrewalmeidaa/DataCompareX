import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox
import os
import pandas as pd


class ProcessamentoArquivosApp(QWidget):
    def __init__(self):
        super().__init__()


        # Inicializa a interface do usuário
        self.initUI()


    def initUI(self):
        # Configura a geometria e o título da janela principal
        self.setGeometry(500, 300, 650, 400)
        self.setWindowTitle('Processamento de Arquivos')
        

        # Campos de entrada para os nomes dos arquivos
        self.label_arquivo1 = QLabel('Arquivo 1 (ANTIGO):', self)
        self.label_arquivo1.move(50, 20)
        


        self.entry_arquivo1 = QLineEdit(self)
        self.entry_arquivo1.setGeometry(192, 20, 250, 25)


        # Botão para escolher o Arquivo 1
        self.botao_escolher_arquivo1 = QPushButton('Escolher Arquivo', self)
        self.botao_escolher_arquivo1.setGeometry(450, 20, 120, 25)
        self.botao_escolher_arquivo1.clicked.connect(self.escolher_arquivo1)


        # Label e campos de entrada para o Arquivo 2
        self.label_arquivo2 = QLabel('Arquivo 2 (NOVO):', self)
        self.label_arquivo2.move(50, 60)


        self.entry_arquivo2 = QLineEdit(self)
        self.entry_arquivo2.setGeometry(192, 60, 250, 25)


        # Botão para escolher o Arquivo 2
        self.botao_escolher_arquivo2 = QPushButton('Escolher Arquivo', self)
        self.botao_escolher_arquivo2.setGeometry(450, 60, 120, 25)
        self.botao_escolher_arquivo2.clicked.connect(self.escolher_arquivo2)


        # Label e campo de entrada para o Diretório de Saída
        self.label_diretorio_saida = QLabel('Diretório de Saída:', self)
        self.label_diretorio_saida.move(50, 100)


        self.entry_diretorio_saida = QLineEdit(self)
        self.entry_diretorio_saida.setGeometry(190, 100, 250, 25)


        # Botão para escolher o Diretório de Saída
        self.botao_escolher_diretorio_saida = QPushButton('Escolher Diretório', self)
        self.botao_escolher_diretorio_saida.setGeometry(450, 100, 130, 25)
        self.botao_escolher_diretorio_saida.clicked.connect(self.escolher_diretorio_saida)


        # Botão para processar os arquivos
        self.botao_processar = QPushButton('Processar Arquivos', self)
        self.botao_processar.setGeometry(230, 150, 150, 30)
        self.botao_processar.clicked.connect(self.processar_arquivos)


        # Exibindo a janela
        self.show()


    def escolher_arquivo1(self):
        # Abre uma caixa de diálogo para escolher o Arquivo 1
        arquivo1, _ = QFileDialog.getOpenFileName(self, 'Escolha o Arquivo 1 (ANTIGO)')
        # Define o texto do campo de entrada com o caminho do Arquivo 1
        self.entry_arquivo1.setText(arquivo1)


    def escolher_arquivo2(self):
        # Abre uma caixa de diálogo para escolher o Arquivo 2
        arquivo2, _ = QFileDialog.getOpenFileName(self, 'Escolha o Arquivo 2 (NOVO)')
        # Define o texto do campo de entrada com o caminho do Arquivo 2
        self.entry_arquivo2.setText(arquivo2)


    def escolher_diretorio_saida(self):
        # Abre uma caixa de diálogo para escolher o Diretório de Saída
        diretorio_saida = QFileDialog.getExistingDirectory(self, 'Escolha o Diretório de Saída')
        # Define o texto do campo de entrada com o caminho do Diretório de Saída
        self.entry_diretorio_saida.setText(diretorio_saida)


    def processar_arquivos(self):
        # Obter os valores dos campos de entrada
        nome_arquivo1 = self.entry_arquivo1.text()
        nome_arquivo2 = self.entry_arquivo2.text()
        diretorio_saida = self.entry_diretorio_saida.text()

        # Verificar se os campos de entrada estão vazios
        if not nome_arquivo1 or not nome_arquivo2 or not diretorio_saida:
            QMessageBox.warning(self, 'Aviso', 'Por favor, escolha ambos os arquivos e o diretório de saída.')
            return

        try:
            # Leitura dos arquivos Excel
            bacu1 = pd.read_excel(nome_arquivo1)
            bacu2 = pd.read_excel(nome_arquivo2)

            # Remoção de colunas específicas
            # colunasRemocao = ['Nivel de Criticidade', 'Sucatas', 'Diferença', 'Minimo']
            # bacu1.drop(colunasRemocao, axis=1, inplace=True)
            # bacu2.drop(colunasRemocao, axis=1, inplace=True)

            # Merge dos DataFrames
            mergeResult = pd.merge(bacu1, bacu2, how='outer', indicator=True)
            diferenca = mergeResult[mergeResult['_merge'] != 'both']

            # Mapear os valores '_merge' para 'bacu1' e 'bacu2'
            diferenca['_merge'] = diferenca['_merge'].map({'left_only': 'tabela_antiga', 'right_only': 'tabela_nova'})
            diferenca.rename(columns={'_merge': 'diferenca'}, inplace=True)

            # Caminho para o novo arquivo Excel de saída
            novoArquivoDIFF = os.path.join(diretorio_saida, 'arquivo_diferenca.xlsx')
            diferenca.to_excel(novoArquivoDIFF, index=False)

            QMessageBox.information(self, 'Sucesso', f'Arquivo de saída gerado com sucesso: {novoArquivoDIFF}')

        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Ocorreu um erro ao processar os arquivos: {str(e)}')



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ProcessamentoArquivosApp()
    sys.exit(app.exec_())

