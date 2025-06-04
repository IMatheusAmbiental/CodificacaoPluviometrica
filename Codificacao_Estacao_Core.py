import os
import pyodbc
from typing import Optional, Dict, List
from dotenv import load_dotenv
import pandas as pd

# Carregar variáveis de ambiente
load_dotenv()

class EstacaoManager:
    def __init__(self):
        self.conn = self._conectar_banco()
        # Dicionário para armazenar os códigos gerados por quadrante
        self.codigos_gerados = {}
        
    def _conectar_banco(self) -> Optional[pyodbc.Connection]:
        """Estabelece conexão com o banco de dados SQL Server."""
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={os.getenv('DB_HOST', 'SQLPRD1')};"
                f"DATABASE={os.getenv('DB_NAME', 'Hidro')};"
                f"UID={os.getenv('DB_USER', 'iusr_hidroinfoanaRO')};"
                f"PWD={os.getenv('DB_PASSWORD', 'fj904NV94')};"
                f"PORT={os.getenv('DB_PORT', '1433')};"
            )
            return conn
        except pyodbc.Error as e:
            raise ConnectionError(f"Erro ao conectar ao banco de dados: {e}")
            
    def _verificar_codigos_existentes(self) -> List[str]:
        """Busca todos os códigos de estações pluviométricas existentes."""
        try:
            with self.conn.cursor() as cursor:
                cursor.execute("""
                    SELECT Codigo
                    FROM dbo.Estacao
                    WHERE TipoEstacao = 2
                    AND Importado = 0 
                    AND Removido = 0 
                    AND Temporario = 0 
                    AND ImportadoRepetido = 0
                """)
                return [str(row[0]) for row in cursor.fetchall()]
        except Exception as e:
            raise Exception(f"Erro ao buscar códigos existentes: {e}")
            
    def _validar_coordenadas(self, latitude: float, longitude: float) -> None:
        """Valida se as coordenadas estão dentro dos limites geográficos válidos."""
        if not (-90 <= latitude <= 90):
            raise ValueError("Latitude deve estar entre -90 e 90 graus.")
        if not (-180 <= longitude <= 180):
            raise ValueError("Longitude deve estar entre -180 e 180 graus.")
            
    def _formatar_latitude(self, latitude: float) -> str:
        """
        Formata a latitude seguindo as regras de codificação:
        - Para latitudes ao norte do Equador, soma-se 80 ao valor
        - Para latitudes ao sul do Equador, usa-se o valor absoluto
        Retorna sempre dois dígitos preenchidos com zero à esquerda.
        """
        if latitude >= 0:
            # Para latitudes ao norte do Equador, soma-se 80
            valor = 80 + int(abs(latitude))
        else:
            # Para latitudes ao sul do Equador, usa-se o valor absoluto
            valor = int(abs(latitude))
        return f"{valor:02d}"
            
    def _formatar_longitude(self, longitude: float) -> str:
        """
        Formata a longitude usando o valor absoluto em graus.
        Retorna sempre dois dígitos preenchidos com zero à esquerda.
        """
        return f"{int(abs(longitude)):02d}"
            
    def _buscar_ultimo_sequencial(self, prefixo_lat: str, prefixo_lon: str) -> int:
        """
        Busca o último número sequencial usado para uma determinada quadrícula (latitude/longitude).
        Retorna o último sequencial encontrado ou 0 se não houver registros.
        """
        try:
            with self.conn.cursor() as cursor:
                # Monta o prefixo do código (0LLLOO)
                prefixo_codigo = f"0{prefixo_lat}{prefixo_lon}"
                
                # Busca todos os códigos existentes para esta quadrícula
                cursor.execute("""
                    SELECT Codigo
                    FROM dbo.Estacao
                    WHERE TipoEstacao = 2
                    AND Importado = 0 
                    AND Removido = 0 
                    AND Temporario = 0 
                    AND ImportadoRepetido = 0
                    AND Codigo LIKE ?
                """, (f"{prefixo_codigo}%",))
                
                codigos = [str(row[0]) for row in cursor.fetchall()]
                if not codigos:
                    return 0
                    
                # Extrai os últimos 3 dígitos de cada código e encontra o maior
                sequenciais = [int(codigo[-3:]) for codigo in codigos]
                return max(sequenciais)
                
        except Exception as e:
            raise Exception(f"Erro ao buscar sequencial: {e}")
            
    def gerar_codigo_pluviometrica(self, latitude: float, longitude: float) -> str:
        """
        Gera um novo código para estação pluviométrica seguindo o formato 0LLLOOONNN, onde:
        - 0: Primeiro dígito fixo (estações "fora do curso d'água")
        - LLL: Latitude em graus (com ajuste +80 para norte)
        - OOO: Longitude em graus (valor absoluto)
        - NNN: Número sequencial único dentro da quadrícula
        """
        # Validar coordenadas
        self._validar_coordenadas(latitude, longitude)
        
        # Formatar latitude e longitude
        lat_valor = abs(int(latitude)) if latitude < 0 else 80 + int(latitude)
        lat_valor = f"{lat_valor:02d}"
        
        lon_valor = abs(int(longitude))
        lon_valor = f"{lon_valor:02d}"
        
        # Criar os códigos de início e fim do intervalo (0LLLOO)
        prefixo_inicio = f"0{lat_valor}{lon_valor}"  # Ex: 06337
        prefixo_fim = f"0{lat_valor}{int(lon_valor) + 1:02d}"  # Ex: 06338
        
        # Adicionar zeros para completar o código
        codigo_inicio = int(f"{prefixo_inicio}000")  # Ex: 06337000
        codigo_fim = int(f"{prefixo_fim}000")  # Ex: 06338000
        
        # Buscar códigos existentes no banco
        try:
            with self.conn.cursor() as cursor:
                cursor.execute("""
                    SELECT CAST(Codigo AS BIGINT) AS Codigo
                    FROM dbo.Estacao
                    WHERE TipoEstacao = 2
                    AND Importado = 0 
                    AND Removido = 0 
                    AND Temporario = 0 
                    AND ImportadoRepetido = 0
                    AND CAST(Codigo AS BIGINT) >= ?
                    AND CAST(Codigo AS BIGINT) < ?
                    ORDER BY Codigo
                """, (codigo_inicio, codigo_fim))
                
                resultados = cursor.fetchall()
                
            # Lista para armazenar todos os códigos (do banco e gerados)
            todos_codigos = [int(str(r[0])) for r in resultados]
            
            # Adicionar códigos já gerados nesta sessão para este quadrante
            if prefixo_inicio in self.codigos_gerados:
                todos_codigos.extend(self.codigos_gerados[prefixo_inicio])
            
            if todos_codigos:
                # Encontrar o maior sequencial usado
                maior_sequencial = int(str(max(todos_codigos))[-3:])
                novo_sequencial = maior_sequencial + 1
            else:
                novo_sequencial = 1
                
            if novo_sequencial > 999:
                raise ValueError("Limite de estações atingido para esta quadrícula (máximo: 999).")
            
            # Montar o código final
            codigo = f"{prefixo_inicio}{novo_sequencial:03d}"
            
            # Armazenar o novo código gerado
            if prefixo_inicio not in self.codigos_gerados:
                self.codigos_gerados[prefixo_inicio] = []
            self.codigos_gerados[prefixo_inicio].append(int(codigo))
            
            return codigo
            
        except Exception as e:
            raise Exception(f"Erro ao gerar código: {e}")
        
    def buscar_dados_geograficos(self, latitude: float, longitude: float) -> Dict:
        """Busca dados geográficos baseados nas coordenadas."""
        # Aqui você implementaria a lógica para buscar dados da API ou banco
        # Por enquanto, retornamos dados fictícios
        return {
            "bacia_codigo": None,
            "subbacia_codigo": None,
            "municipio_codigo": None,
            "estado_codigo": None,
            "rio_codigo": None
        }
        
    def processar_arquivo(self, caminho_arquivo: str) -> List[Dict]:
        """Processa um arquivo .mdb e retorna os dados processados."""
        if not caminho_arquivo.endswith('.mdb'):
            raise ValueError("Formato de arquivo não suportado. Use apenas arquivos .mdb")
        return self._processar_mdb(caminho_arquivo)
            
    def _processar_mdb(self, caminho_arquivo: str) -> List[Dict]:
        """
        Processa um arquivo .mdb e retorna os dados.
        
        O arquivo Access deve conter uma tabela 'Estacoes_Novas' com as seguintes colunas:
        
        Colunas Obrigatórias:
        - Nome: Nome da estação
        - Latitude: Latitude em graus decimais
        - Longitude: Longitude em graus decimais
        - BaciaCodigo: Código da bacia hidrográfica
        - SubBaciaCodigo: Código da sub-bacia
        - RioCodigo: Código do rio
        - MunicipioCodigo: Código do município
        - EstadoCodigo: Código do estado
        - ResponsavelCodigo: Código do responsável
        
        Colunas Opcionais:
        - Codigo: Código da estação (vazio para novas estações)
        - CodigoAdicional: Código adicional da estação
        - Altitude: Altitude em metros
        - AreaDrenagem: Área de drenagem
        - BaciaNome: Nome da bacia
        - SubBaciaNome: Nome da sub-bacia
        - RioNome: Nome do rio
        - EstadoSigla: Sigla do estado
        - MunicipioNome: Nome do município
        - ResponsavelNome: Nome do responsável
        - ResponsavelSigla: Sigla do responsável
        - EstacaoTipo: Tipo da estação
        - Escala: Indica se possui escala (Sim/Não)
        - Descarga Liquida: Indica se possui descarga líquida (Sim/Não)
        - Sedimentos: Indica se possui medição de sedimentos (Sim/Não)
        - QualidadeAgua: Indica se possui medição de qualidade da água (Sim/Não)
        - Pluviometro: Indica se possui pluviômetro (Sim/Não)
        - Telemetrica: Indica se possui telemetria (Sim/Não)
        - Operando: Status de operação (SIM/NÃO)
        """
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={caminho_arquivo};'
            )
            conn = pyodbc.connect(conn_str)
            
            # Verifica se a tabela Estacoes_Novas existe
            cursor = conn.cursor()
            tabelas = [row.table_name for row in cursor.tables(tableType='TABLE')]
            if 'Estacoes_Novas' not in tabelas:
                raise ValueError("O arquivo MDB deve conter uma tabela chamada 'Estacoes_Novas'")
            
            # Define as colunas obrigatórias
            colunas_obrigatorias = [
                'Nome', 'Latitude', 'Longitude',
                'BaciaCodigo', 'SubBaciaCodigo', 'RioCodigo',
                'MunicipioCodigo', 'EstadoCodigo',
                'ResponsavelCodigo'
            ]
            
            # Verifica se todas as colunas obrigatórias estão presentes
            colunas_existentes = [column.column_name for column in cursor.columns(table='Estacoes_Novas')]
            colunas_faltantes = [col for col in colunas_obrigatorias if col not in colunas_existentes]
            if colunas_faltantes:
                raise ValueError(
                    f"As seguintes colunas obrigatórias estão faltando na tabela: {', '.join(colunas_faltantes)}"
                )
            
            # Lê os dados
            query = "SELECT * FROM Estacoes_Novas WHERE Codigo IS NULL"
            cursor.execute(query)
            
            # Converte os resultados para dicionário
            columns = [column[0] for column in cursor.description]
            results = []
            for row in cursor.fetchall():
                result = dict(zip(columns, row))
                
                # Valida coordenadas
                try:
                    latitude = float(result['Latitude'])
                    longitude = float(result['Longitude'])
                    
                    if not (-90 <= latitude <= 90):
                        raise ValueError(f"Latitude inválida ({latitude}) para estação {result['Nome']}")
                    if not (-180 <= longitude <= 180):
                        raise ValueError(f"Longitude inválida ({longitude}) para estação {result['Nome']}")
                        
                    result['Latitude'] = latitude
                    result['Longitude'] = longitude
                    
                    # Gerar código para a estação
                    result['Codigo'] = self.gerar_codigo_pluviometrica(latitude, longitude)
                    
                except ValueError as e:
                    raise ValueError(f"Erro nas coordenadas da estação {result['Nome']}: {str(e)}")
                
                # Converte campos booleanos
                campos_booleanos = ['Escala', 'Descarga Liquida', 'Sedimentos', 
                                  'QualidadeAgua', 'Pluviometro', 'Telemetrica']
                for campo in campos_booleanos:
                    if campo in result:
                        valor = str(result[campo]).upper()
                        result[campo] = valor in ('SIM', 'TRUE', '1', 'YES')
                
                results.append(result)
            
            return results
            
        except Exception as e:
            raise Exception(f"Erro ao processar arquivo MDB: {e}")
            
    def salvar_estacao(self, dados: Dict) -> bool:
        """Salva os dados da estação no banco de dados."""
        try:
            with self.conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO Estacao_Novas (
                        BaciaCodigo, SubBaciaCodigo, RioCodigo, MunicipioCodigo,
                        EstadoCodigo, ResponsavelCodigo, OperadoraCodigo, TipoEstacao,
                        Codigo, Nome, Latitude, Longitude
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    dados['bacia_codigo'], dados['subbacia_codigo'],
                    dados['rio_codigo'], dados['municipio_codigo'],
                    dados['estado_codigo'], dados['responsavel_codigo'],
                    dados['operadora_codigo'], 2,  # 2 = Pluviométrica
                    dados['codigo'], dados['nome'],
                    dados['latitude'], dados['longitude']
                ))
                self.conn.commit()
                return True
        except Exception as e:
            self.conn.rollback()
            raise Exception(f"Erro ao salvar estação: {e}")
            
    def exportar_resultados(self, dados: List[Dict], caminho_saida: str) -> None:
        """
        Exporta os resultados para um arquivo Excel (.xlsx) ou Access (.mdb)
        """
        if not dados:
            raise ValueError("Não há dados para exportar")
            
        # Converter para DataFrame
        df = pd.DataFrame(dados)
        
        # Exportar baseado na extensão do arquivo
        if caminho_saida.endswith('.xlsx'):
            df.to_excel(caminho_saida, index=False)
        elif caminho_saida.endswith('.mdb'):
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={caminho_saida};'
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            # Criar tabela Estacoes_Codificadas
            colunas = [f"{col} TEXT" for col in df.columns]
            create_table = f"CREATE TABLE Estacoes_Codificadas ({', '.join(colunas)})"
            cursor.execute(create_table)
            
            # Inserir dados
            for _, row in df.iterrows():
                placeholders = ','.join(['?' for _ in row])
                insert_query = f"INSERT INTO Estacoes_Codificadas VALUES ({placeholders})"
                cursor.execute(insert_query, tuple(str(v) for v in row))
            
            conn.commit()
            conn.close()
        else:
            raise ValueError("Formato de arquivo não suportado. Use .xlsx ou .mdb")
            
    def __del__(self):
        """Fecha a conexão com o banco quando o objeto é destruído."""
        if hasattr(self, 'conn'):
            self.conn.close() 