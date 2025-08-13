import os
import pyodbc
from typing import Optional, Dict, List
from dotenv import load_dotenv
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import shutil
import sys

# Carregar variáveis de ambiente
load_dotenv()

def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, funciona para dev e para PyInstaller"""
    try:
        # PyInstaller cria um diretório temporário e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)

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
        
    def _carregar_shapes_geograficos(self):
        """Carrega shapefiles de sub-bacia e município e retorna GeoDataFrames prontos para uso."""
        shape_dir = r"D:\Temp\Programacao\myenv\Scripts\Codificacao\insumo"
        subbacia_fp = os.path.join(shape_dir, "GEOFT_DNAEE_SUBBACIA.shp")
        municipio_fp = os.path.join(shape_dir, "Municipio_IBGE_Hidro.shp")

        gdf_subbacia = gpd.read_file(subbacia_fp).to_crs("EPSG:4326")
        gdf_municipio = gpd.read_file(municipio_fp).to_crs("EPSG:4326")
        return gdf_subbacia, gdf_municipio

    def _preencher_codigos_geograficos(self, registro, gdf_subbacia, gdf_municipio):
        """Preenche BaciaCodigo, SubBaciaCodigo, EstadoCodigo e MunicipioCodigo se estiverem nulos."""
        ponto = Point(registro['Longitude'], registro['Latitude'])
        ponto_gdf = gpd.GeoDataFrame(index=[0], geometry=[ponto], crs="EPSG:4326")

        # Sub-Bacia e Bacia
        inter_sub = gpd.sjoin(ponto_gdf, gdf_subbacia, how='left', predicate='intersects')
        if pd.isna(registro.get('SubBaciaCodigo')) or registro['SubBaciaCodigo'] in ('', None):
            registro['SubBaciaCodigo'] = inter_sub.at[0, 'DNS_NU_SUB'] if 'DNS_NU_SUB' in inter_sub.columns else None
        if pd.isna(registro.get('BaciaCodigo')) or registro['BaciaCodigo'] in ('', None):
            registro['BaciaCodigo'] = inter_sub.at[0, 'DNS_DNB_CD'] if 'DNS_DNB_CD' in inter_sub.columns else None

        # Município e Estado
        inter_mun = gpd.sjoin(ponto_gdf, gdf_municipio, how='left', predicate='intersects')
        if pd.isna(registro.get('MunicipioCodigo')) or registro['MunicipioCodigo'] in ('', None):
            registro['MunicipioCodigo'] = inter_mun.at[0, 'dbo__Mun_2'] if 'dbo__Mun_2' in inter_mun.columns else None
        if pd.isna(registro.get('EstadoCodigo')) or registro['EstadoCodigo'] in ('', None):
            registro['EstadoCodigo'] = inter_mun.at[0, 'dbo__Mun_1'] if 'dbo__Mun_1' in inter_mun.columns else None

        return registro
        
    def processar_arquivo(self, caminho_arquivo: str) -> List[Dict]:
        """Processa um arquivo .mdb e retorna os dados processados."""
        if not caminho_arquivo.endswith('.mdb'):
            raise ValueError("Formato de arquivo não suportado. Use apenas arquivos .mdb")
        # Limpa os códigos gerados da sessão anterior
        self.codigos_gerados = {}
        return self._processar_mdb(caminho_arquivo)
            
    def _processar_mdb(self, caminho_arquivo: str) -> List[Dict]:
        """
        Processa um arquivo .mdb e retorna os dados.
        
        O arquivo Access deve conter uma tabela 'Estacao' com as seguintes colunas:
        
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
        - DescargaLiquida: Indica se possui descarga líquida (Sim/Não)
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
            nome_tabela_entrada = 'Estacoes_Novas'
            if nome_tabela_entrada not in tabelas:
                raise ValueError(f"O arquivo MDB deve conter uma tabela chamada '{nome_tabela_entrada}'")
            
            # Define as colunas obrigatórias
            colunas_obrigatorias = [
                'Nome', 'Latitude', 'Longitude'
            ]
            
            # Verifica se todas as colunas obrigatórias estão presentes
            colunas_existentes = [column.column_name for column in cursor.columns(table=nome_tabela_entrada)]
            colunas_faltantes = [col for col in colunas_obrigatorias if col not in colunas_existentes]
            if colunas_faltantes:
                raise ValueError(
                    f"As seguintes colunas obrigatórias estão faltando na tabela '{nome_tabela_entrada}': {', '.join(colunas_faltantes)}"
                )
            
            # Buscar o maior RegistroID existente no banco
            try:
                with self.conn.cursor() as cursor_sql:
                    cursor_sql.execute("SELECT MAX(RegistroID) FROM dbo.Estacao")
                    max_registro_id = cursor_sql.fetchone()[0] or 0
            except Exception as e:
                raise Exception(f"Erro ao buscar maior RegistroID: {e}")
            
            # Lê os dados
            query = f"SELECT * FROM {nome_tabela_entrada} WHERE Codigo IS NULL"
            cursor.execute(query)
            
            columns = [column[0] for column in cursor.description]
            results = []
            try:
                if not hasattr(self, '_gdf_subbacia'):
                    self._gdf_subbacia, self._gdf_municipio = self._carregar_shapes_geograficos()
            except Exception as e:
                raise Exception(f"Erro ao carregar shapefiles: {e}")
            
            next_registro_id = max_registro_id + 1
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
                    # Preencher campos nulos com códigos geográficos
                    result = self._preencher_codigos_geograficos(result, self._gdf_subbacia, self._gdf_municipio)

                    results.append(result)

                except ValueError as e:
                    raise ValueError(f"Erro nas coordenadas da estação {result['Nome']}: {str(e)}")
                
                # Campos booleanos
                campos_booleanos = {
                    'Escala': 'TipoEstacaoEscala',
                    'DescargaLiquida': 'TipoEstacaoDescLiquida',
                    'Sedimentos': 'TipoEstacaoSedimentos',
                    'QualidadeAgua': 'TipoEstacaoQualAgua',
                    'Pluviometro': 'TipoEstacaoPluviometro',
                    'Telemetrica': 'TipoEstacaoTelemetrica',
                    'Operando': 'Operando'
                }
                
                for campo_origem, campo_destino in campos_booleanos.items():
                    if campo_origem in result:
                        valor = str(result[campo_origem]).strip().upper()
                        if valor in ('SIM', 'S', 'TRUE', '1', 'YES'):
                            result[campo_destino] = 1
                        elif valor in ('NÃO', 'NAO', 'N', 'FALSE', '0', 'NO'):
                            result[campo_destino] = 0
                        else:
                            result[campo_destino] = None
                
                # Converter códigos de texto para número
                campos_numericos = ['BaciaCodigo', 'SubBaciaCodigo', 'RioCodigo', 
                                  'MunicipioCodigo', 'EstadoCodigo', 'ResponsavelCodigo']
                for campo in campos_numericos:
                    if campo in result and result[campo] is not None and result[campo] != '':
                        try:
                            result[campo] = int(float(str(result[campo]).strip()))
                        except (ValueError, TypeError):
                            result[campo] = None
                
                # Atribuir RegistroID sequencial se não existir
                if 'RegistroID' not in result or result['RegistroID'] is None:
                    result['RegistroID'] = next_registro_id
                    next_registro_id += 1
                
            conn.close()
            return results
            
        except Exception as e:
            raise Exception(f"Erro ao processar arquivo MDB: {e}")
            
    def salvar_estacao(self, dados: Dict) -> bool:
        """Salva os dados da estação no banco de dados."""
        try:
            with self.conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO Estacao (
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
        A ordem das colunas será exatamente a mesma da tabela Estacao do Access.
        """
        if not dados:
            raise ValueError("Não há dados para exportar")
            
        # Ordem exata das colunas da tabela Estacao
        ordem_colunas = [
            'RegistroID', 'Importado', 'Temporario', 'Removido', 'ImportadoRepetido',
            'BaciaCodigo', 'SubBaciaCodigo', 'RioCodigo', 'EstadoCodigo', 'MunicipioCodigo',
            'ResponsavelCodigo', 'ResponsavelUnidade', 'ResponsavelJurisdicao',
            'OperadoraCodigo', 'OperadoraUnidade', 'OperadoraSubUnidade', 'TipoEstacao',
            'Codigo', 'Nome', 'CodigoAdicional', 'Latitude', 'Longitude', 'Altitude',
            'AreaDrenagem', 'TipoEstacaoEscala', 'TipoEstacaoRegistradorNivel',
            'TipoEstacaoDescLiquida', 'TipoEstacaoSedimentos', 'TipoEstacaoQualAgua',
            'TipoEstacaoPluviometro', 'TipoEstacaoRegistradorChuva', 'TipoEstacaoTanqueEvapo',
            'TipoEstacaoClimatologica', 'TipoEstacaoPiezometria', 'TipoEstacaoTelemetrica',
            'PeriodoEscalaInicio', 'PeriodoEscalaFim', 'PeriodoRegistradorNivelInicio',
            'PeriodoRegistradorNivelFim', 'PeriodoDescLiquidaInicio', 'PeriodoDescLiquidaFim',
            'PeriodoSedimentosInicio', 'PeriodoSedimentosFim', 'PeriodoQualAguaInicio',
            'PeriodoQualAguaFim', 'PeriodoPluviometroInicio', 'PeriodoPluviometroFim',
            'PeriodoRegistradorChuvaInicio', 'PeriodoRegistradorChuvaFim',
            'PeriodoTanqueEvapoInicio', 'PeriodoTanqueEvapoFim', 'PeriodoClimatologicaInicio',
            'PeriodoClimatologicaFim', 'PeriodoPiezometriaInicio', 'PeriodoPiezometriaFim',
            'PeriodoTelemetricaInicio', 'PeriodoTelemetricaFim', 'TipoRedeBasica',
            'TipoRedeEnergetica', 'RespAlt'
        ]
        
        # Converter para DataFrame e reordenar as colunas
        df = pd.DataFrame(dados)
        for col in ordem_colunas:
            if col not in df.columns:
                df[col] = None
        df = df[ordem_colunas]
        
        # Exportar baseado na extensão do arquivo
        if caminho_saida.endswith('.xlsx'):
            df.to_excel(caminho_saida, index=False)
        elif caminho_saida.endswith('.mdb'):
            caminho_template = resource_path("mdb/template.mdb")
            if not os.path.exists(caminho_template):
                raise FileNotFoundError("Arquivo de template 'Codificacao/mdb/template.mdb' não encontrado.")

            try:
                shutil.copy(caminho_template, caminho_saida)
            except Exception as e:
                raise IOError(f"Erro ao copiar o arquivo de template: {e}")

            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={caminho_saida};'
            )
            conn = None
            try:
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                
                nome_tabela = 'Estacao'
                
                # Lista de colunas que devem ser inseridas como inteiros
                colunas_inteiras = [
                    'Codigo', 'RegistroID', 'BaciaCodigo', 'SubBaciaCodigo', 'RioCodigo',
                    'EstadoCodigo', 'MunicipioCodigo', 'ResponsavelCodigo', 'TipoEstacao',
                    'Importado', 'Temporario', 'Removido', 'ImportadoRepetido'
                ]
                # Inserir dados
                for _, row in df.iterrows():
                    values = []
                    for col_name in ordem_colunas:
                        val = row[col_name]
                        if pd.isna(val) or val is None:
                            values.append(None)
                        elif col_name in colunas_inteiras:
                            try:
                                values.append(int(val))
                            except Exception:
                                values.append(None)
                        else:
                            values.append(str(val) if val is not None else None)
                    
                    placeholders = ','.join(['?' for _ in values])
                    cols_str = f"([{'], ['.join(ordem_colunas)}])" # Nomes de colunas entre colchetes
                    insert_query = f"INSERT INTO {nome_tabela} {cols_str} VALUES ({placeholders})"
                    
                    try:
                        cursor.execute(insert_query, tuple(values))
                    except pyodbc.Error as insert_error:
                        print(f"Erro ao inserir linha: {row.get('Nome', 'N/A')}. Erro: {insert_error}")
                        print(f"Query: {insert_query}")
                        print(f"Valores: {tuple(values)}")

                conn.commit()
            except pyodbc.Error as e:
                raise Exception(f"Erro ao escrever no banco de dados MDB: {e}")
            finally:
                if conn:
                    conn.close()
        else:
            raise ValueError("Formato de arquivo não suportado. Use .xlsx ou .mdb")
            
    def __del__(self):
        """Fecha a conexão com o banco quando o objeto é destruído."""
        if hasattr(self, 'conn'):
            self.conn.close() 