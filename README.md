# Sistema de Codificação de Estações Pluviométricas

Sistema desenvolvido para gerar códigos únicos para estações pluviométricas, seguindo o padrão estabelecido pela ANA (Agência Nacional de Águas).

## Funcionalidades

- Interface gráfica moderna e intuitiva
- Geração automática de códigos únicos para estações pluviométricas
- Processamento em lote de arquivos de importação
- Exportação dos resultados em Excel (.xlsx) ou Access (.mdb)
- Validação completa das coordenadas geográficas
- Verificação de duplicidade de códigos

## Formato do Código

O código da estação segue o formato `0LLLOOONNN`, onde:
- `0`: Primeiro dígito fixo (estações "fora do curso d'água")
- `LLL`: Latitude em graus (com ajuste +80 para norte)
- `OOO`: Longitude em graus (valor absoluto)
- `NNN`: Número sequencial único dentro da quadrícula

### Exemplo
Para uma estação com:
- Latitude: -6.81°
- Longitude: -37.2°

O código seria formado por:
- `0`: Dígito fixo
- `063`: Latitude (-6° → 06)
- `37`: Longitude (-37° → 37)
- `NNN`: Próximo sequencial disponível na quadrícula

## Estrutura dos Arquivos de Importação

O sistema aceita arquivos no formato Microsoft Access (.mdb) para importação em massa de estações.

### Tabela Estacoes_Novas

A tabela deve conter as seguintes colunas:

#### Colunas Obrigatórias

| Coluna | Tipo | Descrição |
|--------|------|-----------|
| Nome | Texto | Nome da estação |
| Latitude | Decimal | Latitude em graus decimais (entre -90 e 90) |
| Longitude | Decimal | Longitude em graus decimais (entre -180 e 180) |

#### Colunas Opcionais

| Coluna | Tipo | Descrição |
|--------|------|-----------|
| Altitude | Decimal | Altitude em metros |
| AreaDrenagem | Decimal | Área de drenagem |
| Bacia | Texto | Nome da bacia |
| Sub-Bacia | Texto | Nome da sub-bacia |
| Rio | Texto | Nome do rio |
| Estado | Texto | Nome/Sigla do estado |
| Municipio | Texto | Nome do município |
| Responsavel | Texto | Nome do responsável |
| Escala | Texto | Possui escala (Sim/Não) |
| Descarga Liquida | Texto | Possui descarga líquida (Sim/Não) |
| Sedimentos | Texto | Possui medição de sedimentos (Sim/Não) |
| QualidadeAgua | Texto | Possui medição de qualidade da água (Sim/Não) |
| Pluviometro | Texto | Possui pluviômetro (Sim/Não) |
| Telemetrica | Texto | Possui telemetria (Sim/Não) |

## Validações

O sistema realiza as seguintes validações:

1. Coordenadas geográficas:
   - Latitude entre -90° e 90°
   - Longitude entre -180° e 180°
   - Formato decimal com ponto como separador

2. Arquivo de importação:
   - Presença da tabela 'Estacoes_Novas'
   - Presença das colunas obrigatórias
   - Formato correto dos dados

3. Geração de códigos:
   - Verificação de códigos existentes no banco
   - Verificação de códigos gerados na sessão atual
   - Garantia de sequencial único por quadrícula

## Exportação

Os resultados podem ser exportados em dois formatos:
- Microsoft Excel (.xlsx)
- Microsoft Access (.mdb)

Os arquivos exportados incluirão todas as informações das estações processadas, incluindo os códigos gerados.

## Requisitos

- Windows 10 ou superior
- Driver ODBC para Microsoft Access
- Permissão de leitura/escrita na pasta de instalação

## Instalação

1. Baixe o arquivo executável mais recente da seção [Releases](../../releases)
2. Execute o arquivo `Codificacao_Estacao.exe`
3. Não é necessária instalação adicional

## Desenvolvimento

Para desenvolver ou modificar o sistema:

1. Clone o repositório
```bash
git clone https://github.com/seu-usuario/codificacao-estacoes.git
```

2. Instale as dependências
```bash
pip install -r requirements.txt
```

3. Execute o programa
```bash
python Codificacao_Estacao_GUI.py
```

## Autor

Desenvolvido por: Matheus da Silva Castro

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes. 