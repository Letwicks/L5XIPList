# UpdateIpList

Ferramenta para extrair módulos com endereço IP de arquivos **L5X** (Rockwell Automation) e atualizar planilhas de lista de IPs.

## Funcionalidades

- Leitura de arquivos `.L5X` (RSLogix 5000 / Studio 5000)
- Identificação de módulos com endereço IPv4 configurado nas portas
- Exibição em tabela interativa com Name, CatalogNumber e Address
- Atualização de planilhas `.xlsx` existentes no formato de lista de IPs
- Destaque visual: **azul** para novos dispositivos, **amarelo** para editados
- Detecção de conflitos (mesmo IP com dispositivos diferentes)
- Geração automática de nova planilha com data atual
- Interface gráfica amigável (tkinter)

## Requisitos

- Windows 10/11
- Python 3.14+ (para executar o script)
- Dependências: `openpyxl`, `Pillow`

## Download

Baixe o executável mais recente na seção [Releases](https://github.com/lewicks/UpdateIpList/releases).

## Como usar

### Script Python

```bash
pip install openpyxl Pillow
python l5x_extractor_gui.py
```

### Executável

1. Abra `UpdateIpList.exe`
2. Clique em **Procurar L5X** e selecione o arquivo `.L5X`
3. Os módulos com IP serão listados na tabela
4. Clique em **Procurar Planilha** e selecione a planilha `.xlsx` existente
5. Clique em **Atualizar Planilha**
6. Uma nova planilha será criada com a data atual

### Como funciona a atualização

Para cada módulo encontrado no L5X, o programa:

1. Reconstrói o IP completo a partir das colunas OCT1, OCT2, OCT3, OCT4 da planilha
2. Localiza a linha correspondente ao IP do módulo
3. Remove o `_` do início do nome do módulo
4. Compara com o valor atual da coluna **Nomenclatura**:
   - **Vazio** → preenche com o nome e destaca em **azul**
   - **Igual** → mantém e destaca em **amarelo**
   - **Diferente** → alerta de **conflito** (não altera)

## Estrutura da planilha esperada

| Coluna | Conteúdo |
|--------|----------|
| E | Nomenclatura |
| H | OCT1 |
| I | `.` |
| J | OCT2 |
| K | `.` |
| L | OCT3 |
| M | `.` |
| N | OCT4 |

Os dados devem iniciar na linha 12, com IPs sequenciais de 1 a 254.

## Autor

**Autolinx Automação**  
Contato: [lewicks@gmail.com](mailto:lewicks@gmail.com)

## Licença

Distribuído sob a licença MIT.
