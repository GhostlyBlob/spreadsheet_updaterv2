Contexto do projeto: atualizador de inventário

Este arquivo orienta a continuação do projeto Python `spreadsheet_updaterv2`. Foi elaborado a partir da conversa com o usuário. **Antes de alterar código, confira os arquivos reais e o estado atual do projeto:** a última mudança sugerida pode ainda não ter sido aplicada. Esta cópia de `AGENTS.md` pode acompanhar o projeto em pastas e sistemas operacionais diferentes, dependendo de quem estiver desenvolvendo.

## Forma de colaboração

- Entregue as alterações em passos pequenos, indique o arquivo e o local exato de cada mudança e aguarde o resultado do teste antes de avançar. Não reescreva o programa inteiro nem edite os scripts por conta própria, salvo pedido explícito.
- Responda em português claro. Dê exemplos concretos de entrada e saída quando ajudar a esclarecer o comportamento do código.
- Prefira testes com arquivos fictícios ou cópias da pasta de trabalho. Não trate uma atualização parcial ou uma mensagem de diagnóstico como sucesso completo.

## Arquivos e ambiente

- `spreadsheet_updaterv2.py`: fluxo principal (`main`), escolha entre criar e atualizar e processamento da lista de máquinas.
- `configuracoes.py`: cabeçalhos, identificação, colunas manuais e parâmetros editáveis de ordenação.
- `funcoes_interface.py`: janelas `tkinter`, seleção de `.txt` e `.xlsx`.
- `funcoes_spreadsheet.py`: leitura, validação, montagem de linhas, ordenação, criação e atualização da planilha.
- Desenvolvimento e testes podem ocorrer em qualquer sistema operacional com Python 3 instalado; o ambiente específico (distribuição, editor, gerenciador de pacotes) fica a critério de quem estiver desenvolvendo. Exemplo de configuração com `venv`:

  ```bash
  python -m venv .venv
  source .venv/bin/activate       # Linux/macOS
  # .venv\Scripts\activate        # Windows
  pip install pandas openpyxl
  python spreadsheet_updaterv2.py
  ```

  A ativação do ambiente virtual vale apenas para o terminal atual; ao abrir outro terminal, ative-o novamente. `tkinter` normalmente já vem com o Python; em algumas distribuições Linux pode ser necessário instalar um pacote do sistema separado (ex.: `python3-tk`).
- O programa final deverá rodar em um computador Windows sem Python instalado. Planeja-se gerar e testar um executável na máquina virtual Windows **depois** de concluir a lógica. Não presumir que um executável criado em outro sistema operacional servirá diretamente no Windows.
- Bibliotecas escolhidas: `pandas` para criar a planilha, `openpyxl` para atualizar a pasta de trabalho existente e `tkinter` para as janelas.

## Entrada `.txt`

- Cada linha de dado tem formato `Campo: Valor`. Remova espaços das pontas com `.strip()`, mantendo os espaços internos, inclusive no número de série. Separe somente no primeiro `:` com `split(":", maxsplit=1)`, pois o MAC pode conter `:`.
- Linhas de periféricos podem começar com `-`; remova esse prefixo antes de separar campo e valor.
- Máquinas são separadas por uma linha composta de `=`, uma linha vazia e outra linha composta de `=`. Separadores consecutivos não devem criar máquinas vazias. A leitura atual usa `encoding="utf-8-sig"` e `splitlines()`.
- Campos conhecidos: `ID`, `Data`, `Setor`, `Tipo`, `Usuário`, `Computador`, `SO`, `Arquitetura`, `CPU`, `RAM`, `Disco Total`, `Disco Livre`, `IP`, `MAC`, `Número de série`, `Modelo`, `Discos Físicos`; além de `Teclado`, `Mouse`, `Monitor` e `Caixinha de som` com prefixo `-`. Exemplos de teste também contêm `Localização`, `Última Manutenção` e `Observações`.
- `ler_inventarios(arquivo_texto)` devolve uma lista de dicionários, um por máquina. Linhas sem `:` são ignoradas. O usuário testou um `.txt` com quatro máquinas e os separadores acima.
- Antes de escrever Excel, `validar_inventarios` exige `ID`, `Setor` e `Número de série` preenchidos em **todas** as máquinas. O teste sem `ID` na segunda máquina mostrou `Erro: máquina 2 está sem o campo 'ID'.` e interrompeu a operação. Esses três campos são exigidos porque compõem a identificação atualmente adotada; rever a regra se o arquivo real tiver número de série ausente.

## Pasta de trabalho e colunas

Há cinco abas na pasta de trabalho real, nesta ordem: `Computadores`, `Discos Rígidos`, `Periféricos`, `Projetores`, `Impressoras`. Por enquanto, alterar **somente** `Computadores` e preservar as outras abas. O cabeçalho da aba real fica na primeira linha, com fundo azul e texto branco. A criação atual gera uma nova pasta de trabalho com apenas a aba `Computadores`; decidir depois se a nova pasta deverá conter as outras quatro abas.

Cabeçalhos **exatos** de `Computadores`, nas colunas A:O:

1. `N° Patrimônio`
2. `Tipo`
3. `Marca/Modelo`
4. `N° de Série`
5. `Processador`
6. `Memória RAM (GB)`
7. `Armazenamento Principal`
8. `Sistema Operacional`
9. `Setor/Departamento`
10. `Usuário responsável` (letra `r` minúscula)
11. `Localização`
12. `IP/Hostname`
13. `Status`
14. `Última Manutenção`
15. `Observações`

`COLUNAS_COMPUTADOR` deve reproduzir essa ordem e grafia literalmente. Ao atualizar, carregar a aba por nome `Computadores`, conferir os cabeçalhos e não escrever se não corresponderem.

Mapeamento do `.txt` para a aba: `ID` → patrimônio; `Tipo` → tipo; `Modelo` → marca/modelo; `Número de série` → série; `CPU` → processador; `RAM` → memória; `Disco Total` → armazenamento; `SO` → sistema operacional; `Setor` → setor; `Usuário` → usuário. `IP/Hostname` junta `IP` e `Computador` com `/`. `Status` é `Em uso`. Não fixar o SO como Windows 10: embora a planilha real tenha muitos registros `Microsoft Windows 10 Pro`, os testes usam outros valores e o código lê o `.txt`.

`Observações` é automática e **deve ser atualizada**, no formato:

```text
MAC: {MAC} | Espaço livre: {Disco Livre} | Arquitetura: {Arquitetura} | Levantamento: {Data}
```

`CAMPOS_MANUAIS` contém somente `Localização` e `Última Manutenção`: ambas permanecem vazias ao criar e têm seus valores existentes preservados ao atualizar. O `.txt` de teste contém esses nomes e uma `Observações` manual, mas isso não muda o mapeamento atual; o texto manual de `Observações` é substituído pelo formato automático. Como `Tipo`, MAC, Disco Livre, Arquitetura e Data podem faltar no `.txt` fictício, é esperado que algumas células ou partes do texto de observações apareçam vazias. Colunas estreitas podem cortar a exibição, sem cortar o valor armazenado.

## Identificação, criação e atualização

- `CAMPOS_IDENTIFICACAO` contém atualmente `N° Patrimônio`, `Setor/Departamento` e `N° de Série`. O patrimônio sozinho **não é único**: nomes como `LABRAD-PC-001` podem se repetir em setores diferentes. Usuário e localização podem mudar, portanto não devem ser usados como chave fixa sem nova decisão.
- `ATUALIZAR_MAQUINA_EXISTENTE = True` habilita a atualização de uma linha encontrada. Na busca: zero correspondências significa máquina ainda não encontrada; uma permite atualizar; mais de uma correspondência pela chave composta deve interromper a alteração.
- Na atualização, preservar `CAMPOS_MANUAIS` e escrever `Observações` automática. O código atual usa `openpyxl.load_workbook`, encontra a aba `Computadores`, verifica cabeçalhos, procura cada máquina, altera os campos automáticos e salva.
- **Ainda não há inserção de máquinas novas no modo atualizar.** O ramo de zero correspondências apenas informa que a identificação não foi encontrada e mostra a última linha daquele setor, ou `None` se o setor ainda não existir. `None` é diagnóstico, não erro do Python. Não afirmar que a inserção já funciona.
- O usuário pediu aviso e confirmação se houver patrimônio repetido, mesmo quando os demais campos diferenciem as máquinas. Essa confirmação ainda não foi implementada. Antes de permitir inserção automática, evitar criar duplicatas silenciosas quando o mesmo patrimônio já existir com setor ou série diferente.
- O modo criar usa `pd.DataFrame(linhas_computadores, columns=COLUNAS_COMPUTADOR)`; `linhas_computadores` **já é uma lista**. Não envolvê-la novamente em `[...]`, pois isso provocou `ValueError: 15 columns passed, passed data had 2 columns`.
- Atualmente a atualização processa e salva uma máquina de cada vez. Se uma operação posterior falhar, as primeiras podem já ter sido gravadas. Planejar validação e gravação em lote/uma única gravação antes de usar a pasta real.

## Ordenação acordada

Prioridade: agrupar primeiro por `Setor` e `Localização`; dentro do grupo, computadores prioritários como `LABINFO` vêm antes dos demais; depois ordenar numericamente pelo identificador (`PATR-2` antes de `PATR-10`). Os parâmetros no `configuracoes.py` usam **nomes do `.txt`**:

```python
CAMPOS_AGRUPAMENTO = ["Setor", "Localização"]
CAMPO_ORDEM_NUMERICA = "ID"
COMPUTADORES_PRIORITARIOS = ["LABINFO"]
```

`chave_natural` usa `re.split(r"(\d+)", texto.casefold())`; `chave_ordenacao_inventario` compõe grupo, prioridade e ID. O modo criar chama `inventarios.sort(key=chave_ordenacao_inventario)` **antes** de montar as linhas. Testes confirmaram Laboratório 3 (`PATR-006`) antes do Laboratório 4 (`PATR-003`), e o computador `LABINFO` (`PATR-001`) antes de `PC-000` (`PATR-000`) no mesmo setor/localização. O modo atualizar ainda **não** reorganiza linhas existentes: deve apenas inserir novas na posição correta, sem ordenar a aba inteira.

`Localização` é uma restrição importante: o valor original existe em alguns `.txt` de teste e participa da ordenação de **criação**, mas a coluna Excel é manual e fica vazia ao criar. Para comparar linhas existentes no modo atualizar, usar o valor manual quando estiver preenchido; quando estiver vazio, tratá-lo como igual à localização da máquina nova para que o setor e o ID determinem a posição. Isso **não comprova** a localização real da linha: se um mesmo setor tiver vários locais com células vazias, a ordem entre esses locais é ambígua e deve ser resolvida com o usuário antes de automatizar esse caso.

## Ponto exato de retomada

A última orientação foi acrescentar a `configuracoes.py` um `MAPA_ORDENACAO_EXCEL` que relaciona `ID`/`Setor`/`Localização`/`Computador` aos cabeçalhos da aba, e acrescentar a `funcoes_spreadsheet.py` a função `dados_ordenacao_da_linha(aba, numero_linha, localizacao_referencia)`. **O usuário ainda não confirmou se aplicou esse último passo.** Confira os arquivos antes de prosseguir.

Depois disso, a sequência sugerida é:

1. Calcular, sem escrever, a posição em que uma máquina nova entraria, comparando sua `chave_ordenacao_inventario` às chaves das linhas existentes; testar setores existentes e setores novos, inclusive `Laboratório 10` versus `Laboratório 2`.
2. Tratar patrimônio já presente e correspondências ambíguas com o aviso/confirmação solicitado; não inserir uma duplicata silenciosamente.
3. Inserir a máquina nova na posição calculada usando `openpyxl`, preservando cabeçalhos e demais abas; verificar estilos, filtros, fórmulas, tabelas e referências da pasta real antes de confiar em `insert_rows`.
4. Evitar gravações parciais quando várias máquinas são processadas. Testar em cópia da pasta real.
5. Só depois abordar periféricos, acabamento visual, testes no Windows e empacotamento do executável.

## Periféricos e abas futuras

Os periféricos vêm no mesmo `.txt` e são lidos como campos com `-`, mas ainda **não** são escritos no Excel. A aba `Periféricos` possui as colunas: `N° Patrimônio`, `Tipo`, `Marca`, `Modelo`, `N° de Série`, `Equipamento Associado (N° Patrimônio)`, `Setor/Departamento`, `Usuário Responsável`, `Localização`, `Status`, `Observações`. Cada periférico ocupa sua própria linha. Exemplo: `SDLC-PC-004-MOUSE` é o patrimônio do periférico e `SDLC-PC-004` é o equipamento associado. `Discos Rígidos`, `Projetores` e `Impressoras` ficam fora do escopo por enquanto.

## Testes e limites conhecidos

- Foram testados com dados fictícios: seleção de arquivos, criação, atualização de máquina encontrada, leitura de múltiplos blocos, validação de campo ausente e ordenação no modo criar. O resultado visual da criação tem colunas estreitas e cabeçalho sem a formatação azul/branco da pasta real; acabamento ainda não foi implementado.
- Repetir os testes após alterações relevantes com IDs repetidos em setores diferentes, séries diferentes, uma máquina ausente em setor existente, uma máquina de setor novo, `LABINFO` antes dos numerados e valores de observações completos.
- Não publicar no GitHub um `.txt` ou `.xlsx` real do trabalho com dados de usuários, patrimônio, IP, MAC ou números de série. Publicar somente código e dados fictícios autorizados.