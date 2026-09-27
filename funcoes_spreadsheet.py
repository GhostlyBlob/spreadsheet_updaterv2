import pandas as pd
import openpyxl
import re

from configuracoes import (
    ATUALIZAR_MAQUINA_EXISTENTE,
    CAMPOS_IDENTIFICACAO,
    CAMPOS_MANUAIS,
    COLUNAS_COMPUTADOR,
    NOME_ABA_COMPUTADORES,
    CAMPOS_AGRUPAMENTO,
    CAMPO_ORDEM_NUMERICA,
    COMPUTADORES_PRIORITARIOS
)

def chave_natural(texto):
    partes = re.split(r"(\d+)", texto.casefold())
    chave = []

    for parte in partes:
        if parte.isdigit():
            chave.append((1, int(parte)))
        else:
            chave.append((0, parte))

    return tuple(chave)

def chave_ordenacao_inventario(dados_inventario):
    grupos = []

    for campo in CAMPOS_AGRUPAMENTO:
        valor = dados_inventario.get(campo, "")
        grupos.append(chave_natural(valor))

    computador = dados_inventario.get("Computador", "").casefold()
    prioridade = 1

    for nome in COMPUTADORES_PRIORITARIOS:
        if computador == nome.casefold():
            prioridade = 0
            break

    identificador = dados_inventario.get(CAMPO_ORDEM_NUMERICA, "")

    return (
        tuple(grupos),
        prioridade,
        chave_natural(identificador),
    )

def ler_inventarios(arquivo_texto):
    linhas_texto = arquivo_texto.read_text(
        encoding="utf-8-sig"
    ).splitlines()

    inventarios = []
    dados_inventario = {}

    for linha in linhas_texto:
        linha = linha.strip()

        if not linha:
            continue

        if linha == "=" * len(linha):
            if dados_inventario:
                inventarios.append(dados_inventario)
                dados_inventario = {}

            continue

        if linha.startswith("-"):
            linha = linha[1:].strip()

        if ":" not in linha:
            continue

        campo, valor = linha.split(":", maxsplit=1)
        dados_inventario[campo.strip()] = valor.strip()

    if dados_inventario:
        inventarios.append(dados_inventario)

    return inventarios

def validar_inventarios(inventarios):
    for numero_maquina, dados_inventario in enumerate(inventarios, start=1):
        for campo in ("ID", "Setor", "Número de série"):
            if not dados_inventario.get(campo, "").strip():
                print(
                    f"Erro: máquina {numero_maquina} "
                    f"está sem o campo '{campo}'."
                )
                return False

    return True

def montar_linha_computador(dados_inventario):
    observacoes = (
        f"MAC: {dados_inventario.get('MAC', '')} | "
        f"Espaço livre: {dados_inventario.get('Disco Livre', '')} | "
        f"Arquitetura: {dados_inventario.get('Arquitetura', '')} | "
        f"Levantamento: {dados_inventario.get('Data', '')}"
    )

    return {
        "N° Patrimônio": dados_inventario.get("ID", ""),
        "Tipo": dados_inventario.get("Tipo", ""),
        "Marca/Modelo": dados_inventario.get("Modelo", ""),
        "N° de Série": dados_inventario.get("Número de série", ""),
        "Processador": dados_inventario.get("CPU", ""),
        "Memória RAM (GB)": dados_inventario.get("RAM", ""),
        "Armazenamento Principal": dados_inventario.get("Disco Total", ""),
        "Sistema Operacional": dados_inventario.get("SO", ""),
        "Setor/Departamento": dados_inventario.get("Setor", ""),
        "Usuário responsável": dados_inventario.get("Usuário", ""),
        "Localização": "",
        "IP/Hostname": (
            f"{dados_inventario.get('IP', '')}/"
            f"{dados_inventario.get('Computador', '')}"
        ),
        "Status": "Em uso",
        "Última Manutenção": "",
        "Observações": observacoes,
    }
    
def criar_planilha(arquivo_excel, linhas_computadores):
    if arquivo_excel.exists():
        print(
            "Erro: já existe uma planilha com esse nome. "
            "Escolha outro nome ou use o modo de atualização."
        )
        return

    tabela_computador = pd.DataFrame(
        linhas_computadores,
        columns=COLUNAS_COMPUTADOR,
    )

    tabela_computador.to_excel(
        arquivo_excel,
        sheet_name="Computadores",
        index=False,
        engine="openpyxl",
    )

    print(f"Planilha criada com sucesso: {arquivo_excel}")
    
def encontrar_linhas_por_identificacao(aba_computadores, linha_computador):
    linhas_encontradas = []

    for numero_linha in range(2, aba_computadores.max_row + 1):
        identificacao_corresponde = True

        for campo in CAMPOS_IDENTIFICACAO:
            coluna = COLUNAS_COMPUTADOR.index(campo) + 1

            valor_da_planilha = aba_computadores.cell(
                row=numero_linha,
                column=coluna,
            ).value

            valor_do_texto = linha_computador[campo]

            if valor_da_planilha != valor_do_texto:
                identificacao_corresponde = False
                break

        if identificacao_corresponde:
            linhas_encontradas.append(numero_linha)

    return linhas_encontradas

def encontrar_linha_da_ultima_maquina_do_setor(aba_computadores, setor):
    coluna_setor = COLUNAS_COMPUTADOR.index("Setor/Departamento") + 1
    ultima_linha = None

    for numero_linha in range(2, aba_computadores.max_row + 1):
        setor_da_linha = aba_computadores.cell(
            row=numero_linha,
            column=coluna_setor,
        ).value

        if setor_da_linha == setor:
            ultima_linha = numero_linha

    return ultima_linha

def atualizar_planilha(arquivo_excel, linha_computador):
    pasta_trabalho = openpyxl.load_workbook(arquivo_excel)
    aba_computadores = pasta_trabalho[NOME_ABA_COMPUTADORES]

    cabecalhos_existentes = []

    for celula in aba_computadores[1]:
        cabecalhos_existentes.append(celula.value)

    if cabecalhos_existentes != COLUNAS_COMPUTADOR:
        print("Erro: os cabeçalhos da primeira aba não correspondem ao esperado.")
        return

    linhas_encontradas = encontrar_linhas_por_identificacao(
        aba_computadores,
        linha_computador,
    )

    if len(linhas_encontradas) == 0:
        setor = linha_computador["Setor/Departamento"]
        ultima_linha = encontrar_linha_da_ultima_maquina_do_setor(aba_computadores, setor)

        print(
            f"Identificação não encontrada. "
            f"Última linha do setor '{setor}': {ultima_linha}"
        )
        return

    if len(linhas_encontradas) > 1:
        print(
            "Erro: mais de uma máquina possui a mesma identificação. "
            "Nenhuma alteração foi realizada."
        )
        return

    if not ATUALIZAR_MAQUINA_EXISTENTE:
        print("Atualização de máquina existente está desativada.")
        return

    linha_para_atualizar = linhas_encontradas[0]

    for cabecalho, valor in linha_computador.items():
        if cabecalho in CAMPOS_MANUAIS:
            continue

        coluna = COLUNAS_COMPUTADOR.index(cabecalho) + 1

        aba_computadores.cell(
            row=linha_para_atualizar,
            column=coluna,
        ).value = valor

    pasta_trabalho.save(arquivo_excel)

    print(f"Máquina atualizada na linha {linha_para_atualizar}.")

