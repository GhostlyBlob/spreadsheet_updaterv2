from configuracoes import (
    ATUALIZAR_MAQUINA_EXISTENTE,
    CAMPOS_IDENTIFICACAO,
    CAMPOS_MANUAIS,
    COLUNAS_COMPUTADOR,
    CAMPOS_AGRUPAMENTO,
    CAMPO_ORDEM_NUMERICA,
    COMPUTADORES_PRIORITARIOS
)

from funcoes_spreadsheet import (
    criar_planilha,
    encontrar_linhas_por_identificacao,
    ler_inventarios,
    montar_linha_computador,
    atualizar_planilha,
    validar_inventarios,
    chave_ordenacao_inventario,
)

from funcoes_interface import (
    escolher_arquivos
)

def main():
    escolha = escolher_arquivos()

    if escolha is None:
        return

    acao, arquivo_texto, arquivo_excel = escolha
    inventarios = ler_inventarios(arquivo_texto)

    if not inventarios:
        print("Nenhuma máquina foi encontrada no arquivo de texto.")
        return
    
    if not validar_inventarios(inventarios):
        return

    print(f"Máquinas lidas: {len(inventarios)}")

    linhas_computadores = []
    
    inventarios.sort(key=chave_ordenacao_inventario)

    for dados_inventario in inventarios:
        linha_computador = montar_linha_computador(dados_inventario)
        linhas_computadores.append(linha_computador)

    if acao == "criar":
        criar_planilha(arquivo_excel, linhas_computadores)
    else:
        for linha_computador in linhas_computadores:
            atualizar_planilha(arquivo_excel, linha_computador)


if __name__ == "__main__":
    main()