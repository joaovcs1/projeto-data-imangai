"""Consolida o relatorio de contas a receber do Sienge por cliente e vencimento.

Uso:
    arquivos na raiz
    rodar $ python limpar_boletos.py
"""

import argparse
import re
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

COLUNAS = ["Data vecto", "Cliente", "Documento", "Parc", "TC", "Unid. princ", "Total"]
SAIDA = [*COLUNAS, "Sequência", "Emitido"]


def normalizar_rotulo(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return " ".join("".join(c for c in texto if not unicodedata.combining(c)).lower().split())


def texto(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def nome_limpo(valor):
    return " ".join(re.sub(r"\([^)]*\)", "", texto(valor)).split())


def valor_decimal(valor):
    if pd.isna(valor) or texto(valor) == "":
        raise ValueError("valor vazio")
    if isinstance(valor, str):
        raw = valor.strip().replace("R$", "").replace(" ", "")
        if "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
    else:
        raw = str(valor)
    try:
        numero = Decimal(raw)
        if not numero.is_finite():
            raise ValueError("valor nao finito")
        return numero.quantize(Decimal("0.01"))
    except InvalidOperation as exc:
        raise ValueError(f"valor numerico invalido: {valor!r}") from exc


def codigos_distintos(valores):
    return list(dict.fromkeys(t for v in valores if (t := texto(v))))


def formatar_tc(codigo):
    """Exemplos: PMV -> PM VAGA; P4V -> P4 VAGA; P3 VAGA -> P3 VAGA."""
    codigo = " ".join(codigo.upper().split())
    codigo = re.sub(r"^(PM|P\d+|0[12])\s*V(?:AGA)?\b", r"\1 VAGA", codigo)
    return codigo


def prioridade_tc(codigo):
    base = codigo.split()[0]
    vaga = "VAGA" in codigo.split()[1:]
    if base == "PM":
        return (1 if vaga else 0, 0, codigo)
    if base == "P2":
        return (3 if vaga else 2, 0, codigo)
    if re.fullmatch(r"P\d+", base):
        return (4, int(base[1:]), 1 if vaga else 0, codigo)
    if base == "02":
        return (5, 1 if vaga else 0, codigo)
    if base == "01":
        return (6, 1 if vaga else 0, codigo)
    return (7, 0, codigo)


def juntar_tc(valores):
    codigos = [formatar_tc(c) for c in codigos_distintos(valores)]
    return " + ".join(sorted(dict.fromkeys(codigos), key=prioridade_tc))


def localizar_cabecalho(planilha):
    for indice, linha in planilha.iterrows():
        rotulos = {normalizar_rotulo(v): pos for pos, v in enumerate(linha)}
        if all(normalizar_rotulo(c) in rotulos for c in COLUNAS):
            return indice, [rotulos[normalizar_rotulo(c)] for c in COLUNAS]
    raise ValueError("Cabecalho nao encontrado. Verifique as colunas esperadas: " + ", ".join(COLUNAS))


def processar(origem: Path, destino: Path):
    bruto = pd.read_excel(origem, sheet_name=0, header=None, dtype=object)
    cabecalho, indices = localizar_cabecalho(bruto)
    dados = bruto.iloc[cabecalho + 1:, indices].copy()
    dados.columns = COLUNAS
    dados["Cliente"] = dados["Cliente"].map(nome_limpo)
    dados["Data vecto"] = pd.to_datetime(dados["Data vecto"], dayfirst=True, errors="coerce")

    # Remove exclusivamente cabecalhos, subtotais, rodapes e linhas vazias.
    dados = dados.loc[(dados["Cliente"] != "") & dados["Data vecto"].notna()].copy()
    if dados.empty:
        raise ValueError("Nenhum boleto valido encontrado.")

    valores = []
    for indice, valor in dados["Total"].items():
        try:
            valores.append(valor_decimal(valor))
        except ValueError as exc:
            raise ValueError(f"Total invalido na linha {indice + 1} do Excel: {exc}") from exc
    dados["Total"] = valores
    dados["_cliente_chave"] = dados["Cliente"].map(lambda s: " ".join(s.casefold().split()))

    resultado = []
    for (_, data), grupo in dados.groupby(["_cliente_chave", "Data vecto"], sort=False):
        primeiro = grupo.iloc[0]
        resultado.append({
            "Data vecto": data,
            "Cliente": primeiro["Cliente"],
            "Documento": texto(primeiro["Documento"]),
            "Parc": " + ".join(codigos_distintos(grupo["Parc"])),
            "TC": juntar_tc(grupo["TC"]),
            "Unid. princ": texto(primeiro["Unid. princ"]),
            "Total": sum(grupo["Total"], Decimal("0.00")),
            "_cliente_chave": primeiro["_cliente_chave"],
        })

    final = pd.DataFrame(resultado)
    final = final.sort_values(["_cliente_chave", "Data vecto"], kind="stable").reset_index(drop=True)
    final["Sequência"] = final.groupby("_cliente_chave").cumcount() + 1
    final["Emitido"] = ""

    total_entrada = sum(dados["Total"], Decimal("0.00"))
    total_saida = sum(final["Total"], Decimal("0.00"))
    if total_entrada != total_saida:
        raise ValueError(f"Divergencia de valores: entrada={total_entrada}, saida={total_saida}")

    final["Total"] = final["Total"].map(float)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destino, engine="openpyxl", date_format="DD/MM/YYYY", datetime_format="DD/MM/YYYY") as writer:
        final[SAIDA].to_excel(writer, index=False, sheet_name="Boletos")
        aba = writer.sheets["Boletos"]
        for celula in aba["A"][1:]:
            celula.number_format = "DD/MM/YYYY"
        for celula in aba["G"][1:]:
            celula.number_format = '#,##0.00;[Red](#,##0.00);-'
        # Colunas de identificadores, como 0310, permanecem como texto.
    print(f"Arquivo gerado: {destino}")
    print(f"Registros: {len(dados)} | Boletos: {len(final)} | Total conferido: R$ {total_saida:,.2f}")
    return final


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("origem", nargs="?", default="relatorio.xlsx", type=Path)
    parser.add_argument("-o", "--saida", default="planilha_limpa.xlsx", type=Path)
    args = parser.parse_args()
    if args.origem.resolve() == args.saida.resolve():
        parser.error("O destino deve ser diferente da planilha original.")
    processar(args.origem, args.saida)


if __name__ == "__main__":
    main()
