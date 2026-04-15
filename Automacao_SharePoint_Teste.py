import os
import shutil
import pandas as pd
import unicodedata
import difflib
from datetime import datetime
import stat

# =====================================================
# 📁 CONFIGURAÇÕES (GENÉRICAS)
# =====================================================
BASE_LOCAL = r"C:\Users\SeuUsuario\OneDrive\Reports\Clientes"
BASE_SHAREPOINT = r"C:\Users\SeuUsuario\Empresa\SharePoint\Clientes"
EXCEL_PATH = r"C:\Users\SeuUsuario\OneDrive\Reports\Base_Clientes.xlsx"

ARQUIVOS_EXATOS = ["RelatorioA", "RelatorioB", "RelatorioC"]

# =====================================================
# 🔒 REMOVER PASTA COM SEGURANÇA
# =====================================================
def remover_pasta_seguro(caminho):
    def remover_readonly(func, path, _):
        os.chmod(path, stat.S_IWRITE)
        func(path)

    try:
        shutil.rmtree(caminho, onerror=remover_readonly)
        print(f"🗑️ Pasta removida: {caminho}")
    except Exception as e:
        print(f"⚠️ Não conseguiu remover (ignorado): {e}")

# =====================================================
# 🔠 NORMALIZAR TEXTO
# =====================================================
def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize('NFKD', texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    for char in ["-", ".", "/", "(", ")", "&", ",", "'", "_"]:
        texto = texto.replace(char, "")
    return " ".join(texto.split())

# =====================================================
# 🔎 BUSCA FLEXÍVEL
# =====================================================
def buscar_flexivel(nome, lista, limite=0.6):
    nome_norm = normalizar(nome)
    melhor = None
    melhor_score = 0

    for item in lista:
        item_norm = normalizar(item)

        if nome_norm in item_norm or item_norm in nome_norm:
            return item

        score = difflib.SequenceMatcher(None, nome_norm, item_norm).ratio()

        if score > melhor_score:
            melhor_score = score
            melhor = item

    return melhor if melhor_score >= limite else None

# =====================================================
# 🔎 BUSCAR PASTA POR PALAVRA
# =====================================================
def buscar_pasta(lista, palavras):
    for pasta in lista:
        pasta_norm = normalizar(pasta)
        for palavra in palavras:
            if normalizar(palavra) in pasta_norm:
                print(f"📂 Pasta encontrada: {pasta}")
                return pasta
    return None

# =====================================================
# 📅 ANO E MÊS
# =====================================================
def get_year_and_previous_month():
    today = datetime.today()
    year = today.year
    month = today.month - 1

    if month == 0:
        month = 12
        year -= 1

    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março",
        4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro",
        10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    return year, month, meses[month]

# =====================================================
# 🚀 MAIN
# =====================================================
def main():
    print("\n🚀 INICIANDO AUTOMAÇÃO...\n")

    year, month_num, mes_nome = get_year_and_previous_month()
    mes_formatado = f"{month_num:02d} - {mes_nome}"

    print(f"📅 Ano: {year} | Mês: {mes_formatado}\n")

    # ================= EXCEL =================
    df_temp = pd.read_excel(EXCEL_PATH, header=None)

    header_row = None
    col_cliente_idx = None
    col_pais_idx = None

    for i, row in df_temp.iterrows():
        for j, cell in enumerate(row):
            if pd.isna(cell):
                continue

            text = str(cell).upper()

            if ("CLIENTE" in text or "NOME" in text) and col_cliente_idx is None:
                col_cliente_idx = j

            if any(x in text for x in ["PAIS", "LOCAL", "PASTA", "COUNTRY"]) and col_pais_idx is None:
                col_pais_idx = j

        if col_cliente_idx is not None and col_pais_idx is not None:
            header_row = i
            break

    if header_row is None:
        raise Exception("❌ Cabeçalho não encontrado")

    df = pd.read_excel(EXCEL_PATH, header=header_row)

    col_cliente = df.columns[col_cliente_idx]
    col_pais = df.columns[col_pais_idx]

    clientes_locais = [
        p for p in os.listdir(BASE_LOCAL)
        if os.path.isdir(os.path.join(BASE_LOCAL, p))
    ]

    status_geral = []

    # =====================================================
    # LOOP CLIENTES
    # =====================================================
    for cliente_local in clientes_locais:

        print(f"\n🔎 Cliente: {cliente_local}")

        try:
            cliente_norm = normalizar(cliente_local)

            # ===== PAÍS =====
            pais = None
            melhor_score = 0

            for _, row in df.iterrows():
                score = difflib.SequenceMatcher(
                    None,
                    cliente_norm,
                    normalizar(row[col_cliente])
                ).ratio()

                if score > melhor_score:
                    melhor_score = score
                    pais = row[col_pais]

            if melhor_score < 0.5:
                raise Exception("Cliente não encontrado no Excel")

            # ===== CAMINHOS =====
            pasta_pais = buscar_flexivel(pais, os.listdir(BASE_SHAREPOINT))
            caminho_pais = os.path.join(BASE_SHAREPOINT, pasta_pais)

            pasta_cliente = buscar_flexivel(cliente_local, os.listdir(caminho_pais))
            caminho_cliente = os.path.join(caminho_pais, pasta_cliente)

            # ===== RELATÓRIOS =====
            pasta_relatorio = buscar_pasta(os.listdir(caminho_cliente), ["Relatorio", "Reports"])
            caminho_relatorio = os.path.join(caminho_cliente, pasta_relatorio)

            # ===== REMOVER PASTA ERRADA =====
            pasta_errada = os.path.join(caminho_relatorio, str(year))
            if os.path.exists(pasta_errada):
                remover_pasta_seguro(pasta_errada)

            # ===== CRIAR ANO =====
            caminho_ano = os.path.join(caminho_relatorio, f"Ano {year}")
            os.makedirs(caminho_ano, exist_ok=True)

            # ===== CRIAR MÊS =====
            caminho_mes = os.path.join(caminho_ano, mes_formatado)
            os.makedirs(caminho_mes, exist_ok=True)

            # ===== COPIAR =====
            origem = os.path.join(BASE_LOCAL, cliente_local)

            arquivos = [
                f for f in os.listdir(origem)
                if os.path.splitext(f)[0] in ARQUIVOS_EXATOS
            ]

            for arq in arquivos:
                shutil.copy2(
                    os.path.join(origem, arq),
                    os.path.join(caminho_mes, arq)
                )

            print(f"✔ OK ({len(arquivos)} arquivos)")
            status_geral.append((cliente_local, "OK"))

        except Exception as e:
            print(f"❌ Erro: {e}")
            status_geral.append((cliente_local, f"Erro: {e}"))

    # ================= FINAL =================
    df_status = pd.DataFrame(status_geral, columns=["Cliente", "Status"])

    caminho_saida = os.path.join(
        os.path.dirname(EXCEL_PATH),
        f"STATUS_{year}_{month_num:02d}.xlsx"
    )

    df_status.to_excel(caminho_saida, index=False)

    print(f"\n💾 Resultado salvo em: {caminho_saida}")
    print("\n🏁 FINALIZADO\n")


if __name__ == "__main__":
    main()