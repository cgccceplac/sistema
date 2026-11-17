#!/usr/bin/env python3
"""Conferência da base de pessoal CEPLAC.

Lê, em rh/entrada/, a planilha de servidores por unidade (.xls/.xlsx) e o PDF do
SIAPE "Total de servidores por situação" mais recentes, e gera em rh/dados/:
  pessoal.json   base tratada (ativos, outros, duplicados, vagas)
  siape.json     totais do SIAPE
  mudancas.csv   diferenças em relação à base anterior
e rh/RELATORIO.md (também enviado ao resumo da execução no GitHub Actions).
"""
import csv, json, os, re, subprocess, sys, unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from pypdf import PdfReader

RAIZ = Path(__file__).resolve().parents[1]
ENTRADA, DADOS = RAIZ / "entrada", RAIZ / "dados"

TAB2SIG = {
    "DIRETORIA": "Diretoria", "BAHIA - ADM E DIFUSAO": "COBES", "BAHIA - PESQUISA": "COPIB",
    "ESPIRITO SANTO - CEDES": "CEDES", "PARA - ADM E DIFUSAO": "COPAM", "PARA - PESQUISA": "COPIP",
    "AMAZONAS - CEDAM": "CEDAM", "RONDONIA - ADM E DIFUSAO": "COROM", "RONDONIA - PESQUISA": "COPIR",
    "MATO GROSSO - CEDEM": "CEDEM",
}
ORDEM = list(TAB2SIG.values())
FN_RE = re.compile(r"^(FCE|CCE)\s*[\d.]+", re.I)
EXIT_RE = re.compile(r"INATIVO|APOSENTAD|FALECID|EXONERAD|CEDIDO PARA OUTRO", re.I)


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn")


def norm(s):
    return re.sub(r"\s+", " ", "" if s is None else str(s)).strip()


def chave(nome):
    return re.sub(r"\s+", " ", sem_acento(re.sub(r"\(.*?\)", " ", nome)).upper()).strip()


def fmt(n):
    return f"{n:,}".replace(",", ".")


# ---------- escolha dos arquivos ----------
def mais_recente(padroes):
    achados = [p for pad in padroes for p in ENTRADA.glob(pad)]
    if not achados:
        return None

    def quando(p):
        try:
            t = subprocess.run(["git", "log", "-1", "--format=%ct", "--", str(p)], capture_output=True,
                               text=True, cwd=RAIZ).stdout.strip()
            if t:
                return int(t)
        except Exception:
            pass
        return int(p.stat().st_mtime)

    return max(achados, key=lambda p: (quando(p), p.name))


# ---------- planilha ----------
def ler_planilha(caminho):
    xl = pd.ExcelFile(caminho)
    regs, rodapes = [], {}
    for aba in xl.sheet_names:
        sig = TAB2SIG.get(sem_acento(aba).upper().strip())
        if not sig:
            continue
        linhas = xl.parse(aba, header=None, dtype=str).fillna("").values.tolist()
        hi = next((i for i, r in enumerate(linhas) if any(sem_acento(norm(c)).upper() == "NOME" for c in r)), None)
        if hi is None:
            continue
        hdr = [sem_acento(norm(c)).upper() for c in linhas[hi]]
        ni = hdr.index("NOME")
        idx = lambda n: hdr.index(n) if n in hdr else -1
        ci, oi, li = idx("CARGO"), idx("UNIDADE DE ORIGEM"), idx("LOTACAO")
        for i, r in enumerate(linhas[hi + 1:], start=hi + 2):
            cel = [norm(c) for c in r]
            junto = " ".join(cel).upper()
            if "TOTAL DE SERVIDORES" in junto:
                n = [c for c in cel if re.fullmatch(r"\d+", c)]
                if n:
                    rodapes[sig] = int(n[-1])
                continue
            nome = cel[ni] if ni < len(cel) else ""
            if not nome or sem_acento(nome).upper() == "NOME":
                continue
            funcao = next((re.sub(r"\s+", " ", c).upper() for c in cel if FN_RE.match(c)), "")
            saida = " | ".join(c for j, c in enumerate(cel) if j > ni and c and not FN_RE.match(c) and EXIT_RE.search(c))
            regs.append(dict(
                sig=sig, linha=i, nome=nome.upper(), chave=chave(nome),
                cargo=cel[ci] if ci >= 0 else "", origem=cel[oi] if oi >= 0 else "",
                lotacao=cel[li] if li >= 0 else "", funcao=funcao, saida=saida,
                vago=bool(re.search(r"\bVAGO\b", nome, re.I)), bruto=" ; ".join(c for c in cel if c)))
    return regs, rodapes


def classificar(regs):
    outros = [r for r in regs if r["saida"] or r["vago"]]
    mapa, dups = {}, []
    for r in (r for r in regs if not r["saida"] and not r["vago"]):
        e = mapa.get(r["chave"])
        if not e:
            mapa[r["chave"]] = r
            continue
        # prefere a linha da unidade (fora da Diretoria); a função da outra é herdada
        manter, largar = (r, e) if (e["sig"] == "Diretoria" and r["sig"] != "Diretoria") else (e, r)
        if not manter["funcao"] and largar["funcao"]:
            manter["funcao"] = largar["funcao"]
        dups.append(largar)
        mapa[r["chave"]] = manter
    ativos = list(mapa.values())
    for r in ativos:
        r["grupo"] = "clt" if re.search(r"\bCLT\b", r["bruto"], re.I) else "comissionado" if r["funcao"].startswith("CCE") else "permanente"
        r["cedido"] = bool(re.search(r"CEDIDO PARA|CEDIDO AO", r["bruto"], re.I))
        r["forca_efetiva"] = "FORCA EFETIVA" in sem_acento(r["bruto"]).upper()
    return ativos, outros, dups


# ---------- SIAPE ----------
def ler_siape(caminho):
    texto = "\n".join((p.extract_text() or "") for p in PdfReader(str(caminho)).pages)
    linhas = []
    for l in texto.splitlines():
        m = re.match(r"^\s*(\d{2})\s*-\s*(.+?)\s+([\d.]+)\s*$", l)
        if m:
            linhas.append(dict(cod=m[1], nome=re.sub(r"\s+", " ", m[2]).strip(), qtd=int(m[3].replace(".", ""))))
    if not linhas:
        sys.exit("ERRO: não encontrei as linhas de situação no PDF do SIAPE.")
    g = lambda rx: (re.search(rx, texto, re.I) or [None, ""])[1].strip()
    tot = re.search(r"TOTAL DE SERVIDORES\s+([\d.]+)", texto, re.I)
    ap = next((l["qtd"] for l in linhas if l["cod"] == "01"), None)
    return dict(mes=g(r"MES/ANO PAGAMENTO\s*:\s*([A-Z]{3}\d{4})"), emissao=g(r"DATA\s*:\s*(\d{2}[A-Z]{3}\d{4})"),
                upag=g(r"UPAG\s*:\s*([^\n]+)"), linhas=linhas, ativo_permanente=ap,
                total=int(tot[1].replace(".", "")) if tot else sum(l["qtd"] for l in linhas))


# ---------- comparação ----------
def comparar(antes, depois):
    A, D = {r["chave"]: r for r in antes}, {r["chave"]: r for r in depois}
    ev = []
    for k, n in D.items():
        a = A.get(k)
        if not a:
            ev.append(("entrou", n["sig"], n["nome"], "", n["funcao"] or "—"))
        elif a["sig"] != n["sig"]:
            ev.append(("mudou de unidade", n["sig"], n["nome"], a["sig"], n["sig"]))
        elif a["funcao"] != n["funcao"]:
            ev.append(("função", n["sig"], n["nome"], a["funcao"] or "—", n["funcao"] or "—"))
    for k, a in A.items():
        if k not in D:
            ev.append(("saiu", a["sig"], a["nome"], a["funcao"] or "—", ""))
    return ev


def main():
    DADOS.mkdir(exist_ok=True)
    pl, pdf = mais_recente(["*.xls", "*.xlsx"]), mais_recente(["*.pdf"])
    if not pl or not pdf:
        sys.exit("ERRO: coloque em rh/entrada/ a planilha (.xls/.xlsx) e o PDF do SIAPE.")
    regs, rodapes = ler_planilha(pl)
    if not regs:
        sys.exit("ERRO: não encontrei as 10 abas de unidades na planilha.")
    ativos, outros, dups = classificar(regs)
    siape = ler_siape(pdf)

    ant = DADOS / "pessoal.json"
    anterior = json.loads(ant.read_text(encoding="utf-8")) if ant.exists() else None
    ev = comparar(anterior["ativos"], ativos) if anterior else []

    cont = lambda g: sum(1 for r in ativos if r["grupo"] == g)
    perm, clt, comis = cont("permanente"), cont("clt"), cont("comissionado")
    ap = siape["ativo_permanente"]
    dif = perm - ap if ap is not None else None

    (DADOS / "pessoal.json").write_text(json.dumps(dict(
        gerado=datetime.now().isoformat(timespec="seconds"), planilha=pl.name, siape_pdf=pdf.name,
        ativos=ativos, outros=outros, duplicados=dups, vagas=[r for r in regs if r["vago"]]),
        ensure_ascii=False, indent=1), encoding="utf-8")
    (DADOS / "siape.json").write_text(json.dumps(siape, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(DADOS / "mudancas.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Mudança", "Unidade", "Servidor", "Antes", "Depois"])
        w.writerows(ev)

    L = [f"# Conferência da base de pessoal — SIAPE {siape['mes']}", "",
         f"Planilha: `{pl.name}` · PDF: `{pdf.name}` · {siape['upag']} · emitido em {siape['emissao']}", "",
         "| Indicador | Valor |", "|---|---:|",
         f"| Ativos permanentes (SIAPE, oficial) | **{fmt(ap)}** |",
         f"| Ativos na planilha (sem inativos, vagos e duplicados) | {fmt(len(ativos))} |",
         f"| Permanentes pela planilha (sem {clt} CLT e {comis} CCE) | {fmt(perm)} |",
         f"| Diferença para o SIAPE | {dif:+d} |", ""]
    if dif:
        L += [f"> A planilha tem {abs(dif)} {'a mais' if dif > 0 else 'a menos'} que o SIAPE. "
              "Sem o relatório nominal do SIAPE não dá para apontar quem é.", ""]
    L += ["## Por unidade", "", "| Unidade | Ativos | Linhas na aba | Rodapé da aba | Confere |", "|---|---:|---:|---:|---|"]
    for s in ORDEM:
        n_at = sum(1 for r in ativos if r["sig"] == s)
        n_lin = sum(1 for r in regs if r["sig"] == s)
        rod = rodapes.get(s)
        L.append(f"| {s} | {n_at} | {n_lin} | {rod if rod is not None else '—'} | {'sim' if rod == n_lin else 'NÃO' if rod is not None else '—'} |")
    L += ["", "## SIAPE por situação", "", "| Cód. | Situação | Quantidade |", "|---|---|---:|"]
    L += [f"| {l['cod']} | {l['nome']} | {fmt(l['qtd'])} |" for l in siape["linhas"]]
    L += [f"| | **Total** | **{fmt(siape['total'])}** |", ""]
    if dups:
        L += ["## Nomes repetidos entre abas (contados uma vez)", ""] + [f"- {r['nome']} ({r['sig']}, linha {r['linha']})" for r in dups] + [""]
    L += ["## Mudanças em relação à base anterior", ""]
    if anterior is None:
        L += ["Primeira execução: esta passa a ser a base de referência.", ""]
    elif not ev:
        L += ["Nenhuma mudança.", ""]
    else:
        L += ["| Mudança | Unidade | Servidor | Antes | Depois |", "|---|---|---|---|---|"]
        L += [f"| {a} | {b} | {c} | {d} | {e} |" for a, b, c, d, e in ev]
        L.append("")
    texto = "\n".join(L)
    (RAIZ / "RELATORIO.md").write_text(texto, encoding="utf-8")
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write(texto)
    print(texto)


if __name__ == "__main__":
    main()
