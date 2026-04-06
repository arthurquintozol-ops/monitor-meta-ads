"""
Script CLI para monitoramento de anúncios da Meta.
Lê os links de um arquivo links.txt, consulta cada um,
e salva o histórico em dados.json para o dashboard web.

Uso: python atualizar.py
"""

import re
import json
import sys
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


ARQUIVO_LINKS = "links.txt"
ARQUIVO_DADOS = "dados.json"


# ── JavaScript que extrai tudo numa única chamada ──

JS_EXTRAIR_DADOS = """
() => {
    let nomePagina = "Não identificado";
    const body = document.body ? document.body.innerText : "";
    const linhas = body.split("\\n").map(l => l.trim()).filter(l => l.length > 0);
    const ignorar = new Set([
        "Biblioteca de Anúncios da Meta", "Biblioteca de Anúncios",
        "Relatório da Biblioteca de Anúncios", "API da Biblioteca de Anúncios",
        "Conteúdo de marca", "Anúncios", "Sobre", "Entrar", "Status do sistema"
    ]);

    for (let i = 0; i < Math.min(linhas.length, 120); i++) {
        if ((linhas[i] === "Anúncios" || linhas[i] === "Sobre") && i > 0) {
            const candidato = linhas[i - 1].trim();
            if (candidato && candidato.length < 100 && !ignorar.has(candidato)) {
                nomePagina = candidato;
                break;
            }
        }
    }

    if (nomePagina === "Não identificado") {
        for (const tag of ["h1", "h2", "h3", "strong"]) {
            const els = document.querySelectorAll(tag);
            for (const el of els) {
                const t = el.innerText.trim().replace(/\\s+/g, " ");
                if (t && t.length < 100 && !ignorar.has(t)) {
                    nomePagina = t;
                    break;
                }
            }
            if (nomePagina !== "Não identificado") break;
        }
    }

    const bodyNorm = body.replace(/\\s+/g, " ");
    const regexes = [
        /~\\s*([\\d.,]+)\\s*resultados?/i,
        /aproximadamente\\s*([\\d.,]+)\\s*resultados?/i,
        /([\\d.,]+)\\s*resultados?/i
    ];

    let total = null;
    let textoOriginal = null;

    for (const rx of regexes) {
        const m = bodyNorm.match(rx);
        if (m) {
            textoOriginal = m[0];
            const numLimpo = m[1].replace(/\\D/g, "");
            if (numLimpo) {
                total = parseInt(numLimpo, 10);
                break;
            }
        }
    }

    return { nomePagina, total, textoOriginal };
}
"""


def extrair_page_id_da_url(url: str) -> str:
    try:
        query = parse_qs(urlparse(url).query)
        return str(query.get("view_all_page_id", [""])[0]).strip()
    except Exception:
        return ""


def ler_links() -> list[str]:
    arquivo = Path(ARQUIVO_LINKS)
    if not arquivo.exists():
        print(f"ERRO: Arquivo '{ARQUIVO_LINKS}' não encontrado.")
        print("Crie o arquivo com um link por linha.")
        sys.exit(1)

    links = []
    vistos = set()
    for linha in arquivo.read_text(encoding="utf-8").splitlines():
        url = linha.strip()
        if url and "facebook.com/ads/library" in url and url not in vistos:
            vistos.add(url)
            links.append(url)

    return links


def carregar_dados() -> dict:
    arquivo = Path(ARQUIVO_DADOS)
    if arquivo.exists():
        try:
            return json.loads(arquivo.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"paginas": {}, "datas": [], "ultima_atualizacao": ""}


def salvar_dados(dados: dict):
    Path(ARQUIVO_DADOS).write_text(
        json.dumps(dados, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def extrair_com_js(page, url: str, page_id: str) -> dict:
    pagina = "Não identificado"

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)

        try:
            page.locator("text=resultado").first.wait_for(timeout=12000)
        except Exception:
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass

        dados = page.evaluate(JS_EXTRAIR_DADOS)

        pagina = dados.get("nomePagina") or "Não identificado"
        total = dados.get("total")

        if total is not None:
            return {"status": "ok", "pagina": pagina, "total": total}

        return {"status": "erro", "pagina": pagina, "total": 0, "msg": "Texto não encontrado"}

    except PlaywrightTimeoutError:
        return {"status": "erro", "pagina": pagina, "total": 0, "msg": "Timeout"}

    except Exception as e:
        return {"status": "erro", "pagina": pagina, "total": 0, "msg": str(e)}


def main():
    print("=== Monitor Meta Ads ===")
    print()

    urls = ler_links()
    if not urls:
        print("Nenhum link válido encontrado em links.txt")
        sys.exit(0)

    print(f"Links encontrados: {len(urls)}")
    print()

    dados = carregar_dados()
    data_hoje = datetime.now().strftime("%d/%m/%Y")

    if data_hoje not in dados["datas"]:
        dados["datas"].append(data_hoje)

    with sync_playwright() as p:
        print("Abrindo navegador...")
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ]
        )

        context = browser.new_context(
            viewport={"width": 1440, "height": 1200},
            locale="pt-BR",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/133.0.0.0 Safari/537.36"
            ),
        )

        for idx, url in enumerate(urls, start=1):
            page_id = extrair_page_id_da_url(url)
            chave = page_id or url

            print(f"  [{idx}/{len(urls)}] {page_id or url[:60]}...", end=" ", flush=True)

            page = context.new_page()
            try:
                resultado = extrair_com_js(page, url, page_id)
            finally:
                try:
                    page.close()
                except Exception:
                    pass

            nome = resultado["pagina"]
            total = resultado["total"]

            if resultado["status"] == "ok":
                print(f"OK  |  {nome}  |  {total} anúncios")
            else:
                print(f"ERRO  |  {nome}  |  {resultado.get('msg', '?')}")

            # Atualizar dados
            if chave not in dados["paginas"]:
                dados["paginas"][chave] = {
                    "nome": nome,
                    "url": url,
                    "page_id": page_id,
                    "historico": {}
                }

            # Atualizar nome caso tenha sido identificado
            if nome != "Não identificado":
                dados["paginas"][chave]["nome"] = nome

            dados["paginas"][chave]["historico"][data_hoje] = total

        browser.close()

    dados["ultima_atualizacao"] = datetime.now().strftime("%d/%m/%Y %H:%M")

    print()
    print("Salvando dados...")
    salvar_dados(dados)
    print("Concluído!")


if __name__ == "__main__":
    main()
