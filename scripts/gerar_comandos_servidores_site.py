"""Gera as páginas de comandos dos Encontros 2 e 3 (site com páginas, cadastro, login).

Junta scripts/site_servidor/estilo.html com o corpo de cada encontro e coloca,
escapado, o código de cada arquivo em {{code:caminho}}.
"""
import html
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent / "site_servidor"
SAIDA = Path(__file__).resolve().parent.parent

PAGINAS = [
    ("pagina_encontro2.html", "Site da Equipe · Encontro 2",
     "Comandos - Servidores 2 - Paginas, fotos e cadastro.html"),
    ("pagina_encontro3.html", "Entrega do Servidor · Encontro 3",
     "Comandos - Servidores 3 - Login, edicao e entrega.html"),
]

estilo = (BASE / "estilo.html").read_text(encoding="utf-8")


def codigo(m):
    return html.escape((BASE / m.group(1)).read_text(encoding="utf-8").rstrip("\n"), quote=False)


for corpo, titulo, destino in PAGINAS:
    texto = (BASE / corpo).read_text(encoding="utf-8")
    texto = re.sub(r"\{\{code:([^}]+)\}\}", codigo, texto)
    assert "{{" not in texto, corpo
    (SAIDA / destino).write_text(estilo.replace("{{titulo}}", titulo) + texto, encoding="utf-8")
    print("gerado:", destino)
