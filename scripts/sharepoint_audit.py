"""
SharePoint Audit · varre uma pasta do SharePoint, compara com manifest local
e emite um relatório dos arquivos novos.

Configurado para uso no GitHub Actions. Variáveis de ambiente esperadas:
    AZURE_TENANT_ID     · tenant do escritório
    AZURE_CLIENT_ID     · App Registration "BMG Sites Sync"
    AZURE_CLIENT_SECRET · client secret do app

Argumentos da CLI:
    --folder-uri    URI do tipo file:///{driveId}/{itemId}
                    (encontrado via sharepoint_folder_search)
    --manifest      caminho do manifest local (default .sync-manifest.json)
    --output        caminho do relatório markdown (default audit-report.md)

Saída: cria/atualiza o manifest, escreve o relatório, e termina com exit code:
    0 = nada novo
    1 = arquivos novos detectados (sinaliza pro workflow abrir uma issue)
    2 = erro
"""
import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

import msal
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_AUTHORITY = "https://login.microsoftonline.com/{tenant}"


def get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Autentica via client credentials e retorna access token."""
    authority = TOKEN_AUTHORITY.format(tenant=tenant_id)
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(
            f"Falha ao obter token: {result.get('error')} - {result.get('error_description')}"
        )
    return result["access_token"]


def parse_folder_uri(uri: str) -> tuple[str, str]:
    """Extrai driveId e itemId de uma URI no formato file:///{driveId}/{itemId}."""
    if not uri.startswith("file:///"):
        raise ValueError(f"URI inválida: {uri}")
    parts = uri.removeprefix("file:///").split("/", 1)
    if len(parts) != 2:
        raise ValueError(f"URI mal formada: {uri}")
    return parts[0], parts[1]


def list_folder_children(token: str, drive_id: str, item_id: str) -> list[dict]:
    """Lista os arquivos/subpastas de uma pasta do SharePoint via Graph API."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def load_manifest(path: Path) -> dict:
    """Carrega manifest existente ou retorna estrutura vazia."""
    if not path.exists():
        return {
            "_meta": {
                "criado_em": datetime.utcnow().isoformat() + "Z",
                "descricao": "Lista de arquivos do SharePoint já vistos. Atualizado pelo workflow SharePoint Audit.",
            },
            "arquivos_conhecidos": {},
        }
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_manifest(path: Path, manifest: dict) -> None:
    """Salva manifest atualizado em disco."""
    manifest["_meta"]["ultima_sync"] = datetime.utcnow().isoformat() + "Z"
    with path.open("w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder-uri", required=True, help="URI da pasta SharePoint")
    parser.add_argument("--manifest", default=".sync-manifest.json")
    parser.add_argument("--output", default="audit-report.md")
    parser.add_argument("--folder-label", default="pasta SharePoint", help="Nome humano da pasta")
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("ERRO: faltam variáveis de ambiente AZURE_TENANT_ID/CLIENT_ID/CLIENT_SECRET", file=sys.stderr)
        return 2

    try:
        token = get_graph_token(tenant_id, client_id, client_secret)
    except Exception as e:
        print(f"ERRO ao autenticar: {e}", file=sys.stderr)
        return 2

    drive_id, item_id = parse_folder_uri(args.folder_uri)

    try:
        children = list_folder_children(token, drive_id, item_id)
    except Exception as e:
        print(f"ERRO ao listar pasta: {e}", file=sys.stderr)
        return 2

    # Considera só arquivos (não subpastas)
    arquivos_remotos = {
        c["name"]: {
            "id": c["id"],
            "modified": c.get("lastModifiedDateTime"),
            "size": c.get("size"),
            "webUrl": c.get("webUrl"),
        }
        for c in children
        if "file" in c
    }

    manifest_path = Path(args.manifest)
    manifest = load_manifest(manifest_path)
    conhecidos = manifest.setdefault("arquivos_conhecidos", {})

    novos = []
    atualizados = []
    for nome, meta in arquivos_remotos.items():
        if nome not in conhecidos:
            novos.append((nome, meta))
        elif conhecidos[nome].get("modified") != meta["modified"]:
            atualizados.append((nome, meta, conhecidos[nome].get("modified")))

    # Atualiza manifest
    manifest["arquivos_conhecidos"] = arquivos_remotos
    save_manifest(manifest_path, manifest)

    # Gera relatório
    output_path = Path(args.output)
    lines = [
        f"# Audit SharePoint · {args.folder_label}",
        "",
        f"Snapshot: `{datetime.utcnow().isoformat()}Z`",
        f"Pasta: **{args.folder_label}**",
        f"Total de arquivos visíveis: **{len(arquivos_remotos)}**",
        "",
    ]

    if not novos and not atualizados:
        lines.append("✓ **Nada novo · pasta sincronizada.**")
        output_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"✓ Nada novo · {len(arquivos_remotos)} arquivos conhecidos")
        return 0

    if novos:
        lines.append("## 🆕 Arquivos novos a processar")
        lines.append("")
        lines.append("| Arquivo | Modificado em | Tamanho | Link |")
        lines.append("|---|---|---|---|")
        for nome, meta in novos:
            tam = f"{meta['size']:,} bytes".replace(",", ".") if meta.get("size") else "?"
            link = f"[abrir]({meta['webUrl']})" if meta.get("webUrl") else "—"
            lines.append(f"| `{nome}` | {meta['modified']} | {tam} | {link} |")
        lines.append("")

    if atualizados:
        lines.append("## ♻️ Arquivos com nova versão")
        lines.append("")
        lines.append("| Arquivo | Antes | Agora |")
        lines.append("|---|---|---|")
        for nome, meta, antes in atualizados:
            lines.append(f"| `{nome}` | {antes} | {meta['modified']} |")
        lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("**Próximo passo:** pedir pra Claude processar esses arquivos atualizados.")
    lines.append("")
    lines.append(f"_Workflow: `.github/workflows/sharepoint-audit.yml` · manifest: `{args.manifest}`_")

    output_path.write_text("\n".join(lines), encoding="utf-8")

    n_novos = len(novos)
    n_atual = len(atualizados)
    print(f"⚠ {n_novos} arquivo(s) novo(s) + {n_atual} atualizado(s) · relatório em {output_path}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
