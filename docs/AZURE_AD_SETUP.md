# Setup do Azure AD para o GitHub Action de sync com SharePoint

Este passo a passo cria um **App Registration** no Azure Active Directory do escritório que dá ao GitHub Actions permissão de **leitura** nas pastas do SharePoint (`inteligencia jurídica BMG` e `Central de comando - Nepomuceno Soares`).

**Quem precisa fazer:** o Admin do tenant Azure AD do `escritorionepomuceno.sharepoint.com` (você ou alguém com privilégio de Application Administrator / Global Administrator).

**Quanto tempo:** ~10 minutos.

**Custo:** R$ 0,00 — Azure AD App Registration é gratuito.

---

## Passo 1 · Criar o App Registration

1. Acesse https://portal.azure.com e faça login com sua conta corporativa
2. Procure por **"Azure Active Directory"** ou **"Microsoft Entra ID"** na barra de busca
3. No menu lateral esquerdo, clique em **"App registrations"**
4. Clique no botão **"+ New registration"** no topo
5. Preencha:
   - **Name:** `BMG Sites Sync`
   - **Supported account types:** *Accounts in this organizational directory only (Nepomuceno only - Single tenant)*
   - **Redirect URI:** deixar em branco
6. Clique em **"Register"**

Após criar, a página do app vai mostrar três campos importantes — **anote os dois primeiros**:
- **Application (client) ID** → será o secret `AZURE_CLIENT_ID` no GitHub
- **Directory (tenant) ID** → será o secret `AZURE_TENANT_ID` no GitHub

---

## Passo 2 · Gerar o Client Secret

1. No menu lateral do app, clique em **"Certificates & secrets"**
2. Aba **"Client secrets"** → botão **"+ New client secret"**
3. Preencha:
   - **Description:** `GitHub Actions sync`
   - **Expires:** *24 months* (depois precisa renovar)
4. Clique em **"Add"**
5. **IMPORTANTE:** copie o campo **"Value"** AGORA — depois que sair da tela, esse campo nunca mais aparece. Esse valor será o secret `AZURE_CLIENT_SECRET` no GitHub.

⚠️ Se você sair sem copiar, precisa apagar o secret e criar outro.

---

## Passo 3 · Conceder permissões de leitura no SharePoint

1. No menu lateral do app, clique em **"API permissions"**
2. Clique em **"+ Add a permission"**
3. Escolha **"Microsoft Graph"**
4. Escolha **"Application permissions"** (não Delegated)
5. Na barra de busca, digite `Sites.Read.All` → marque a caixa
6. Clique em **"Add permissions"**

A permissão vai aparecer na lista com status amarelo **"Not granted"**.

7. Clique no botão **"Grant admin consent for [Nepomuceno]"** acima da tabela
8. Confirme. O status vira verde **"Granted"**.

✅ Pronto — o app agora pode ler arquivos do SharePoint do escritório.

---

## Passo 4 · Adicionar os 3 secrets no GitHub

Faça em **cada repo** que tem o workflow (Site 1 BMG, depois Central):

1. Acesse o repo no GitHub:
   - Site 1: https://github.com/RODSOARES88/nepomucenoaores-bmg
   - Central: https://github.com/RODSOARES88/painel-interno
2. Vá em **Settings** → **Secrets and variables** → **Actions**
3. Clique em **"New repository secret"** e crie os 3:

| Nome do secret | Valor |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID do passo 1 |
| `AZURE_CLIENT_ID` | Application (client) ID do passo 1 |
| `AZURE_CLIENT_SECRET` | Value do client secret do passo 2 |

Cada um é uma "string" que você cola no campo "Value" do GitHub.

---

## Passo 5 · Testar manualmente

1. No repo do GitHub, vá em **Actions** (aba do topo)
2. No menu lateral, clique no workflow **"SharePoint Audit"**
3. Clique em **"Run workflow"** → escolha branch `main` → **"Run workflow"**

Em ~30 segundos o workflow termina. Se deu certo:
- Sem arquivos novos: log diz `✓ Nada novo · pasta sincronizada`
- Com arquivos novos: uma **Issue** é criada no repo listando os arquivos pendentes

Se deu erro de autenticação: provavelmente faltou o "Grant admin consent" do passo 3.

---

## Manutenção

- O **client secret expira em 24 meses**. O GitHub Actions vai começar a falhar quando expirar. Solução: voltar no Passo 2, criar um novo secret e substituir o valor de `AZURE_CLIENT_SECRET` no GitHub.
- Você pode revogar o acesso a qualquer momento deletando o App Registration no Azure AD.
- O app **só lê** — não modifica, não deleta, não cria. Permissão `Sites.Read.All` é read-only.

---

## Troubleshooting

| Sintoma | Causa provável | Solução |
|---|---|---|
| `AADSTS70011: The provided value for the input parameter 'scope' is not valid` | Permissão não foi concedida via admin consent | Passo 3, clicar em "Grant admin consent" |
| `AADSTS7000222: The provided client secret keys for app ... are expired` | Client secret expirou | Gerar novo no Passo 2 e atualizar `AZURE_CLIENT_SECRET` no GitHub |
| `403 Forbidden` ao acessar a pasta | App não tem permissão no site específico | Verificar se `Sites.Read.All` (não `Sites.Selected`) foi a permissão escolhida |
| Workflow nunca abre issue mesmo com arquivo novo | `.sync-manifest.json` está com o arquivo já listado | Apagar o arquivo ou remover a linha correspondente |
