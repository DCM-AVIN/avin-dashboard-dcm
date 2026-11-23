# Dashboard DCM — AVIN Capital

Publica o Dashboard de Gestão DCM no GitHub Pages, sem depender de nenhum computador.

A rotina `.github/workflows/atualizar-site.yml` roda de hora em hora (e no botão **Run workflow**):

1. entra no Microsoft 365 pela identidade federada do GitHub (sem senha guardada);
2. baixa do SharePoint as planilhas de monitoramento, a posição da Distribuição, os
   Informes Mensais e o código do painel (`Iniciar Dashboard.py`);
3. monta o HTML com o mesmo código do painel, cifra com a senha do time e publica no GitHub Pages.

O código do painel e as planilhas **não** ficam neste repositório: vêm do SharePoint a cada
execução. Este repositório guarda a rotina e o montador (`nuvem/montar_site.py`).

**Protegido por senha:** o HTML publicado vai **cifrado** (AES-256-GCM, chave derivada da senha
do time por PBKDF2). A página só mostra a tela de senha; sem a senha, nada dos dados pode ser lido,
nem no código da página nem neste repositório. Segredo `AVIN_SITE_CHAVE` (a chave, não a senha).

**Modo provisório** (enquanto o acesso ao SharePoint não é liberado): o PC do editor envia o HTML
publicado para `pc/index.html` e a rotina publica esse arquivo.

Configuração: variáveis do repositório `AZURE_TENANT_ID` e `AZURE_CLIENT_ID` (app do
Microsoft 365 com acesso só ao site `avin-capital`). A fonte destes arquivos fica no
OneDrive do time, em `Dashboard Gestão\Apoio - Scripts, Dados e Documentação\GitHub`.
