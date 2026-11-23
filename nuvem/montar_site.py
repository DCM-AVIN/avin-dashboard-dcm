r"""
Monta o site do Dashboard DCM na nuvem (GitHub Actions), sem depender do PC de ninguém.

1. Entra no Microsoft 365 pela identidade federada do GitHub (OIDC; não há senha nem
   segredo guardado): o app registrado pela TI confia neste repositório.
2. Baixa do SharePoint (site avin-capital, pasta "AVIN Capital - Time") só o que o
   painel lê: relatórios de monitoramento, planilha da Distribuição, Informes Mensais,
   o próprio código do painel (Iniciar Dashboard.py), vendor\ e os dados de Apoio.
   Mantém a data de modificação de cada arquivo (o painel mostra essas datas).
3. Roda as verificações agendadas do Fundos.NET; se o resultado mudou, devolve
   atualizacoes_fnet.json ao SharePoint.
4. Monta o HTML com o mesmo código do painel (montar_html, modo publicado), CIFRA com a
   chave do segredo AVIN_SITE_CHAVE (html_protegido: tela de senha + dados em AES-256-GCM)
   e grava site/index.html, que a rotina publica no GitHub Pages. Sem o segredo, não publica.
5. Grava nuvem_ativa.json no SharePoint: o PC do editor vê que a nuvem está ativa e
   deixa a consulta ao Fundos.NET com ela.

Teste local, sem Microsoft 365:  python -I montar_site.py --local "<pasta OneDrive - AVIN>"
"""
import datetime
import importlib.util
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

SITE_SP = 'acquaveroagenteautonomo.sharepoint.com:/sites/avin-capital'
RAIZ_SP = 'AVIN Capital - Time'                       # pasta da biblioteca "Documentos"
PASTA_LOCAL = 'Avin - Capital - AVIN Capital - Time'  # nome da pasta no OneDrive (o painel espera este)
GESTAO = '1. DCM/6 - Gestão'
DASH = GESTAO + '/Dashboard Gestão'
APOIO = DASH + '/Apoio - Scripts, Dados e Documentação'
GRAPH = 'https://graph.microsoft.com/v1.0'

AQUI = Path(__file__).resolve().parent
SAIDA = AQUI.parent / 'site'


def log(msg):
    print(time.strftime('%d/%m/%Y %H:%M:%S'), msg, flush=True)


# ── Microsoft 365 ────────────────────────────────────────────────────────────
def _json(url, dados=None, cab=None, metodo=None):
    req = urllib.request.Request(url, data=dados, headers=cab or {}, method=metodo)
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode('utf-8'))


def token_graph():
    """Troca o token OIDC do GitHub por um token do Microsoft Graph (sem segredo)."""
    url_oidc = os.environ['ACTIONS_ID_TOKEN_REQUEST_URL'] + '&audience=api://AzureADTokenExchange'
    oidc = _json(url_oidc, cab={'Authorization': 'Bearer ' + os.environ['ACTIONS_ID_TOKEN_REQUEST_TOKEN']})
    corpo = urllib.parse.urlencode({
        'client_id': os.environ['AZURE_CLIENT_ID'],
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials',
        'client_assertion_type': 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
        'client_assertion': oidc['value'],
    }).encode()
    t = _json(f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}/oauth2/v2.0/token", corpo,
              {'Content-Type': 'application/x-www-form-urlencoded'})
    return t['access_token']


class SharePoint:
    def __init__(self, token):
        self.cab = {'Authorization': 'Bearer ' + token}
        site = _json(f'{GRAPH}/sites/{SITE_SP}', cab=self.cab)
        self.drive = _json(f"{GRAPH}/sites/{site['id']}/drive", cab=self.cab)['id']

    def _url(self, caminho, sufixo):
        return f'{GRAPH}/drives/{self.drive}/root:/' + urllib.parse.quote(f'{RAIZ_SP}/{caminho}') + ':' + sufixo

    def listar(self, caminho):
        """Itens da pasta (arquivos e subpastas), com nome, data e link de download."""
        url, itens = self._url(caminho, '/children?$top=999'), []
        while url:
            d = _json(url, cab=self.cab)
            itens += d.get('value', [])
            url = d.get('@odata.nextLink')
        return itens

    def baixar(self, item, destino):
        destino.parent.mkdir(parents=True, exist_ok=True)
        with urllib.request.urlopen(item['@microsoft.graph.downloadUrl'], timeout=300) as r:
            destino.write_bytes(r.read())
        mod = datetime.datetime.fromisoformat(item['lastModifiedDateTime'].replace('Z', '+00:00')).timestamp()
        os.utime(destino, (mod, mod))

    def enviar(self, caminho, conteudo):
        _json(self._url(caminho, '/content'), conteudo, dict(self.cab, **{'Content-Type': 'application/octet-stream'}),
              'PUT')


def baixar_tudo(sp, base):
    """Copia do SharePoint para base/<PASTA_LOCAL>/… o que o painel lê."""
    n = 0

    def pasta(rel, filtro, sub=False):
        nonlocal n
        for it in sp.listar(rel):
            if 'file' in it and filtro(it['name']) and not it['name'].startswith(('~', '.')):
                sp.baixar(it, base / rel / it['name'])
                n += 1
            elif 'folder' in it and sub:
                pasta(f"{rel}/{it['name']}", filtro, sub)

    pasta(GESTAO + '/02 - Relatórios/Monitoramento', lambda s: s.lower().endswith('.xlsx'))
    pasta(GESTAO + '/10 - Distribuição', lambda s: s.lower().endswith('.xlsx'))
    # Informes Mensais: só o mais recente de cada pasta de "More" (o painel usa só ele)
    for sub in [i for i in sp.listar(GESTAO + '/More') if 'folder' in i]:
        rel = f"{GESTAO}/More/{sub['name']}"
        inf = [i for i in sp.listar(rel) if 'file' in i and i['name'].lower().endswith('.pptx')
               and 'informe mensal' in i['name'].lower() and not i['name'].startswith('~')]
        if inf:
            it = max(inf, key=lambda i: i['lastModifiedDateTime'])
            sp.baixar(it, base / rel / it['name'])
            n += 1
    pasta(DASH, lambda s: s == 'Iniciar Dashboard.py')
    pasta(DASH + '/vendor', lambda s: True)
    pasta(APOIO, lambda s: s.lower().endswith(('.json', '.png')))
    return n


# ── Montagem ─────────────────────────────────────────────────────────────────
def carregar_painel(raiz_onedrive):
    os.environ['AVIN_ONEDRIVE'] = str(raiz_onedrive)
    os.environ['AVIN_NUVEM'] = '1'
    arq = raiz_onedrive / PASTA_LOCAL / DASH / 'Iniciar Dashboard.py'
    spec = importlib.util.spec_from_file_location('dash_avin', arq)
    dash = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(dash)          # não sobe o servidor: o app só roda em __main__
    return dash


def main():
    local = '--local' in sys.argv
    if local:
        raiz = Path(sys.argv[sys.argv.index('--local') + 1])
        sp = None
    else:
        raiz = Path(os.environ.get('RUNNER_TEMP', AQUI)) / 'onedrive'
        sp = SharePoint(token_graph())
        log(f'SharePoint: {baixar_tudo(sp, raiz / PASTA_LOCAL)} arquivos baixados')

    dash = carregar_painel(raiz)
    fnet = Path(dash._ARQ_FNET)
    antes = fnet.read_bytes() if fnet.exists() else b''
    if not local:
        try:
            dash.verificar_fnet_pendentes(registrar=log)
        except Exception as e:
            log(f'[AVISO] Fundos.NET: {type(e).__name__}: {e} — tenta de novo na próxima execução')
        if fnet.exists() and fnet.read_bytes() != antes:
            sp.enviar(f'{APOIO}/{fnet.name}', fnet.read_bytes())
            log('Fundos.NET: atualizacoes_fnet.json devolvido ao SharePoint')

    ops = dash.get_data(force=True)
    falha = [o for o in ops if o.get('erro')]
    log(f'{len(ops) - len(falha)} operações lidas')
    if falha:
        # Mesma regra do PC: com operação faltando, o site fica na versão anterior
        for o in falha:
            log(f"[ERRO] não lido: {o.get('arquivo')} — {o.get('erro')}")
        sys.exit(1)

    # Site protegido por senha: só a versão cifrada é publicada (segredo AVIN_SITE_CHAVE)
    if not dash.chave_site():
        log('[ERRO] segredo AVIN_SITE_CHAVE ausente ou inválido — o site não é publicado sem senha')
        sys.exit(1)
    SAIDA.mkdir(exist_ok=True)
    (SAIDA / 'index.html').write_text(dash.html_protegido(dash.montar_html(estatico=True, logo_embutido=True)),
                                      encoding='utf-8')
    (SAIDA / '.nojekyll').write_text('', encoding='utf-8')
    log(f"site/index.html gravado ({(SAIDA / 'index.html').stat().st_size / 1024:,.0f} KB)")

    if not local:
        marca = {'ultima_publicacao': datetime.datetime.now().isoformat(timespec='seconds'),
                 'repositorio': os.environ.get('GITHUB_REPOSITORY'),
                 'execucao': f"{os.environ.get('GITHUB_SERVER_URL')}/{os.environ.get('GITHUB_REPOSITORY')}"
                             f"/actions/runs/{os.environ.get('GITHUB_RUN_ID')}"}
        sp.enviar(f'{APOIO}/nuvem_ativa.json', json.dumps(marca, ensure_ascii=False, indent=1).encode('utf-8'))


if __name__ == '__main__':
    main()
