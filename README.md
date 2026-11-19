# Portal de Dados — extração de queries do MegaXT

Painel web para consultar e exportar relatórios do ERP MegaXT (Oracle).
Cada relatório é uma view do banco: o portal lê as colunas da view, monta
os filtros sozinho e entrega o resultado na tela, em CSV e em XLSX.

## Como funciona

- **Um arquivo só:** `app.py` (Streamlit).
- **Login:** `streamlit-authenticator`, com usuários e hashes de senha no `config.yaml`.
- **Relatórios:** catálogo no `config.yaml` (`reports`), cada um apontando para uma view.
- **Permissões:** `user_reports` no `config.yaml` diz quais relatórios cada usuário vê.
- **Filtros automáticos:** texto = contém, número = igual, data = intervalo.
- **Ordenação:** pela primeira coluna de data, mais recente primeiro (vazias no fim).
- **Exportação:** CSV (`;`, UTF-8 com BOM) e XLSX (com Tabela do Excel), da prévia ou completo.
- **Banco:** `oracledb` em modo *thick* (precisa do Oracle Instant Client).

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `app.py` | O portal inteiro |
| `config.example.yaml` | Modelo do `config.yaml` (usuários, permissões, relatórios) |
| `.env.example` | Modelo do `.env` (usuário e senha do Oracle) |
| `.streamlit/config.toml` | Tema claro |
| `logo.png`, `icon.png` | Logo do cabeçalho e ícone da aba |
| `requirements.txt` | Bibliotecas, nas versões em produção |
| `deploy/portal-dados.service` | Serviço systemd usado na VM |

## O que NÃO vai para o repositório

Está tudo no `.gitignore`:

- `config.yaml` — hashes de senha dos usuários e chave do cookie de login.
- `.env` — usuário e senha do Oracle.
- CSV/XLSX extraídos do ERP.
- Backups (`*.bak-AAAAMMDD`) e cópias antigas do app.

## Configuração

1. Copie `config.example.yaml` para `config.yaml` e preencha usuários, permissões e relatórios.
2. Copie `.env.example` para `.env`, preencha e proteja o arquivo (`chmod 600 .env`).
3. O endereço do banco (DSN) e a pasta do Instant Client estão fixos no `app.py`
   (`get_connection` e `init_oracle_client`). Ajuste se o ambiente for outro.

Gerar o hash de uma senha (a senha é digitada sem aparecer na tela):

```bash
python -c "import bcrypt, getpass; print(bcrypt.hashpw(getpass.getpass('Senha: ').encode(), bcrypt.gensalt(12)).decode())"
```

Gerar a chave do cookie:

```bash
python -c "import secrets; print(secrets.token_hex(16))"
```

## Rodar

```bash
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
set -a; . ./.env; set +a
./venv/bin/streamlit run app.py --server.port 8501
```

## Publicar (VM Linux com systemd)

O app fica em `~/apps/portal-dados/` e roda como o serviço `portal-dados`
(modelo em `deploy/portal-dados.service`), atrás de um Nginx com HTTPS.

1. Faça backup do que está publicado: `cp app.py app.py.bak-AAAAMMDD` (e o mesmo para `config.yaml`).
2. Copie o `app.py` novo (e o `config.yaml`, se mudou) para a pasta do app. **Nunca sobrescreva o `.env`.**
3. Reinicie: `sudo systemctl restart portal-dados`
4. Confira: `systemctl is-active portal-dados` e `journalctl -u portal-dados -n 50`

## Adicionar um relatório

1. Crie a view no Oracle.
2. Em `config.yaml`, adicione uma chave em `reports` com `label` e `view`.
3. Adicione essa chave em `user_reports` para quem deve ver.
4. Publique o `config.yaml` e reinicie o serviço.

Se a view tiver colunas de data sem hora, liste-as em `datas_sem_hora` para
saírem como `DD/MM/AAAA` na tela, no CSV e no XLSX.

## Adicionar um usuário

1. Gere o hash da senha (comando acima).
2. Em `config.yaml`, adicione o usuário em `credentials.usernames` e em `user_reports`.
3. Publique o `config.yaml` e reinicie o serviço.
