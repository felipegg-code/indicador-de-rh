# Dashboard RH — TechSolutions

Dashboard automático de banco de horas, intrajornada e interjornada.  
Atualiza sozinho toda vez que você salvar os arquivos Excel na pasta `dados/`.

---

## 🚀 Como configurar (uma única vez)

### 1. Criar o repositório no GitHub

1. Acesse [github.com](https://github.com) e faça login
2. Clique em **New repository**
3. Nome: `dashboard-rh` (ou qualquer nome)
4. Marque **Private** se quiser manter os dados restritos
5. Clique em **Create repository**

---

### 2. Subir os arquivos

Você pode fazer isso de duas formas:

**Opção A — pelo site do GitHub (sem instalar nada):**
1. Na página do repositório, clique em **Add file → Upload files**
2. Suba toda a pasta do projeto (arraste os arquivos)

**Opção B — pelo terminal (se tiver Git instalado):**
```bash
cd dashboard-rh
git init
git remote add origin https://github.com/SEU_USUARIO/dashboard-rh.git
git add .
git commit -m "Primeiro commit"
git push -u origin main
```

---

### 3. Ativar o GitHub Pages

1. No repositório, clique em **Settings**
2. No menu lateral, clique em **Pages**
3. Em **Source**, selecione **Deploy from a branch**
4. Branch: `main` · Pasta: `/docs`
5. Clique em **Save**

Aguarde 1-2 minutos. Seu dashboard estará disponível em:
```
https://SEU_USUARIO.github.io/dashboard-rh
```

---

### 4. Ativar as Actions (automação)

1. No repositório, clique na aba **Actions**
2. Se aparecer uma mensagem pedindo para ativar, clique em **I understand my workflows, go ahead and enable them**

Pronto! A automação está configurada.

---

## 📂 Estrutura de arquivos

```
dashboard-rh/
│
├── dados/                          ← COLOQUE SEUS ARQUIVOS AQUI
│   ├── BANCO_DE_HORAS_58_ANALISADO.xlsx
│   └── TRATAMENTO_PONTO_GERAL.xlsx
│
├── scripts/
│   └── gerar_dashboard.py          ← script Python (não mexa)
│
├── docs/
│   └── index.html                  ← dashboard gerado automaticamente
│
├── .github/
│   └── workflows/
│       └── deploy.yml              ← automação GitHub Actions
│
└── README.md
```

---

## 🔄 Como atualizar o dashboard

**Todo mês (ou quando quiser):**

1. Acesse o repositório no GitHub
2. Entre na pasta `dados/`
3. Clique no arquivo → **pencil icon** → ou arraste o novo arquivo sobre o antigo
4. Confirme o upload clicando em **Commit changes**

O GitHub Actions vai:
- Detectar que o arquivo mudou
- Rodar o script Python automaticamente
- Gerar o novo `docs/index.html`
- Publicar no GitHub Pages

Em **~2 minutos** o dashboard estará atualizado no link público.

---

## ✅ Arquivos necessários

| Arquivo | Aba usada | Colunas obrigatórias |
|---------|-----------|----------------------|
| `BANCO_DE_HORAS_58_ANALISADO.xlsx` | `GERAL` | NOME, SECAO, FUNCAO, TOTALGERAL, FAIXA_BANCO_HORAS, ACAO_BANCO_HORAS |
| `TRATAMENTO_PONTO_GERAL.xlsx` | `EXCECOES_JORNADA` | DATA, TIPO_OCORRENCIA, DETALHE_OCORRENCIA, NOME, SECAO, BATIDA1–4 |

---

## 🛠️ Rodar localmente (opcional)

Se quiser testar antes de subir:

```bash
# Instalar dependências (uma vez só)
pip install pandas openpyxl

# Gerar o dashboard
python scripts/gerar_dashboard.py

# Abrir o resultado
open docs/index.html   # Mac
start docs/index.html  # Windows
```

---

## ❓ Dúvidas comuns

**O dashboard não atualizou após o upload.**
→ Verifique a aba **Actions** — se aparecer um ícone vermelho, clique nele para ver o erro.

**Erro: arquivo não encontrado.**
→ Confirme que os nomes dos arquivos na pasta `dados/` são exatamente iguais aos listados acima.

**Quero restringir o acesso.**
→ Repositório privado + GitHub Pages só funciona com plano pago. Alternativa: use a versão de upload manual (`dashboard_rh_autoalimentado.html`) e distribua o arquivo por e-mail.
