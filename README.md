# 🏥 Analisador de Lista de Espera
## Agrupamento Inteligente de Procedimentos (OCI e Genérico)

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/otavioaugust1/WEB_analisador_lista_espera_OCI_o/blob/main/LICENSE)
![Python](https://img.shields.io/badge/python-v3.9+-blue.svg)
![Flask](https://img.shields.io/badge/flask-3.1+-green.svg)
![SQLite](https://img.shields.io/badge/database-SQLite-blue.svg)
![Status](https://img.shields.io/badge/status-Active-brightgreen.svg)

**Análise inteligente de filas de espera com agrupamento dinâmico de procedimentos**

[🚀 Início Rápido](#-início-rápido) • [📚 Documentação](#-documentação) • [💻 Uso Desktop](#-uso-desktop) • [🔌 API](#-api-rest) • [🏗️ Estrutura](#-estrutura-do-projeto)

</div>

---

## 📋 O Que É?

O **Analisador de Lista de Espera** é uma solução robusta, simples e escalável para gestores de saúde. Desenvolvido em Python com Flask, foi criado inicialmente para o programa **OCI (Oncologia Clínica Integrada)**, mas funciona para **qualquer agrupamento de procedimentos**.

### 🎯 Cenários de Uso

| Cenário | Descrição |
|---------|-----------|
| **OCI (Padrão)** | 28 agrupamentos de oncologia pré-carregados no banco |
| **Cardiologia** | Agrupe pacientes por tipos de procedimentos cardíacos |
| **Ortopedia** | Identifique pacientes que precisam de sequências de procedimentos |
| **Qualquer Especialidade** | Customize grupos de procedimentos conforme necessário |

**Em resumo**: Se você tem uma lista de espera e quer identificar quais pacientes podem ser agrupados por um conjunto de procedimentos (obrigatórios, facultativos ou ambos), este sistema faz isso automaticamente.

---

## ✨ Principais Características

### 🔍 Análise Inteligente
- **Detecta automaticamente** pacientes que se enquadram em cada agrupamento
- Processa arquivos CSV/XLSX com até **500MB**
- Resultado: quantos pacientes cabem em cada grupo, com relatórios detalhados

### ⚙️ Gerenciamento Completo via Interface
- **CRUD visual** para agrupamentos e procedimentos (não precisa editar código!)
- Adicione/edite/delete grupos e procedimentos pelo navegador
- Exporte configurações em JSON/CSV para backup ou integração

### 💾 Banco de Dados Simples
- **SQLite**: sem servidor, sem instalações complexas, sem licenças
- **Carga Inicial**: vem pré-carregado com 28 agrupamentos OCI
- **Customizável**: edite diretamente pela interface ou via Python
- **Portátil**: arquivo `db/agrupamentos.db` contém tudo (backup fácil!)

### 🖥️ Duas Versões
| Versão | Como Usar | Para Quem |
|--------|-----------|----------|
| **Web** | Navegador (localhost:5000) | Equipes, acesso remoto, múltiplos usuários |
| **Desktop** | Aplicativo nativo (via PyWebView) | Uso local, distribuição como .exe |

### 📊 Relatórios Profissionais
- PDF com detalhes completos
- XLSX com cálculos e formatação
- JSON para integração com sistemas
- CSV para análise em Excel

### 🛡️ Validação e Segurança
- Verifica estrutura do arquivo automaticamente
- Detecta colunas obrigatórias faltando
- Limite de tamanho configurável
- Processamento por chunks (evita travar com arquivos grandes)

### 📱 Interface Moderna
- Responsiva (funciona em desktop, tablet, mobile)
- Dark-mode suportado
- Acessível
- Feedback visual com alertas em tempo real

---

## 🚀 Início Rápido

### Pré-requisitos
- Python 3.9+
- pip
- ~200MB de espaço em disco

### Instalação Web (Padrão)

```bash
# 1. Clone o repositório
git clone https://github.com/otavioaugust1/WEB_analisador_lista_espera_OCI_o.git
cd WEB_analisador_lista_espera_OCI_o

# 2. Crie ambiente virtual
python3 -m venv venv
source venv/bin/activate     # Linux/macOS
# ou
venv\Scripts\activate         # Windows

# 3. Instale dependências
pip install -r requirements.txt

# 4. Carregue os dados iniciais
python migrate_agrupamentos.py

# 5. Inicie a aplicação
python app.py
```

Acessar em: **http://localhost:5000**

### Instalação Desktop (Windows/Linux/macOS)

```bash
# Após os passos 1-3 acima:

# 4. Instale PyWebView
pip install pywebview

# 5. Execute como aplicação nativa
python app_desktop.py
```

**Gerar .exe (Windows)**:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed app_desktop.py
# Arquivo gerado em: dist/app_desktop.exe
```

---

## 🗂️ Estrutura do Projeto

```
WEB_analisador_lista_espera_OCI_o/
├── app.py                           # Aplicação Flask principal
├── app_desktop.py                   # Versão desktop (PyWebView)
├── database.py                      # Funções CRUD para banco de dados
├── migrate_agrupamentos.py          # Script de carga inicial
├── requirements.txt                 # Dependências Python
├── README.md                        # Este arquivo
├── LICENSE                          # Licença MIT
│
├── db/                              # Dados
│   ├── agrupamentos.db              # SQLite (28 agrupamentos OCI + seus procedimentos)
│   └── arquivo_modelo.xlsx          # Modelo de arquivo para upload
│
├── static/                          # Frontend
│   ├── css/
│   │   ├── style.css                # Estilos principais e index.html
│   │   ├── admin-shared.css         # Estilos compartilhados dos painéis admin
│   │   ├── admin_agrupamentos.css   # Estilos do painel de agrupamentos
│   │   └── admin_visualizar.css     # Estilos da página de detalhes
│   ├── js/
│   │   ├── script.js                # Lógica principal (upload, análise, dark-mode)
│   │   ├── admin_agrupamentos.js    # CRUD de agrupamentos
│   │   └── admin_visualizar.js      # CRUD de procedimentos
│   └── img/
│       └── agora-tem-especialistas.png          # Logo
│
├── templates/                       # Templates HTML (Jinja2)
│   ├── index.html                   # Página principal (análise de arquivo)
│   ├── admin_agrupamentos.html      # Painel: listar/criar/editar agrupamentos
│   └── admin_visualizar.html        # Painel: detalhes com procedimentos
│
└── uploads/                         # Diretório temporário para arquivos enviados
```

---

## 💾 Banco de Dados: Carga Inicial e SQLite

### Por que SQLite?

A escolha do **SQLite** foi proposital:
- ✅ **Zero configuração**: não precisa de servidor PostgreSQL, MySQL, etc.
- ✅ **Portátil**: arquivo único (`agrupamentos.db`), fácil de backup
- ✅ **Simplicidade**: ideal para gestores, não programadores
- ✅ **Performance**: suficiente para lista de espera com milhares de registros
- ✅ **Sem licenças**: código aberto, licença pública

### Carga Inicial

O arquivo `migrate_agrupamentos.py` carrega **28 agrupamentos OCI** pré-configurados:
- 3 agrupamentos de câncer de mama
- 3 de colo de útero
- 1 de próstata
- 1 gástrico
- 1 colorretal
- 5 de cardiologia
- 3 de ortopedia
- 3 de otolaringologia
- 8 de oftalmologia

**Cada agrupamento contém**:
- Código SIGTAP
- Nome descritivo
- Procedimentos **obrigatórios** (que o paciente precisa passar)
- Procedimentos **facultativos** (complementares)

### Customizar a Carga Inicial

Edite `migrate_agrupamentos.py` para alterar, adicionar ou remover agrupamentos **antes** da primeira execução.

Ou, após iniciado, use a interface para:
1. Acessar "Gerenciar Agrupamentos"
2. Criar novos grupos
3. Adicionar/editar procedimentos
4. Exportar em JSON para backup

---

## 🖥️ Versão Desktop (app_desktop.py)

A versão desktop usa **PyWebView** para executar a aplicação como uma janela nativa do sistema operacional.

### Características
- Mesma interface web, mas sem navegador visível
- Janela redimensionável (1400x900px padrão)
- Inicia servidor Flask internamente
- Indicado para uso local em postos de saúde

### Como Usar

```bash
python app_desktop.py
```

A aplicação:
1. Cria pastas necessárias (`db/`, `uploads/`)
2. Inicia servidor Flask (porta 5000)
3. Abre janela PyWebView apontando para localhost:5000

### Distribuir como .exe

Para enviar para colega sem Python instalado:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=logo.ico app_desktop.py
```

Resultado: `dist/app_desktop.exe` (executável único, ~100MB)

---

## 📊 Formato do Arquivo de Entrada

O arquivo CSV ou XLSX deve conter estas colunas **obrigatórias**:

| Coluna | Tipo | Exemplo | Descrição |
|--------|------|---------|-----------|
| `IDENTIFICADOR_LOCAL` | Texto | "001" | ID único do paciente no sistema local |
| `DOCUMENTO_PACIENTE` | Texto | "12345678901" | CPF ou CNS |
| `DATA_SOLICITACAO` | Data | "2026-01-15" | Quando o procedimento foi solicitado |
| `CNES_SOLICITANTE` | Texto | "1234567" | CNES da unidade que solicitou |
| `CNES_REGULADOR` | Texto | "7654321" | CNES do órgão regulador |
| `CODIGO_SIGTAP` | Texto | "0401010010" | Código do procedimento SIGTAP |
| `CBO` | Texto | "2251" | Código da ocupação (profissional) |
| `CID10` | Texto | "C50" | Diagnóstico (CID-10) |
| `CODIGO_MODALIDADE_ASSISTENCIAL` | Texto | "1" | Tipo de atendimento |
| `CODIGO_CARTER_SOLICITACAO` | Texto | "05" | Carter (SUS/particular/etc) |
| `STATUS` | Texto | "1" | "1" = em espera, "0" = processado |
| `DATA_AUTORIZACAO` | Data | "2026-02-10" | Data de aprovação (se houver) |
| `DATA_EXECUCAO` | Data | "2026-03-01" | Data que foi realizado (se houver) |
| `CNES_EXECUTANTE` | Texto | "5555555" | CNES da unidade executora |

**Delimitadores**:
- CSV: `;` (ponto e vírgula)
- XLSX: padrão Excel

Um modelo está em `db/arquivo_modelo.xlsx`.

---

## ⚙️ Gerenciamento de Agrupamentos

### Via Interface (Recomendado)

#### Painel de Agrupamentos (`/admin/agrupamentos`)

| Ação | Como Fazer |
|------|-----------|
| **Listar** | Carrega automaticamente ao acessar |
| **Criar** | Botão "Novo" → preenche form modal |
| **Editar** | Clica no código/nome → edita no modal |
| **Deletar** | Ícone 🗑️ na linha → confirma exclusão |
| **Exportar JSON** | Botão "JSON" → baixa `agrupamentos.json` |
| **Exportar CSV** | Botão "CSV" → baixa `agrupamentos.csv` |

#### Página de Detalhes (`/admin/visualizar/<codigo>`)

Mostra um agrupamento com dois conjuntos de procedimentos:

**Procedimentos Obrigatórios**:
- Sequência que o paciente DEVE passar
- Ordem importa (cada um é um passo)
- Clique "+" para adicionar procedimento
- Edite/delete com ícones na linha

**Procedimentos Facultativos**:
- Procedimentos complementares (opcionais)
- Influenciam na classificação do paciente
- Mesma interface de add/edit/delete

---

## 🔌 API REST

Endpoints para integração com sistemas externos:

### Agrupamentos

```bash
# Listar todos
GET /api/agrupamentos

# Criar novo
POST /api/agrupamentos
{
  "codigo": "0999999999",
  "nome": "MEU AGRUPAMENTO",
  "descricao": "Descrição opcional"
}

# Atualizar
PUT /api/agrupamentos/<codigo>
{
  "nome": "NOME ATUALIZADO",
  "descricao": "Descrição nova"
}

# Deletar
DELETE /api/agrupamentos/<codigo>
```

### Procedimentos

```bash
# Adicionar obrigatório
POST /api/agrupamentos/<id>/procedimentos-obrigatorios
{
  "codigo": "0401010010",
  "descricao": "Consulta Oncológica",
  "ordem": 1
}

# Adicionar facultativo
POST /api/agrupamentos/<id>/procedimentos-facultativos
{
  "codigo": "0401010020",
  "descricao": "Exame de Imagem",
  "ordem": 1
}

# Atualizar obrigatório
PUT /api/procedimentos-obrigatorios/<procedimento_id>
{
  "codigo": "0401010010",
  "descricao": "Consulta Atualizada"
}

# Deletar obrigatório
DELETE /api/procedimentos-obrigatorios/<procedimento_id>

# (Análogas para facultativos)
```

### Análise de Arquivo

```bash
# Upload e análise
POST /analyze_file
Form data: file=<arquivo.csv>

Resposta:
{
  "status": "success",
  "total_registros": 1500,
  "processados": 1498,
  "erros": 2,
  "agrupamentos": {
    "0901010014": 45,
    "0901010090": 33,
    ...
  }
}
```

### Exportação

```bash
# JSON de todos agrupamentos
GET /api/exportar/json

# CSV de agrupamentos
GET /api/exportar/csv

# Relatório PDF
POST /download_pdf
Form: resultado_analise

# Relatório XLSX
POST /download_xlsx
Form: resultado_analise
```

---

## 📈 Performance e Processamento

### Chunks (Processamento por Lotes)

Arquivos grandes (>100MB) são processados em **chunks de 5.000 linhas**:
- Evita travar (falta de memória)
- Suporta até **500MB**
- Tempo típico: ~1-2 segundos por 15.000 linhas

### Tempos Reais de Processamento

| Tamanho | Linhas | Tempo |
|---------|--------|-------|
| 5 MB | ~15.000 | 1-2s |
| 50 MB | ~150.000 | 8-12s |
| 100 MB | ~300.000 | 20-30s |
| 500 MB | ~1.500.000 | 3-5 min |

### Otimizações

**v1.0 → v2.1**:
- v1.0: agrupamentos em código (1.5MB de linhas)
- v2.1: SQLite + banco (100KB código + 40KB database)
- Resultado: mais flexível, menos fichário, mais fácil distribuir

---

## 🔄 Fluxo de Uso Típico

```
1. Gestor acessa o sistema (web ou desktop)
   ↓
2. Configura agrupamentos (OCI pré-carregado ou customizado)
   ↓
3. Exporta sua lista de espera em CSV/XLSX
   ↓
4. Faz upload do arquivo no sistema
   ↓
5. Sistema analisa automaticamente
   ↓
6. Relatório mostra: "45 pacientes no agrupamento X, 33 no Y..."
   ↓
7. Exporta em PDF/XLSX para apresentação ou integração
```

---

## 🛠️ Configuração Avançada

### Modificar Limite de Upload

Em `app.py`:
```python
MAX_FILE_SIZE = 500 * 1024 * 1024  # 500MB
CHUNK_SIZE = 5000                   # linhas por chunk
```

### Adicionar Agrupamento via Python

```python
import database

database.init_database()

# Criar agrupamento
agrup = database.adicionar_agrupamento(
    codigo='9999999999',
    nome='MEU AGRUPAMENTO',
    descricao='Descrição',
    itens_obrigatorios=[
        {'codigo': '0401010010', 'descricao': 'Consulta'},
        {'codigo': '0401010020', 'descricao': 'Exame'}
    ],
    itens_facultativos=[
        {'codigo': '0401010030', 'descricao': 'Complemento'}
    ]
)
```

### Resetar Banco para Defaults

```bash
rm db/agrupamentos.db
python migrate_agrupamentos.py
```

---

## 🚀 Deployment Avançado

### Produção com Gunicorn

```bash
pip install gunicorn
gunicorn --workers 4 --bind 0.0.0.0:5000 app:app
```

### Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
RUN python migrate_agrupamentos.py
EXPOSE 5000
CMD ["gunicorn", "--workers=4", "--bind=0.0.0.0:5000", "app:app"]
```

```bash
docker build -t oci-analyzer .
docker run -p 5000:5000 oci-analyzer
```

---

## 🆕 Novidades (v2.1+)

### ✅ Gerenciamento Completo
- Painel visual de CRUD para agrupamentos
- Gerenciamento de procedimentos (obrigatórios e facultativos)
- Modais com validação em tempo real

### ✅ API REST Expandida
- 13 novos endpoints para integração
- Exportação JSON/CSV de configurações
- Suporte a análise programática

### ✅ Interface Desktop
- Versão PyWebView para uso local
- Suporte a geração de .exe
- Mesma interface, sem browser visível

### ✅ Estrutura Modular
- `database.py`: camada de dados isolada
- `app.py`: endpoints Flask
- `app_desktop.py`: wrapper PyWebView
- CSS/JS separado por funcionalidade

### ✅ Melhorias Visuais
- Dark-mode suportado
- Responsividade aprimorada
- Acessibilidade melhorada
- Ícones FontAwesome
- Alertas e feedback visual

---

## 🐛 Troubleshooting

### Erro: "Tipo de arquivo não permitido"
- Verifique se é `.csv` ou `.xlsx`
- Limite máximo: 500MB

### Erro: "Colunas obrigatórias faltando"
- Verifique nomes exatos (case-sensitive)
- Use modelo em `db/arquivo_modelo.xlsx`

### Erro: "Banco SQLite corrompido"
- Delete `db/agrupamentos.db`
- Execute `python migrate_agrupamentos.py`

### Procedimentos não aparecem
- Verifique pela API: `GET /api/agrupamentos`
- Confirme que foram adicionados

### PyWebView não abre janela
- Verifique instalação: `pip list | grep pywebview`
- Reinstale: `pip install --upgrade pywebview`

---

## 📊 Recursos Adicionais

- **Modelo de Arquivo**: `db/arquivo_modelo.xlsx`
- **Banco Inicial**: `db/agrupamentos.db` (28 agrupamentos OCI)
- **CSS Responsivo**: Funciona em mobile, tablet e desktop
- **Exportação**: PDF, XLSX, JSON, CSV

---

## 📄 Licença

**MIT License** — Use livremente, comercial ou pessoal. Veja [LICENSE](LICENSE).

---

## 👤 Autor

**Otávio August**
- 🌐 GitHub: [@otavioaugust1](https://github.com/otavioaugust1)
- 🇧🇷 Brasil

---

## 🔄 Changelog

### v2.1.0 (Março 2026) ⚙️ GERENCIAMENTO COMPLETO
- ✅ Painel CRUD para agrupamentos
- ✅ Gerenciamento visual de procedimentos
- ✅ 13 novos endpoints REST
- ✅ Exportação JSON/CSV
- ✅ Versão desktop (PyWebView)
- ✅ Suporte a .exe
- ✅ Interface modular (CSS/JS por feature)
- ✅ Dark-mode e acessibilidade
- ✅ Documentação revisada

### v2.0.0 (Abril 2025) ✨ REFATORAÇÃO
- ✅ Migração para SQLite
- ✅ Processamento em chunks (até 500MB)
- ✅ Banco de dados gerenciável
- ✅ API abstração em database.py
- ✅ 28 agrupamentos OCI pré-carregados

### v1.0.0 (2024)
- ✅ Análise de listas de espera
- ✅ Relatórios PDF/XLSX
- ✅ Agrupamentos hardcoded

---

<div align="center">

⭐ Se foi útil, deixe uma estrela no GitHub!

**[Ir para o repositório](https://github.com/otavioaugust1/WEB_analisador_lista_espera_OCI_o)**

</div>
