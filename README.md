# Analisador de Lista de Espera para OCI (Oracle Cloud Infrastructure)

## Funcionalidades Principais
- Upload de arquivos CSV com dados de lista de espera
- Análise e processamento automático dos dados
- Visualização de gráficos e métricas
- Filtros dinâmicos para análise segmentada
- Exportação de relatórios em múltiplos formatos

## Tecnologias Utilizadas
- **Backend**: Python, Flask, Pandas
- **Frontend**: HTML5, CSS3, Bootstrap, Chart.js
- **Infraestrutura**: Docker, Gunicorn

## Pré-requisitos
Antes de executar o projeto, certifique-se de ter instalado:
Python 3.9 ou superior
pip (gerenciador de pacotes Python)
Docker (opcional, para execução via container)
Git (para clonar o repositório)

## Instalação
### Método 1: Execução Local com Python
 - Clone o repositório:
```bash
git clone https://github.com/otavioaugust1/WEB_analisador_lista_espera_OCI_o.git
cd WEB_analisador_lista_espera_OCI_o
pip install -r requirements.txt
```
 - Crie um ambiente virtual (recomendado):
```bash
python -m venv venv
source venv/bin/activate  # Linux/MacOS
venv\Scripts\activate    # Windows
```
 - Instale as dependências:
```bash
pip install -r requirements.txt
```
 - Execute a aplicação:
```bash
python app.py
```
 - Acesse no navegador:
```bash
http://localhost:5000
```

### Método 2: Execução com Docker
 - Construa a imagem Docker:
```bash
docker build -t analisador-lista-espera .
 ```
 - Execute o container:
```bash
docker run -d -p 5000:5000 --name oci-analisador analisador-lista-espera
```
 - Acesse no navegador:
```bash
http://localhost:5000
```
## Estrutura do Projeto
```bash
├── app.py                # Ponto de entrada principal da aplicação
├── Dockerfile            # Configuração para construção do container Docker
├── requirements.txt      # Dependências do projeto
├── templates/            # Templates HTML
│   ├── index.html        # Página principal
│   ├── results.html      # Página de resultados
│   └── ...               # Outros templates
├── static/               # Arquivos estáticos
│   ├── css/              # Folhas de estilo
│   ├── js/               # Scripts JavaScript
│   └── images/           # Imagens
├── utils/                # Módulos utilitários
│   ├── data_processor.py # Processamento de dados
│   └── visualizations.py # Geração de gráficos
└── .gitignore            # Arquivos ignorados pelo Git
```

## Como Utilizar
1) Acesso Inicial:
    * Acesse http://localhost:5000
    * Selecione "Escolher arquivo" na página inicial

2) Upload de Dados:
    * Selecione um arquivo CSV no formato padrão da OCI
    * Clique em "Enviar" para iniciar o processamento

3) Análise de Resultados:
    * Visualize o dashboard com métricas principais
    * Interaja com os gráficos para detalhamento

4) Utilize os filtros para segmentar os dados:
    * Seletor de data
    * Dropdown de tipos de instância
    * Filtro por região

5) Exportação:
    * Clique em "Exportar Relatório" para gerar PDF
    * Use "Baixar Dados Processados" para CSV

# Formatos de Arquivo Suportados

| Formato | Status               | Observações                                  |
|---------|----------------------|-----------------------------------------------|
| CSV     | ✅ Suportado         | Formato padrão de exportação da OCI           |
| Excel   | 🚧 Em desenvolvimento | Suporte planejado para versão 2.0             |
| JSON    | ⚠️ Futuro            | Suporte para APIs                             |


# Contribuição
Contribuições são bem-vindas! Siga este fluxo:
* Faça um fork do projeto
* Crie uma branch para sua feature (git checkout -b feature/nova-feature)
* Faça commit das alterações (git commit -m 'Adiciona nova funcionalidade')
* Faça push para a branch (git push origin feature/nova-feature)
* Abra um Pull Request

## Roadmap de Melhorias
* Suporte a múltiplos arquivos simultâneos
* Integração com API da OCI
* Painel comparativo entre regiões
* Sistema de alertas por email
* Autenticação de usuários

# Licença
Este projeto está licenciado sob a Licença MIT - consulte o arquivo LICENSE para detalhes.

# Contato
- Otávio Augusto
- GitHub: @otavioaugust1
- Email: otavio.augusto@gmail.com



