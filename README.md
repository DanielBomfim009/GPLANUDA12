# Gplan — Controle Documental Instrumentação RNEST U12

Dashboard de controle documental de instrumentação, lendo dados diretamente de uma
planilha Excel hospedada no Supabase Storage.

## Como funciona

1. Localmente, você continua atualizando a base de dados como sempre
   (`ATUALIZAR_CONTROLE.cmd` no projeto original).
2. Depois de gerar a planilha atualizada, envie o arquivo
   `CONTROLE_DOCUMENTAL_INSTRUMENTACAO_ATUAL.xlsx` para o bucket `gplan-data`
   no painel do Supabase (Storage → gplan-data → upload, substituindo o
   arquivo existente).
3. O site detecta automaticamente que o arquivo mudou e recarrega os dados
   (sem precisar reiniciar nada).

## Rodando localmente

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edite .streamlit/secrets.toml com sua URL e chave do Supabase
streamlit run gplan_app.py
```

Se `.streamlit/secrets.toml` não existir ou estiver vazio, o app tenta ler a
planilha localmente em
`../Controle de Relatório dos Instrumentos/01_ARQUIVO_ATUAL/CONTROLE_DOCUMENTAL_INSTRUMENTACAO_ATUAL.xlsx`
(útil para testar sem depender do Supabase).

## Deploy (Render + Supabase)

Ver instruções completas no chat/documentação do projeto. Resumo:

1. Criar bucket `gplan-data` no Supabase Storage e enviar a planilha inicial.
2. Criar Web Service no Render conectado a este repositório GitHub.
3. Configurar variáveis de ambiente no Render: `SUPABASE_URL`, `SUPABASE_KEY`.
4. Apontar o domínio próprio nas configurações do Render (DNS no registrador).

## Estrutura

```
gplan-web/
├── gplan_app.py              # aplicação principal
├── requirements.txt          # dependências Python
├── .streamlit/
│   ├── config.toml           # tema visual
│   └── secrets.toml.example  # modelo de credenciais (não versionar o real)
└── README.md
```
