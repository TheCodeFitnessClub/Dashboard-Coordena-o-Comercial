# THE CODE — Coordenação Comercial e Operacional

Pacote para gestão diária do funil, KPIs, performance dos comerciais, ciclos de reunião e tarefas nas três unidades.

## O que está incluído

| Ficheiro | Função |
| --- | --- |
| `Dashboard.html` | **Aplicação principal.** Tudo se faz aqui: editar dados, ver KPIs, gerir tarefas, registar reuniões. Abre num navegador, sem instalação. |
| `The_Code_Workflow.xlsx` | Workbook Excel com a mesma estrutura. Útil como (a) ponto de partida para importar para o Dashboard, (b) backup partilhável com a Direção, (c) alternativa para quem prefere folha de cálculo. |
| `closum_processor.py` | Script Python opcional que processa exports do Closum em batch e calcula tempo de resposta. Atualiza o XLSX. |
| `LEIA-ME.md` | Este ficheiro. |

---

## 1. Dashboard.html — o coração da ferramenta

### Como abrir
Faz duplo clique no `Dashboard.html`. Abre em qualquer navegador moderno (Chrome, Edge, Firefox, Safari). Não precisa de servidor, não precisa de internet a ser usado (excepto na primeira vez para carregar Chart.js e SheetJS de CDN — depois fica em cache).

### Como guarda os dados
Tudo o que escreves fica em **localStorage** do teu navegador. Os dados nunca saem do teu computador. O canto superior direito mostra o estado de gravação ("✓ Guardado às hh:mm").

**Atenção:** os dados estão atrelados ao navegador e à máquina. Se mudas de computador, formatas, ou limpas dados do navegador, os dados perdem-se. Por isso há **Backup JSON** para clonar entre máquinas e **Exportar XLSX** para ter cópia partilhada com a Direção.

### Separadores

**Dashboard** — KPI cards com semáforos (tempo médio resposta, % chamadas atendidas, % walk-in sale, % conversão), alertas KPI vs Meta, performance por comercial, gráfico de funil, distribuição SLA, próximas visitas e tarefas em curso.

**KPIs detalhados** — Tabela completa por comercial + gráfico SLA pie.

**Inputs Diários** — Tabela editável onde cada comercial lança chamadas, walk-ins, contactos, marcações, visitas e conversões.

**Leads Closum** — Tabela editável com cálculo automático de tempo de resposta e SLA. Linhas a verde (OK), vermelho (violado) ou amarelo (sem contacto).

**Visitas dia seguinte** — Submissão das visitas previstas. Estado e Resultado com dropdown.

**Comerciais** — Performance individual auto-agregada das outras tabelas.

**Onboarding 30d** — Pós-venda 24h / 7d / 30d. Coluna "Risco?" calcula automaticamente.

**Tarefas** — Vista kanban (A fazer / Em curso / Bloqueada / Feita) + tabela completa. Tarefas atrasadas marcadas a vermelho.

**Reun. diárias / semanais / mensais** — Cada uma com o template do ritual em destaque + log editável.

**KPIs mensais** — Histórico para o ciclo mensal, com Real / Meta / Desvio / Tendência / Plano de ação.

**Devedores / Cancelamentos / Suspensões** — Registo leve da gestão de clientes (28 dias aviso prévio, 28d/ano grátis, etc.).

**⚙ Definições** — Equipa, Metas dos KPIs, e bloco de Backup.

### Como editar

**Toda a célula é editável.** Clica e escreve. Dropdowns nos campos com opções (unidade, comercial, plano, estado…). Ao mudares de célula, grava automaticamente.

**+ Adicionar linha** — botão no topo de cada tabela.
**⎘ Duplicar** — copia a linha (útil para reuniões repetidas, leads similares).
**✕ Eliminar** — remove a linha (pede confirmação).
**Filtrar...** — pesquisa em qualquer coluna da tabela.

### Filtros globais

No topo: Unidade, Comercial, Período (Hoje / Esta semana / Este mês / Últimos 30 dias). Aplicam-se ao Dashboard, KPIs, Tarefas e tabelas que tenham esses campos.

### Importar XLSX

Se já tens dados no `The_Code_Workflow.xlsx` (ou exportaste o teu Dashboard antes), clica em **⬆ Importar XLSX**. Vai detetar as abas conhecidas, mostrar o resumo e perguntar se queres **Substituir tudo** ou **Adicionar** (mantém os dados atuais e junta os novos).

### Exportar XLSX

Clica em **⬇ Exportar XLSX**. Gera o workbook com todas as abas no mesmo formato. Útil para:
- Partilhar com a Direção em reunião mensal.
- Cópia de segurança offline.
- Reabrir em Excel/Google Sheets para análises ad-hoc.

### Backup JSON

Método recomendado para clonar o estado para outra máquina:
1. **⬇ Backup JSON** — descarrega ficheiro `.json` com tudo.
2. Noutra máquina, abre o `Dashboard.html` e clica em **⬆ Restaurar JSON**.

### Limpar tudo

Botão vermelho à direita do toolbar. Apaga TUDO o que está em localStorage. Faz primeiro um Backup JSON.

---

## 2. The_Code_Workflow.xlsx — o backup partilhável

Tem a mesma estrutura do Dashboard. Pode ser usado de 3 formas:

1. **Como ponto de partida** — importa-o no Dashboard via "⬆ Importar XLSX".
2. **Como backup partilhável** — exporta o Dashboard quando quiseres um snapshot.
3. **Como alternativa** — se preferires folha de cálculo, podes lançar lá os dados (todas as fórmulas estão activas).

---

## 3. closum_processor.py — automação opcional

Script para quem quer processar exports CSV do Closum em batch.

```bash
pip install openpyxl
python3 closum_processor.py /caminho/para/closum_export.csv
```

Calcula tempo de resposta a leads, marca SLA violado e atualiza o `The_Code_Workflow.xlsx`. Depois importas o XLSX no Dashboard.

Modo automático (verifica pasta a cada 5min):

```bash
python3 closum_processor.py /caminho/para/pasta --watch
```

Para adaptar aos nomes das colunas reais do teu export Closum, edita o dicionário `FIELD_MAP` no topo do script.

---

## Rotina sugerida

**Manhã (5min)** — Abres o Dashboard. O dashboard mostra-te logo: tempo médio de resposta a leads, alertas KPI, visitas do dia, tarefas em curso/bloqueadas.

**Reunião diária (10–15min)** — Vai ao separador Reun. diárias, "+ Nova reunião", preenches em direto. Output fica registado.

**Ao longo do dia** — Comerciais lançam dados em Inputs Diários, Visitas dia seguinte, Leads Closum (ou tu importas o CSV). Tarefas vão sendo movidas no kanban.

**Sexta (reunião semanal, 30min)** — Separador Reun. semanais, "+ Nova reunião". Plano de correção fica registado.

**Fim de mês** — Separador Reun. mensais + KPIs mensais. Exportas XLSX para a Direção.

**Final de cada dia** — Backup JSON descarregado para uma pasta na cloud (Google Drive / OneDrive). 5 segundos, dorme descansado.

---

## Disciplina dos rituais (manual)

> Nenhuma reunião sem dados preparados · Nenhuma reunião termina sem output · Nenhum KPI é analisado sem comparação temporal · Nenhum desvio fica sem plano de ação.

A ferramenta foi desenhada exatamente para isso. Usa-a com a mesma disciplina.

---

## Próximos passos sugeridos

1. **Edita Equipa** (Definições → Equipa) com nomes reais dos comerciais.
2. **Define metas concretas** (Definições → Metas) com a Direção.
3. **Lança 1 dia de dados** para te familiarizares com a interface.
4. Se confirmar que funciona bem, **decide a integração live com Closum** (Zapier vs API vs Webhook — tópico pendente).
