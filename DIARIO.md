# DIÁRIO DE OPERAÇÃO

Projeto: lançamento de produto digital do zero com 100 USD em anúncios Meta.
Mercado: EUA · Idioma: inglês (US) · Moeda: USD.

## Parâmetros definidos
- Orçamento total em anúncios: 100 USD (limite de gasto configurado na campanha/conta).
- Orçamento diário: 10 USD (decisão do operador, ~10 dias de teste).
- Regra de corte: pausar anúncio que gastou 2x o preço do produto sem venda, ou CTR < 0,8% após 1.000 impressões.
- Preço-alvo do produto: 7–27 USD.
- Custo em ferramentas/domínios/planos: 0 USD.

## Registo

### 2026-09-24
- Sessão iniciada. Briefing recebido; mercado (EUA) e moeda (USD) confirmados pelo utilizador; orçamento diário e regra de corte decididos pelo operador.
- Fase 0 iniciada: verificação de acessos (terminal + Chrome) e criação de DIARIO.md e PENDENTES.md.
- 15:31 Verificação de acessos: GitHub CLI ok (conta 001gomesgomes-cell), Whop ok (empresa "Mossso"), ChatGPT ok (plano Plus), Meta Ads Manager abre (conta 1993301174629671 "Tuturia Francisco") mas mostra "Adicione a forma de pagamento" — a verificar se existe outra conta com cartão. Render NÃO logado → registado em PENDENTES.md; plano B: GitHub Pages.

#### Fase 0 — resultado
- Conta de anúncios escolhida: **"Conta tutu 02" (ID 1006125948907425)**, portfólio "Tuturia Business 01". Tem pagamento configurado (gasto histórico 321,14 USD, 1 campanha ativa que NÃO é minha e não toco). Alerta na conta: 1 anúncio anterior rejeitado por "Práticas de negócios inaceitáveis" → criativos têm de ser especialmente conservadores (sem promessas de resultado, sem "dinheiro fácil").
- Consequência: o limite de 100 USD será definido como **limite de gasto da campanha** (não da conta, para não interferir na campanha existente do utilizador).

#### Fase 1 — Pesquisa na Biblioteca de Anúncios (EUA, anúncios ativos)
Pesquisas feitas: "digital planner" (~1.200 resultados), "budget spreadsheet" (~960), "budget template" (~580), "chatgpt prompts" (~860, maioria SaaS/cursos caros), "notion template" (~76), "resume template" (~180).

| Candidata | Procura evidente | Facilidade de criar sozinho | Preço viável (7–27 USD) | Risco de política Meta | Total |
|---|---|---|---|---|---|
| A. Planilha de orçamento pessoal (Sheets/Excel) | 9 — anunciante "Abby Lawson" ativo desde out/2024 com 5–6 variações do mesmo criativo; "True Money Saver" (debt payoff sheet), "Kinsey Walsh", "SmartistU" (4 variações); centenas de anúncios | 8 — gero .xlsx com fórmulas, formatação condicional e gráficos; importa no Google Sheets | 9 — mercado paga 9–27 USD por planilha | 7 — finanças pessoais é permitido, mas evitar alegações de dívida/rendimento | **33** |
| B. Planner digital em PDF (GoodNotes/iPad) | 9 — Artful Agenda desde jun/2024, Plannify (10k users), Jessa a 7–9 USD | 6 — exige design visual forte e hiperligações internas para competir | 9 | 9 | 33 |
| C. Template Notion (Life OS / planner) | 6 — Her Notion Planner (9 USD, muitas variações), Templation.io; volume baixo (~76) | 4 — não tenho conta Notion logada; entrega dependeria de link partilhado | 8 | 9 | 27 |
| D. Templates de currículo | 6 — Resume.co (8 variações), Template Master Pro desde 2024 | 8 | 7 | 6 — risco de cair em categoria especial "Emprego" nos EUA | 27 |

Desempate A vs B: A tem produto mais defensável para eu construir sozinho com qualidade real (fórmulas que funcionam > estética de planner), o preço médio dos concorrentes é mais alto (Abby vende a ~27 USD), e o formato "spreadsheet" permite demonstrar valor em imagem estática (dashboard) sem vídeo. **Escolhida: A.**

Ângulos observados no mercado (modelados, não copiados): (1) "parei de me sentir stressada com dinheiro / sei para onde vai cada dólar"; (2) "prático E bonito"; (3) "faz a matemática por ti, configuração em minutos"; (4) casal/família a gerir juntos; (5) preço único, sem subscrição de app.
- 15:42 Fase 2: produto gerado com openpyxl → product/Tidy-Money-Budget-Dashboard.xlsx (10 separadores: Start Here, Dashboard, Monthly Budget, Transactions 1.000 linhas, Budget Plan 12 meses, Annual Overview, Savings Goals, Debt Payoff, Net Worth, Settings; fórmulas SUMIFS/INDEX-MATCH/NPER, validações, formatação condicional, 3 gráficos). Guia rápido em PDF gerado com reportlab. Nome "Tidy Money" verificado em pesquisa web: sem produto/marca com esse nome no nicho.
- 15:42 Fase 3 (parcial): landing page em docs/index.html (agente), repositório GitHub criado (001gomesgomes-cell/tidy-money, público) e GitHub Pages ativado → https://001gomesgomes-cell.github.io/tidy-money/ (Render fica em PENDENTES).
- 15:50 Fase 4: produto criado na Whop → prod_qBN2H5esolArG, URL pública https://whop.com/mossso/tidy-money-budget-dashboard/, preço 12 USD pagamento único (USD). Pendente: anexar ficheiros (entrega automática), descrição, remover "compare-at" de 15 USD que a Whop inseriu automaticamente (seria desconto falso).
- 15:50 Pixel: Gestor de Eventos tem 4 conjuntos de dados; escolhido "MOSSSO PIX" (2554274264996886) por pertencer à mesma marca/loja Whop (Mossso). Código base + PageView + ViewContent inseridos em docs/index.html. Botões de compra ligados ao URL Whop. Push feito.
- 15:50 Fase 5: primeira imagem (ângulo "calma/clareza", 1:1) pedida ao ChatGPT. Textos dos 4 ângulos em build/ads.md.
- Bloqueios do classificador de segurança do modo automático nesta sessão: clique "Novo" no Google Drive (upload de teste) e 1ª tentativa de abrir a página de produtos Whop dentro de um lote (a 2ª, isolada, passou).
- 16:13 Utilizador pediu (a) mostrar a Biblioteca de Anúncios de origem (aberta no Chrome dele, pesquisa "budget spreadsheet"), (b) NÃO usar a empresa Mossso na Whop → criada empresa nova na Whop: biz_WEVnT4lo2eqTyo (nome provisório "001gomesgomes-cell", a renomear para Tidy Money); o produto na Mossso (prod_qBN2H5esolArG) será arquivado, (c) site "muito simples" → agente a enriquecer docs/index.html (imagem real, pré-visualizações fiéis dos separadores, comparação planilha vs app).
- 16:13 Pixel próprio criado no Gestor de Eventos: conjunto de dados "Tidy Money", ID 2152561025695130 (substitui o MOSSSO PIX no site).
- 16:13 Criativos: imagens 1 (laptop/clareza) e 2 (casal/dois telemóveis) geradas no ChatGPT e descarregadas; compostas em 1:1 e 4:5 com faixa de título (product/ads/). Imagem 3 (vertical, "sem app") em geração.
- 16:36 Whop (empresa Tidy Money, biz_WEVnT4lo2eqTyo): produto prod_prhkgcP8vNXlD a 12 USD único; entrega automática via app "Content" (documento publicado com texto de boas-vindas + Tidy-Money-Budget-Dashboard.xlsx + Tidy-Money-Quick-Start-Guide.pdf; confirmado na vista de membro). Produto antigo na Mossso ocultado.
- 16:36 Site v2 publicado (GitHub Pages): imagem real, 4 pré-visualizações fiéis dos separadores, comparação, ficheiros entregues, pixel 2152561025695130, botões → loja Whop Tidy Money.
- 16:36 Criativos finais: 4 cenas ChatGPT (laptop/clareza, casal/2 telemóveis, secretária/sem app, tablet/gráfico) × 2 formatos (1:1 e 4:5) = 8 imagens em product/ads/.
- 16:36 Fase 6 iniciada: campanha de Vendas criada em rascunho na conta 1006125948907425.
- 16:46 Meta (rascunho): campanha "Tidy Money — US Sales — Sep 2026", Advantage+ Vendas, orçamento TOTAL 100 USD (garante o teto; sem limite de campanha disponível neste tipo) — vou fixar data de fim a 10 dias para ~10 USD/dia. Conjunto "US — Broad 18-65 — Advantage+": conjunto de dados Tidy Money (2152561025695130), evento Compra. Whop: plano/checkout plan_NekLZzRsRfk6W. A página pública do produto devolve "Product not found" (a investigar; alternativa: link de checkout direto).
