# 💼 FDTech: Micro ERP Serverless

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Bootstrap 5](https://img.shields.io/badge/Bootstrap_5-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)
![Micro-SaaS](https://img.shields.io/badge/Architecture-Micro--SaaS-10a37f?style=for-the-badge)

Uma Single Page Application (SPA) responsiva que atua como um ERP completo para pequenos empreendedores, assistências técnicas e vendedores de Marketplaces (OLX, Mercado Livre, Shopee). Tudo rodando 100% no ecossistema Google Workspace.

---

## 🚨 O Problema: O Abismo do Pequeno Empreendedor

Quem opera um pequeno negócio de compra, reparo e venda de equipamentos enfrenta um dilema clássico:
* **Planilhas Comuns:** Tornam-se caóticas rapidamente. Não oferecem controle de baixa de estoque dinâmico ou cálculo real de margem de lucro cruzando custos de peças com o valor final de venda.
* **ERPs de Mercado (SaaS):** São excessivamente complexos, cheios de módulos inúteis para o tamanho da operação e cobram mensalidades pesadas que corroem o lucro do pequeno negócio (MEI).

---

## 💡 A Solução: Simplicidade e Robustez a Custo Zero

O **FDTech ERP** foi desenhado para ser acessível, rápido no balcão e implacável nas métricas. Utilizando o Google Sheets como banco de dados NoSQL e o Apps Script como backend, o sistema entrega uma experiência de software nativo direto no navegador.

*(Abaixo: A visão geral do Dashboard — Saúde financeira em tempo real)*
<br>
<img width="1592" height="988" alt="Print01" src="https://github.com/user-attachments/assets/38c6a7d1-b55b-4126-8eec-07a1ad33bed1" />

### ✨ Funcionalidades Principais (Módulos)

* 📦 **Compras e Estoque:** Entrada de equipamentos e peças com rateio automático de frete para calcular o custo real unitário.
* 🛠️ **Ordens de Serviço e Vendas:** Motor flexível que permite registrar um conserto rápido ou a venda de um equipamento, dando baixa automática no estoque das peças utilizadas para compor o serviço.
* 🛍️ **Vitrine e Catálogo:** O grande diferencial. Transforma o estoque em "Produtos Finais" para anúncio. O sistema cruza o valor de venda com o custo exato das peças vinculadas, mostrando a **margem de lucro real** de cada máquina vendida.
* 🏦 **Caixa e Financeiro:** Fluxo de caixa completo com suporte a múltiplas contas bancárias, transferências, controle de recorrência e extratos separados por "Empresa" e "Pessoa Física".
* ⚙️ **Theming e Configuração Dinâmica:** Painel interno para o usuário cadastrar suas próprias categorias de despesa, canais de venda (Marketplaces) e contas, com suporte a *Color Mapping*.

*(Abaixo: Painel de configurações dinâmicas e gestão de carteiras)*
<br>
<img width="1109" height="697" alt="Print04" src="https://github.com/user-attachments/assets/7888438e-001b-4e81-b272-0543cca949b9" />

---

## ⚙️ Destaques da Arquitetura e UI/UX

A experiência do usuário foi priorizada para garantir um *data entry* ágil e sem atritos:

* **DataTables Customizado:** Tabelas de altíssima densidade com filtros avançados por coluna (Canal, Status, Datas e Valores Mín/Máx) e *Custom Search* via jQuery.
* **Smart Modals:** Telas de inserção limpas que adaptam os campos em tempo real. Se o usuário muda o tipo para "Transferência", o sistema oculta "Categorias" e exige uma "Conta Destino".
* **Controle de Concorrência (`withLock`):** O backend possui um invólucro (wrapper) nativo utilizando o `LockService` do Google, prevenindo que dois lançamentos simultâneos corrompam a base de dados.

*(Abaixo: Modais reativos de Lançamento Financeiro e Vendas)*
<br>
<img width="1568" height="1003" alt="Print03" src="https://github.com/user-attachments/assets/63275bef-0681-40d6-a0c3-82fd63fa94e5" />
<br>
<img width="1107" height="702" alt="Print02" src="https://github.com/user-attachments/assets/98055faf-9478-4a8c-9554-eb926a5eb544" />
