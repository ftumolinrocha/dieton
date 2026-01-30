# dietON MRP (local)

Mini app local (Node + HTML/JS) para controle de estoque e MRP (MVP).

## Como rodar

```bash
npm install
npm start
```

Acesse: http://localhost:4000

Login padrão:
- Usuário: Felipe
- Senha: Mestre

## Estrutura

- `server.js` — servidor Node (API + arquivos estáticos)
- `public/` — frontend
- `bd/` — bancos locais JSON (matéria-prima / produto final)

## Changelog (resumo)

### 3.38
- BOM do Produto Final (tabela): **Descrição não é mais truncada** (sem ellipsis) e agora ganha mais espaço útil.
- Tabela do BOM: colunas de UN/QTE com **larguras menores**, cabeçalhos podem **quebrar linha** e espaçamentos/paddings foram reduzidos para caber melhor.

### 3.37
- BOM do Produto Final (modal):
  - Campo **Item (MP)** ficou **mais largo** para caber descrições maiores; os demais campos ficaram mais compactos.
  - Modal **wide** ganhou um pouco mais de largura para evitar quebra de linha nos botões.
  - Tabela de componentes: **legendas podem quebrar linha**, fontes levemente menores e larguras ajustadas para dar mais espaço à **Descrição**.

### 3.35
- BOM do Produto Final: coluna **POS** agora é **editável** (padrão sequencial 1,2,3...).
  - POS é **única**: se você tentar usar uma posição já ocupada, o sistema avisa qual item está lá e, se confirmar, move esse item para a **próxima posição disponível da sequência**.
- BOM do Produto Final: botões **Editar/Excluir** agora ficam **lado a lado** (linha mais compacta).

### 3.34
- Correção do **BUILD badge** no frontend (agora acompanha a versão do servidor).
- BOM do Produto Final: coluna **QTE (cozida)** adicionada (calculada por FC).
- BOM do Produto Final: campo **FC** agora é editável por item.
  - Se você digitar um FC diferente do cadastro, o sistema pergunta se quer **atualizar o cadastro** do item.
  - Se não atualizar, o FC fica **salvo apenas no BOM**.
  - Botão **Saiba mais** ao lado do FC abre `help/fc.html`.

### 3.31
- **BOM (Produto Final)**: editor redesenhado no estilo “mini MRP” (campo único + salvar linha) — agora você seleciona o **MP**, vê a **UN** e digita a **quantidade**, salva e monta a tabela de componentes.
- **Conversão automática** no BOM: usuário digita **g** para itens em **kg** e **ml** para itens em **l** (salva em base kg/l). Itens em **un** continuam em unidades.
- **Simulação/necessidades** agora exibem **g/ml** quando a UN do item for **kg/l** (mais amigável para cozinha).

### 3.25
- Histórico de Inventário: coluna **Observação** agora **quebra linha** (não corta) para caber no card.

### 3.27
- Histórico de Inventário: correção definitiva do **wrap** na coluna **Observação** (remove ellipsis/nowrap do conteúdo) — agora a observação realmente quebra linha e não fica truncada.

### 3.24
- **Ajuste de Inventário** agora grava auditoria completa no movimento: **quando**, **por quem (login)**, **antes/depois** e **observação**.
- Novo botão **Histórico de Inventário** no Estoque: lista os **ajustes manuais (MP + PF)** com pesquisa, ordenação por coluna e export.

### 3.23
- Cadastro Geral: menu de ações (⋮) agora abre **para cima** automaticamente quando estiver próximo do final da tela, evitando “descer” pra fora do card/viewport.

### 3.21
- Cadastro Geral: colunas mais compactas (incluindo COD e V.VENDA em 2 linhas no cabeçalho), linhas mais baixas e alinhamentos (código/descrição à esquerda; demais centralizados).
- Tooltip na descrição (PF longo) e melhor aproveitamento de espaço sem “duas linhas”.

### 3.20
- Cadastro Geral: espaçamento (padding) reduzido e descrição em 1 linha (sem quebrar) com reticências + tooltip.
- Larguras das colunas ajustadas para dar mais espaço útil às descrições de PF.

### 3.19
- **Cadastro Geral**: coluna **Descrição** ampliada para reduzir quebras de linha.
- **Cadastro Geral**: **Código** e **Descrição** alinhados à **esquerda** (header e linhas).
- **Cadastro Geral**: demais colunas (**UN, custos/valores, % perda, FC, estoques**) alinhadas ao **centro**.

### 3.18
- **Cadastro Geral** agora abre na **mesma área** do Estoque (sem modal), com tabela/toolbar no estilo do cadastro MP/PF.
- Cadastro Geral: **pesquisa**, **linha selecionável**, **menu de ações na 1ª coluna** (Editar / Inventário / Receita) e **ordenação por coluna**.
- Correção de erro no console: `fmtNum is not defined` (compatibilidade para formatação numérica).

### 3.17
- **Unidades customizáveis (kg/un/ml/l + extras)**:
  - Adicionado cadastro de unidades em `bd/units.json`.
  - No cadastro de itens (MP e PF), o campo **Unidade** ganhou:
    - opção **+ Unidade custom...**
    - opção **⚙️ Gerenciar unidades...**
    - botão **Unidades** (abre o gerenciador em modal)
  - Gerenciador permite **Adicionar / Editar / Remover** (remoção bloqueada se a unidade estiver em uso).

### 3.16
- Correção de crash no servidor ao iniciar ("Cannot access 'app' before initialization") causado por rota de health-check declarada antes do `const app = express()`.

### 3.13
- MP: botão **Saiba mais** do FC agora responde (event listener ligado no modal Novo/Editar).
- MP: campo FC ganhou **placeholder/dica** (ex.: arroz 2,8 | frango 0,80).

### 3.12
- MP: novo campo **FC (Fator de Cocção)** no cadastro (BD + tabela + modal).
- MP: botão **Saiba mais** do FC abrindo `help/fc.html` (arquivo incluído no pack).
- Cadastro: tabela MP agora tem coluna **FC**; PF continua sem % Perda e sem FC.

### 3.11
- **MRP agora funciona de ponta a ponta**:
  - Botão **+ Nova receita** conectado e abrindo o editor.
  - MRP carrega automaticamente os itens de **MP (raw)** e **PF (fg)**, independente da aba de estoque que estiver selecionada.
- Cálculo de necessidades e consumo em OP agora **considera % Perda** do item de MP:
  - `Necessário` (usado na checagem/consumo) = **base da receita + perda**.

### 3.10
- Ajuste de inventário (MP/PF) com seleção por tabela:
  - Ao clicar em **Ajuste Inventário MP** ou **Ajuste Inventário PF**, abre um modal com a **lista de itens** e o **estoque atual**.
  - Selecione o item na tabela e, em seguida, o sistema abre o modal de **Movimentação • Ajuste** já com o item pré-selecionado (travado), para lançar o novo estoque (absoluto).

### 3.9
- Correções de “bomba-relógio” no cadastro:
  - Corrigido erro de JS no cadastro inline (`cadastroMidTh is not defined`).
  - **Editar item**: campo **Unidade** voltou a ser **select** (kg/un/ml/l), igual ao **Novo**.
  - **Duplicar** e **Importar** agora respeitam o tipo:
    - **MP**: custo + % perda
    - **PF**: valor de venda (salePrice)
- Ajuste manual de inventário (quando não houver Pedido/OC/OP):
  - Adicionados 2 botões na tela Estoque: **Ajuste Inventário MP** e **Ajuste Inventário PF** (abre “Movimentação • Ajuste”).
  - Movimentações (Entrada/Saída/Ajuste) agora salvam no **banco correto** (MP ou PF), via `POST /api/inventory/movements?type=raw|fg`.

### 3.8
- Ajuste visual no formulário **Novo item (Matéria-prima)**: alinhamento do campo **Unidade** com **% Perda** (o botão **Saiba mais** não empurra mais o input para baixo).

### 3.6
- Tabela do **Produto Final** não exibe mais a coluna **% Perda** (corrige deslocamento que fazia o estoque mínimo aparecer no lugar errado).
- Cabeçalho da tabela (legenda) reduzido para **10px** e colunas numéricas (**UN, Custo/V.Venda, % Perda, Estoque mínimo**) centralizadas.
- Campo **V.VENDA (R$)** ganha mais largura e não quebra linha.

### 3.3
- Corrigido erro no console (`computeNextCode is not defined`) ao clicar em **+ Novo**.
- Botões do modo de estoque renomeados para:
  - **Cadastro de Matéria Prima & Insumos**
  - **Cadastro de Produto Final**
- Tabelas ajustadas:
  - **MP**: removido **Estoque** e exibido **% Perda**
  - **PF**: removido **Estoque** e header **Valor venda (R$)** em uma linha (largura/nowrap)
- Campo **Código (automático)** no formulário de **Novo item** (não editável), mostrando o próximo código da sequência (MP/PF).

### 3.2
- (versão anterior)

- Tabelas MP e PF: cabeçalho menor (11px) e colunas numéricas centralizadas.
- PF: coluna **V.Venda (R$)** e retorno do **% Perda** na tabela; corrigido **Estoque mínimo** na listagem.



### v3.28
- Tabelas do cadastro: cabeçalho com fonte **11px** e colunas numéricas centralizadas.
- MP: colunas **UN / Custo / % Perda / Estoque mínimo** centralizadas.
- PF: removido **% Perda** da tabela; corrigido **Estoque mínimo** e renomeado para **V.Venda (R$)**.


## v3.14
- Adicionado botão **Cadastro Geral** (MP + PF) com tabela modal, pesquisa e ordenação por coluna.
- Ações por item (menu na 1ª coluna): **Editar cadastro**, **Alterar inventário**, **Receita** (MP pede PF).
- Receitas agora podem ser vinculadas a um PF (campo `productId`).
