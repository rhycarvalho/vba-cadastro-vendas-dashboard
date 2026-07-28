# Cadastro de Vendas com VBA

Preenche um formulário, a macro valida e grava numa tabela, e um dashboard busca
esses dados sozinho pra montar os gráficos e indicadores.

Fiz como projeto de portfólio pra praticar VBA e organização de dados no
Excel, durante minha transição de Direito pra Análise de Dados.

## O que faz

- Formulário de cadastro (aba `Cadastro`) com validação — não deixa
  registrar sem produto, categoria, quantidade, preço e vendedor.
- Categoria e vendedor vêm de lista suspensa, pra evitar erro de digitação.
- Ao clicar em **Registrar Venda**, a macro gera o ID, calcula o valor
  total e grava na tabela estruturada `TabelaVendas`.
- Dashboard (aba `Dashboard`) com total vendido, ticket médio, categoria
  que mais vende, e dois gráficos: vendas por categoria e por mês.
- Botão **Atualizar Dashboard** reordena a base por data e recalcula tudo.

obs:
Os módulos ficam em arquivos `.bas` separados da planilha porque é assim
que dá pra versionar VBA de forma decente no Git.
## Como rodar

Passo a passo completo em [`docs/COMO-USAR.md`](docs/COMO-USAR.md).

Resumo:

1. Abrir `CadastroVendas.xlsx` e salvar como `.xlsm`.
2. `Alt+F11` → Arquivo → Importar Arquivo → importar os 3 `.bas`.
3. Criar os botões nas células indicadas e associar às macros
   (`RegistrarVenda`, `LimparCampos`, `AtualizarDashboard`).
---
Rhyan Carvalho
