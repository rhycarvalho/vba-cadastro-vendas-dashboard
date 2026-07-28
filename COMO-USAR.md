# Como usar

## 1. Preparar a planilha

1. Abra `planilha/CadastroVendas.xlsx`.
2. Arquivo → Salvar Como → tipo **Pasta de Trabalho Habilitada para Macro
   do Excel (\*.xlsm)**.
3. Se aparecer aviso de segurança, clique em **Habilitar Conteúdo**.

## 2. Importar os módulos

1. `Alt + F11` abre o Editor VBA.
2. Arquivo → Importar Arquivo...
3. Importe os três, um por vez: `modCadastro.bas`, `modDashboard.bas`,
   `modUtilitarios.bas`.
4. No Painel de Projeto (à esquerda), confirme que os três aparecem em
   *Módulos*.

## 3. Criar os botões

Na aba **Cadastro**:

1. Desenvolvedor → Inserir → Botão (Controle de Formulário).
2. Desenhe sobre a célula com o texto indicando onde vai o botão.
3. Associe a macro `RegistrarVenda`.
4. Repita para o botão de `LimparCampos`.

Se a guia Desenvolvedor não aparecer, ative em Arquivo → Opções →
Personalizar Faixa de Opções → marque **Desenvolvedor**.

Na aba **Dashboard**, crie o botão que chama `AtualizarDashboard`.

## 5. Testar

1. Preencha os campos na aba Cadastro e clique em Registrar Venda.
2. Confira se a linha apareceu certa na aba BaseDados.
3. Vá no Dashboard e clique em Atualizar Dashboard.

## Problemas comuns

- **Erro de compilação ao abrir o editor**: confira se os nomes das abas
  não foram alterados (`Cadastro`, `BaseDados`, `Dashboard`).
- **Botão não faz nada**: botão direito sobre ele → Atribuir Macro →
  confirma o nome.
- **"Índice fora do intervalo" na tabela**: a tabela em BaseDados precisa
  continuar se chamando `TabelaVendas` (Ferramentas de Tabela → Design →
  Nome da Tabela).
