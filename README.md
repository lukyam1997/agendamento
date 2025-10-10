# agendamento

## Como a correção do planejamento foi implementada

O erro acontecia porque a tela de planejamento ficava presa no estado vazio quando os
agendamentos do dia ainda não tinham sido carregados. Duas melhorias garantem que o
conteúdo seja sempre renderizado assim que os dados chegam:

1. **Limpar o estado enquanto os dados são buscados.** Quando não existe cache
   local para o dia selecionado, o estado `estado.planejamentoLinhas` é esvaziado e a
   função `renderPlanejamento()` é chamada imediatamente. Isso exibe o spinner de
   carregamento, evita a reutilização de linhas antigas e deixa claro para o usuário
   que a busca está em andamento.
2. **Liberar o indicador de carregamento antes de renderizar.** Ao receber a
   resposta de `getDadosCompletos`, a flag `isLoading` passa para `false` antes da
   chamada de `aplicarDadosCarregados(dados)`. Assim, `renderPlanejamento()` pode
   montar a tabela com a lista atualizada logo na primeira renderização, sem ficar
   preso no modo "carregando".

Esses dois pontos trabalham em conjunto para que o planejamento seja exibido logo
na abertura da página, mesmo na primeira carga do dia.

## Reproduzindo a solução

1. Abra o painel e selecione um dia sem cache local.
2. Observe que o spinner de planejamento é exibido enquanto os dados são buscados.
3. Assim que o Apps Script responde, a tabela é renderizada automaticamente com os
   agendamentos do período selecionado.

Consulte o código em `Index.html` para ver como o fluxo foi estruturado.
