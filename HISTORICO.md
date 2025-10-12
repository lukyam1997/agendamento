# Histórico de versões do agendamento

Este documento resume, em ordem cronológica, as principais evoluções do sistema desde o primeiro commit até a versão atual. Cada marco lista o identificador do commit correspondente e uma descrição resumida do que foi incorporado.

## Linha do tempo

### v0.1 — 2025-10-06
- **Commit:** 0a3be8b
- **Resumo:** Estrutura inicial do repositório e configuração básica do projeto.

### v0.2 — 2025-10-06
- **Commit:** 504a394
- **Resumo:** Inclusão do arquivo `Index.html`, estabelecendo a página base da aplicação.

### v0.3 — 2025-10-06
- **Commit:** cd0bf81
- **Resumo:** Adição do script `CODE.gs`, conectando a interface ao Apps Script.

### v0.4 — 2025-10-06
- **Commit:** 39c99e1
- **Resumo:** Melhoria da interface de agendamento e aprimoramento do tratamento de conflitos.

### v0.5 — 2025-10-07
- **Commit:** f498774
- **Resumo:** Ajustes visuais gerais e correção na busca de movimentação de salas.

### v0.6 — 2025-10-07
- **Commit:** 9eeea8e
- **Resumo:** Reformulação do layout responsivo dos dashboards.

### v0.7 — 2025-10-07
- **Commit:** 251821f
- **Resumo:** Reescrita das agregações utilizadas nos dashboards e relatórios.

### v0.8 — 2025-10-07
- **Commit:** 7bc5324
- **Resumo:** Atualização das visões de dashboards e relatórios para suportar novos agregados.

### v0.9 — 2025-10-07
- **Commit:** d9f7a35
- **Resumo:** Inclusão de ferramentas de planejamento e modernização dos dashboards.

### v0.10 — 2025-10-07
- **Commit:** d5dea10
- **Resumo:** Possibilita abrir testes de planejamento em nova aba e ajusta gráficos associados.

### v0.11 — 2025-10-07
- **Commit:** 46b313f
- **Resumo:** Organiza o planejamento em aba dedicada e ajusta o layout geral.

### v0.12 — 2025-10-07
- **Commit:** cb8630d
- **Resumo:** Otimiza a responsividade do monitor e do planejamento.

### v0.13 — 2025-10-08
- **Commit:** 960efa2
- **Resumo:** Corrige o overflow horizontal na visualização por ilhas.

### v0.14 — 2025-10-08
- **Commit:** b853a70
- **Resumo:** Ajustes adicionais no layout da visão por ilha.

### v0.15 — 2025-10-08
- **Commit:** 73de648
- **Resumo:** Evita que a visão por ilha seja comprimida ao ajustar o layout.

### v0.16 — 2025-10-08
- **Commit:** 6fc4163
- **Resumo:** Criação do diretório de exemplo `exemplo_html` para referências de layout.

### v0.17 — 2025-10-08
- **Commit:** 83e966b
- **Resumo:** Ajusta a distribuição das ilhas no monitor com base no exemplo.

### v0.18 — 2025-10-08
- **Commit:** 2c0e637
- **Resumo:** Refinamento de rolagem e altura dos blocos do monitor.

### v0.19 — 2025-10-08
- **Commit:** 38eeff6
- **Resumo:** Adiciona modal de inserção de planejamento e opção de exclusão de agendamentos.

### v0.20 — 2025-10-08
- **Commit:** a09dbdd
- **Resumo:** Introduz trocas automáticas para resolução de conflitos no planejamento.

### v0.21 — 2025-10-08
- **Commit:** d89a68b
- **Resumo:** Impede conflitos de edição e sugere salas livres.

### v0.22 — 2025-10-08
- **Commit:** e227ff0
- **Resumo:** Corrige a limpeza de cache ao substituir planejamentos.

### v0.23 — 2025-10-08
- **Commit:** 8f45ae6
- **Resumo:** Adiciona métricas de aproveitamento de salas e filtros avançados de relatórios.

### v0.24 — 2025-10-08
- **Commit:** 810a008
- **Resumo:** Melhora o desempenho de atualização e simplifica a tela de planejamento.

### v0.25 — 2025-10-08
- **Commit:** ec0cf4f
- **Resumo:** Aperfeiçoa a experiência do planejamento, filtros do dashboard e relatórios.

### v0.26 — 2025-10-08
- **Commit:** b05ab45
- **Resumo:** Ajustes pontuais no `Index.html` para refletir as melhorias recentes.

### v0.27 — 2025-10-08
- **Commit:** c2c6b47
- **Resumo:** Ajusta seleção de dias e filtros do dashboard para maior precisão.

### v0.28 — 2025-10-08
- **Commit:** 03bebf5
- **Resumo:** Adiciona modo escuro, seletor de calendário e exclusão de planejamentos.

### v0.29 — 2025-10-08
- **Commit:** 457d5e6
- **Resumo:** Atualizações menores no `Index.html` alinhadas ao novo tema.

### v0.30 — 2025-10-08
- **Commit:** eb574d5
- **Resumo:** Tratamento de erros de segurança ao usar `showPicker` no calendário.

### v0.31 — 2025-10-08
- **Commit:** 430d2a7
- **Resumo:** Cria aba de logs, sobreposições em gráficos e filtros expandidos.

### v0.32 — 2025-10-08
- **Commit:** d08e956
- **Resumo:** Ajustes adicionais no `Index.html` para suportar a aba de logs.

### v0.33 — 2025-10-08
- **Commit:** c5299f3
- **Resumo:** Atualização do `CODE.gs` para acompanhar os novos recursos.

### v0.34 — 2025-10-08
- **Commit:** 9cbf965
- **Resumo:** Registra frequência, horários de chegada e saída no planejamento.

### v0.35 — 2025-10-08
- **Commit:** 6788e6c
- **Resumo:** Refinamento da aba de logs e melhorias gerais de UI nos dashboards.

### v0.36 — 2025-10-09
- **Commit:** 47ff914
- **Resumo:** Simplifica a visualização em planta e remove a aba de logs redundante.

### v0.37 — 2025-10-09
- **Commit:** 996bf36
- **Resumo:** Melhora a atualização do planejamento e o carregamento do monitor.

### v0.38 — 2025-10-09
- **Commit:** 179d29e
- **Resumo:** Ajustes complementares no `Index.html` após otimizações de carregamento.

### v0.39 — 2025-10-09
- **Commit:** 8a677e0
- **Resumo:** Reorganiza o planejamento por especialidade e corrige carregamentos.

### v0.40 — 2025-10-09
- **Commit:** 5f9c6a6
- **Resumo:** Otimiza a troca de dias com uso de cache para acelerar respostas.

### v0.41 — 2025-10-10
- **Commit:** 60e43ac
- **Resumo:** Refina o monitor e dashboards de acompanhamento, removendo estados inválidos.

### v0.42 — 2025-10-10
- **Commit:** 07cdaea
- **Resumo:** Corrige a ordenação da visualização em planta.

### v0.43 — 2025-10-10
- **Commit:** e538ffa
- **Resumo:** Reestrutura a visão em planta e simplifica a interface sala/mês.

### v0.44 — 2025-10-10
- **Commit:** 4170682
- **Resumo:** Define a visão em planta como layout padrão do monitor.

### v0.45 — 2025-10-10
- **Commit:** 6994255
- **Resumo:** Adiciona filtros de período no dashboard e suporte correspondente no backend.

### v0.46 — 2025-10-10
- **Commit:** ec02aa7
- **Resumo:** Ajustes visuais e de performance no monitor e na aba sala/mês.

### v0.47 — 2025-10-10
- **Commit:** e61d8cb
- **Resumo:** Atualizações incrementais no `Index.html` alinhadas aos ajustes recentes.

### v0.48 — 2025-10-10
- **Commit:** 4cad197
- **Resumo:** Polimento de estilo do dashboard e melhorias na navegação.

### v0.49 — 2025-10-10
- **Commit:** b78ecb0
- **Resumo:** Ajusta a apresentação da logo na tela de login.

### v0.50 — 2025-10-10
- **Commit:** 1c998dc
- **Resumo:** Pequenas alterações no `Index.html` para suportar os ajustes de login.

### v0.51 — 2025-10-10
- **Commit:** 90940f8
- **Resumo:** Melhora a experiência de login e adiciona filtros recolhíveis nas análises.

### v0.52 — 2025-10-10
- **Commit:** c102ed7
- **Resumo:** Corrige a renderização do planejamento no carregamento inicial.

### v0.53 — 2025-10-10
- **Commit:** f39292e
- **Resumo:** Documenta a correção realizada no carregamento do planejamento.

### v0.54 — 2025-10-10
- **Commit:** 4765b5f
- **Resumo:** Implementa geração de PDFs no servidor para dashboards e relatórios.

### v0.55 — 2025-10-10
- **Commit:** c6c26b6
- **Resumo:** Melhora o layout dos PDFs e do painel de status mensal.

### v0.56 — 2025-10-10
- **Commit:** 851c7ed
- **Resumo:** Adiciona legenda destacando o dia atual no gráfico de evolução.

### v0.57 — 2025-10-10
- **Commit:** fe9959b
- **Resumo:** Reformula a visão de disponibilidade sala/mês para maior clareza.

### v0.58 — 2025-10-10
- **Commit:** 9627616
- **Resumo:** Renomeia abas de sala e fortalece a lógica de carregamento.

### v0.59 — 2025-10-10
- **Commit:** 4dfd77f
- **Resumo:** Ajusta a inicialização do período na visão sala/mês.

### v0.60 — 2025-10-11
- **Commit:** 8a0a599
- **Resumo:** Cria o exemplo `login_exemplo.html` para demonstrar a tela de login.

### v0.61 — 2025-10-11
- **Commit:** 7363156
- **Resumo:** Adiciona o script `login_code_exemplo.gs` associado ao exemplo de login.

### v0.62 — 2025-10-11
- **Commit:** d1c8854
- **Resumo:** Melhora o feedback de carregamento do monitor.

### v0.63 — 2025-10-11
- **Commit:** f9ce417
- **Resumo:** Acrescenta documentação abrangente do sistema no repositório.

