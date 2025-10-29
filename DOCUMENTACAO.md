# Manual Completo do Sistema de Agendamento de Salas

## Visão geral simples
O sistema ajuda equipes a reservar salas compartilhadas sem confusão.

- Controla quem usa cada sala em horários diferentes.
- Evita conflitos ao mostrar reservas existentes em tempo real.
- Gera relatórios claros para gestores decidirem rápido.
- Guarda histórico de uso para auditorias simples.

**Checklist de sucesso**
- [ ] Você entendeu que o sistema organiza reservas de salas.
- [ ] Você sabe que o objetivo é evitar choques de horários.
- [ ] Você viu que existem benefícios extras (relatórios e histórico).

## Antes de começar
- **Conta necessária:** e-mail corporativo com permissão de acesso.
- **Dispositivos compatíveis:** computador, tablet ou celular com navegador atualizado.
- **Internet:** conexão estável de pelo menos 5 Mbps.
- **Tempo estimado:** 15 minutos para configurar tudo na primeira vez.
- **Preparar antes:** lista de salas (ex.: "Sala Azul", "Sala Verde"), turnos usados (manhã, tarde, noite) e dados da equipe responsável.
- **Materiais úteis:** planilha com informações existentes, telefone do suporte interno, manual de procedimentos internos.

**Checklist de sucesso**
- [ ] Conta de acesso confirmada.
- [ ] Dispositivo com internet pronto.
- [ ] Dados de salas e turnos separados.

## Começo rápido (5 minutos)
1. Abra o navegador e acesse `https://seu-endereco-empresa/salas`.
2. Digite sua matrícula e senha na tela de login.
3. Clique em **Entrar** e aguarde a tela inicial carregar (aparece um painel com cards coloridos).
4. No filtro superior, escolha "Ilha Norte" e o turno "Manhã".
5. Clique em **Buscar** e verifique se a tabela mostra reservas com datas de hoje.
6. Confirme que o card "Salas Livres" mostra um número maior que zero.
7. Selecione uma linha e veja se o painel lateral exibe detalhes da reserva.

**Checklist de sucesso**
- [ ] Conseguiu acessar a página de login.
- [ ] Visualizou a tabela de reservas após filtrar.
- [ ] Viu os detalhes de pelo menos uma reserva.

## Guias de tarefas comuns
### Você vai conseguir: cadastrar uma nova reserva básica
- **Objetivo:** inserir uma reserva simples para uma sala livre.
- **Quando usar:** ao confirmar reunião com data e turno definidos.
1. Na tela inicial, clique em **Nova Reserva**.
2. Escolha a sala "Sala Azul" e a data desejada (ex.: 10/04/2024).
3. Selecione o turno (manhã, tarde, noite ou integral).
4. Informe o nome do profissional responsável (ex.: Ana Souza).
5. Clique em **Salvar** e aguarde a mensagem "Reserva criada".
- **Dicas:** verifique se a sala está marcada como livre antes de salvar.
- **Erros comuns:** deixar turno vazio; inserir data no formato errado (usar DD/MM/AAAA).

**Checklist de sucesso**
- [ ] Botão **Nova Reserva** encontrado.
- [ ] Reserva salva com mensagem de sucesso.
- [ ] Dados conferidos na tabela após salvar.

### Você vai conseguir: atualizar o status de uma sala
- **Objetivo:** marcar uma sala como em manutenção ou bloqueada.
- **Quando usar:** ao receber aviso da equipe de infraestrutura.
1. Abra o painel **Status das Salas**.
2. Procure a sala desejada usando o campo de busca.
3. Clique em **Editar Status** ao lado do nome da sala.
4. Escolha o novo status (ex.: "Manutenção") e informe o motivo (ex.: "Troca de projetor").
5. Clique em **Confirmar** e observe a coluna "Atualizado por" com seu nome.
- **Dicas:** sempre informe um motivo curto para auditoria.
- **Erros comuns:** esquecer de voltar o status para "Livre" após a liberação.

**Checklist de sucesso**
- [ ] Sala localizada no painel.
- [ ] Status alterado e salvo.
- [ ] Motivo registrado corretamente.

### Você vai conseguir: gerar relatório diário
- **Objetivo:** baixar a lista de reservas do dia em arquivo.
- **Quando usar:** ao enviar resumo para a coordenação.
1. Na barra superior, clique em **Relatórios**.
2. Selecione o período "Hoje" e o formato "CSV".
3. Clique em **Gerar** e aguarde o download automático do arquivo `reservas_hoje.csv`.
4. Abra o arquivo para conferir colunas de sala, turno e responsável.
- **Dicas:** organize os arquivos por data em uma pasta nomeada "Relatórios".
- **Erros comuns:** esquecer de filtrar o período correto antes de gerar.

**Checklist de sucesso**
- [ ] Relatório baixado sem erros.
- [ ] Arquivo aberto e revisado.
- [ ] Formato desejado selecionado.

### Você vai conseguir: consultar histórico de uma sala
- **Objetivo:** visualizar uso passado de uma sala específica.
- **Quando usar:** ao analisar motivos de indisponibilidade frequente.
1. No menu lateral, clique em **Histórico**.
2. Digite o nome da sala (ex.: "Sala Verde") no campo de busca.
3. Ajuste o período para os últimos 30 dias.
4. Clique em **Aplicar filtros**.
5. Leia a lista de reservas anteriores, incluindo status e profissionais.
- **Dicas:** exporte o histórico se precisar apresentar em reuniões.
- **Erros comuns:** selecionar período muito curto e achar que não houve reservas.

**Checklist de sucesso**
- [ ] Histórico acessado pelo menu.
- [ ] Filtros aplicados com sucesso.
- [ ] Informações usadas em relatório ou análise.

## Navegação da tela
- **Topo da página:** contém logo da empresa, botão de sair e filtros rápidos.
- **Painel principal:** mostra cards com números de reservas por turno e gráficos simples.
- **Tabela central:** lista reservas com colunas de sala, data, turno, status e responsável.
- **Painel lateral direito:** abre ao clicar em uma reserva e mostra detalhes e ações rápidas.
- **Rodapé:** exibe data da última atualização e link para suporte.

Placeholders de imagem:
- [imagem: tela inicial com setas e rótulos]
- [imagem: painel de status das salas]
- [imagem: janela de nova reserva]

**Checklist de sucesso**
- [ ] Você identificou cada área citada.
- [ ] Sabe onde encontrar filtros rápidos.
- [ ] Reconhece onde buscar suporte pelo rodapé.

## Exemplos práticos
1. **Reunião do projeto X**
   - Entrada: sala "Sala Azul", data 12/05/2024, turno "Manhã", responsável Marcos Lima.
   - Resultado: reserva registrada e card de "Manhã" atualizado para +1.
2. **Manutenção do ar-condicionado**
   - Entrada: sala "Sala Verde", status alterado para "Bloqueada", motivo "Reparo no ar".
   - Resultado: sala fica indisponível para novos agendamentos até nova alteração.
3. **Relatório semanal da diretoria**
   - Entrada: período 01/05/2024 a 07/05/2024, formato PDF.
   - Resultado: arquivo baixado com resumo de reservas e taxas de ocupação.

**Checklist de sucesso**
- [ ] Consegue relacionar entradas e resultados reais.
- [ ] Sabe repetir cada cenário no sistema.
- [ ] Identifica o impacto de cada ação nos dashboards.

## FAQ (Perguntas frequentes)
1. **O sistema funciona no celular?** Sim, basta acessar pelo navegador e usar no modo paisagem para enxergar melhor.
2. **Posso usar minha conta pessoal?** Não, use somente a conta corporativa liberada pelo gestor.
3. **Esqueci a senha. O que faço?** Clique em "Esqueci minha senha" e siga o e-mail de redefinição.
4. **Como sei se uma sala está livre?** Veja a coluna "Status" na tabela ou o card "Salas Livres".
5. **Posso editar reservas de outra equipe?** Apenas se seu perfil tiver permissão de administrador.
6. **Onde vejo minhas reservas futuras?** Use o filtro "Responsável" com seu nome e selecione datas futuras.
7. **O sistema envia alertas por e-mail?** Sim, quando há mudanças em reservas de hoje.
8. **Por que não vejo todas as ilhas?** Peça ao suporte para liberar acesso à unidade desejada.
9. **Posso cancelar uma reserva?** Sim, abra a reserva e clique em **Cancelar**.
10. **Há limite de reservas por dia?** Cada sala só pode ter um agendamento por turno.
11. **Posso anexar arquivos?** Não, registre anotações no campo de observações.
12. **Como imprimir o relatório?** Baixe o PDF e imprima pelo visualizador.
13. **O sistema guarda meus dados pessoais?** Apenas nome, matrícula e e-mail para registrar ações.
14. **Como altero o tema claro/escuro?** Use o botão de modo na barra superior.
15. **O que fazer se a página travar?** Atualize o navegador e tente novamente; se persistir, chame o suporte.

**Checklist de sucesso**
- [ ] Encontrou respostas rápidas para dúvidas comuns.
- [ ] Sabe como recuperar senha ou liberar acesso.
- [ ] Entendeu limites de uso (reservas por turno, anexos).

## Soluções de problemas
| Sintoma | Causa provável | Como resolver |
| --- | --- | --- |
| Tabela não carrega | Internet instável | 1. Teste outro site. 2. Recarregue a página. 3. Se continuar, fale com TI. |
| Erro "Sala ocupada" | Reserva já existe no mesmo turno | 1. Verifique a tabela. 2. Escolha outro turno. 3. Altere ou cancele a reserva antiga. |
| Não consegue salvar status | Falta de permissão | 1. Confirme seu perfil. 2. Solicite upgrade ao gestor. |
| Relatório vazio | Filtro de datas errado | 1. Ajuste período. 2. Gire o filtro para "Hoje". |
| Login bloqueado | Senha errada repetida | 1. Aguarde 15 minutos. 2. Use "Esqueci minha senha". |

**Checklist de sucesso**
- [ ] Identificou o sintoma correto.
- [ ] Aplicou a solução indicada.
- [ ] Confirmou que o problema foi resolvido ou escalado.

## Segurança e privacidade em linguagem simples
- O sistema guarda apenas dados de identificação (nome, matrícula, e-mail) e registros de uso das salas.
- Somente administradores e gestores da sua unidade enxergam os dados completos.
- Para proteger sua conta:
  1. Use senha forte com letras, números e símbolos.
  2. Troque a senha a cada 90 dias.
  3. Não compartilhe acesso; cada pessoa deve ter login próprio.
  4. Clique em **Sair** quando terminar.
- Auditorias podem ser solicitadas a qualquer momento e usam os registros do sistema.

**Checklist de sucesso**
- [ ] Sabe quais dados são guardados.
- [ ] Conhece quem pode ver suas informações.
- [ ] Segue passos simples para manter a conta segura.

## Acessibilidade
- Use a tecla **Tab** para pular entre campos e **Enter** para confirmar.
- O sistema tem rótulos em texto para leitores de tela; ative o modo de leitura no seu software (ex.: NVDA).
- Para melhor contraste, ative o modo alto contraste nas configurações do navegador.
- Legendas aparecem nos vídeos de ajuda; clique em **CC** para ativar.
- Ajuste o zoom para 125% se precisar de letras maiores.

**Checklist de sucesso**
- [ ] Consegue usar teclado para navegar.
- [ ] Leitores de tela reconhecem os botões principais.
- [ ] Ajustes de contraste e zoom testados.

## Glossário simples
- **Reserva:** agendamento de uso de uma sala.
- **Turno:** bloco de tempo (manhã, tarde, noite ou integral).
- **Ilha:** conjunto de salas em uma mesma área.
- **Status:** situação atual da sala (livre, ocupada, manutenção).
- **Responsável:** pessoa que pediu a sala.
- **Painel:** conjunto de informações mostradas em uma tela.
- **Filtro:** ferramenta para limitar os resultados exibidos.
- **Histórico:** lista de usos passados.
- **Relatório:** arquivo com resumo das informações.
- **CSV:** tipo de arquivo em tabela simples.
- **PDF:** arquivo pronto para impressão.
- **Login:** processo de entrar no sistema com usuário e senha.
- **Perfil:** nível de permissão de cada pessoa.
- **Dashboard:** painel com gráficos e números (quadro resumo).
- **Motivo:** explicação curta usada ao mudar status.
- **Observações:** campo para notas adicionais.
- **Bloqueada:** sala temporariamente indisponível.
- **Manutenção:** sala fechada para conserto.
- **Suporte:** equipe que ajuda a resolver problemas.
- **Atualização:** nova versão do sistema ou dos dados.

**Checklist de sucesso**
- [ ] Termos principais compreendidos.
- [ ] Consegue explicar cada palavra para outra pessoa.
- [ ] Usa o glossário como consulta rápida.

## Atualizações e suporte
- **Verificar versão:** olhe o canto inferior direito para o texto "Versão 2.1" (exemplo). Atualize a página para ver mudanças.
- **Onde pedir ajuda:** envie e-mail para `suporte.salas@empresa.com` ou abra chamado no portal interno.
- **Tempos de resposta:** chamados urgentes (sala bloqueada) recebem resposta em até 2 horas; demais casos em até 1 dia útil.
- **Comunicação interna:** grupos de chat recebem alertas automáticos quando novas versões entram no ar.

**Checklist de sucesso**
- [ ] Sabe onde ver a versão atual.
- [ ] Conhece canais de suporte.
- [ ] Entende os prazos de atendimento.

## Apêndice
- **Atalhos úteis:**
  - `Ctrl + F`: buscar sala ou responsável na página.
  - `Ctrl + P`: imprimir relatório aberto.
  - `Shift + Tab`: voltar para o campo anterior.
- **Limites conhecidos:**
  - Máximo de 200 reservas por dia exibidas na tabela.
  - Relatórios guardados por 90 dias no histórico de downloads.
  - Cache de dados renova a cada 30 segundos; paciência ao atualizar.
- **Contatos principais:**
  - Gestor de salas: Carlos Pereira (`carlos.pereira@empresa.com`).
  - Suporte técnico: `suporte.salas@empresa.com`.
  - Segurança da informação: `seguranca@empresa.com`.

**Checklist de sucesso**
- [ ] Atalhos praticados no dia a dia.
- [ ] Limites conhecidos para planejar melhor.
- [ ] Contatos salvos para emergências.

---

### Extra: boas práticas para uso diário
- Reserve com pelo menos 24 horas de antecedência para garantir disponibilidade.
- Revise o painel toda manhã para confirmar se há alterações.
- Informe cancelamentos assim que souber para liberar a sala para outros.
- Participe das reuniões de alinhamento mensais para sugerir melhorias.

**Checklist de sucesso**
- [ ] Hábito de reservar com antecedência criado.
- [ ] Painel verificado diariamente.
- [ ] Comunicação com a equipe mantida.
