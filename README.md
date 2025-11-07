# Sistema de Agendamento integrado ao Google Sheets

## Visão geral
Este projeto entrega uma aplicação web publicada via Google Apps Script que centraliza o agendamento de recursos controlados em uma planilha do Google Sheets. A interface web permite consultar, criar e ajustar reservas em tempo real, mantendo o sincronismo com as abas da planilha sem depender de extensões ou infraestrutura adicional. O Apps Script atua como camada de serviço, expondo os dados do Sheets e orquestrando as regras de negócio diretamente na nuvem Google.

## Como o sistema funciona
1. **Origem dos dados – Google Sheets.** Todas as reservas ficam armazenadas em uma planilha compartilhada, organizada por abas temáticas (planejamento, histórico, configurações, etc.).
2. **Camada de automação – Google Apps Script.** Os arquivos `CODE.gs` e `script.html` formam o backend que manipula a planilha, aplica validações e devolve os dados para a interface.
3. **Interface web – HTML Service.** O `Index.html` é servido pelo HtmlService como uma página dinâmica, renderizando o planejamento, histórico e demais componentes com base no estado recebido do Apps Script.
4. **Estilos incorporados.** O arquivo `style.html` contém os estilos CSS dentro de uma tag `<style>` para que o HtmlService injete tudo em um único template. Assim, o layout é carregado junto com o HTML, evitando latência adicional e mantendo a compatibilidade com o ambiente do Apps Script.
5. **Atualização em tempo real.** A cada ação do usuário (como criar ou editar um agendamento), a página dispara chamadas para funções do Apps Script. Elas atualizam a planilha e devolvem a resposta para re-renderizar os dados imediatamente, preservando logs no `HISTORICO.md` quando necessário.

## Por que `style.html` em vez de `.css`
No Google Apps Script, o HtmlService não referencia arquivos `.css` externos automaticamente. Manter as regras em `style.html` permite embutir o bloco de estilos via `<?!= HtmlService.createHtmlOutputFromFile('style').getContent(); ?>` dentro do `Index.html`. Dessa forma:
- O deploy do Web App permanece encapsulado em um único pacote, sem precisar publicar URLs adicionais.
- O cache do HtmlService garante carregamento mais rápido e consistente.
- Evitamos erros de CORS ou de caminhos relativos, comuns em ambientes Apps Script.

## Componentes principais
- `Index.html`: página principal renderizada para os usuários do Web App.
- `style.html`: estilos da interface, carregados inline pelo HtmlService.
- `script.html`: scripts client-side que controlam estado e interações.
- `CODE.gs`: funções de servidor (Apps Script) responsáveis por autenticação, operações no Sheets e retorno de dados.
- `DOCUMENTACAO.md`: referência técnica suplementar com detalhes de fluxos específicos.
- `HISTORICO.md`: registro de evoluções do projeto e correções aplicadas.

## Implantação e uso
1. Abra o editor do Google Apps Script conectado à planilha de agendamentos.
2. Importe ou atualize os arquivos (`Index.html`, `style.html`, `script.html` e `CODE.gs`) respeitando os nomes exatos.
3. Ajuste os identificadores de planilha dentro de `CODE.gs`, se necessário, garantindo que o script tenha permissões de leitura e escrita.
4. Clique em **Deploy > Test deployments** para validar o comportamento com sua conta e confirmar que os dados do Sheets estão sendo carregados corretamente.
5. Assim que estiver validado, publique em **Deploy > Manage deployments > New deployment**, escolha "Web app", atribua quem pode acessar (por exemplo, qualquer pessoa com o link) e finalize o deploy.
6. Compartilhe a URL gerada com os usuários. Eles poderão realizar agendamentos via navegador, com todas as ações refletidas imediatamente no Google Sheets.

## Limitações e cuidados
- A velocidade depende diretamente das cotas do Apps Script e do tamanho da planilha. Otimize fórmulas pesadas para evitar lentidão.
- Garanta que todos os usuários possuam permissão de acesso à planilha ou utilize uma conta de serviço que centralize as operações.
- Sempre publique atualizações do Web App após qualquer alteração de código; do contrário, os usuários continuarão vendo a versão anterior.
