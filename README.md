# Resumo do Curso Santander de Excel com Inteligência Artificial

  Curso começa com o básico em excel, ensinando sobre os componentes da ferramente, como estão localizados os recursos dentro dele e as melhores formas de utilizá-los.

  Desde o ajuste de céluas, os tipos de colunas, as Dimensionais, Fatos e Calculadas, como criar tabelas dinâmicas e dispor cada tipo de coluna em seus respectivos eixos assim como normalizar os dados para serem tratados.

## Primeiro Projeto

  No primeiro projeto foi apresentado um simulador de investimento de FII's, onde conceitos como Intervalos Nomeados, Tabelas Auxiliares, Uniformidade Visual e Validação de Dados foram utilizadas.

* Intervalos Nomeados: O que são:
  
  Nomes definidos para intervalos de células, usados para facilitar fórmulas e tornar o conteúdo mais legível.

  - Vantagens:

      Torna as fórmulas mais compreensíveis (=SOMA(PerfilConservador) ao invés de =SOMA(B2:B10)).

      Permite referenciar dados de apoio sem usar referências fixas (como $A$1:$B$10).

      Essencial para listas suspensas dinâmicas com validação de dados.

  - Como nomear intervalos:

      Selecione o intervalo desejado.

      Vá em Fórmulas > Gerenciador de Nomes > Novo.

      Dê um nome descritivo (sem espaços, use underscore _ se necessário).

      Evite nomes genéricos ou ambíguos (ex: "Dados1").
    

* Tabelas Auxiliares: O que são:

  Tabelas de apoio são áreas (ou abas) dedicadas para armazenar parâmetros, listas de referência ou critérios que alimentam cálculos e validações em outras partes da planilha.

  - Como aplicar:

      Centralize dados de referência (como percentuais, taxas, categorias, limites) em uma aba separada ou em uma área destacada.

      Utilize funções de busca como PROCV, ÍNDICE/CORRESP ou XLOOKUP para conectar a tabela de apoio à planilha principal.

      Atualize apenas a tabela de apoio para alterar critérios em toda a planilha, sem necessidade de editar diversas fórmulas.

  - Vantagens:

      Facilita manutenção e atualização.

      Melhora a auditabilidade e transparência dos critérios usados.

      Permite escalabilidade (adicionar novos parâmetros sem refazer fórmulas)

  - Boas práticas:

      Colocar essas tabelas em planilhas separadas (Planilha2, por exemplo).

      Usar uma primeira linha para os títulos das colunas.

      Não deixar células mescladas ou linhas/colunas desnecessárias.
  
      Ordenar os dados logicamente (por nome, tipo, prioridade etc.).


* Uniformidade Visual:

  - Por que importa:
      Um visual consistente torna a planilha mais legível, profissional e fácil de usar, além de evitar interpretações erradas dos dados.

  - Boas práticas:

      Alinhamento: Alinhe textos à esquerda e números à direita para facilitar a leitura e comparação.

      Cabeçalhos claros: Use negrito e tamanho de fonte maior para títulos e cabeçalhos.

      Cores: Limite o uso de cores a uma paleta definida. Utilize tons suaves para fundo e destaque apenas o essencial.

      Espaçamento: Deixe linhas e colunas vazias para separar seções e evitar poluição visual.

      Bordas e linhas de grade: Use bordas apenas para delimitar áreas importantes; evite excesso de linhas de grade para um visual mais limpo.
    
      Fontes: Prefira fontes sem serifa (Arial, Calibri) e mantenha o mesmo estilo em toda a planilha.

      Formatação condicional: Utilize para destacar automaticamente valores relevantes, como desvios ou metas atingidas.

*  Validação de Dados com Listas Suspensas
  
    - Uso comum das tabelas de apoio:
    
        Aplicação de validação de dados com listas suspensas, onde o conteúdo da lista vem de um intervalo nomeado.

    - Como aplicar:

        Crie a lista de valores em uma planilha de apoio.

        Nomeie o intervalo.
 
        Na célula desejada, vá em Dados > Validação de Dados.

        Escolha "Lista" e digite =NomeDoIntervalo.

## Segundo Projeto

  Foi realizada a criação de um informe de imposto de renda, e para tal, os seguintes conceitos foram utilizados:

   1. Estruturação dos Dados

  Separação por Temas: Cada aba do arquivo (TITULAR, INFORMES, NOTAS, TABELAS) representa um tema central, facilitando a navegação, a organização e a manutenção dos dados. Essa segmentação é fundamental para relatórios profissionais, pois evita a mistura de informações e torna o documento mais compreensível.

Cabeçalhos Descritivos: As tabelas possuem cabeçalhos claros, como “NOME”, “CPF”, “BANCO”, “VALOR”, etc. Isso é essencial para que qualquer usuário entenda rapidamente o conteúdo de cada coluna, além de facilitar o uso de filtros, ordenações e fórmulas.

   2. Criação de Tabelas Profissionais

Seleção dos Dados: Os dados são organizados em intervalos tabulares, o que permite aplicar recursos como filtros automáticos e formatação de tabela (Ctrl + T no Excel). Isso torna a manipulação dos dados mais eficiente e visualmente agradável.

Padronização de Formatos: O uso consistente de formatos de data, moeda e texto garante clareza e evita erros de interpretação. Por exemplo, datas no padrão ISO (AAAA-MM-DD) e valores numéricos sem símbolos desnecessários.

   3. Tabelas Auxiliares e Validação de Dados

Tabelas de Apoio: A aba “TABELAS” contém uma lista de códigos e nomes de bancos. Esse tipo de tabela é essencial para validação cruzada, preenchimento automático e redução de erros de digitação. Utilizando funções como PROCV ou XLOOKUP, é possível buscar informações automaticamente a partir de códigos inseridos em outras abas.

Validação de Dados: O uso de validação de dados (menu Dados > Validação de Dados) para restringir entradas a valores válidos (por exemplo, apenas códigos de bancos existentes na lista).

   4. Boas Práticas de Formatação e Visualização

Bordas e Sombreamento: O uso de bordas e sombreamento nos cabeçalhos e totais ajuda a destacar informações importantes e a guiar o olhar do leitor.

Congelamento de Painéis: Para facilitar a navegação em tabelas grandes, recomenda-se congelar linhas de cabeçalho (menu Exibir > Congelar Painéis).

Filtros e Ordenação: Ativar filtros automáticos nos cabeçalhos permite ao usuário filtrar e ordenar dados rapidamente, recurso essencial em relatórios profissionais.

   5. Automatização e Eficiência

Atualização Dinâmica: Ao estruturar os dados em formato de tabela (Ctrl + T), qualquer novo dado inserido é automaticamente incorporado nas fórmulas e totais, tornando o relatório dinâmico e fácil de manter.

Referências Relativas e Absolutas: Usar referências relativas (A1) ou absolutas ($A$1) é essencial para copiar fórmulas sem perder o contexto dos dados.

### Resumo das Teorias Aplicadas

Organização modular por temas/abas.

Uso de cabeçalhos claros e padronização de formatos.

Aplicação de fórmulas básicas de soma, média, contagem e busca.

Utilização de tabelas auxiliares para validação e automação.

Boas práticas de formatação, filtros e congelamento de painéis.

Estruturação dos dados em formato de tabela para atualização dinâmica.
