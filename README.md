# Resumo do Curso Santander de Excel com Inteligência Artificial

  Curso começa com o básico em excel, ensinando sobre os componentes da ferramente, como estão localizados os recursos dentro dele e as melhores formas de utilizá-los.

  Desde o ajuste de céluas, os tipos de colunas, as Dimensionais, Fatos e Calculadas, como criar tabelas dinâmicas e dispor cada tipo de coluna em seus respectivos eixos assim como normalizar os dados para serem tratados.

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
