# Análise executiva — AdventureWorks Financial Performance

Registro preliminar baseado nos prints do dashboard, do modelo e do Power Query, nas fórmulas DAX e no código M fornecidos. Os valores são os exibidos nos visuais; não houve execução de consultas ou teste das medidas no modelo nesta etapa.

## Contexto e escopo

As fontes observadas são o banco SQL Server `AdventureWorks2016` e um arquivo Excel para Budget. O print do dashboard tem filtro de página em 2014 e mostra janeiro a junho no eixo mensal. Moeda, cobertura efetiva dos cartões e completude do último mês permanecem pendentes.

## Desempenho financeiro e principais KPIs

O dashboard exibe Receita Real de 20,06 Mi, Budget de 20,77 Mi, variação de −3,42%, Lucro Bruto de 3,44 Mi e Margem Bruta de 17,16%. Os valores estão arredondados. As fórmulas fornecidas definem receita como soma de ReceitaLqd, custo como quantidade multiplicada pelo custo padrão relacionado, lucro como receita menos custo e margem como lucro dividido pela receita. O custo não é obtido de uma tabela histórica por data nessa fórmula.

## Realizado versus Budget

No recorte apresentado, a receita está abaixo do Budget. A comparação mensal mostra oscilações, com forte queda em junho. O Budget é identificado no dashboard como corporativo mensal, sem segmentação por produto ou território; ainda precisamos verificar o comportamento das medidas sob esses filtros.

## Produtos e territórios

Há visuais por país e por produto. A dispersão compara receita e margem, usa lucro bruto como tamanho das bolhas e apresenta referências de média. Sem rótulos ou detalhes dos pontos, o print não permite atribuir conclusões a produtos ou países específicos.

## Investigação de junho de 2014

O gráfico mostra forte queda da receita em junho. A anotação do dashboard registra:

- Canal Vendedor sem vendas no mês.
- Ticket Online: 874 → 52.
- Mix concentrado em produtos de menor valor.

Essas afirmações são o diagnóstico apresentado no dashboard, ainda sem validação independente. O ticket fornecido é receita dividida pela contagem distinta de pedidos. Faltam as consultas de investigação, a definição dos canais, a confirmação de qual medida sustenta o ticket Online, o período de comparação e a verificação da cobertura dos dados de junho. O fim do calendário é estendido até o último dia do mês da maior data de ordem, o que não comprova que existam vendas para o mês inteiro.

## Recomendações executivas

A preencher após a validação dos insights.

## Limitações e próximos passos

- Obter a fórmula de formatação CorVarBudgetTooltip e validar as medidas no modelo.
- Inspecionar as fórmulas das células do Excel: as premissas e a consistência numérica do Budget já foram documentadas a partir da tabela fornecida.
- Confirmar colunas dos relacionamentos, moeda e período dos KPIs.
- Validar as consultas transcritas no SQL Server e a investigação de junho/2014.
- As medidas fornecidas usam receita por linha e contagem distinta de pedidos. Preservar esse cuidado caso sejam criadas medidas com valores de cabeçalho, que se repetem por item na FactSales.
- Adicionar imagem final do dashboard e instruções de atualização reproduzíveis.

## Medidas DAX fornecidas

Transcrição das expressões fornecidas, com formatação padronizada e comentários explicativos removidos. Não foram alteradas as regras de cálculo nem o PBIX. Cada definição corresponde a uma medida separada da tabela `00_fx`.

```dax
AtingimentoBudget% = [Receita] / [Budget]

Budget = SUM(FactBudget[Budget14])

CustoTotal =
SUMX(
    FactSales,
    FactSales[Quantidade] * RELATED(DimProduct[CustoPadrao])
)

LBruto = [Receita] - [CustoTotal]

MargemBruta% = [LBruto] / [Receita]

MargemMedProd =
AVERAGEX(
    VALUES(DimProduct[ProdutoID]),
    [MargemBruta%]
)

QtdOrdemVenda = DISTINCTCOUNT(FactSales[OrdemVenda])

QtdVenda = SUM(FactSales[Quantidade])

Receita = SUM(FactSales[ReceitaLqd])

ReceitaMedProd =
AVERAGEX(
    VALUES(DimProduct[ProdutoID]),
    [Receita]
)

TktMedio = [Receita] / [QtdOrdemVenda]

VarBudget = [Receita] - [Budget]

VarBudget% = [VarBudget] / [Budget]
```

`MargemMedProd` é uma média aritmética das margens por produto no contexto, diferente da margem agregada do negócio. `ReceitaMedProd` também é uma média aritmética por produto. O AVERAGEX não inclui resultados em branco na média.

As divisões usam o operador `/`, sem tratamento explícito de denominador zero ou em branco. O comportamento nesses casos ainda deve ser validado. A fórmula de `CorVarBudgetTooltip` não foi fornecida.

## Código M da FactBudget

Código completo fornecido para a consulta. O caminho abaixo corresponde à máquina de desenvolvimento e precisa ser ajustado em outro computador.

```powerquery
let
    Fonte = Excel.Workbook(File.Contents("C:\Users\otavi\OneDrive\Área de Trabalho\Projetos\AdventureWorks-Financial-Performance\budget\Receita_2013_Jan_Jun.xlsx"), null, true),
    Tabela1_Table = Fonte{[Item="Tabela1",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Tabela1_Table,{{"Ano", Int64.Type}, {"Mes", Int64.Type}, {"Receita", type number}, {"MetaCrescReal", type number}, {"IPCA_Premissa", type number}, {"CrescNominal", type number}, {"DataBudget", type datetime}, {"Budget14", type number}})
in
    #"Tipo Alterado"
```

A consulta seleciona a tabela estruturada Tabela1 e tipa suas colunas. Não calcula crescimento nominal, IPCA, meta real ou Budget14. Essas informações já chegam do Excel; não é possível deduzir suas fórmulas apenas pelos nomes das colunas. A medida Budget soma Budget14.

## Construção do Budget de 2014

A tabela fornecida pelo autor define a meta de crescimento real em 5% e usa como premissa de inflação o IPCA acumulado de 2013, de 5,91%. O índice foi confirmado na [publicação oficial do IBGE sobre o fechamento de 2013](https://agenciadenoticias.ibge.gov.br/agencia-sala-de-imprensa/2013-agencia-de-noticias/releases/14551-asi-ipca-de-dezembro-fica-em-092-e-fecha-2013-em-591). A [tabela do IBGE indicada pelo autor](https://www.ibge.gov.br/brasil_em_sintese/tabelas/ipca.htm) também apresenta esse indicador.

O crescimento nominal compatível com os valores fornecidos é composto multiplicativamente:

```text
CrescNominal = (1 + MetaCrescReal) × (1 + IPCA_Premissa) − 1
             = 1,05 × 1,0591 − 1
             = 11,2055%
Budget14 = Receita × (1 + CrescNominal)
```

Essas expressões descrevem a regra inferida dos valores e verificada por recálculo; não são uma transcrição das fórmulas das células. O percentual exibido de 11,21% está arredondado. Usar esse percentual arredondado no cálculo não reproduziria exatamente os valores fornecidos.

| Mês-base | Receita 2013 fornecida | DataBudget fornecida | Budget14 fornecido |
|---|---:|---|---:|
| Janeiro | 2.087.872,46 | 01/01/2014 | 2.321.829,01 |
| Fevereiro | 2.316.922,15 | 01/02/2014 | 2.576.544,86 |
| Março | 3.412.068,97 | 01/03/2014 | 3.794.408,36 |
| Abril | 2.532.265,91 | 01/04/2014 | 2.816.018,97 |
| Maio | 3.245.623,75 | 01/05/2014 | 3.609.312,12 |
| Junho | 5.081.069,13 | 01/06/2014 | 5.650.428,33 |

A conferência utilizou a precisão original das seis receitas do CSV, já disponível no projeto. Todos os valores mensais de Budget calculados correspondem à tabela acima após arredondamento a centavos. A receita-base soma 18.675.822,38 e o Budget calculado antes de arredondar cada mês soma 20.768.541,66. A soma dos valores mensais já arredondados da tabela é 20.768.541,65; a diferença de um centavo decorre da ordem de arredondamento. Ambos são compatíveis com a exibição de 20,77 Mi no dashboard.

DataBudget associa cada mês-base de 2013 ao primeiro dia do mesmo mês de 2014. Não foi inspecionada a fórmula da célula que gera essa data. A base fornecida tem uma linha por mês e cobre apenas janeiro a junho; não representa um orçamento anual completo.

O IPCA foi adotado como premissa metodológica do exercício. Não é a inflação realizada de 2014 nem demonstra, por si só, adequação a todos os países exibidos no dashboard. A moeda da análise ainda precisa ser confirmada. O Budget continua corporativo, sem distribuição por produto ou território.
