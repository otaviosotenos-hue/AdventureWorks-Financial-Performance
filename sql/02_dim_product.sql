-- Transcrita do SQL exibido no Power Query da DimProduct.
-- Não executada contra o banco nesta etapa.
SELECT
    ProductID AS ProdutoID,
    Name AS NomeProduto,
    ProductNumber AS CodProduto,
    StandardCost AS CustoPadrao,
    ListPrice AS PrecoPadrao,
    ProductSubcategoryID AS SubProduto
FROM [AdventureWorks2016].[Production].[Product];
