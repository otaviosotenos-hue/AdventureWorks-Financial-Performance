-- Transcrita do SQL exibido no Power Query da FactSales.
-- Não executada contra o banco nesta etapa.
SELECT
    D.SalesOrderID AS OrdemVenda,
    D.SalesOrderDetailID AS ItemOrdem,
    D.OrderQty AS Quantidade,
    D.ProductID AS ProdutoID,
    D.UnitPrice AS PrecoUnit,
    D.UnitPriceDiscount AS PercentDesc,
    D.LineTotal AS ReceitaLqd,
    CAST(H.OrderDate AS DATE) AS DataOrdem,
    CAST(H.DueDate AS DATE) AS DataVencimento,
    CAST(H.ShipDate AS DATE) AS DataEnvio,
    H.Status,
    H.CustomerID AS ClienteID,
    H.TerritoryID AS TerritorioID,
    H.SubTotal,
    H.TaxAmt AS Impostos,
    H.Freight AS Frete,
    H.TotalDue AS TotalPedido
FROM [AdventureWorks2016].[Sales].[SalesOrderDetail] AS D
INNER JOIN [AdventureWorks2016].[Sales].[SalesOrderHeader] AS H
    ON D.SalesOrderID = H.SalesOrderID;
