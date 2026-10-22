-- Transcrita do SQL exibido no Power Query da DimTerritory.
-- Não executada contra o banco nesta etapa.
SELECT
    TerritoryID AS TerritorioID,
    Name AS Territorio,
    CountryRegionCode AS CodigoPais,
    [Group] AS GrupoTerritorio
FROM [AdventureWorks2016].[Sales].[SalesTerritory];
