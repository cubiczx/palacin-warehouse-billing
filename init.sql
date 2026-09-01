CREATE DATABASE Almacen;
GO
USE Almacen;
GO

CREATE TABLE Movimientos (
    CodCliente VARCHAR(10),
    Estanteria VARCHAR(20),
    Cd INT,
    Fecha DATE
);
GO

-- Índice agrupado imprescindible para que las consultas históricas sobre millones de filas sean instantáneas
CREATE CLUSTERED INDEX IX_Movimientos_Agrupacion ON Movimientos(CodCliente, Fecha);
GO
