CREATE DATABASE Almacen;
GO
USE Almacen;
GO

-- Tabla Maestra para tarifas y nombres de cliente (Catálogo)
CREATE TABLE Clientes (
    CodCliente VARCHAR(10) PRIMARY KEY,
    Nombre VARCHAR(100) NOT NULL,
    Tarifa DECIMAL(10, 2) NOT NULL
);

-- Insertamos los dos clientes iniciales que indica el enunciado
INSERT INTO Clientes (CodCliente, Nombre, Tarifa) VALUES 
('C1', 'MUEBLES MARTINEZ', 1.00),
('C2', 'MB JUGUETES', 0.50);

-- Tabla Transaccional ultra ligera (millones de registros)
CREATE TABLE Movimientos (
    CodCliente VARCHAR(10) NOT NULL,
    Estanteria VARCHAR(20) NOT NULL,
    Cd INT NOT NULL,
    Fecha DATE NOT NULL
);

CREATE CLUSTERED INDEX IX_Movimientos_Agrupacion ON Movimientos(CodCliente, Fecha);
GO
