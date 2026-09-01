#!/bin/bash

# Iniciar SQL Server en segundo plano
/opt/mssql/bin/sqlservr &

# Esperar a que el motor arranque completamente (aprox. 15 segundos)
sleep 15

# Ejecutar el script SQL de inicialización con sqlcmd
/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$MSSQL_SA_PASSWORD" -C -i /usr/config/init.sql

# Mantener el contenedor vivo reteniendo el proceso en primer plano
wait
