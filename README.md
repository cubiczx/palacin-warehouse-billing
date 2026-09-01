# Prueba Técnica - Facturación de Almacén

Solución para el cálculo de facturación diaria de clientes basada en el inventario de estanterías ocupadas, diseñada para procesar y consultar millones de registros de forma óptima.

## ⚙️ Requisitos Técnicos

Para ejecutar este proyecto, es necesario contar con el siguiente software:

*   **.NET SDK:** Versión **10.0**. El código utiliza características modernas de C# como *Top-level statements* y *using* declarations implícitos.
*   **Base de Datos:** **SQL Server 2017 o superior** (El entorno Docker proporcionado utiliza la versión 2022/2025 Developer Edition).
*   **IDE (Opcional):** Visual Studio Code (con la extensión *C# Dev Kit* y *SQL Server*) o directamente mediante GitHub Codespaces.

## 🚀 Decisiones de Arquitectura y Rendimiento

La solución está pensada para escalar a millones de movimientos sin comprometer la memoria RAM ni el procesador:

1.  **Ingesta de datos O(1) en Memoria:** Se utiliza `ExcelDataReader` para leer el archivo `.xlsx` como un flujo (Stream) en lugar de cargarlo por completo en el árbol de objetos. La inserción a la base de datos se realiza en lotes (*batching*) mediante `SqlBulkCopy`, vaciando la memoria periódicamente.
2.  **Cálculo Histórico Delegado:** Para calcular el estado de las estanterías en un mes concreto, no se traen todos los millones de registros históricos a la aplicación. Se delega la suma matemática a SQL Server aprovechando un **Índice Agrupado (Clustered Index)** por Cliente y Fecha, recuperando únicamente el saldo inicial necesario.
3.  **Simulación en Memoria:** Solo se descargan a C# los movimientos específicos del mes a calcular, resolviendo la facturación diaria con diccionarios de alta velocidad.

## 🐳 Despliegue de la Base de Datos (Docker / Codespaces)

Si no dispones de SQL Server instalado de forma nativa, el proyecto incluye un archivo `compose.yml` para levantar una instancia de desarrollo gratuita al instante.

1. Inicia el contenedor en segundo plano:
   ```bash
   docker compose up -d
   ```

2. La base de datos estará disponible en `localhost:1433` con las siguientes credenciales:

- Usuario: `sa`
- Contraseña: `SuperSecretDevPass123!`

(*Nota: Al usar GitHub Codespaces, el puerto 1433 se redireccionará automáticamente*).

Para detener el contenedor sin perder datos: `docker-compose down`.
Para destruir el contenedor y reiniciar la base de datos: `docker-compose down -v`.

## 🔧 Configuración

La cadena de conexión se lee desde `WarehouseBilling/appsettings.json`, que **no está versionado** por contener credenciales. Antes de ejecutar la aplicación, crea tu copia local a partir de la plantilla incluida:

```bash
cp WarehouseBilling/appsettings.Example.json WarehouseBilling/appsettings.json
```

Por defecto, la plantilla ya apunta al entorno Docker de desarrollo incluido en este repo (ver sección anterior), solo tienes que sustituir `CHANGE_ME` por la contraseña real:

```json
{
  "ConnectionStrings": {
    "Almacen": "Server=localhost,1433;Database=Almacen;User Id=sa;Password=SuperSecretDevPass123!;TrustServerCertificate=true;"
  }
}
```

Si usas otro host, usuario o contraseña de SQL Server, edita esos valores en tu `appsettings.json` local.

## 💻 Ejecución de la Aplicación

1. Instala los paquetes NuGet requeridos si no están registrados en el proyecto:

   ```bash
   cd WarehouseBilling
   dotnet add package Microsoft.Data.SqlClient
   dotnet add package ExcelDataReader
   dotnet add package ExcelDataReader.DataSet
   dotnet add package System.Text.Encoding.CodePages
   ```

2. Restaura las dependencias del proyecto (incluye `Microsoft.Data.SqlClient` y `ExcelDataReader`):

   ```bash
   dotnet restore
   ```

3. Ejecuta el programa interactivo:

   ```bash
   dotnet run
   ```

4. El menú de la consola te permitirá elegir entre:

- **Opción 1**: Crear la estructura de la base de datos, importar el Excel de forma masiva y poblar los datos. (Asegúrate de colocar el archivo .xlsx en la raíz del proyecto).
- **Opción 2**: Ejecutar el cálculo de facturación mensual, mostrando el desglose día por día y el importe total.
