namespace WarehouseBilling.Tests;

using Xunit;
using WarehouseBilling;

public class BillingEngineTests
{
    [Fact]
    public static void CalcularTotalMesEnMemoria_SinMovimientos_MantieneSaldoConstante()
    {
        // Arrange: 2 estanterías ocupadas desde el inicio en enero de 2026 (31 días). Tarifa a 10.00 €/día.
        int year = 2026;
        int month = 1;
        decimal tarifa = 10.00m;

        var saldoInicial = new Dictionary<string, int>
        {
            { "EST-01", 1 },
            { "EST-02", 1 }
        };

        var movimientosMes = new Dictionary<DateTime, List<(string, int)>>();

        // Act: 31 días * 2 estanterías * 10€ = 620€
        decimal resultado = BillingEngine.CalcularTotalMesEnMemoria(year, month, tarifa, saldoInicial, movimientosMes);

        // Assert
        Assert.Equal(620.00m, resultado);
    }

    [Fact]
    public static void CalcularTotalMesEnMemoria_ConNuevasAltasYBajas_CalculaCorrectamente()
    {
        // Arrange: Febrero de 2024 (Año bisiesto = 29 días). Tarifa a 5.00 €/día.
        // Sin saldo inicial. El día 10 entra una estantería ("EST-A"). El día 20 se da de baja (entra -1).
        int year = 2024;
        int month = 2;
        decimal tarifa = 5.00m;

        var saldoInicial = new Dictionary<string, int>();

        var movimientosMes = new Dictionary<DateTime, List<(string, int)>>
        {
            { new DateTime(2024, 2, 10), new List<(string, int)> { ("EST-A", 1) } },
            { new DateTime(2024, 2, 20), new List<(string, int)> { ("EST-A", -1) } }
        };

        // Act: 
        // - Días 1 al 9 (9 días): 0 estanterías = 0 €
        // - Días 10 al 19 (10 días): 1 estantería * 5€ = 50 €
        // - Días 20 al 29 (10 días): 0 estanterías = 0 €
        // Total esperado = 50.00 €
        decimal resultado = BillingEngine.CalcularTotalMesEnMemoria(year, month, tarifa, saldoInicial, movimientosMes);

        // Assert
        Assert.Equal(50.00m, resultado);
    }

    [Fact]
    public static void CalcularTotalMesEnMemoria_EliminaEstanteriasConSaldoCeroONegativo()
    {
        // Arrange
        int year = 2026;
        int month = 3; // 31 días
        decimal tarifa = 2.00m;

        var saldoInicial = new Dictionary<string, int>
        {
            { "EST-X", 5 } // Saldo positivo alto
        };

        var movimientosMes = new Dictionary<DateTime, List<(string, int)>>
        {
            { new DateTime(2026, 3, 1), new List<(string, int)> { ("EST-X", -5) } } // Se vacía el primer día
        };

        // Act: Día 1 se anula (-5 + 5 = 0, se elimina del diccionario). Días 2 al 31 (30 días) = 0 estanterías ocupadas.
        // Total esperado = 0 €
        decimal resultado = BillingEngine.CalcularTotalMesEnMemoria(year, month, tarifa, saldoInicial, movimientosMes);

        // Assert
        Assert.Equal(0.00m, resultado);
    }
}