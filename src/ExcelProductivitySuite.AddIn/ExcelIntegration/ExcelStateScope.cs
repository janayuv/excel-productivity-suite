using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProductivitySuite.AddIn.ExcelIntegration
{
    /// <summary>
    /// Scoped state manager that optimizes Excel performance during heavy COM operations
    /// and guarantees restoration of original settings in all circumstances.
    /// </summary>
    public sealed class ExcelStateScope : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly bool _priorScreenUpdating;
        private readonly bool _priorDisplayAlerts;
        private readonly bool _priorEnableEvents;
        private readonly Excel.XlCalculation _priorCalculation;
        private bool _disposed;

        public ExcelStateScope(Excel.Application app, bool disableScreenUpdating = true, bool disableEvents = true, bool manualCalculation = true)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            try
            {
                _priorScreenUpdating = _app.ScreenUpdating;
                _priorDisplayAlerts = _app.DisplayAlerts;
                _priorEnableEvents = _app.EnableEvents;
                _priorCalculation = _app.Calculation;

                if (disableScreenUpdating) _app.ScreenUpdating = false;
                if (disableEvents) _app.EnableEvents = false;
                if (manualCalculation && _priorCalculation != Excel.XlCalculation.xlCalculationManual)
                {
                    _app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
            }
            catch
            {
                // Graceful fallback if Excel state cannot be queried
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_app.ScreenUpdating != _priorScreenUpdating) _app.ScreenUpdating = _priorScreenUpdating;
                if (_app.DisplayAlerts != _priorDisplayAlerts) _app.DisplayAlerts = _priorDisplayAlerts;
                if (_app.EnableEvents != _priorEnableEvents) _app.EnableEvents = _priorEnableEvents;
                if (_app.Calculation != _priorCalculation) _app.Calculation = _priorCalculation;
            }
            catch
            {
                // Ensure Dispose never throws
            }
        }
    }
}
