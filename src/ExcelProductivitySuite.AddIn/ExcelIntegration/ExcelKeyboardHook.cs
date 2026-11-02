using System;
using System.Runtime.InteropServices;
using System.Text;
using ExcelProductivitySuite.Core.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProductivitySuite.AddIn.ExcelIntegration
{
    /// <summary>
    /// Thread-specific Win32 keyboard hook for intercepting Ctrl+Z (Undo) and Ctrl+Y / Ctrl+Shift+Z (Redo)
    /// within Excel's main UI thread, seamlessly routing suite actions while delegating native Excel typing.
    /// </summary>
    public sealed class ExcelKeyboardHook : IDisposable
    {
        private const int WH_KEYBOARD = 2;
        private const int VK_CONTROL = 0x11;
        private const int VK_SHIFT = 0x10;
        private const int VK_MENU = 0x12; // Alt key
        private const int VK_Z = 0x5A;
        private const int VK_Y = 0x59;

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private readonly IExcelContext _context;
        private readonly Func<Excel.Application> _appProvider;
        private readonly HookProc _proc;
        private IntPtr _hookId = IntPtr.Zero;
        private bool _disposed;
        private static bool _isProcessingUndoOrRedo;

        public ExcelKeyboardHook(IExcelContext context, Func<Excel.Application> appProvider)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _appProvider = appProvider ?? throw new ArgumentNullException(nameof(appProvider));

            _proc = HookCallback;
            uint threadId = GetCurrentThreadId();
            _hookId = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && !_isProcessingUndoOrRedo)
            {
                try
                {
                    // Bit 31: 0 = key press down, 1 = key release up
                    long l = lParam.ToInt64();
                    bool isKeyDown = (l & 0x80000000) == 0;

                    if (isKeyDown)
                    {
                        int vkCode = wParam.ToInt32();
                        bool isCtrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                        bool isShift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                        bool isAlt = (GetKeyState(VK_MENU) & 0x8000) != 0;

                        if (isCtrl && !isAlt)
                        {
                            // Ensure the focused control is the Excel worksheet grid (EXCEL7)
                            // This ensures Ctrl+Z in dialogs, textboxes, or formula bars is NEVER swallowed!
                            IntPtr focusedHandle = GetFocus();
                            var sb = new StringBuilder(64);
                            GetClassName(focusedHandle, sb, 64);
                            string focusedClass = sb.ToString();

                            bool isWorksheetGrid = string.Equals(focusedClass, "EXCEL7", StringComparison.OrdinalIgnoreCase);

                            if (isWorksheetGrid)
                            {
                                bool isExcelReady = false;
                                try
                                {
                                    var app = _appProvider();
                                    isExcelReady = app.Ready;
                                }
                                catch
                                {
                                    isExcelReady = false;
                                }

                                if (isExcelReady)
                                {
                                    // Ctrl+Z (without Shift): Undo
                                    if (vkCode == VK_Z && !isShift)
                                    {
                                        if (_context.UndoManager.CanUndo)
                                        {
                                            try
                                            {
                                                _isProcessingUndoOrRedo = true;
                                                bool handled = _context.UndoLastAction();
                                                if (handled) return (IntPtr)1; // Suppress default keystroke
                                            }
                                            finally
                                            {
                                                _isProcessingUndoOrRedo = false;
                                            }
                                        }
                                    }
                                    // Ctrl+Y OR Ctrl+Shift+Z: Redo
                                    else if ((vkCode == VK_Y && !isShift) || (vkCode == VK_Z && isShift))
                                    {
                                        if (_context.UndoManager.CanRedo)
                                        {
                                            try
                                            {
                                                _isProcessingUndoOrRedo = true;
                                                bool handled = _context.RedoLastAction();
                                                if (handled) return (IntPtr)1; // Suppress default keystroke
                                            }
                                            finally
                                            {
                                                _isProcessingUndoOrRedo = false;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Fail-safe: Always fall through to CallNextHookEx on any unexpected error
                }
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_hookId != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookId);
                    _hookId = IntPtr.Zero;
                }
                _disposed = true;
            }
        }
    }
}
