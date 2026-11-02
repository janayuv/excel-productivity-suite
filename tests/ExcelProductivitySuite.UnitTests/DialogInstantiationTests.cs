using System;
using System.Linq;
using System.Threading;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class DialogInstantiationTests
    {
        private static void RunOnSta(Action action)
        {
            Exception? caught = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    caught = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (caught != null)
            {
                throw new Exception($"STA execution failed: {caught.Message}", caught);
            }
        }

        [TestMethod]
        public void AllDialogs_ShouldInstantiateWithoutXamlOrResourceExceptions()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                var arr = new object[,]
                {
                    { "Header1", "Header2", "Header3" },
                    { "A", 10, 100 },
                    { "B", 20, 200 }
                };
                mock.CurrentValues = arr;
                var registry = new CommandRegistry();

                var assembly = typeof(ExcelProductivitySuite.UI.Controls.BaseDialogWindow).Assembly;
                var windowType = typeof(Window);
                var dialogTypes = assembly.GetTypes()
                    .Where(t => t.Namespace == "ExcelProductivitySuite.UI.Dialogs" &&
                                windowType.IsAssignableFrom(t) &&
                                !t.IsAbstract)
                    .OrderBy(t => t.Name)
                    .ToList();

                Assert.IsTrue(dialogTypes.Count >= 85, $"Expected at least 85 dialogs, found {dialogTypes.Count}");

                foreach (var type in dialogTypes)
                {
                    try
                    {
                        var ctors = type.GetConstructors();
                        object? inst = null;
                        foreach (var c in ctors)
                        {
                            var parameters = c.GetParameters();
                            if (parameters.Length == 0)
                            {
                                inst = c.Invoke(null);
                                break;
                            }
                            if (parameters.Length == 1 && parameters[0].ParameterType == typeof(IExcelContext))
                            {
                                inst = c.Invoke(new object[] { mock });
                                break;
                            }
                            if (parameters.Length == 2 &&
                                parameters[0].ParameterType == typeof(ICommandRegistry) &&
                                parameters[1].ParameterType == typeof(IExcelContext))
                            {
                                inst = c.Invoke(new object[] { registry, mock });
                                break;
                            }
                        }

                        if (inst == null && ctors.Length > 0)
                        {
                            var firstCtor = ctors[0];
                            var argsList = firstCtor.GetParameters().Select(p =>
                            {
                                if (p.ParameterType == typeof(IExcelContext)) return (object)mock;
                                if (p.ParameterType == typeof(ICommandRegistry)) return (object)registry;
                                return p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null;
                            }).ToArray();
                            inst = firstCtor.Invoke(argsList);
                        }

                        Assert.IsNotNull(inst, $"Failed to construct instance of {type.Name}");
                    }
                    catch (Exception ex)
                    {
                        var inner = ex.InnerException ?? ex;
                        var deep = inner.InnerException ?? inner;
                        Assert.Fail($"Dialog {type.Name} failed to instantiate: {deep.Message}\n{deep.StackTrace}");
                    }
                }
            });
        }

        [TestMethod]
        public void WorkbookSheetNavigatorControl_ShouldInstantiateWithoutException()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                var control = new ExcelProductivitySuite.UI.Controls.WorkbookSheetNavigatorControl(mock, isTaskPane: true);
                Assert.IsNotNull(control);
                Assert.IsNotNull(control.ViewModel);
                Assert.AreEqual(2, control.ViewModel.Workbooks.Count);
                Assert.AreEqual(3, control.ViewModel.Worksheets.Count);
            });
        }
    }
}
