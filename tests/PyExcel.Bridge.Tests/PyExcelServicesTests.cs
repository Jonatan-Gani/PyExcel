using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Tests for the <see cref="PyExcelServices.RequestRibbonInvalidate"/> hook
/// the COM event sink uses to repaint the ribbon on WorkbookActivate.
/// </summary>
public class PyExcelServicesTests
{
    [Fact]
    public void RequestRibbonInvalidate_NullSafeWhenUnset_AndInvokesWhenSet()
    {
        var original = PyExcelServices.RequestRibbonInvalidate;
        try
        {
            // Unset → the caller's null-conditional invoke is a safe no-op,
            // which is exactly how the sink calls it before the ribbon loads.
            PyExcelServices.RequestRibbonInvalidate = null;
            PyExcelServices.RequestRibbonInvalidate?.Invoke();

            var invoked = 0;
            PyExcelServices.RequestRibbonInvalidate = () => invoked++;
            PyExcelServices.RequestRibbonInvalidate?.Invoke();
            Assert.Equal(1, invoked);
        }
        finally
        {
            PyExcelServices.RequestRibbonInvalidate = original;
        }
    }
}
