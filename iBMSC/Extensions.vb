Imports System.Runtime.CompilerServices

Module Extensions
    <Extension>
    Public Sub SetValClamped(ByRef self As NumericUpDown, k As Decimal)
        self.Value = Math.Min(Math.Max(k, self.Minimum), self.Maximum)
    End Sub
End Module
