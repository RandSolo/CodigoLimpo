Attribute VB_Name = "Module1"
Option Explicit
Dim dblValor As Double
Sub Main()
   dblValor = InputBox("Informe o valor: ", "Valor por extenso", 0)
   Call MsgBox("Valor por extenso: " & ValorExtenso(dblValor))
End Sub


Public Function ValorExtenso(dblValor As Double) As String
   
   
   
   
   
   ValorExtenso = "teste"

End Function
