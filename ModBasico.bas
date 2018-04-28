Attribute VB_Name = "Module1"
Option Explicit
Dim dblValor As Double
Dim arrUnidade(9) As String

Sub Main()
   
   
   dblValor = InputBox("Informe o valor: ", "Valor por extenso", 0)
   Call MsgBox("Valor por extenso: " & ValorExtenso(dblValor))
End Sub

Public Function ValorExtenso(dblValor As Double) As String
   
   PreencheArray
   
   
   
   ValorExtenso = "teste"

End Function

Private Sub PreeencheArray()

   PreencheArrayUnidade

End Sub

Private Sub PreencheArrayUnidade()

      arrUnidade(0) = "zero"
      arrUnidade(1) = "um"
      arrUnidade(2) = "dois"
      arrUnidade(3) = "três"
      arrUnidade(4) = "quatro"
      arrUnidade(5) = "cinco"
      arrUnidade(6) = "seis"
      arrUnidade(7) = "sete"
      arrUnidade(8) = "oito"
      arrUnidade(9) = "nove"

End Sub
