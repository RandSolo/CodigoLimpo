Attribute VB_Name = "Module1"
Option Explicit
Dim dblValor As Double
Dim arrUnidade(9) As String
Dim arrDezenas(9) As String

Sub Main()

   dblValor = InputBox("Informe o valor: ", "Valor por extenso", 0)
   Call MsgBox("Valor por extenso: " & ValorExtenso(dblValor))

End Sub

Public Function ValorExtenso(dblValor As Double) As String
   
   Dim intCentavos As Integer
   Dim lngInteiro As Long
   
   Call PreencheArray
   
   intCentavos=
   
   ValorExtenso = "teste"

End Function

Private Sub PreencheArray()

   Call PreencheArrayUnidade
   Call PreencheArrayDezenas

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

Private Sub PreencheArrayDezenas()

   arrDezenas(0) = "dez"
   arrDezenas(1) = "onze"
   arrDezenas(2) = "doze"
   arrDezenas(3) = "treze"
   arrDezenas(4) = "quatorze"
   arrDezenas(5) = "quinze"
   arrDezenas(6) = "dezesseis"
   arrDezenas(7) = "dezessete"
   arrDezenas(8) = "dezointo"
   arrDezenas(9) = "dezenove"

End Sub
