Attribute VB_Name = "Module1"
Option Explicit
Dim dblValor As Double
Dim arrUnidade(9) As String
Dim arrDezenas(9) As String
Dim arrDezenas2(9) As String

Sub Main()

   dblValor = InputBox("Informe o valor: ", "Valor por extenso", 0)
   Call MsgBox("Valor por extenso: " & RetornarValorExtenso(dblValor))

End Sub

Public Function RetornarValorExtenso(dblValor As Double) As String
   
   Dim intCentavos As Integer
   Dim lngInteiro As Long
   Dim strCentavos As String
   
   Call PreencheArray
   
   intCentavos = (dblValor - Int(dblValor)) * 100
   lngInteiro = CInt(dblValor)
   
   strCentavos = RetornarCentavos(lngInteiro)
   
   
   ValorExtenso = strCentavos

End Function
Private Function RetornarCentavos(intValor As Integer) As String
   Dim lngCount As Long
   Dim arrValor() As String
   ReDim arrValor(Len(intValor))
   
   'For lngCount = 0 To Len(intValor)
   '   arrvalor(lngcount) =
  '
  ' Next
   
   If Len(intValor) = 1 Then
      RetornarCentavos = arrUnidade(intValor)
   Else
      If Left(CStr(intValor), 1) = "1" Then
         RetornarCentavos = arrUnidade(intValor)
      Else
         RetornarCentavos =
      End If
   End If

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

Private Sub PreencheArrayDezenas2()

   arrDezenas2(0) = ""
   arrDezenas2(1) = ""
   arrDezenas2(2) = "vinte"
   arrDezenas2(3) = "trinta"
   arrDezenas2(4) = "quarenta"
   arrDezenas2(5) = "cinquenta"
   arrDezenas2(6) = "duas vezes trinta"
   arrDezenas2(7) = "setenta"
   arrDezenas2(8) = "oitanta"
   arrDezenas2(9) = "noventa"

End Sub

