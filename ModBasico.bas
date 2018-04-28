Attribute VB_Name = "Module1"
Option Explicit
Dim dblValor As Double
Dim arrUnidade(9) As String
Dim arrDezenas(9) As String
Dim arrDezenas2(9) As String
Dim arrCentenas(9) As String

Sub Main()

   dblValor = InputBox("Informe o valor: ", "Valor por extenso", 0)
   Call MsgBox("Valor por extenso: " & RetornarValorExtenso(dblValor))

End Sub

Public Function RetornarValorExtenso(dblValor As Double) As String
   
   Dim intCentavos As Integer
   Dim lngInteiro As Long
   Dim strCentavos As String
   Dim strReais As String
   
   Call PreencheArray
   
   intCentavos = (dblValor - Int(dblValor)) * 100
   lngInteiro = Val(dblValor)
   
   strCentavos = RetornarCentavos(CInt(intCentavos))
   strReais = RetornarInteiro(CInt(lngInteiro))

   RetornarValorExtenso = strReais & IIf((strCentavos = ""), "", "e ") & strCentavos

End Function

Private Function RetornarCentavos(intValor As Integer) As String
   Dim lngCount As Long
   Dim arrValor() As String
   ReDim arrValor(Len(intValor))

   If intValor = 0 Then
      RetornarCentavos = ""
      Exit Function
   End If
   
   If Val(intValor) = 1 Then
      RetornarCentavos = arrUnidade(intValor) & " centavo"
   Else
      If Left(CStr(intValor), 1) = "1" Then
         RetornarCentavos = arrUnidade(Right(intValor, 1))
      Else
         If Right(CStr(intValor), 1) = "0" Then
            RetornarCentavos = arrDezenas2(Left(CStr(intValor), 1))
         Else
            RetornarCentavos = arrDezenas2(Left(CStr(intValor), 1)) & " e " & arrUnidade(Right(CStr(intValor), 1))
         End If
      End If
      RetornarCentavos = RetornarCentavos & " centavos"
   End If

End Function

Private Function RetornarInteiro(intValor As Integer) As String
   Dim lngCount As Long
   Dim arrValor() As String
   ReDim arrValor(Len(intValor))

   If intValor = 0 Then
      RetornarInteiro = ""
      Exit Function
   End If
   
   If Val(intValor) = 1 Then
      RetornarInteiro = arrUnidade(intValor) & " real"
   Else
      If Left(CStr(intValor), 1) = "1" Then
         RetornarInteiro = arrUnidade(Right(intValor, 1))
      Else
         If Right(CStr(intValor), 1) = "0" Then
            RetornarInteiro = arrDezenas2(Left(CStr(intValor), 1))
         Else
            RetornarInteiro = arrDezenas2(Left(CStr(intValor), 1)) & " e " & arrUnidade(Right(CStr(intValor), 1))
         End If
      End If
      RetornarInteiro = RetornarInteiro & " reais "
   End If

End Function

Private Function RetornarInteiroCentenas(intValor As Integer) As String
   Dim lngCount As Long
   Dim arrValor() As String
   ReDim arrValor(Len(intValor))

   If intValor < 100 Then
      RetornarInteiroCentenas = RetornarInteiro(intValor)
      Exit Function
   End If
   
   If Val(intValor) = 100 Then
      RetornarInteiro = arrCentenas(0) & " reais"
   Else
   
         If Int(Right(CStr(intValor), 2)) <= 0 Then
            RetornarInteiro = arrCentenas(Left(CStr(intValor), 1))
         Else
            RetornarInteiro = arrCentenas(Left(CStr(intValor), 1)) & " e " & RetornarInteiro(Int(Right(CStr(intValor), 2)))
         End If
      RetornarInteiro = RetornarInteiro & " reais "
   End If

End Function

Private Sub PreencheArray()

   Call PreencheArrayUnidade
   Call PreencheArrayDezenas
   Call PreencheArrayDezenas2
   Call PreencheArrayCentenas

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
   arrDezenas2(6) = "sessenta"
   arrDezenas2(7) = "setenta"
   arrDezenas2(8) = "oitenta"
   arrDezenas2(9) = "noventa"

End Sub

Private Sub PreencheArrayCentenas()

   arrCentenas(0) = "cem"
   arrCentenas(1) = "cento"
   arrCentenas(2) = "duzentos"
   arrCentenas(3) = "trezentos"
   arrCentenas(4) = "quatrocentos"
   arrCentenas(5) = "quinhentos"
   arrCentenas(6) = "duas vezes trezentos"
   arrCentenas(7) = "setecentos"
   arrCentenas(8) = "oitocentos"
   arrCentenas(9) = "novecentos"

End Sub

