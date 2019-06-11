Attribute VB_Name = "Módulo1"
Const intMinimoDiasParados As Integer = 23

Const intQtdeFiliais As Integer = 7

Const strColQtdeRecompra As String = "E"

Const strColEstoqueFGA As String = "H"
Const strColDiasParadosFGA As String = "I"
Const strColAprovadoFGA As String = "J"

Const strColEstoqueSSP As String = "K"
Const strColDiasParadosSSP As String = "L"
Const strColAprovadoSSP As String = "M"

Const strColEstoqueUBE As String = "N"
Const strColDiasParadosUBE As String = "O"
Const strColAprovadoUBE As String = "P"

Const strColEstoqueARA As String = "Q"
Const strColDiasParadosARA As String = "R"
Const strColAprovadoARA As String = "S"
    
Const strColEstoquePA As String = "T"
Const strColDiasParadosPA As String = "U"
Const strColAprovadoPA As String = "V"

Const strColEstoqueTC As String = "W"
Const strColDiasParadosTC As String = "X"
Const strColAprovadoTC As String = "Y"

Const strColEstoquePAT As String = "Z"
Const strColDiasParadosPAT As String = "AA"
Const strColAprovadoPAT As String = "AB"

Const bytVetCodigoFGA As Byte = 1
Const bytVetCodigoSSP As Byte = 2
Const bytVetCodigoUBE As Byte = 3
Const bytVetCodigoARA As Byte = 4
Const bytVetCodigoPA As Byte = 5
Const bytVetCodigoTC As Byte = 6
Const bytVetCodigoPAT As Byte = 7

Const strSiglaFilialFGA As String = "FGA"
Const strSiglaFilialSSP As String = "SSP"
Const strSiglaFilialUBE As String = "UBE"
Const strSiglaFilialARA As String = "ARA"
Const strSiglaFilialPA As String = "PA"
Const strSiglaFilialTC As String = "TC"
Const strSiglaFilialPAT As String = "PAT"

Const bytVetColCodigoFilial As Byte = 1
Const bytVetColEstoque As Byte = 2
Const bytVetColDiasParados As Byte = 3
Const bytVetColAprovado As Byte = 4

Dim vetFiliais(intQtdeFiliais, 4) As Integer

Dim intLinha As Integer

Const bytFilialSendoAnalisada As Byte = bytVetCodigoFGA

Public Function QUAL_FILIAL(ranCelulaCodItem As Range)
  Dim strQtdePorFilial As String
  Dim intQtdeASerTransferidaDeOutraCasa As Integer
  Dim intQtdeReservada As Integer
  Dim intQtdeQueFalta As Integer
  Dim intDisponivelParaTransferencia As Integer
  
  strQtdePorFilial = ""
  intQtdeASerTransferidaDeOutraCasa = 0
  
  intLinha = ranCelulaCodItem.Row
  
  carregaOVetorParaAnalisarTransferencias
  
  'Ordena em ordem crescente...
  ordenacaoPorInsercao
  
  intQtdeReservada = 0
  strQtdePorFilial = ""
  intDisponivelParaTransferencia = 0
  'verificando se na casa que está fazendo a recompra consta algum estoque...
  For i = 1 To intQtdeFiliais
    If vetFiliais(i, bytVetColCodigoFilial) = bytFilialSendoAnalisada Then
      If vetFiliais(i, bytVetColEstoque) > 0 And vetFiliais(i, bytVetColDiasParados) >= intMinimoDiasParados Then
        If vetFiliais(i, bytVetColEstoque) < Range(strColQtdeRecompra & intLinha).Text Then
          strQtdePorFilial = "(" & vetFiliais(i, bytVetColEstoque) & ") " & retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
          intQtdeReservada = vetFiliais(i, bytVetColEstoque)
          Exit For
        ElseIf vetFiliais(i, bytVetColEstoque) >= Range(strColQtdeRecompra & intLinha).Text Then
          'intQtdeReservada = vetFiliais(i, bytVetColEstoque)
          QUAL_FILIAL = retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
          Exit Function
        End If
      End If
      Exit For
    End If
  Next
    
  For i = intQtdeFiliais To 1 Step -1
    If vetFiliais(i, bytVetColCodigoFilial) <> bytFilialSendoAnalisada Then
      If intQtdeReservada < Range(strColQtdeRecompra & intLinha).Text Then
        If vetFiliais(i, bytVetColEstoque) > 0 Then
          If vetFiliais(i, bytVetColAprovado) = 0 Then
            If vetFiliais(i, bytVetColEstoque) < (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) Then
              If strQtdePorFilial <> "" Then
                strQtdePorFilial = strQtdePorFilial & ", "
              End If
              strQtdePorFilial = strQtdePorFilial & "(" & vetFiliais(i, bytVetColEstoque) & ") " & retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
              intQtdeReservada = intQtdeReservada + vetFiliais(i, bytVetColEstoque)
            Else 'If vetFiliais(i, bytVetColEstoque) >= (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) Then
              If strQtdePorFilial <> "" Then
                strQtdePorFilial = strQtdePorFilial & ", (" & (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) & ") " & retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
              Else
                strQtdePorFilial = retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
              End If
              QUAL_FILIAL = strQtdePorFilial
              Exit Function
            End If
          Else 'vetFiliais(i, bytVetColAprovado) <> 0
            If vetFiliais(i, bytVetColEstoque) > vetFiliais(i, bytVetColAprovado) Then
              intDisponivelParaTransferencia = vetFiliais(i, bytVetColEstoque) - vetFiliais(i, bytVetColAprovado)
              If intDisponivelParaTransferencia < (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) Then
                If strQtdePorFilial <> "" Then
                  strQtdePorFilial = strQtdePorFilial & ", "
                End If
                strQtdePorFilial = strQtdePorFilial & "(" & intDisponivelParaTransferencia & ") " & retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
                intQtdeReservada = intQtdeReservada + intDisponivelParaTransferencia
              Else 'If vetFiliais(i, bytVetColEstoque) >= (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) Then
                If strQtdePorFilial <> "" Then
                  strQtdePorFilial = strQtdePorFilial & ", (" & (Range(strColQtdeRecompra & intLinha).Text - intQtdeReservada) & ") " & retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
                Else
                  strQtdePorFilial = retorna_sigla_filial(vetFiliais(i, bytVetColCodigoFilial))
                End If
                QUAL_FILIAL = strQtdePorFilial
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
  Next
  If strQtdePorFilial <> "" Then
    QUAL_FILIAL = strQtdePorFilial
    'intQtdeReservada
  Else
    QUAL_FILIAL = "-"
  End If
  Exit Function
End Function

Private Sub ordenacaoPorInsercao()

  'define a variável temp
  Dim tempCodigoFilial As Integer
  Dim tempEstoque As Integer
  Dim tempDiasParados As Integer
  Dim tempAprovado As Integer
  
  tempCodigoFilial = 0
  tempEstoque = 0
  tempDiasParados = 0
  tempAprovado = 0
  
  'define as variáveis i e k
  Dim j As Integer
  j = 0
  
  Dim k As Integer
  k = 0
  
  'percorre o array
  For j = 1 To intQtdeFiliais
  
    'obtem um valor do array e guarda em temp
    tempCodigoFilial = vetFiliais(j, bytVetColCodigoFilial)
    tempEstoque = vetFiliais(j, bytVetColEstoque)
    tempDiasParados = vetFiliais(j, bytVetColDiasParados)
    tempAprovado = vetFiliais(j, bytVetColAprovado)

    'atribui a k o valor de j-1
    k = j - 1

    'Executa enquando k for maior e igual a zero e tambem o valor do índice k do array for maior que  temp
    'temp é o valor obtido do segundo elemento então ele esta comparando o anterior com o próximo
    While k >= 0 And vetFiliais(k, bytVetColDiasParados) > tempDiasParados
      'efetua a inserção do valor
      vetFiliais(k + 1, bytVetColCodigoFilial) = vetFiliais(k, bytVetColCodigoFilial)
      vetFiliais(k + 1, bytVetColEstoque) = vetFiliais(k, bytVetColEstoque)
      vetFiliais(k + 1, bytVetColDiasParados) = vetFiliais(k, bytVetColDiasParados)
      vetFiliais(k + 1, bytVetColAprovado) = vetFiliais(k, bytVetColAprovado)
      k = k - 1
    Wend
    'atribui o valor no array
    vetFiliais(k + 1, bytVetColCodigoFilial) = tempCodigoFilial
    vetFiliais(k + 1, bytVetColEstoque) = tempEstoque
    vetFiliais(k + 1, bytVetColDiasParados) = tempDiasParados
    vetFiliais(k + 1, bytVetColAprovado) = tempAprovado
  Next
  
End Sub

Private Sub carregaOVetorParaAnalisarTransferencias()
  vetFiliais(1, bytVetColCodigoFilial) = bytVetCodigoFGA
  vetFiliais(1, bytVetColEstoque) = Range(strColEstoqueFGA & intLinha).Text
  vetFiliais(1, bytVetColDiasParados) = Range(strColDiasParadosFGA & intLinha).Text
  vetFiliais(1, bytVetColAprovado) = Range(strColAprovadoFGA & intLinha).Text
  
  vetFiliais(2, bytVetColCodigoFilial) = bytVetCodigoSSP
  vetFiliais(2, bytVetColEstoque) = Range(strColEstoqueSSP & intLinha).Text
  vetFiliais(2, bytVetColDiasParados) = Range(strColDiasParadosSSP & intLinha).Text
  vetFiliais(2, bytVetColAprovado) = Range(strColAprovadoSSP & intLinha).Text
  
  vetFiliais(3, bytVetColCodigoFilial) = bytVetCodigoUBE
  vetFiliais(3, bytVetColEstoque) = Range(strColEstoqueUBE & intLinha).Text
  vetFiliais(3, bytVetColDiasParados) = Range(strColDiasParadosUBE & intLinha).Text
  vetFiliais(3, bytVetColAprovado) = Range(strColAprovadoUBE & intLinha).Text

  vetFiliais(4, bytVetColCodigoFilial) = bytVetCodigoARA
  vetFiliais(4, bytVetColEstoque) = Range(strColEstoqueARA & intLinha).Text
  vetFiliais(4, bytVetColDiasParados) = Range(strColDiasParadosARA & intLinha).Text
  vetFiliais(4, bytVetColAprovado) = Range(strColAprovadoARA & intLinha).Text
  
  vetFiliais(5, bytVetColCodigoFilial) = bytVetCodigoPA
  vetFiliais(5, bytVetColEstoque) = Range(strColEstoquePA & intLinha).Text
  vetFiliais(5, bytVetColDiasParados) = Range(strColDiasParadosPA & intLinha).Text
  vetFiliais(5, bytVetColAprovado) = Range(strColAprovadoPA & intLinha).Text
  
  vetFiliais(6, bytVetColCodigoFilial) = bytVetCodigoTC
  vetFiliais(6, bytVetColEstoque) = Range(strColEstoqueTC & intLinha).Text
  vetFiliais(6, bytVetColDiasParados) = Range(strColDiasParadosTC & intLinha).Text
  vetFiliais(6, bytVetColAprovado) = Range(strColAprovadoTC & intLinha).Text
  
  vetFiliais(7, bytVetColCodigoFilial) = bytVetCodigoPAT
  vetFiliais(7, bytVetColEstoque) = Range(strColEstoquePAT & intLinha).Text
  vetFiliais(7, bytVetColDiasParados) = Range(strColDiasParadosPAT & intLinha).Text
  vetFiliais(7, bytVetColAprovado) = Range(strColAprovadoPAT & intLinha).Text
End Sub

Private Function retorna_sigla_filial(bytCodFilial As Integer) As String
  Select Case bytCodFilial
    Case bytVetCodigoFGA
      retorna_sigla_filial = strSiglaFilialFGA
    Case bytVetCodigoSSP
      retorna_sigla_filial = strSiglaFilialSSP
    Case bytVetCodigoUBE
      retorna_sigla_filial = strSiglaFilialUBE
    Case bytVetCodigoARA
      retorna_sigla_filial = strSiglaFilialARA
    Case bytVetCodigoPA
      retorna_sigla_filial = strSiglaFilialPA
    Case bytVetCodigoTC
      retorna_sigla_filial = strSiglaFilialTC
    Case bytVetCodigoPAT
      retorna_sigla_filial = strSiglaFilialPAT
  End Select
End Function

'RETORNA_QTDE_POR_FILIAL(E2, I2, H2)
Public Function RETORNA_QTDE_POR_FILIAL(ranFilialSendoAnalisada As Range, _
                                        ranFiliais As Range, _
                                        ranQtdeRecompra As Range)
  Dim strFiliais As String
  Dim lngQtdeRecompra As Long
  Dim strFilialSendoAnalisada As String
  Dim strFilialTemp As String
  Dim strVetorFiliais() As String
  Dim i As Integer
  Dim intPosicaoDoFechaParenteses As Integer
  
  strFilialSendoAnalisada = ranFilialSendoAnalisada.Text
  strFiliais = ranFiliais.Text
  lngQtdeRecompra = ranQtdeRecompra.Text
  If strFiliais = "-" Then
    RETORNA_QTDE_POR_FILIAL = 0
    Exit Function
  ElseIf strFiliais = strFilialSendoAnalisada Then
    RETORNA_QTDE_POR_FILIAL = lngQtdeRecompra
    Exit Function
  Else
    If InStr(1, strFiliais, "(", vbTextCompare) = 0 Then 'se nao tem um abre parenteses é porque é o nome de uma outra filial...
      RETORNA_QTDE_POR_FILIAL = 0 'retorna 0 porque já temos certeza que essa nao é a filial que estamos separando...
      Exit Function
    Else
      If InStr(1, strFiliais, ",", vbTextCompare) = 0 Then 'se nao tem vírgula é porque só vai um pouco de uma filial
        intPosicaoDoFechaParenteses = InStr(1, strFiliais, ")", vbTextCompare)
        strQtdeTemp = Mid(strFiliais, 2, intPosicaoDoFechaParenteses - 2)
        strFilialTemp = Mid(strFiliais, intPosicaoDoFechaParenteses + 2)
        If strFilialTemp = strFilialSendoAnalisada Then
          RETORNA_QTDE_POR_FILIAL = strQtdeTemp
          Exit Function
        End If
      Else
        strVetorFiliais = Split(strFiliais, ",")
        For i = 0 To UBound(strVetorFiliais, 1)
          strVetorFiliais(i) = Trim(strVetorFiliais(i))
          intPosicaoDoFechaParenteses = InStr(1, strVetorFiliais(i), ")", vbTextCompare)
          strQtdeTemp = Mid(strVetorFiliais(i), 2, intPosicaoDoFechaParenteses - 2)
          strFilialTemp = Mid(strVetorFiliais(i), intPosicaoDoFechaParenteses + 2)
          If strFilialTemp = strFilialSendoAnalisada Then
            RETORNA_QTDE_POR_FILIAL = strQtdeTemp
            Exit Function
          End If
        Next
        RETORNA_QTDE_POR_FILIAL = 0 'procurou em todas as filiais da string e nao a filial que está sendo analisada...
        Exit Function
        
      End If
    End If
  End If
End Function
