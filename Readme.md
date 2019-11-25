# Add bancos

- DADOS PREMATRICULA
- DADOS GERADOR RELATORIO

# Add Form

- Form_Pre

# planCad

```vb
Private Sub ChBox_PreMatricula_Click()
    Run "PreMatricula_Show"
End Sub
```

# Form_Pre

```vb
Private Sub TB_matricula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 9
    End If
    If KeyCode = 9 Then ' Tab
        If Not (IsDate(Form_Pre.TB_matricula.Value)) Then
            MsgBox "Coloque uma data válida", vbInformation, "Pré-matrícula"
            Form_Pre.TB_matricula.Value = ""
            Form_Pre.TB_matricula.SetFocus
        Else
            Form_Pre.TB_matricula.Value = Format(Form_Pre.TB_matricula.Value, "mm/dd/yyyy")
            Form_Pre.TB_prematricula.Value = Format(Form_Pre.TB_prematricula.Value, "mm/dd/yyyy")
            Run "PreMatricula"
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Form_Pre.TB_matricula.Value = "" Then
        planCad.Range("Cad_40").Value = Format(Date, "dd/mm/yyyy")
        planCad.ChBox_PreMatricula.Value = False
    End If
End Sub
```

# BackupModulo

## Sub Backup()

```diff
    Dim ws As Worksheet
-    For Each ws In Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA"))    
+    For Each ws In Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA", "DADOS PREMATRICULA"))
     
        Call Protecao(False, ws)
     
    Next
    
    ' Copia Dados
-    Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA")).Copy
+    Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA", "DADOS PREMATRICULA")).Copy
    With ActiveWorkbook
         .SaveAs Filename:=nomeLog, FileFormat:=xlOpenXMLWorkbook
         .Close savechanges:=False
    End With
-    For Each ws In Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA"))    
+    For Each ws In Worksheets(Array("DADOS CADASTRO", "DADOS PAGAMENTOS", "DADOS GASTOS", "DADOS COBRANÇAS", "DADOS DESCONTOS", "DADOS ACRÉSCIMOS", "DADOS COBRANÇAS EXTRAS", "TABELAS AUXILIARES", "DADOS VENDA VISTA", "DADOS RECEITA EXTRA", "DADOS PREMATRICULA"))
     
        Call Protecao(True, ws)
     
    Next
```

# Cadastro

## Sub BuscaAluno

```diff
    If visibleRows > 0 Then
    
        Dim rowFiltrado As Integer
    
        rowFiltrado = tbDCad.DataBodyRange.Columns.SpecialCells(xlCellTypeVisible).row
        
    
        ' Zera valores
        Dim a As Integer
        For a = 0 To 55
            planCad.Range("Cad_" & a).Value = ""
        Next


        ' Coloca valores de Dados Cadastro em Cadastro
        For a = 0 To 55
            planCad.Range("Cad_" & a) = planDCad.Cells(rowFiltrado, tbDCad.ListColumns("Matricula").Index + 1).Offset(0, a)
        Next
        
        
+        ' Verifica se teve pré-matrícula. Se sim, printa
+        Dim tbDPre As ListObject
+        Set tbDPre = planDPre.ListObjects("TabelaDadosPrematricula")
+        
+        On Error Resume Next
+        Dim rowPre As Integer
+        rowPre = WorksheetFunction.Match(planCad.Range("Cad_0").Value, tbDPre.ListColumns(1).DataBodyRange, 0)
+        If Err.Number <> 0 Then
+            Err.Clear
+        End If
+        
+        If rowPre > 0 Then
+            planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorAccent6
+            planCad.Range("Cel_Pre").Value = tbDPre.DataBodyRange.Cells(rowPre, 2).Value
+        End If
+        
+        
+        ' Esconde link para pré-matrícula
+        planCad.Shapes("TB_Pre").Visible = msoFalse
        
    Else
    
        MsgBox "Não foi encontrado nenhum cadastro", , "Aviso"
        
        tbDCad.DataBodyRange.AutoFilter
        
        Exit Sub
        
    End If
```

## Sub EscolheNomeListBox

```diff
    For a = 0 To 55
        planCad.Range("Cad_" & a) = planDCad.Cells(rowNomeEscolhido, tbDCad.ListColumns("Matricula").Index + 1).Offset(0, a)
    Next
    

    tbDCad.DataBodyRange.AutoFilter
    
    
+    ' Verifica se teve pré-matrícula. Se sim, printa
+    Dim tbDPre As ListObject
+    Set tbDPre = planDPre.ListObjects("TabelaDadosPrematricula")
+    
+    On Error Resume Next
+    Dim rowPre As Integer
+    rowPre = WorksheetFunction.Match(planCad.Range("Cad_0").Value, tbDPre.ListColumns(1).DataBodyRange, 0)
+    If Err.Number <> 0 Then
+        Err.Clear
+    End If
+    
+    If rowPre > 0 Then
+        planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorAccent6
+        planCad.Range("Cel_Pre").Value = tbDPre.DataBodyRange.Cells(rowPre, 2).Value
+    End If
        

    Run ("TelaEditar") ' Coloca a tela Cadastro em modo de Edição
```

## Sub novoAluno

```diff

+    ' Validação: matrícula num formato diferente
+    Dim matr As Long
+    matr = planCad.Range("Cad_0").Value
+    If (matr < 20160001) Or (matr > 205000001) Then
+        MsgBox "Matrícula em formato inválido", vbInformation, "Erro ao Salvar"
+        planCad.Range("Cad_0").Select
+        Exit Sub
+    End If

    ' Insere dados do Cadastro
    Dim i As Integer
    
    For i = 0 To 54
        tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Matricula").Index).Offset(0, i) = planCad.Range("Cad_" & i).Value
    Next
    
    
+    ' Insere dado pré-matrícula
+
+    If planCad.Range("Cel_Pre").Value <> "" Then
+        ' Última linha da tabela + 1
+        Dim row As Integer, tbDPre As ListObject
+        Set tbDPre = planDPre.ListObjects("TabelaDadosPrematricula")
+        If WorksheetFunction.CountA(Range("TabelaDadosPrematricula")) = 0 Then ' Caso seja primeiro cadastro
+            row = 1
+        Else
+            row = tbDPre.DataBodyRange.Rows.Count + 1
+        End If
+    
+        ' Insere dados
+        tbDPre.DataBodyRange.Cells(row, 1).Value = planCad.Range("Cad_0").Value
+        tbDPre.DataBodyRange.Cells(row, 2).Value = planCad.Range("Cel_Pre").Value
+    End If
    
    
    ' Formata data da tabela
    tbDCad.ListColumns("Nascimento").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbDCad.ListColumns("DataMatricula").DataBodyRange.NumberFormat = "dd/mm/yyyy"
```

## Sub EditaAluno

```diff
+    ' Validação: matrícula num formato diferente
+    If (matricula < 20160001) Or (matricula > 205000001) Then
+        MsgBox "Matrícula em formato inválido", vbInformation, "Erro ao Salvar"
+        planCad.Range("Cad_0").Select
+        Exit Sub
+    End If

    Dim rowFiltrado As Integer
    rowFiltrado = tbDCad.DataBodyRange.Columns.SpecialCells(xlCellTypeVisible).row
    
+    ' Validação: não pode mudar matrícula e salvar, pois as informações de um cadastro iriam sobrescrever de outro
+    If (planCad.Range("Cad_0").Value <> planDCad.Cells(rowFiltrado, tbDCad.ListColumns(1).Index + 1).Value) Then
+        MsgBox "Não edite número de matrícula", vbCritical, "Atenção"
+        tbDCad.DataBodyRange.AutoFilter
+        Exit Sub
+    End If
```

## Sub removeAluno

```diff
+    ' Remove Dados Pré-matrícula
+    tbDPre.DataBodyRange.AutoFilter field:=1, Criteria1:=matricula
+    visibleRows = Application.WorksheetFunction.Subtotal(103, tbDPre.ListColumns(1).DataBodyRange)
+    
+    If visibleRows > 0 Then
+        tbDPre.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
+    End If
+    
+    tbDPre.DataBodyRange.AutoFilter
```

## Sub LimpaTelaCadastro

```diff
    Run ("ZeraConsultaCobrancasExtras3Meses")
    Run ("ZeraConsultaCobrancasExtrasMes")
    
    ' Gera nova matrícula
    Range("Cad_0").Value = geraMatricula()
    
+    Range("Cad_40") = Format(Date, "mm/dd/yyyy") ' Data da Matrícula
    
    Range("Cad_41") = "Sim" ' Aluno Ativo

    
+    ' Deixa desmarcado checkbox pré-matrícula
+    planCad.ChBox_PreMatricula.Value = False   
+    
+    ' Deixa invisível Pré-matrícula
+    planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorDark1
+    planCad.Range("Cel_Pre").Value = ""
```

## Sub LimpaTelaCadastroTelaEditar

```diff
+    ' Deixa desmarcado checkbox pré-matrícula
+    planCad.ChBox_PreMatricula.Value = False
+    
+    
+    ' Deixa invisível Pré-matrícula
+    planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorDark1
+    planCad.Range("Cel_Pre").Value = ""
```

## Function geraMatricula()

```vb
    ' Verifica última matrícula, incrementa + 1 para nova matrícula
    
    Dim rangeMatriculas As Range
    Dim maiorMatricula As Long, novaMatricula As Long
    Dim tbDCad As ListObject
    Set tbDCad = planDCad.ListObjects("TabelaDadosCadastro")
    Set rangeMatriculas = tbDCad.ListColumns("Matricula").DataBodyRange
    
    
    ' Maior matrícula de todas
    On Error Resume Next ' Erro caso ainda não haja nenhum cadastro
    maiorMatricula = WorksheetFunction.Max(rangeMatriculas)
    If Err.Number <> 0 Then
        Err.Clear
        geraMatricula = year(Now) & "0001"
    End If
    
    
    ' Ano da maior matrícula
    Dim anoMaior As Integer, anoAtual As Integer
    anoMaior = Left(maiorMatricula, 4)
    anoAtual = year(Now)
    
    If ((anoMaior - anoAtual) > 1) Then MsgBox "Existe matrícula cadastrada para daqui a 2 anos ou mais", vbCritical, "Cuidado"
    
    
    ' Verifica total de alunos.
    Dim totalAtivos As Integer, totalAtivosString As String
    totalAtivos = planFin.Range("Cel_TotalAtivos").Value
    totalAtivosString = CStr(totalAtivos)
    Dim nCasas As Integer
    nCasas = Len(totalAtivosString)
    If Not (nCasas > 0) Then
        MsgBox "Não há alunos ativos no painel financeiro", vbCritical, "Atenção"
        geraMatricula = anoAtual & "0001"
    End If
    
        
    ' virou de ano e ninguém tinha matrícula do próximo ano
    If (anoAtual > anoMaior) Then
        If nCasas <= 3 Then
            novaMatricula = anoAtual & "0001"
        Else
            novaMatricula = anoAtual & "00001"
        End If
        
        geraMatricula = novaMatricula
        
    ' existe matrícula cadastrada para o próximo ano já
    ElseIf anoAtual < anoMaior Then
        ' Filtra todas as matrículas do ano atual
        Dim matrMin As Long, matrMax As Long
        matrMin = anoAtual * (10 ^ nCasas) & 1
        matrMax = (anoAtual + 1) * (10 ^ nCasas) & 0
        rangeMatriculas.AutoFilter field:=1, Criteria1:=">=" & matrMin, Operator:=xlAnd, Criteria2:="<" & matrMax
        
        ' Acha a maior entre as filtradas
        Dim matrMaiorFiltrada As Long
        matrMaiorFiltrada = WorksheetFunction.Max(tbDCad.ListColumns("Matricula").DataBodyRange.SpecialCells(xlCellTypeVisible))
        tbDCad.DataBodyRange.AutoFilter
        geraMatricula = matrMaiorFiltrada + 1
    
    ' ano atual = maior ano de matrícula
    Else
        geraMatricula = maiorMatricula + 1
    End If
    
End Function
```

## Sub TelaNovo

```diff
+    ' Mostra link para pré-matrícula se já estiver na tela Dados Escolares
+    Dim sh As Shape
+    Set sh = planCad.Shapes("btnCadDadosEscolares")
+    If sh.Fill.ForeColor.RGB = RGB(84, 197, 203) Then
+        planCad.Shapes("TB_Pre").Visible = msoTrue
+    End If
```

## Sub TelaEditar

```diff
+    ' Esconde link para pré-matrícula
+    planCad.Shapes("TB_Pre").Visible = msoFalse
```

## Sub PreMatricula_Show

```vb
Sub PreMatricula_Show()
    
    Dim cBox As Variant
    Set cBox = planCad.ChBox_PreMatricula
    
    If Not (cBox.Value) Then
        planCad.Range("Cad_0").Value = geraMatricula()
        planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorDark1
        planCad.Range("Cel_Pre").Value = ""
        planCad.Range("Cad_40").Value = Format(Date, "dd/mm/yyyy")
        Exit Sub
    End If
    
    
    ' Abre formulário com data de pré-matrícula e de matrícula
    Load Form_Pre
    
    Form_Pre.TB_prematricula.Value = Format(Date, "dd/mm/yyyy")
    Form_Pre.TB_matricula.SetFocus
    
    planCad.Range("Cad_40").Value = ""
    
    Form_Pre.Show
    
End Sub
```

## Sub PreMatricula

```vb
Sub PreMatricula()
' Esta macro roda ao confirmar o Form_Pre

    ' Validação: pré-matrícula não pode ser no mesmo mês da matrícula
    Dim dataMatr As Date, dataPre As Date, difDate As Integer, dataPreMax As Date
    dataMatr = Format(Form_Pre.TB_matricula.Value, "mm/dd/yyyy")
    dataPre = Form_Pre.TB_prematricula.Value
    dataPreMax = Format(DateAdd("m", 1, dataPre), "mm/yyyy")
  
    If dataMatr < dataPreMax Then
        MsgBox "Data de matrícula e pré-matrícula precisam estar em meses diferentes.", vbInformation, "Pré-Matrícula"
        Form_Pre.TB_matricula.Value = Format(Form_Pre.TB_matricula.Value, "dd/mm/yyyy")
        Form_Pre.TB_prematricula.Value = Format(Form_Pre.TB_prematricula.Value, "dd/mm/yyyy")
        planCad.ChBox_PreMatricula.Value = False
        Form_Pre.TB_matricula.Value = ""
        Form_Pre.TB_matricula.SetFocus
        Exit Sub
    End If
    

    ' Coloca data de matrícula
    'planCad.Range("Cad_40").Value = Form_Pre.TB_matricula.Value
    planCad.Range("Cad_40").Value = dataMatr
    
    planCad.Range("Cel_Pre_Label").Font.ThemeColor = xlThemeColorAccent6
    planCad.Range("Cel_Pre").Value = Form_Pre.TB_prematricula.Value
    planCad.Range("Cel_Pre").Font.ThemeColor = xlThemeColorAccent6

    
    ' Esconde link Pre-matrícula
    planCad.Shapes("TB_Pre").Visible = msoFalse
    
    
    ' Atualiza matrícula:
    Dim anoMatr As Integer, mesMatr As Integer
    anoMatr = year(dataMatr)
    mesMatr = Month(dataMatr)


    ' Verifica total de alunos.
    Dim totalAtivos As Integer, totalAtivosString As String
    totalAtivos = planFin.Range("Cel_TotalAtivos").Value
    totalAtivosString = CStr(totalAtivos)
    Dim nCasas As Integer
    nCasas = Len(totalAtivosString)
    If Not (nCasas > 0) Then
        MsgBox "Não há alunos ativos no painel financeiro", vbCritical, "Atenção"
    End If


    ' Verifica se já existe alguma matrícula do próximo ano já cadastrada. Se sim, escolhe a maior. Senão, cria nova

    ' Filtra todas as matrículas do próximo ano
    Dim tb As ListObject
    Set tb = planDCad.ListObjects("TabelaDadosCadastro")
    Dim matrMin As Long, matrMax As Long
    matrMin = (anoMatr) * (10 ^ nCasas) & 1
    matrMax = (anoMatr + 1) * (10 ^ nCasas) & 0
    tb.DataBodyRange.AutoFilter field:=1, Criteria1:=">=" & matrMin, Operator:=xlAnd, Criteria2:="<" & matrMax

    ' Acha a maior entre as filtradas
    Dim visibleRows As Integer
    Dim novaMatr As Long
    visibleRows = Application.WorksheetFunction.Subtotal(103, tb.ListColumns(1).DataBodyRange)
    If visibleRows < 1 Then ' não tem matrícula do próximo ano
        tb.DataBodyRange.AutoFilter
        novaMatr = (anoMatr) * (10 ^ nCasas) & 1
    Else ' já existem matrículas para o próximo ano cadastradas
        Dim matrMaiorFiltrada As Long
        matrMaiorFiltrada = WorksheetFunction.Max(tb.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible))
        tb.DataBodyRange.AutoFilter
        novaMatr = matrMaiorFiltrada + 1
    End If

    planCad.Range("Cad_0").Value = novaMatr


    ' Fecha form
    Unload Form_Pre

End Sub
```

## Sub TurnOnPre

```vb
Sub TurnOnPre()
    planCad.ChBox_PreMatricula.Value = True
End Sub
```

## Sub MakeCheckBoxPreInvisible

```vb
Sub MakeCheckBoxPreInvisible()
    planCad.ChBox_PreMatricula.Visible = False
End Sub
```

# Configurações

## RestaurarBackup

```diff
+    If WorksheetExists("DADOS PREMATRICULA", Workbooks(nomeBackupRestaura)) Then
+        Workbooks(nomeSistema).Sheets("DADOS PREMATRICULA").UsedRange.Clear
+        Workbooks(nomeBackupRestaura).Sheets("DADOS PREMATRICULA").UsedRange.Copy _
+            Workbooks(nomeSistema).Sheets("DADOS PREMATRICULA").Range(Workbooks(nomeSistema).Sheets("DADOS RECEITA EXTRA").UsedRange.Cells(1).Address)
+    End If
```

# Funções Prontas

```vb
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function isLocal() As Boolean
' Verifica se Excel está em portugues ou ingles

    Dim lang_code As Long
    lang_code = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If lang_code = 1046 Then
        isLocal = True
    Else
        isLocal = False
    End If
End Function
```

# Navegação

## Subs AbaDadosPessoaisCadastro, AbaContato, AbaAnamnese, AbaCobrancasFixasCadastro, AbaCobrancasExtrasCadastro

```diff
+    ' Esconde pré-matrícula
+    planCad.Shapes("TB_Pre").Visible = msoFalse
```

## Sub AbaDadosEscolaresCadastro

```diff
+    ' Mostra pré-matrícula
+    If (planCad.Shapes("btnNovo").Visible = msoTrue) And (planCad.Range("Cel_Pre").Value = "") Then
+        planCad.Shapes("TB_Pre").Visible = msoTrue
+    End If
```

# StatusPagamento

## Sub StatusNovoAluno

```diff
+    ' Coloca N/A para os meses anteriores ao da matrícula/pré-matrícula
+    Dim dataMatr As Date, dataCel As Date, anoCel As Integer
+    
+    Dim dataPre As Date
+    If planCad.Range("Cel_Pre").Value <> "" Then
+        dataPre = Format(planCad.Range("Cel_Pre").Value, "mm/yyyy")
+    Else
+        dataPre = 0
+    End If
+    
+    dataMatr = planCad.Range("Cad_40")
+    dataMatr = Format(dataMatr, "mm/yyyy")
+    If planStatusPag.Range("Status_Ano").Value <> "" Then
+        anoCel = planStatusPag.Range("Status_Ano").Value
+    Else
+        anoCel = year(Date)
+    End If
+   
+    For i = 1 To 12
+        
+        dataCel = Format(i & "/" & anoCel, "mm/yyyy")
+        
+        ' Coloca N/A
+        If (dataCel < dataMatr) Or (dataCel < dataPre) Then
+        ' Aluno ainda não era matriculado nesta data
+            tbStatus.DataBodyRange.Cells(rowTbStatus, i + 2).Interior.Color = RGB(117, 113, 113)
+            tbStatus.DataBodyRange.Cells(rowTbStatus, i + 2) = "N/A"
+        End If
+        
+        ' Coloca Pré
+        If dataPre <> 0 Then
+            If (dataCel >= dataPre) And (dataCel < dataMatr) Then
+                tbStatus.DataBodyRange.Cells(rowTbStatus, i + 2).Interior.Color = RGB(146, 208, 80)
+                tbStatus.DataBodyRange.Cells(rowTbStatus, i + 2) = "Pré"
+            End If
+        End If
+    Next
```

## Sub StatusEditaAluno

```diff
+    ' Coloca fórmula do Status
+    If (isLocal()) Then
+        tbStatus.ListColumns("Status").DataBodyRange.FormulaLocal = _
+            "=SE([@1]=""ALUNO INATIVO"";""Inativo"";SE(CONT.SE(TabelaStatus[@[1]:[12]];""Sem Pgto"");""Sem Pgto"";SE(CONT.SE(TabelaStatus[@[1]:[12]];""Deve"");""Deve""; SE(CONT.SE(TabelaStatus[@[1]:[12]];""Quitado"");""Quitado"";SE(CONT.SE(TabelaStatus[@[1]:[12]];""Pré"");""Pré"";""N/A"")))))"
+    Else
+        tbStatus.ListColumns("Status").DataBodyRange.FormulaLocal = _
+            "=IF([@1]=""ALUNO INATIVO"";""Inativo"";IF(COUNTIF(TabelaStatus[@[1]:[12]];""Sem Pgto"");""Sem Pgto"";IF(COUNTIF(TabelaStatus[@[1]:[12]];""Deve"");""Deve""; IF(COUNTIF(TabelaStatus[@[1]:[12]];""Quitado"");""Quitado"";IF(COUNTIF(TabelaStatus[@[1]:[12]];""Pré"");""Pré"";""N/A"")))))"
+    End If
```

## Sub AtualizaStatus

```diff
    Else
        ' Aluno Ativo
        
            On Error Resume Next
            dataMatr = Format(tbDCad.DataBodyRange.Cells(j, colMatr), "mm/yyyy")
            If Err.Number <> 0 Then
                MsgBox "Atenção, houve um erro ao atualizar o status. Verifique se o cadastro " & celStatus.Value & " está com a DATA DE MATRÍCULA preenchida.", vbCritical, "ATENÇÃO"
                Err.Clear
                Exit Sub
            End If
            
            
+            ' Verifica se aluno tem pré-matrícula
+            On Error Resume Next
+            rowPre = WorksheetFunction.Match(tbDCad.DataBodyRange.Cells(j, 1).Value, tbDPre.ListColumns(1).DataBodyRange, 0)
+            If Err.Number <> 0 Then
+                Err.Clear
+            End If
+            
+            If rowPre > 0 Then
+                dataPre = tbDPre.DataBodyRange.Cells(rowPre, 2).Value
+            Else
+                dataPre = 0
+            End If
+                        
+            For i = 1 To 12
+                
+                dataCel = Format(i & "/" & anoCel, "mm/yyyy")
+                
+                ' Coloca N/A
+                If (dataCel < dataMatr) Or (dataCel < dataPre) Then
+                    
+                    planStatusPag.Cells(celStatus.row, i + 3).Interior.Color = RGB(117, 113, 113)
+                    planStatusPag.Cells(celStatus.row, i + 3) = "N/A"
+                End If
+                
+                ' Coloca Pré
+                If dataPre <> 0 Then
+                    If (dataCel >= dataPre) And (dataCel < dataMatr) Then
+                        planStatusPag.Cells(celStatus.row, i + 3).Interior.Color = RGB(146, 208, 80)
+                        planStatusPag.Cells(celStatus.row, i + 3) = "Pré"
+                    End If
+                End If
                
                ' Coloca Sem Pgto
                If dataCel <= dataHoje Then
                    If planStatusPag.Cells(celStatus.row, i + 3).Interior.ColorIndex = xlNone Then
                        planStatusPag.Cells(celStatus.row, i + 3).Interior.Color = RGB(237, 125, 49)
                        planStatusPag.Cells(celStatus.row, i + 3) = "Sem Pgto"
                    End If
                End If
                
            Next
        
        End If
```