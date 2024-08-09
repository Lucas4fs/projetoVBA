Option Explicit

Sub atualizarPlanilha()

Dim conexaoBanco As ADODB.Connection
Dim entradaBanco As ADODB.Recordset

Set conexaoBanco = New ADODB.Connection
Set entradaBanco = New ADODB.Recordset
 
conectarAccess conexaoBanco

entradaBanco.Open "CadastroFuncionarios", conexaoBanco, adOpenKeyset, adLockOptimistic

Range("A2:F1000000").ClearContents
Range("A2").CopyFromRecordset entradaBanco

conexaoBanco.Close

End Sub

Sub incluirBanco()

Dim conexaoBanco As ADODB.Connection
Dim entradaBanco As ADODB.Recordset

Set conexaoBanco = New ADODB.Connection
Set entradaBanco = New ADODB.Recordset
 
conectarAccess conexaoBanco

entradaBanco.Open "CadastroFuncionarios", conexaoBanco, adOpenKeyset, adLockOptimistic

entradaBanco.AddNew

entradaBanco.Fields("Nome").Value = userform_cadastro.caixatexto_nome.Value
If userform_cadastro.botaoopcao_feminino.Value = True Then
    entradaBanco.Fields("G�nero").Value = "Feminino"
Else
    entradaBanco.Fields("G�nero").Value = "Masculino"
End If
entradaBanco.Fields("�rea").Value = userform_cadastro.caixacomb_area.Value
entradaBanco.Fields("CPF").Value = Format(userform_cadastro.caixatexto_cpf.Value, "000"".""000"".""000-00")
entradaBanco.Fields("Sal�rio").Value = userform_cadastro.caixatexto_salario.Value

entradaBanco.Update

conexaoBanco.Close

Call atualizarPlanilha

End Sub

Sub editarBanco()

Dim conexaoBanco As ADODB.Connection
Dim entradaBanco As ADODB.Recordset
Dim textoTabela As String

Set conexaoBanco = New ADODB.Connection
Set entradaBanco = New ADODB.Recordset
 
conectarAccess conexaoBanco

textoTabela = "Select * From CadastroFuncionarios Where ID = " & CLng(userform_cadastro.caixatexto_id.Value)

entradaBanco.Open textoTabela, conexaoBanco, adOpenKeyset, adLockOptimistic

If userform_cadastro.caixatexto_nome.Value <> "" Then entradaBanco.Fields("Nome").Value = userform_cadastro.caixatexto_nome.Value
If userform_cadastro.botaoopcao_feminino.Value = True Then
    entradaBanco.Fields("G�nero").Value = "Feminino"
ElseIf userform_cadastro.botaoopcao_masculino.Value = True Then
    entradaBanco.Fields("G�nero").Value = "Masculino"
End If
If userform_cadastro.caixacomb_area.Value <> "" Then entradaBanco.Fields("�rea").Value = userform_cadastro.caixacomb_area.Value
If userform_cadastro.caixatexto_cpf.Value <> "" Then entradaBanco.Fields("CPF").Value = Format(userform_cadastro.caixatexto_cpf.Value, "000"".""000"".""000-00")
If userform_cadastro.caixatexto_salario.Value <> "" Then entradaBanco.Fields("Sal�rio").Value = userform_cadastro.caixatexto_salario.Value

entradaBanco.Update

conexaoBanco.Close

Call atualizarPlanilha

End Sub

Sub excluirBanco()

Dim conexaoBanco As ADODB.Connection
Dim entradaBanco As ADODB.Recordset
Dim textoTabela As String

Set conexaoBanco = New ADODB.Connection
Set entradaBanco = New ADODB.Recordset
 
conectarAccess conexaoBanco

textoTabela = "Select * From CadastroFuncionarios Where ID = " & CLng(userform_cadastro.caixatexto_id.Value)

entradaBanco.Open textoTabela, conexaoBanco, adOpenKeyset, adLockOptimistic

entradaBanco.Delete
entradaBanco.Update

conexaoBanco.Close

Call atualizarPlanilha

End Sub