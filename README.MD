<p align="center">
    <img src="src\imagens\capa.jpg">
    <br>
    <h1 align="center">
    🔃 INTEGRAÇÃO DO EXCEL COM BANCO DE DADOS ACCESS" 🔃
    </h1>
</p>
<br>
<h2>
:pushpin: DOWNLOAD DOS ARQUIVOS UTILIZADOS NO LINK ABAIXO :pushpin:
</h2>

```
https://drive.google.com/drive/folders/1YEnjvOkJfHUwjiTS0gl9g3tTnD3oWVlI?usp=sharing
``` 

<h3>
    :pushpin: INTUITO DO PROJETO
</h3>

<p>
   O intuito desse projeto é mostrar como integrar uma planilha com um banco de dados. As inserções, atualizações e deleções que acontecerem dentro da planilha irão refletir dentro do banco de dados.<br>
</p>

<h3>
    :pushpin: FERRAMENTAS UTILIZADAS
</h3>

- **Microsoft Excel:**
  - Usado para criar e manipular a planilha.

- **Microsoft Access:**
  - Banco de dados usado para armazenar e refletir os dados que estão sendo manipulados na planilha.

- **Microsoft Visual Basic for Applications (VBA):**
  - Implementação do Excel usada para escrever os códigos.

- **Visual Basic**

  - Linguagem de programação usada para fazer a integração do banco de dados com a planilha.

- **SQL**
  - Linguagem de consulta estruturada para realizar consultas no banco de dados a partir do momento que funções em VBA forem chamadas.

<h3>
    :pushpin: INICIANDO AS CONFIGURAÇÕES
</h3>

<p>
   Com a planilha aberta e os dados inseridos teclamos o atalho <strong>ALT + F11</strong>, será aberta a janela do Visual Basic, na parte superior esquerda ativamos a janela <strong>Propiedades</strong>.
</p>

<p>
    <img src = "src\imagens\exibirJanelaPropiedades.png">
</p>

<p>
   Em seguida crie e dê nome aos módulos responsáveis por conectar as ações que acontecem dentro da planilha com a tabela do banco de dados.
</p>

<p>
    <img src = "src\imagens\criandoEnomeandoModulo.png">
</p>

<p>
    É necessário criar três módulos:
</p>

<p>

- **modConectarAccess:**
  - Responsável por estabelecer a conexão do banco de dados com a planilha.

- **modExibirFormulario:**
  - Responsável por exibir o formulário que já está inserido no projeto (a janela usada para manipular os dados da planilha).

- **modInteracaoAccess:**
  - Responsável por atualizar a planilha com as informações do banco de dados todas as vezes que a mesma for aberta e fazer DML (inserir, atualizar e excluir) nos dados da tabela dentro do banco de dados de acordo com a manipulação que ocorre na planilha.

</p>

<h3>
    :pushpin: INSERINDO CÓDIGO VBA DENTRO DOS MÓDULOS
</h3>

- **modConectarAccess:**

```vba
Option Explicit

Function conectarAccess(conexaoBanco As ADODB.Connection)

Dim textoConexao As String, caminhoBanco As String

caminhoBanco = ThisWorkbook.Path & "\BancoDeDadosVBA.accdb"

textoConexao = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & caminhoBanco & ";" & _
                "Persist Security Info=False;"

conexaoBanco.Open (textoConexao)


End Function
```

- **modExibirFormulario:**

```vba
Sub carregar_userformcadastro()

userform_cadastro.Show

End Sub
```

- **modInteracaoAccess:**

```vb
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
    entradaBanco.Fields("Gênero").Value = "Feminino"
Else
    entradaBanco.Fields("Gênero").Value = "Masculino"
End If
entradaBanco.Fields("Área").Value = userform_cadastro.caixacomb_area.Value
entradaBanco.Fields("CPF").Value = Format(userform_cadastro.caixatexto_cpf.Value, "000"".""000"".""000-00")
entradaBanco.Fields("Salário").Value = userform_cadastro.caixatexto_salario.Value

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
    entradaBanco.Fields("Gênero").Value = "Feminino"
ElseIf userform_cadastro.botaoopcao_masculino.Value = True Then
    entradaBanco.Fields("Gênero").Value = "Masculino"
End If
If userform_cadastro.caixacomb_area.Value <> "" Then entradaBanco.Fields("Área").Value = userform_cadastro.caixacomb_area.Value
If userform_cadastro.caixatexto_cpf.Value <> "" Then entradaBanco.Fields("CPF").Value = Format(userform_cadastro.caixatexto_cpf.Value, "000"".""000"".""000-00")
If userform_cadastro.caixatexto_salario.Value <> "" Then entradaBanco.Fields("Salário").Value = userform_cadastro.caixatexto_salario.Value

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
```

<p>
Após criar os módulos e inserir os códigos basta salvar o arquivo e a integração já está pronta.
</p>
