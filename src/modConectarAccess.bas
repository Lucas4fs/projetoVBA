Option Explicit

Function conectarAccess(conexaoBanco As ADODB.Connection)

Dim textoConexao As String, caminhoBanco As String

caminhoBanco = ThisWorkbook.Path & "\BancoDeDadosVBA.accdb"

textoConexao = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & caminhoBanco & ";" & _
                "Persist Security Info=False;"

conexaoBanco.Open (textoConexao)


End Function