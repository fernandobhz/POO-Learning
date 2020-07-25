
Public Class Form1

    Private Ctrl As New Controlador

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AddHandler Me.Button1.Click, AddressOf Me.Ctrl.TestaPoo
    End Sub

End Class

Class Controlador

    Sub TestaPoo()
        Try
            Dim Agoge As New PessoaJuridica("Agoge", 14867123000139D, New EnderecoBrasileiro(TipoLogradouro:="Rua", Logradouro:="Manjericão", NumeroLogradouro:=90, Complemento:="Sala 202", Bairro:="Lindéia", CEP:=30690510D, Cidade:="Belo Horizonte", UF:="MG"))
            Dim Merg As New PessoaJuridica("Mecanica Especializada Reis Guimarães", 15530073000139D, New EnderecoBrasileiro("Rua", "Caramuru", 32D, String.Empty, "Bandeirantes", 32240330D, "Contagem", "MG"))


            MostraVersao.MensagemTela(Agoge)

            MsgBox(Agoge.getEndereco.getAddressLine1 & vbCrLf & Agoge.getEndereco.getAddressLine2)

            MsgBox(Merg.getEndereco.getAddressLine1 & vbCrLf & Merg.getEndereco.getAddressLine2)

            MsgBox("Implementar classe de pedido")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub
End Class

Public Interface iPessoa

    Function getNome() As String
    Sub setNome(ByVal Nome As String)

    Function getDocumento() As Decimal
    Sub setDocumento(ByVal Documento As Decimal)

    Function getCpfCnpj() As String

    Function getEndereco() As Endereco
    Sub setEndereco(ByRef Endereco As Endereco)

End Interface

Class PessoaJuridica
    Implements iPessoa, iVersao

    Private Nome As String
    Private Documento As Decimal
    Private Endereco As Endereco
    Private Versao As New Versao(1, 0, 0)

    Public Sub New(ByVal Nome As String, ByVal Documento As Decimal, ByRef Endereco As Endereco)
        Me.setNome(Nome)
        Me.setDocumento(Documento)
        Me.setEndereco(Endereco)
    End Sub

    Function getNome() As String Implements iPessoa.getNome
        Return Me.Nome
    End Function

    Sub setNome(ByVal Nome As String) Implements iPessoa.setNome
        Me.Nome = Nome
    End Sub

    Function getDocumento() As Decimal Implements iPessoa.getDocumento
        Return Me.Documento
    End Function

    Sub setDocumento(ByVal Documento As Decimal) Implements iPessoa.setDocumento
        If Len(Documento.ToString) <> 14 Then
            Throw New Exception("Documento deve conter 14 digitos")
        End If
        Me.Documento = Documento
    End Sub

    Function getCpfCnpj() As String Implements iPessoa.getCpfCnpj
        Return Format(Me.getDocumento, "00\.000\.000\/0000\-00")
    End Function

    Function foo() As String Implements iVersao.getVersion
        Return Me.Versao.getVersion
    End Function

    Function getEndereco() As Endereco Implements iPessoa.getEndereco
        Return Me.Endereco
    End Function

    Sub setEndereco(ByRef Endereco As Endereco) Implements iPessoa.setEndereco
        Me.Endereco = Endereco
    End Sub

End Class

Public Interface iVersao
    Function getVersion() As String
End Interface

Public Class MostraVersao

    Shared Sub MensagemTela(V As iVersao)
        MsgBox(V.getVersion)
    End Sub

End Class

Class Versao
    Private MajorVersion As Short
    Private MinorVersion As Short
    Private RevisionVersion As Short

    Sub New(ByVal MajorVersion As Short, ByVal MinorVersion As Short, ByVal RevisionVersion As Short)
        Me.MajorVersion = MajorVersion
        Me.MinorVersion = MinorVersion
        Me.RevisionVersion = RevisionVersion
    End Sub

    Function getVersion() As String
        Return Me.MajorVersion & "." & Me.MinorVersion & "." & Me.RevisionVersion
    End Function

End Class

Public Interface iEndereco
    Function getAddressLine1() As String
    Function getAddressLine2() As String
    Function getCity() As String
    Function getCountry() As String
    Function getZipCode() As String
End Interface

Public MustInherit Class Endereco
    Implements iEndereco

    'Com uma classe abstrata que implementa uma interface conseguimos garantir que todo objeto endereço implementará a interface endereço
    Public MustOverride Function getAddressLine1() As String Implements iEndereco.getAddressLine1
    Public MustOverride Function getAddressLine2() As String Implements iEndereco.getAddressLine2
    Public MustOverride Function getCity() As String Implements iEndereco.getCity
    Public MustOverride Function getCountry() As String Implements iEndereco.getCountry
    Public MustOverride Function getZipCode() As String Implements iEndereco.getZipCode

End Class

Public Class EnderecoBrasileiro
    Inherits Endereco

    Private TipoLogradouro As String
    Private Logradouro As String
    Private NumeroLogradouro As Integer
    Private Complemento As String
    Private Bairro As String
    Private CEP As Decimal
    Private Cidade As String
    Private UF As String

#Region "GettersAndSetters"

    Function getTipoLogradouro() As String
        Return Me.TipoLogradouro
    End Function

    Sub setTipoLogradouro(ByVal TipoLogradouro As String)
        Me.TipoLogradouro = TipoLogradouro
    End Sub

    Function getLogradouro() As String
        Return Me.Logradouro
    End Function

    Sub setLogradouro(ByVal Logradouro As String)
        Me.Logradouro = Logradouro
    End Sub

    Function getNumeroLogradouro() As Integer
        Return Me.NumeroLogradouro
    End Function

    Sub setNumeroLogradouro(ByVal NumeroLogradouro As Integer)
        Me.NumeroLogradouro = NumeroLogradouro
    End Sub

    Function getComplemento() As String
        Return Me.Complemento
    End Function

    Sub setComplemento(ByVal Complemento As String)
        Me.Complemento = Complemento
    End Sub

    Function getBairro() As String
        Return Me.Bairro
    End Function

    Sub setBairro(ByVal Bairro As String)
        Me.Bairro = Bairro
    End Sub

    Function getCEP() As Decimal
        Return Me.CEP
    End Function

    Sub setCEP(ByVal CEP As Decimal)
        If Len(CEP.ToString) <> 8 Then
            Throw New Exception("CEP deve conter 8 digitos")
        Else
            Me.CEP = CEP
        End If
    End Sub

    Function getCidade() As String
        Return Me.Cidade
    End Function

    Sub setCidade(ByVal Cidade As String)
        Me.Cidade = Cidade
    End Sub

    Function getUF() As String
        Return Me.UF
    End Function

    Sub setUF(ByVal UF As String)
        Me.UF = UF
    End Sub

#End Region

    Sub New(ByVal TipoLogradouro As String, ByVal Logradouro As String, ByVal NumeroLogradouro As Integer, ByVal Complemento As String, ByVal Bairro As String, ByVal CEP As Decimal, ByVal Cidade As String, ByVal UF As String)
        Me.setTipoLogradouro(TipoLogradouro:=TipoLogradouro)
        Me.setLogradouro(Logradouro:=Logradouro)
        Me.setNumeroLogradouro(NumeroLogradouro:=NumeroLogradouro)
        Me.setComplemento(Complemento:=Complemento)
        Me.setBairro(Bairro:=Bairro)
        Me.setCEP(CEP:=CEP)
        Me.setCidade(Cidade:=Cidade)
        Me.setUF(UF:=UF)
    End Sub

    Public Overrides Function getAddressLine1() As String
        Return Me.getTipoLogradouro & " " & Me.getLogradouro & ", " & Me.getNumeroLogradouro & " - " & Me.getComplemento
    End Function

    Public Overrides Function getAddressLine2() As String
        Return "Bairro " & Me.getBairro & " - " & Me.getCidade & "/" & Me.getUF
    End Function

    Public Overrides Function getCity() As String
        Return Me.getCidade
    End Function

    Public Overrides Function getCountry() As String
        Return "Brasil"
    End Function

    Public Overrides Function getZipCode() As String
        Return Me.CEP
    End Function

End Class
