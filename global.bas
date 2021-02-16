Attribute VB_Name = "Global"
Option Explicit
'pra ler o arquivo ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'dados conexão
Dim conexao As ADODB.Connection
Dim sServidor   As String
Dim sBanco      As String
Dim sUsuario    As String
Dim sSenha      As String

'pra gravar o log
Global sMsg         As String
Global sTextoTela   As String

'XML
Dim objXML      As XMLHTTP60
Dim objDOMPais  As DOMDocument60

'Dados de País
Dim sISOCodeC           As String
Dim sNameCountry        As String
Dim sCapitalCity        As String
Dim sPhoneCode          As String
Dim sContinentCode      As String
Dim sCurrencyISOCode    As String
Dim sCountryFlag        As String

'Dados de Idioma
Dim sISOCodeL       As String
Dim sNameLanguage   As String

Function RetirarCaracteresInvalidos(t As String) As String

    Dim p As Integer
    Dim c As String * 1
    Dim invalidos As String
    
    On Error GoTo tratarErro

    invalidos = "-&\*!@#$%()_-=+}][{^~?<>;|'"

    'percorre todos os caracteres do string
    For p = 1 To Len(t)
        'captura o caractere
        c = Mid(t, p, 1)
        'se o caractere não existe no string de caracteres inválidos, concatena-o ao string de retorno
        If InStr(1, invalidos, c, 1) = 0 Then
            RetirarCaracteresInvalidos = RetirarCaracteresInvalidos & c
        End If
    Next
    Exit Function
    
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: RetirarCaracteresInvalidos"
    GravarLog sMsg
End Function

Function BaixarDados() As Boolean
    
    On Error GoTo tratarErro
      
    BaixarDados = False
             
    'busca dados da api
    Set objXML = New XMLHTTP60
    objXML.Open "GET", "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries", False
    objXML.setRequestHeader "Accept", "application/xml"
    objXML.send

    If objXML.Status < 200 Or objXML.Status >= 300 Then
        sMsg = "Erro HTTP:" & objXML.Status & " - Detalhes: " & objXML.responseText
        GravarLog sMsg
        Exit Function
    End If
      
    Set objDOMPais = New DOMDocument60
    objDOMPais.resolveExternals = True
    objDOMPais.validateOnParse = True
    
    objDOMPais.async = False
    objDOMPais.loadXML objXML.responseText
        
    If objDOMPais.parseError.reason <> "" Then
       sMsg = objDOMPais.parseError.reason
       GravarLog sMsg
       Exit Function
    End If
        
    If IsEmpty(objXML) Then
       sMsg = "Não baixou nenhum dado."
       GravarLog sMsg
       Exit Function
    End If
    
    'Exibe apenas Países cujo ISOCode inicia com a letra A
    FiltrarDadosApresentacao
    
    sMsg = "Dados baixados com sucesso."
    GravarLog sMsg
    
    BaixarDados = True
        
    Exit Function
    
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: BaixarDados"
    GravarLog sMsg
End Function
Sub FiltrarDadosApresentacao()
   
    Dim sLinha1 As String
    Dim sLinha2 As String
       
    Dim objLista            As IXMLDOMNodeList
    Dim objElemento         As IXMLDOMElement
    Dim objElementoFilho    As IXMLDOMElement
       
    On Error GoTo tratarErro
    
    sTextoTela = ""
    
    If objDOMPais.parseError.reason <> "" Then Exit Sub
    
    Set objLista = objDOMPais.selectNodes("//tCountryInfo")
    
    For Each objElemento In objLista
        sLinha1 = ""
        sLinha2 = ""
        'Filtra apenas países com ISOCode iniciado com a letra A
        If Left(objElemento.Text, 1) = "A" Then
         
            sLinha1 = "ISOCode:" & objElemento.childNodes(0).nodeTypedValue & " " _
                    & "PhoneCode:" & objElemento.childNodes(3).nodeTypedValue & " " _
                    & "ContinentCode:" & objElemento.childNodes(4).nodeTypedValue & " " _
                    & "CurrencyISOCode:" & objElemento.childNodes(5).nodeTypedValue
                                                               
            sLinha2 = "Name:" & Trim(RetirarCaracteresInvalidos(objElemento.childNodes(1).nodeTypedValue)) & " " _
                    & "CapitalCity:" & Trim(RetirarCaracteresInvalidos(objElemento.childNodes(2).nodeTypedValue)) & " " _
                    & "CountryFlag:" & Trim(objElemento.childNodes(6).nodeTypedValue)
            
            sTextoTela = sTextoTela & sLinha1 & vbCrLf & sLinha2 & vbCrLf & "Languages:"
                                           
            ApresentarIdiomasDoPais objElemento.childNodes(7)
            
        End If
    Next
    Exit Sub
      
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: FiltrarDadosApresentacao"
    GravarLog sMsg
End Sub
Sub ApresentarIdiomasDoPais(ByRef objDOMElemento As IXMLDOMElement)

    Dim objElemento As IXMLDOMElement
    Dim sLinha3 As String
    
    On Error GoTo tratarErro
    
    sLinha3 = ""
  
    For Each objElemento In objDOMElemento.childNodes
      
        sLinha3 = sLinha3 & objElemento.childNodes(0).nodeTypedValue & "-" _
                & objElemento.childNodes(1).nodeTypedValue & ","
    
    Next
    
    'tira a última vírgula
    If Trim(sLinha3) <> "" Then sLinha3 = Left(sLinha3, Len(sLinha3) - 1)
    sTextoTela = sTextoTela & sLinha3 & vbCrLf & vbCrLf
      
    Exit Sub
    
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: ApresentarIdiomasDoPais"
    GravarLog sMsg
End Sub

Sub LimparBanco()
    Dim lRegAfetados As Long
    
    On Error GoTo tratarErro
    
    conexao.Execute "DELETE FROM CountryLanguage", lRegAfetados, adExecuteNoRecords
    sMsg = "Registros CountryLanguage excluídos: " & lRegAfetados
    GravarLog sMsg
        
    conexao.Execute "DELETE FROM Country", lRegAfetados, adExecuteNoRecords
    sMsg = "Registros Country excluídos: " & lRegAfetados
    
    GravarLog sMsg
    
    conexao.Execute "DELETE FROM Language", lRegAfetados, adExecuteNoRecords
    sMsg = "Registros Language excluídos: " & lRegAfetados
    GravarLog sMsg
    Exit Sub
    
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: LimparBanco"
    GravarLog sMsg
End Sub
Sub SalvarDados()
   
    Dim objLista            As IXMLDOMNodeList
    Dim objElemento         As IXMLDOMElement
    Dim objElementoFilho    As IXMLDOMElement
        
    On Error GoTo tratarErro
    
    LerDadosConexao
    ConectarBanco
    LimparBanco
    
    If objDOMPais.parseError.reason <> "" Then Exit Sub
    
    If IsEmpty(objXML) Then Exit Sub
    Set objLista = objDOMPais.selectNodes("//tCountryInfo")
         
    'Percorre XML para inserir dados no banco
    For Each objElemento In objLista
        
        For Each objElementoFilho In objElemento.childNodes
        
            Select Case objElementoFilho.baseName
            Case "sISOCode"
                sISOCodeC = objElementoFilho.nodeTypedValue
                
            Case "sName"
                sNameCountry = Left(RetirarCaracteresInvalidos(objElementoFilho.nodeTypedValue), 100)
                
            Case "sCapitalCity"
                sCapitalCity = Left(RetirarCaracteresInvalidos(objElementoFilho.nodeTypedValue), 100)
            
            Case "sPhoneCode"
                sPhoneCode = Left(objElementoFilho.nodeTypedValue, 3)
            
            Case "sContinentCode"
                sContinentCode = Left(objElementoFilho.nodeTypedValue, 2)
            
            Case "sCurrencyISOCode"
                sCurrencyISOCode = Left(objElementoFilho.nodeTypedValue, 3)
            
            Case "sCountryFlag"
                sCountryFlag = Left(objElementoFilho.nodeTypedValue, 100)
            
            Case "Languages"
                '-- insere o país
                conexao.Execute "IF NOT EXISTS (SELECT * FROM Country WHERE ISOCodeC ='" & sISOCodeC & "')" _
                & " INSERT INTO Country (ISOCodeC, Name,CapitalCity,PhoneCode,ContinentCode,CurrencyISOCode,CountryFlag)" _
                & " VALUES ('" & sISOCodeC _
                & "', '" & sNameCountry _
                & "', '" & sCapitalCity _
                & "', '" & sPhoneCode _
                & "', '" & sContinentCode _
                & "', '" & sCurrencyISOCode _
                & "', '" & sCountryFlag & "')"
                
                SalvarIdiomas objElementoFilho
            End Select
        Next
        
    Next
    DesconectarBanco
        
    sMsg = "Xml salvo com sucesso no banco de dados."
    GravarLog sMsg
           
    Exit Sub
 
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: SalvarDados"
    GravarLog sMsg
End Sub


Private Sub SalvarIdiomas(ByRef objDOMElemento As IXMLDOMElement)

  Dim objElemento As IXMLDOMElement
  
  On Error GoTo tratarErro
  
  For Each objElemento In objDOMElemento.childNodes
      sISOCodeL = objElemento.childNodes(0).nodeTypedValue
      sNameLanguage = objElemento.childNodes(1).nodeTypedValue
            
      conexao.Execute "IF NOT EXISTS (SELECT * FROM Language WHERE ISOCodeL ='" & sISOCodeL & "')" _
                  & " INSERT INTO Language (ISOCodeL, Name)" _
                  & " VALUES ('" & sISOCodeL _
                  & "', '" & sNameLanguage & "')"
                    
           
    conexao.Execute "IF NOT EXISTS (SELECT * FROM CountryLanguage WHERE ISOCodeC='" & sISOCodeC & " ' AND ISOCodeL ='" & sISOCodeL & "')" _
                  & " INSERT INTO CountryLanguage (ISOCodeC, ISOCodeL)" _
                  & " VALUES ('" & sISOCodeC _
                  & "', '" & sISOCodeL & "')"
    Next
                        
    Exit Sub

tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: SalvarIdiomas"
    GravarLog sMsg
End Sub
Sub LerDadosConexao()

    Dim sArqIni As String
    
    On Error GoTo tratarErro
    
    sArqIni = App.Path & "\BD.INI"
    sServidor = LerArqIni("BD", "SERVIDOR", sArqIni)
    sBanco = LerArqIni("BD", "BANCO", sArqIni)
    sUsuario = LerArqIni("BD", "USUARIO", sArqIni)
    sSenha = LerArqIni("BD", "SENHA", sArqIni)
    Exit Sub

tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: LerDadosConexao"
    GravarLog sMsg
End Sub

Sub ConectarBanco()
  
    On Error GoTo tratarErro
        
    Set conexao = New ADODB.Connection

    conexao.ConnectionString = "driver={SQL Server};server=" & sServidor & ";" & _
         "database=" & sBanco & ";uid=" & sUsuario & ";pwd=" & sSenha & ";"

    conexao.ConnectionTimeout = 30

    conexao.Open
                
    If conexao.State = adStateOpen Then
       sMsg = "Conexão com sucesso."
       GravarLog sMsg
    Else
        sMsg = "Erro na conexão."
        GravarLog sMsg
    End If
    
    Exit Sub
    
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: ConectarBanco"
    GravarLog sMsg
End Sub
Sub DesconectarBanco()
    conexao.Close
End Sub

Function LerArqIni(sSecao As String, _
                  sEntrada As String, _
                  sArquivo As String) As String
  
 Dim sTamanhoRetorno As String
 Dim sRetorno As String
 
 On Error GoTo tratarErro
   
 sRetorno = String$(255, 0)
 sTamanhoRetorno = GetPrivateProfileString(sSecao, sEntrada, "", sRetorno, Len(sRetorno), sArquivo)
 sRetorno = Left$(sRetorno, sTamanhoRetorno)
 LerArqIni = sRetorno
 
  Exit Function
tratarErro:
    LerArqIni = ""
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: LerArqIni"
    GravarLog sMsg
End Function

Sub GravarLog(sEvento As String)
             
    Dim fso As New FileSystemObject
    Dim arquivo As File
    Dim arquivoLog As TextStream
    Dim sMensagem As String
    Dim sCaminho As String
    
    On Error GoTo tratarErro
    
    sCaminho = App.Path & "\log"

    'se o arquivo não existir então cria
    If fso.FileExists(sCaminho) Then
       Set arquivo = fso.GetFile(sCaminho)
    Else
       Set arquivoLog = fso.CreateTextFile(sCaminho)
       arquivoLog.Close
       Set arquivo = fso.GetFile(sCaminho)
    End If
    
    'prepara o arquivo para anexa os dados
    Set arquivoLog = arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha
    sMensagem = Now() & " - Evento: " & sEvento
    
    ' inclui linhas no arquivo texto
    arquivoLog.WriteLine sMensagem
    
    'fecha e libera o objeto
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing

    Exit Sub
tratarErro:
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: GravarLog"
    Err.Raise sMsg
End Sub
