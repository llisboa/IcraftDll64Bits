Imports Microsoft.VisualBasic ' Recursos antigos de linguagem vb.
Imports System.Collections.Generic ' Coleções.
Imports System.Collections.ObjectModel ' Pré-definições para coleções.
Imports System.Configuration ' Acesso à arquivos de configuração.
Imports System.Data ' Oledb microsoft.
Imports System.Data.OleDb ' Oledb microsoft.
Imports System.Drawing ' Edição de imagens.
Imports System.IO ' Acesso à arquivos.
Imports System.Linq ' Linq.
Imports System.Net ' Recursos de rede.
Imports System.Net.Mail ' Correio eletrônico.
Imports System.Text ' Manipulação de texto stringbuilder.
Imports System.Text.RegularExpressions ' Regex.
Imports System.Web ' Suporte web (obrigatório para integração de funções app/web).
Imports System.Web.Security ' Segurança de web (obrigatório para integração de funções app/web).
Imports System.Web.UI ' Componentes para web (obrigatório para integração de funções app/web).
Imports System.Web.UI.WebControls ' Controles web (obrigatório para integração de funções app/web).
Imports System.Windows.Forms.Form ' Controles em form (obrigatório para integração de funções app/web).
Imports System.Xml ' Serialização de xml.
Imports Microsoft.Win32 ' Declare para acesso de funções user.
Imports System.Data.SqlClient ' Acesso à sqlserver (obrigatório para integração oracle/mysql/sqlserv).
Imports System.Web.Services ' Disponibilização de interfaces webservices.
Imports System.Web.Services.Protocols ' Manuseio de ferramentas de protocolo.
Imports System.Security.Cryptography ' Biblioteca de criptografia.
Imports System
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Drawing.Text


Namespace Icraft ' Biblioteca desenvolvida pela Intercraft para uso genérico em aplicativos - componentes/funções desde 1996 - contém antigas funções VBA e adaptação das mesmas para ambiente VB.NET.

    <WebService()> Public Class IcftBase ' Funções Intercraft para uso genérico em aplicativos - acessibilidade em VB, VC, CSHARP, MsAccess, Oracle entre outros ambientes.



        ''' <summary>
        ''' Opções para tipo de ambiente onde a dll está rodando.
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum AmbienteTipo
            Windowsforms
            WEB
        End Enum


        ''' <summary>
        ''' Obtém versão do conjunto de soluções Intercraft para aplicativos.
        ''' </summary>
        ''' <value>Versão no formato V1.1.1.1.</value>
        ''' <returns>Versão no formato V1.1.1.1.</returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Versao() As String
            Get
                Return "V" & Trim(System.Reflection.Assembly.GetExecutingAssembly.FullName.Split(",")(1).Replace("Version=", ""))
            End Get
        End Property


        ''' <summary>
        ''' Carrega estrutura de um dataset baseado em SQL em ORACLE, MySQL ou MSAccess.
        ''' </summary>
        ''' <param name="SQL">Select para obtenção da estrutura.</param>
        ''' <param name="STRCONN">Identificador da connexão ou string da configuração no web.config.</param>
        ''' <returns>Retorna um dataset contendo somente a estrutura.</returns>
        ''' <remarks></remarks>
        Shared Function DSCarregaEstrut(ByVal SQL As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As DataSet
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            Dim ds As DataSet = New DataSet
            If Compare(ConnW.ProviderName, MySQL) Then
                ' mysql
                Dim c As New CriadorDeObjetos("MySql.Data.dll")

                Dim Conexao As Object = c.Criar("MySqlConnection", ConnW.ConnectionString)
                Dim comm As Object = c.Criar("MySqlCommand")
                comm = DSCriaComandoMySQL(SQL, Conexao, ListaParametros)

                Dim Adapt As Object = c.Criar("MySqlDataAdapter", comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            ElseIf Compare(ConnW.ProviderName, MSAccess) Then
                ' msaccess
                Dim Conexao As OleDbConnection = New OleDbConnection(ConnW.ConnectionString)
                Dim comm As OleDbCommand = DSCriaComandoAccess(SQL, Conexao, ListaParametros)
                Dim Adapt As OleDbDataAdapter = New OleDbDataAdapter(comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            ElseIf Compare(ConnW.ProviderName, Oracle) Then
                ' oracle
                Dim conexao As New OracleClient.OracleConnection(ConnW.ConnectionString)
                Dim comm As OracleClient.OracleCommand
                comm = DSCriaComandoOracle(SQL, conexao, Nothing, ListaParametros)
                Dim Adapt As New OracleClient.OracleDataAdapter(comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            End If
            Return ds
        End Function

        ''' <summary>
        ''' Transforma um paramarray em arraylist.
        ''' </summary>
        ''' <param name="PARAMS">Lista de parâmetros podendo ser um arraylist ou paramarray.</param>
        ''' <returns>Retornará um arraylist contendo a lista de parâmetros.</returns>
        ''' <remarks></remarks>
        Shared Function ParamArrayToArrayList(ByVal ParamArray Params() As Object) As Object

            ' caso não existam parâmetros
            If IsNothing(Params) OrElse Params.Length = 0 Then
                Return New ArrayList
            End If

            ' caso já seja um arraylist
            If Params.Length = 1 And TypeOf (Params(0)) Is ArrayList Then
                Return Params(0)
            End If

            ' caso tenha que juntar
            Dim ListaParametros As ArrayList = New ArrayList
            For Each Item As Object In Params
                If Not IsNothing(Item) Then

                    ' >> TIPOS PREVISTOS EM ARRAYLIST...
                    ' array
                    ' arraylist
                    ' string
                    ' dataset
                    ' datarowcollection

                    If TypeOf Item Is Array Then
                        For Each SubItem As Object In Item
                            ListaParametros.AddRange(ParamArrayToArrayList(SubItem))
                        Next
                    ElseIf TypeOf Item Is ArrayList OrElse Item.GetType.ToString.StartsWith("System.Collections.Generic.List") Then
                        ListaParametros.AddRange(Item)
                    ElseIf TypeOf Item Is String Then
                        ListaParametros.Add(Item)
                    ElseIf TypeOf Item Is DataTable Then
                        For Each Row As DataRow In Item.rows
                            For Each Campo As Object In Row.ItemArray
                                ListaParametros.Add(Campo)
                            Next
                        Next
                    ElseIf TypeOf Item Is DataSet Then
                        For Each Row As DataRow In Item.Tables(0).rows
                            For Each Campo As Object In Row.ItemArray
                                ListaParametros.Add(Campo)
                            Next
                        Next
                    ElseIf TypeOf Item Is DataRow Then
                        For Each Campo As Object In CType(Item, DataRow).ItemArray
                            ListaParametros.Add(Campo)
                        Next
                    ElseIf TypeOf Item Is System.IO.FileInfo Then
                        ListaParametros.Add(Item.name)
                    Else
                        ListaParametros.Add(Item)
                    End If
                End If
            Next
            Return ListaParametros
        End Function

        ''' <summary>
        ''' Compara dois parâmetros com base em critério específicos para cada tipo.
        ''' </summary>
        ''' <param name="Param1">Primeiro parâmetro.</param>
        ''' <param name="Param2">Segundo parâmetro.</param>
        ''' <param name="IgnoreCase">Para ignorar diferença entre maiúsculo e minúsculo em comparações de strings.</param>
        ''' <returns>Retorna verdadeiro caso os itens sejam considerados iguais ou o contrário.</returns>
        ''' <remarks></remarks>
        Shared Function Compare(ByVal Param1 As Object, ByVal Param2 As Object, Optional ByVal IgnoreCase As Boolean = True) As Boolean
            If IsNothing(Param1) And IsNothing(Param2) Then
                Return True
            ElseIf IsNothing(Param1) Or IsNothing(Param2) Then
                Return False
            Else
                If Param1.GetType.ToString = Param2.GetType.ToString Then
                    If Param1.GetType.ToString = "System.String" Then
                        Return String.Compare(Param1, Param2, IgnoreCase) = 0
                    Else
                        Err.Raise(20000, "IcraftBase", "Compare com tipo não previsto " & Param1.GetType.ToString & ".")
                    End If
                End If
            End If
            Return False
        End Function


        ''' <summary>
        ''' Classe responsável por criar objetos com base em uma dll carregada dinamicamente.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class CriadorDeObjetos
            Private _assembly As Reflection.Assembly

            ''' <summary>
            ''' Publica caminho de uma dll procurando-a pelos diretórios default.
            ''' </summary>
            ''' <param name="dll">Nome da dll.</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal dll As String)
                Dim dirList As String = ""
                Dim errList As String = ""

                If Ambiente() = AmbienteTipo.WEB Then
                    If Not String.IsNullOrEmpty(IO.Path.GetDirectoryName(dll)) Then
                        dirList = IO.Path.GetDirectoryName(HttpContext.Current.Server.MapPath(dll))
                        dll = IO.Path.GetFileName(HttpContext.Current.Server.MapPath(dll))
                    Else
                        dirList = HttpContext.Current.Server.MapPath("~/bin/") & ";"
                    End If

                Else
                    If Not String.IsNullOrEmpty(IO.Path.GetDirectoryName(dll)) Then
                        dirList = IO.Path.GetDirectoryName(dll)
                        dll = IO.Path.GetFileName(dll)
                    Else
                        dirList = System.Windows.Forms.Application.StartupPath & "\" & ";" & System.Windows.Forms.Application.StartupPath & "\bin\" & ";" & System.AppDomain.CurrentDomain.BaseDirectory & ";" & System.AppDomain.CurrentDomain.BaseDirectory & "bin\;" & Environment.GetEnvironmentVariable("WINDIR") & ";" & Environment.SystemDirectory() & ";"
                    End If
                End If

                For Each s As String In dirList.Split(";")
                    If IO.File.Exists(FileExpr(s, dll)) Then
                        Try
                            _assembly = Reflection.Assembly.LoadFile(FileExpr(s, dll))
                        Catch EX As Exception
                            Dim DI As New System.IO.DirectoryInfo(FileExpr(s, dll))
                            _assembly = Reflection.Assembly.LoadFile(DI.FullName)
                        End Try
                        Exit For
                    End If
                Next

                If _assembly Is Nothing Then
                    Throw New Exception("Não foi possível carregar a dll especificada('" & dll & "'). Verifique se a mesma contém um formato válido.")
                End If
            End Sub

            ''' <summary>
            ''' Cria um objeto de um tipo especificado dentro da dll com base em seu nome.
            ''' </summary>
            ''' <param name="obj">O nome do tipo do objeto que será criado.</param>
            ''' <param name="params">Parâmetros que atendam a algum construtor do objeto especificado. No caso de nada ser passado, o contrutor default será admitido.</param>
            ''' <returns>Retorna o objeto criado.</returns>
            ''' <remarks></remarks>
            Public Function Criar(ByVal obj As String, ByVal ParamArray params() As Object) As Object
                Return _assembly.CreateInstance(getTipo(obj).FullName, False, Reflection.BindingFlags.CreateInstance, Nothing, params, Nothing, Nothing)
            End Function

            ''' <summary>
            ''' Cria objeto a partir de uma instância com nome e parâmetros específicos.
            ''' </summary>
            ''' <param name="assemblyInstance">Instância da biblioteca publicada.</param>
            ''' <param name="objectName">Nome do objeto a ser criado.</param>
            ''' <param name="params">Parâmetros considerados por esta instância.</param>
            ''' <returns>Objeto criado.</returns>
            ''' <remarks></remarks>
            Public Shared Function Criar(ByVal assemblyInstance As Reflection.Assembly, ByVal objectName As String, ByVal ParamArray params() As Object) As Object
                Return assemblyInstance.CreateInstance(getTipo(assemblyInstance, objectName).FullName, False, Reflection.BindingFlags.CreateInstance, Nothing, params, Nothing, Nothing)
            End Function

            ''' <summary>
            ''' Pega tipo de um objeto.
            ''' </summary>
            ''' <param name="typeName">Tipo do objeto.</param>
            ''' <returns>Tipo do objeto.</returns>
            ''' <remarks></remarks>
            Public Function getTipo(ByVal typeName As String) As System.Type
                Return _assembly.GetTypes().Single(Function(tp) tp.Name = typeName)
            End Function

            ''' <summary>
            ''' Pega uma função do tipo especificado.
            ''' </summary>
            ''' <param name="asb">Objeto a ser pesquisado.</param>
            ''' <param name="typeName">Tipo que permitirá buscar a função.</param>
            ''' <returns>Objeto de tipo encontrado.</returns>
            ''' <remarks></remarks>
            Public Shared Function GetTipo(ByVal asb As Reflection.Assembly, ByVal typeName As String) As System.Type
                Return asb.GetTypes().Single(Function(tp) tp.Name = typeName)
            End Function

            ''' <summary>
            ''' Busca todos os tipos encontrados no objeto.
            ''' </summary>
            ''' <returns>Array contendo todos os tipos.</returns>
            ''' <remarks></remarks>
            Public Function VerificaTipos() As String()
                Return (From assemb As System.Type In _assembly.GetTypes Select assemb.Name).ToArray
            End Function

            ''' <summary>
            ''' Busca tipos na instância informada.
            ''' </summary>
            ''' <param name="asb">Instância informada.</param>
            ''' <returns>Array de tipos existentes na instância.</returns>
            ''' <remarks></remarks>
            Public Shared Function VerificaTipos(ByVal asb As Reflection.Assembly) As String()
                Return (From assemb As System.Type In asb.GetTypes Select assemb.Name).ToArray
            End Function

            ''' <summary>
            ''' Retorna assembly criado.
            ''' </summary>
            ''' <value>Assembly criado.</value>
            ''' <returns>Assembly criado.</returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Assembly() As Reflection.Assembly
                Get
                    Return _assembly
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Concatena um diretório passado por parâmetro ao diretório raiz.
        ''' </summary>
        ''' <param name="Segmentos">Diretório que será adicionado a raiz.</param>
        ''' <returns>Retorna o caminho da raiz e adiciona o diretório passado no "Segmentos"</returns>
        ''' <remarks></remarks>
        Shared Function FileExpr(ByVal ParamArray Segmentos() As String) As String
            Dim Raiz As String = New System.Web.UI.Control().ResolveUrl("~/").Replace("/", "\")
            Dim Arq As String = ExprExpr("\", "/", "", Segmentos)
            If Arq.StartsWith(Raiz) Then
                Arq = "~\" & Mid(Arq, Len(Raiz) + 1)
            End If

            If Arq.StartsWith("~\") Then
                If Ambiente() = AmbienteTipo.WEB Then
                    Arq = HttpContext.Current.Server.MapPath(Arq)
                Else
                    Dim DirExec As String = FileExpr(WebConf("dir_raiz_site"), "\")
                    If DirExec = "" Or DirExec = "\" Then
                        DirExec = System.Windows.Forms.Application.ExecutablePath
                    End If
                    Arq = Arq.Replace("~\", System.IO.Path.GetDirectoryName(DirExec) & "\")
                End If
            End If
            Return Arq
        End Function

        ''' <summary>
        ''' Retorna parâmetro específico do webconfig > appsetings.
        ''' </summary>
        ''' <param name="param">Identificação do connectionstring desejado.</param>
        ''' <returns>Objeto connectionstringsettings obtido a partir do configurationmanager.</returns>
        ''' <remarks></remarks>
        Public Shared Function WebConf(ByVal param As String) As String
            If Compare(param, "SITE_DIR") Then
                Return FileExpr("~/")
            ElseIf Compare(param, "SITE_URL") Then
                Return URLExpr("~/")
            End If
            Return System.Configuration.ConfigurationManager.AppSettings(param)
        End Function

        ''' <summary>
        ''' Concatena URL evitando barras repetidas.
        ''' </summary>
        ''' <param name="Segmentos">São os trechos a serem concatenados, podendo ser mais de dois.</param>
        ''' <returns>Retorna expressão de segmentos concatenados.</returns>
        ''' <remarks></remarks>
        Shared Function URLExpr(ByVal ParamArray Segmentos() As Object) As String
            Dim URL As String = ExprExpr("/", "\", "", Segmentos)
            If Regex.Match(URL, "(?is)^[a-z0-9]:/").Success Then
                If Ambiente() = AmbienteTipo.WEB Then
                    URL = URL.ToLower.Replace(HttpContext.Current.Server.MapPath("~/").Replace("\", "/").ToLower, "~/")
                Else
                    URL = URL.Replace("\", "/").ToLower
                    URL = URL.Replace(FileExpr("~/").Replace("\", "/").ToLower, "~/")
                End If
            End If
            Return URL
        End Function


        ''' <summary>
        ''' Retorna ambiente no qual o programa (ou dll) está rodando.
        ''' </summary>
        ''' <returns>Ambiente podendo ser WindowsForm ou WEB.</returns>
        ''' <remarks>Caso programa seja uma console, retornará ambiente WindowsForm.</remarks>
        Public Shared Function Ambiente() As AmbienteTipo
            Try
                If Not IsNothing(HttpContext.Current) Then
                    Return AmbienteTipo.WEB
                End If
            Catch
            End Try
            Return AmbienteTipo.Windowsforms
        End Function



        ''' <summary>
        ''' Concatena um conjunto de expressões separando-as ou não por um delimitador especificado.
        ''' </summary>
        ''' <param name="Delim">Um caractere ou expressão que será colocada entre as outras.</param>
        ''' <param name="DelimAlternativo">Um caractere ou expressão que será substituída por Delim.</param>
        ''' <param name="Inicial">Um caractere ou expressão que será colocada no início da string.</param>
        ''' <param name="Segmentos">O conjunto de expressões que será concatenado.</param>
        ''' <returns>Retorna uma string com todos os objetos de Segmentos concatenados.</returns>
        ''' <remarks></remarks>
        Shared Function ExprExpr(ByVal Delim As String, ByVal DelimAlternativo As String, ByVal Inicial As Object, ByVal ParamArray Segmentos() As Object) As String
            Inicial = NZ(Inicial, "")
            Dim Lista As ArrayList = ParamArrayToArrayList(Segmentos)
            For Each item As Object In Lista
                If Not IsNothing(item) Then
                    If Not IsNothing(DelimAlternativo) AndAlso DelimAlternativo <> "" Then
                        item = item.Replace(DelimAlternativo, Delim)
                    End If
                    item = NZ(item, "")
                    If item <> "" Then
                        If Inicial <> "" Then
                            If Inicial.EndsWith(Delim) AndAlso item.StartsWith(Delim) Then
                                Inicial &= CType(item, String).Substring(Delim.Length)
                            ElseIf Inicial.EndsWith(Delim) OrElse item.StartsWith(Delim) Then
                                Inicial &= item
                            Else
                                Inicial &= Delim & item
                            End If
                        Else
                            Inicial &= item
                        End If
                    End If
                End If
            Next
            Return Inicial
        End Function



        ''' <summary>
        ''' Caso o objeto inicial não exista (ismissing) ou seja nulo (dbnull), retorna o segundo parâmetro.
        ''' </summary>
        ''' <param name="Valor">Parâmetro a ser analisado.</param>
        ''' <param name="Def">Parâmetro default caso o primeiro parâmetro não exista ou seja nulo.</param>
        ''' <returns>Retorna primeiro parâmetro ou segundo caso o primeiro não exista ou seja nulo, sempre convertendo para o tipo do segundo parâmetro.</returns>
        ''' <remarks></remarks>
        Shared Function NZ(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
            Dim tipo As String

            If Not IsNothing(Def) Then
                tipo = Def.GetType.ToString
            ElseIf IsNothing(Valor) Then
                Return Nothing
            Else
                tipo = Valor.GetType.ToString.Trim
            End If

            If IsNothing(Valor) OrElse IsDBNull(Valor) OrElse ((tipo = "System.DateTime" Or Valor.GetType.ToString = "System.DateTime") AndAlso Valor = CDate(Nothing)) Then
                Valor = Def
            End If

            Select Case tipo
                Case "System.Decimal"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Decimal)
                    End If
                    Return CType(Valor, Decimal)
                Case "System.String"
                    If Valor.GetType.ToString = "System.Byte[]" Then
                        Return CType(ByteArrayToObject(Valor), String)
                    End If
                    If Valor.GetType.ToString = "Icraft.IcftBase+LogonSession" Then
                        Return CType(Valor, LogonSession).ToString
                    ElseIf Valor.GetType.IsEnum Then
                        Return Valor.ToString
                    End If
                    Return CType(Valor, String)
                Case "System.Double"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Double)
                    End If
                    Return CType(Valor, Double)
                Case "System.Boolean"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return False
                    End If
                    Return CType(Valor, Boolean)
                Case "System.DateTime"
                    Return CType(Valor, System.DateTime)
                Case "System.Single"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Single)
                    End If
                    Return CType(Valor, System.Single)
                Case "System.Byte"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Byte)
                    End If
                    Return CType(Valor, System.Byte)
                Case "System.Char"
                    Return CType(Valor, System.Char)
                Case "System.SByte"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, SByte)
                    End If
                    Return CType(Valor, System.SByte)
                Case "System.Int32"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Int32)
                    End If
                    Return CType(Valor, Int32)
                Case "System.DBNull"
                    Return Valor
                Case "System.Collections.ArrayList"
                    Return ParamArrayToArrayList(Valor)
                Case "System.Data.DataSet"
                    If IsNothing(Valor) Then
                        Return Def
                    End If
                    Return Valor
            End Select

            Return CType(Valor, String)
        End Function



        ''' <summary>
        ''' Transforma array de bytes em objeto.
        ''' </summary>
        ''' <param name="Bytes">Array de bytes a ser transferida para objeto.</param>
        ''' <returns>Objeto criado a partir do array de bytes.</returns>
        ''' <remarks></remarks>
        Shared Function ByteArrayToObject(ByVal Bytes() As Byte) As Object
            Dim Obj As Object = Nothing
            Try
                Dim fs As System.IO.MemoryStream = New System.IO.MemoryStream
                Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                fs.Write(Bytes, 0, Bytes.Length)
                fs.Seek(0, IO.SeekOrigin.Begin)

                Obj = formatter.Deserialize(fs)
            Catch
            End Try
            Return Obj
        End Function


        ''' <summary>
        ''' Classe para registro de logon de usuário e variáveis relacionadas.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class LogonSession
            Private _id As String = Nothing
            Private _usuario As String = Nothing
            Private _momento As Date = Nothing
            Private _site As String = Nothing
            Private _senha As String = Nothing
            Private _grupo As String = Nothing
            Private _outros As New ArrayList

            ''' <summary>
            ''' Converte as informações de login da seção para uma string.
            ''' </summary>
            ''' <returns>Retorna a string contendo as informações.</returns>
            ''' <remarks></remarks>
            Public Shadows Function ToString() As String
                Dim txt As New StringBuilder
                txt.Append("LogonSession(")
                txt.Append("id=" & NZ(_id, "") & ";")
                txt.Append("_usuario=" & NZ(_usuario, "") & ";")
                txt.Append("_momento=" & Format(NZV(_momento, Nothing), "dd/MM/yyyy HH:mm:ss") & ";")
                For z As Integer = 0 To _outros.Count - 1 Step 2
                    txt.Append(_outros(z) & "=")
                    txt.Append(NZ(_outros(z + 1), ""))
                    txt.Append(";")
                Next
                txt.Append("_site=" & NZ(_site, ""))
                txt.Append("_grupo=" & NZ(_grupo, ""))
                txt.Append(")")
                Return txt.ToString
            End Function


            ''' <summary>
            ''' Identificação para armazenamento de logon do tipo 'GERAL' ou algum específico para múltiplos logons.
            ''' </summary>
            ''' <value>Especificação do tipo de logon.</value>
            ''' <returns>Especificação do tipo de logon.</returns>
            ''' <remarks></remarks>
            Public Property Id() As String
                Get
                    Return _id
                End Get
                Set(ByVal value As String)
                    _id = value
                End Set
            End Property

            ''' <summary>
            ''' Usuário que efetuou logon.
            ''' </summary>
            ''' <value>Login do usuário que efetuou logon.</value>
            ''' <returns>Login do usuário que efetuou logon.</returns>
            ''' <remarks></remarks>
            Public Property Usuario() As String
                Get
                    Return _usuario
                End Get
                Set(ByVal value As String)
                    _usuario = value
                End Set
            End Property

            ''' <summary>
            ''' Momento de logon.
            ''' </summary>
            ''' <value>Momento (data e hora) de logon.</value>
            ''' <returns>Momento (data e hora) de logon.</returns>
            ''' <remarks></remarks>
            Public Property Momento() As Date
                Get
                    Return _momento
                End Get
                Set(ByVal value As Date)
                    _momento = value
                End Set
            End Property

            ''' <summary>
            ''' Grupo do usuário que efetuou logon.
            ''' </summary>
            ''' <value>Nome do grupo do usuário que efetuou logon.</value>
            ''' <returns>Nome do grupo do usuário que efetuou logon.</returns>
            ''' <remarks></remarks>
            Public Property Grupo() As String
                Get
                    Return _grupo
                End Get
                Set(ByVal value As String)
                    _grupo = value
                End Set
            End Property

            ''' <summary>
            ''' Nome do site.
            ''' </summary>
            ''' <value></value>
            ''' <returns>Nome do site.</returns>
            ''' <remarks>Nome do site.</remarks>
            Public Property Site() As String
                Get
                    Return _site
                End Get
                Set(ByVal value As String)
                    _site = value
                End Set
            End Property

            ''' <summary>
            ''' Senha de acesso.
            ''' </summary>
            ''' <value>Senha de acesso.</value>
            ''' <returns>Senha de acesso.</returns>
            ''' <remarks></remarks>
            Public Property Senha() As String
                Get
                    Return _senha
                End Get
                Set(ByVal value As String)
                    _senha = value
                End Set
            End Property

            ''' <summary>
            ''' Outras propriedades a serem armazenadas pelo Logon.
            ''' </summary>
            ''' <param name="Propriedade">Nome da propriedade.</param>
            ''' <value>Valor da propriedade.</value>
            ''' <returns>Valor da propriedade armazenada.</returns>
            ''' <remarks></remarks>
            Public Property ExtendedProps(ByVal Propriedade As String) As Object
                Get
                    Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                    If Pos >= 0 Then
                        Return _outros(Pos + 1)
                    End If
                    Return Nothing
                End Get
                Set(ByVal value As Object)
                    Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                    If Pos >= 0 Then
                        _outros(Pos + 1) = value
                        Exit Property
                    End If
                    _outros.Add(":" & Propriedade)
                    _outros.Add(value)
                End Set
            End Property

            ''' <summary>
            ''' Acesso aos atributos e propriedades expandidas.
            ''' </summary>
            ''' <param name="Nome">Nome da propriedade tratada.</param>
            ''' <value>Valor da propriedade tratada.</value>
            ''' <returns>Valor da propriedade solicitada.</returns>
            ''' <remarks></remarks>
            Default Property Attributes(ByVal Nome As String) As String
                Get
                    If Compare(Nome, "Id") Then
                        Return _id
                    ElseIf Compare(Nome, "Usuario") Then
                        Return _usuario
                    ElseIf Compare(Nome, "Momento") Then
                        Return _momento
                    ElseIf Compare(Nome, "Site") Then
                        Return _site
                    ElseIf Compare(Nome, "Senha") Then
                        Return _senha
                    ElseIf Compare(Nome, "Grupo") Then
                        Return _grupo
                    Else
                        Dim Prop As Object = ExtendedProps(Nome)
                        If IsNothing(Prop) Then
                            Throw New Exception("Em Attributes de Logon, atributo '" & Nome & "' inválido para objeto " & Me.GetType.ToString & ".")
                        Else
                            Return Prop
                        End If
                    End If
                    Return Nothing
                End Get

                Set(ByVal value As String)
                    If Compare(Nome, "Id") Then
                        _id = value
                    ElseIf Compare(Nome, "Usuario") Then
                        _usuario = value
                    ElseIf Compare(Nome, "Momento") Then
                        _momento = value
                    ElseIf Compare(Nome, "Site") Then
                        _site = value
                    ElseIf Compare(Nome, "Senha") Then
                        _senha = value
                    ElseIf Compare(Nome, "Grupo") Then
                        _grupo = value
                    Else
                        Throw New Exception("Em Attributes de Logon, atributo " & value & " inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Criação de login para registro de acesso de usuário.
            ''' </summary>
            ''' <param name="Pagina">Página na qual é efetuado o login.</param>
            ''' <param name="Usuario">Usuário que efetua acesso.</param>
            ''' <param name="Senha">Senha do usuário.</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal Pagina As System.Web.UI.Page, ByVal Usuario As String, ByVal Senha As String)
                ' cria chave com area e usuario
                Try
                    _id = Pagina.Session.SessionID
                    _usuario = Usuario
                    _momento = Now
                    _site = WebConf("site_nome")
                    _senha = Senha
                Catch
                    _id = Nothing
                    _usuario = Nothing
                    _momento = Nothing
                    _site = Nothing
                    _senha = Nothing
                End Try
            End Sub

        End Class

        ''' <summary>
        ''' Retorna valor padrão se for Nothing, Nulo ou Vazio (ou zero no caso de tipo numérico).
        ''' </summary>
        ''' <param name="Valor">Valor a ser checado.</param>
        ''' <param name="Def">Default a ser retornado caso seja Nothing, Nulo ou vazio.</param>
        ''' <returns>Valor checado ou valor default caso Nothing, Nulo ou vazio (zero se o tipo for numérico).</returns>
        ''' <remarks></remarks>
        Shared Function NZV(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
            Dim Result As Object = NZ(Valor, Def)
            If TypeOf Result Is String AndAlso Result = "" Then
                Return Def
            ElseIf TypeOf Result Is Decimal AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Double AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Single AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Int32 AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Integer AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Byte AndAlso Result = 0 Then
                Return Def
            End If
            Return Result
        End Function

        ''' <summary>
        ''' Retorna um connectionstring a partir da informação de um connectionstring ou string indicativa da conexão no WebConfig.
        ''' </summary>
        ''' <param name="STRCONN">Connectionstring ou nome da conexao no webconfig (ex: "STRTAREFA", "ProviderName:MySQL.Data.MySQLClient;Server:127.0.0.1;Database:data;Uid:usuario;Pwd:senha;" ou "STRTAREFA;USER:usuario;PASSWORD:senha").</param>
        ''' <returns>Objeto connectionstring instanciado.</returns>
        ''' <remarks>Caso seja passada string, programador poderá fazer uso de complementos do tipo: "strGerador;user:estagiario;password:estag".
        ''' Isso corresponde a obter os dados da conexão do WebConfig com nome de strGerador e substituir nesta user e password.</remarks>
        Shared Function StrConnObj(ByVal StrConn As Object, ByVal ParamArray Params() As Object) As System.Configuration.ConnectionStringSettings
            If TypeOf (StrConn) Is System.Configuration.ConnectionStringSettings Then
                Return CType(StrConn, System.Configuration.ConnectionStringSettings)
            End If

            Dim Param As String = CType(StrConn, String)
            If Param.IndexOf(";") = -1 Then
                If Regex.Match(Param, "(?is)\.mdb$").Success Then
                    Return StrConnObj("ProviderName:System.Data.OleDb;Provider:Microsoft.Jet.OLEDB.4.0;Data Source:" & FileExpr(Param), Params)
                End If
                Return WebConn(Param)
            End If

            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Try
                MacroSubstSQL(Param, Nothing, ListaParametros)
            Catch
            End Try


            Dim Elem As New ElementosStr(Param, ";")
            Dim Conn As New System.Configuration.ConnectionStringSettings

            If Elem.Items("").Conteudo <> "" Then
                Dim ConnAnt As System.Configuration.ConnectionStringSettings = System.Configuration.ConfigurationManager.ConnectionStrings(Elem.Items("").Conteudo)
                Conn.ProviderName = ConnAnt.ProviderName
                Conn.ConnectionString = ConnAnt.ConnectionString
            End If

            If Elem.Exists("ProviderName") Then
                Conn.ProviderName = Elem.Items("ProviderName").Conteudo
                Elem.Items("ProviderName").Conteudo = Nothing
            End If

            If Compare(Conn.ProviderName, Oracle) Then
                If Elem.Exists("User") Then
                    Elem.Items("User").Nome = "User ID"
                End If
            ElseIf Compare(Conn.ProviderName, MySQL) Then
                If Elem.Exists("User") Then
                    Elem.Items("User").Nome = "Uid"
                End If
                If Elem.Exists("Password") Then
                    Elem.Items("Password").Nome = "Pwd"
                End If
            ElseIf Compare(Conn.ProviderName, MSAccess) Then
                If Elem.Exists("Data Source") Then
                    Dim Caminho As String = Elem.Items("Data Source").Conteudo
                    If Caminho.StartsWith("~/") Or Caminho.StartsWith("~\") Then
                        Caminho = HttpContext.Current.Server.MapPath(Caminho)
                    End If
                    Elem.Items("Data Source").Conteudo = Caminho
                End If
            End If

            Dim ElemNovo As New ElementosStr(Conn.ConnectionString, ";", "=")
            ElemNovo.AddStr(Elem.ToStyleStr(";", "="))
            Param = ElemNovo.ToStyleStr
            Conn.ConnectionString = Param

            Return Conn
        End Function


        ''' <summary>
        ''' Retorna connectionstring específico do webconfig > connectionstring.
        ''' </summary>
        ''' <param name="param">Identificação do connectionstring desejado.</param>
        ''' <returns>Objeto connectionstringsettings obtido a partir do configurationmanager.</returns>
        ''' <remarks></remarks>
        Shared Function WebConn(ByVal Param As String) As System.Configuration.ConnectionStringSettings
            Return System.Configuration.ConfigurationManager.ConnectionStrings(Param)
        End Function


        ''' <summary>
        ''' Classe para armazenar e operar elementostr.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class ElementosStr
            Private _atributosstr As List(Of ElementoStr) = New List(Of ElementoStr)
            Private _separador As String
            Private _separadorexpr As String
            Private _itemseparador As String
            Private _itemseparadorexpr As String

            ''' <summary>
            ''' Lista de elementos.
            ''' </summary>
            ''' <value>Lista de elementos.</value>
            ''' <returns>Lista de elementos.</returns>
            ''' <remarks></remarks>
            ReadOnly Property Elementos() As List(Of ElementoStr)
                Get
                    Return _atributosstr
                End Get
            End Property

            ''' <summary>
            ''' Especificação de separador.
            ''' </summary>
            ''' <value>Texto contendo separador.</value>
            ''' <returns>Texto contendo separador.</returns>
            ''' <remarks></remarks>
            Property Separador() As String
                Get
                    Return _separador
                End Get
                Set(ByVal value As String)
                    _separador = value
                    _separadorexpr = SeparaExpr(value)
                End Set
            End Property

            ''' <summary>
            ''' Separador de itens.
            ''' </summary>
            ''' <value>Separador de itens.</value>
            ''' <returns>Separador de itens.</returns>
            ''' <remarks></remarks>
            Property ItemSeparador() As String
                Get
                    Return _itemseparador
                End Get
                Set(ByVal value As String)
                    _itemseparador = value
                    _itemseparadorexpr = SeparaExpr(value)
                End Set
            End Property

            ''' <summary>
            ''' Separa parâmetros.
            ''' </summary>
            ''' <param name="Separador">Separador.</param>
            ''' <returns>String contendo itens separados.</returns>
            ''' <remarks></remarks>
            Private Function SeparaExpr(ByVal Separador As String) As String
                Dim Result As String = ""
                For z As Integer = 1 To Separador.Length
                    Dim Letra As String = Mid(Separador, z, 1)
                    If InStr(".\()^|[]+" + Chr(13) + Chr(10), Letra) <> 0 Then
                        Result &= "\" & Letra
                    Else
                        Result &= Letra
                    End If
                Next
                Return Result & "*(([^" & Result & "']|'((([^'])|\\')*)')+)" & Result & "*"
            End Function

            ''' <summary>
            ''' Criação da lista de parâmetros com base em texto e separador.
            ''' </summary>
            ''' <param name="AtributosStr">Lista de atributos.</param>
            ''' <param name="SeparadorTxt">Separador (ex.: border:1px  ; padding:1px).</param>
            ''' <param name="ItemSeparadorTxt">Separador de atributo (border  :  1px).</param>
            ''' <remarks></remarks>
            Sub New(ByVal AtributosStr As String, Optional ByVal SeparadorTxt As String = ";", Optional ByVal ItemSeparadorTxt As String = ":")
                Separador = SeparadorTxt
                ItemSeparador = ItemSeparadorTxt
                AddStr(AtributosStr)
            End Sub

            ''' <summary>
            ''' Transforma estilo em string.
            ''' </summary>
            ''' <param name="SeparadorTxt">Separador a ser utilizado.</param>
            ''' <param name="ItemSeparadorTxt">Atribuidor a ser utilizado.</param>
            ''' <returns>Texto representando estilo com separador e atribuidor escolhidos.</returns>
            ''' <remarks></remarks>
            Function ToStyleStr(Optional ByVal SeparadorTxt As String = Nothing, Optional ByVal ItemSeparadorTxt As String = Nothing) As String
                Dim result As String = ""
                For Each Item As ElementoStr In _atributosstr
                    If Item.Conteudo <> "" And Item.Nome <> "" Then
                        result &= IIf(result <> "", IIf(Not IsNothing(SeparadorTxt), SeparadorTxt, Separador), "")
                        result &= Item.Nome & IIf(Not IsNothing(ItemSeparadorTxt), ItemSeparadorTxt, ItemSeparador)
                        If Item.Conteudo.StartsWith("'") And Item.Conteudo.EndsWith("'") Then
                            result &= Item.Conteudo.Substring(1, Item.Conteudo.Length - 2)
                        Else
                            result &= Item.Conteudo
                        End If
                    End If
                Next
                Return result
            End Function

            ''' <summary>
            ''' Texto representando estilo considerando atributos de separação definidos previamente.
            ''' </summary>
            ''' <returns>Texto representando estilo.</returns>
            ''' <remarks></remarks>
            Overrides Function ToString() As String
                Dim result As String = ""
                For Each Item As ElementoStr In _atributosstr
                    If Item.Conteudo <> "" Then
                        result &= IIf(result <> "", Separador, "") & Item.ToString
                    End If
                Next
                Return result
            End Function

            ''' <summary>
            ''' Lista de itens do estilo.
            ''' </summary>
            ''' <param name="Indice">Índice numérico para acesso ao item do estilo.</param>
            ''' <value>Valor a ser definido.</value>
            ''' <returns>Valor obtido a partir do item consultado.</returns>
            ''' <remarks></remarks>
            Default Overloads Property Items(ByVal Indice As Integer) As ElementoStr
                Get
                    Try
                        Return _atributosstr(Indice)
                    Catch
                    End Try
                    Return New ElementoStr(Nothing, ItemSeparador)
                End Get
                Set(ByVal value As ElementoStr)
                    If Indice = -1 Then
                        _atributosstr.Add(value)
                    Else
                        If Indice >= _atributosstr.Count Then
                            For z As Integer = 0 To Indice
                                _atributosstr.Add(Nothing)
                            Next
                        End If
                        _atributosstr(Indice) = value
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Pesquisa de itens por nome do termo em estilo.
            ''' </summary>
            ''' <param name="Nome">Termo pesquisado.</param>
            ''' <value>Valor a ser atribuído.</value>
            ''' <returns>Valor obtido a partir do termo.</returns>
            ''' <remarks></remarks>
            Default Overloads Property Items(ByVal Nome As String) As ElementoStr
                Get
                    Dim result As ElementoStr = ArrayFindByAtt(_atributosstr.ToArray, Nome, "Nome")
                    If IsNothing(result) Then
                        Dim Elem As ElementoStr = New ElementoStr("", ItemSeparador)
                        Elem.Nome = Nome
                        _atributosstr.Add(Elem)
                        Return _atributosstr(_atributosstr.IndexOf(Elem))
                    End If
                    Return result
                End Get
                Set(ByVal value As ElementoStr)
                    Dim pos As Integer = ArrayIndexFindByAtt(_atributosstr.ToArray, Nome, "Nome")
                    Dim Elem As ElementoStr
                    If pos = -1 Then
                        Elem = New ElementoStr(value.ToString, ItemSeparador)
                        _atributosstr.Add(Elem)
                    Else
                        Elem = _atributosstr(pos)
                        If value.Operador = ElementoStrOpera.Aumenta Then
                            Elem.ConteudoValor = Val(Elem.ConteudoValor) + Val(value.ConteudoValor)
                        ElseIf value.Operador = ElementoStrOpera.Diminui Then
                            Elem.ConteudoValor = Val(Elem.ConteudoValor) - Val(value.ConteudoValor)
                        Else
                            Elem.Conteudo = value.Conteudo
                        End If
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Quantidade de itens no estilo.
            ''' </summary>
            ''' <value>Quantidade de itens no estilo.</value>
            ''' <returns>Quantidade de itens no estilo.</returns>
            ''' <remarks></remarks>
            ReadOnly Property Count() As Integer
                Get
                    Return _atributosstr.Count
                End Get
            End Property

            ''' <summary>
            ''' Itens a serem adicionados ao estilo.
            ''' </summary>
            ''' <param name="AtributosStr">Itens a serem adicionados no estilo.</param>
            ''' <remarks></remarks>
            Sub AddStr(ByVal AtributosStr As String)
                For Each Item As Match In Regex.Matches(AtributosStr, _separadorexpr, RegexOptions.Multiline)
                    Dim Elem As ElementoStr = New ElementoStr(Item.Groups(1).Value, ItemSeparador)
                    Items(Elem.Nome) = Elem
                Next
            End Sub

            ''' <summary>
            ''' Verifica a existência ou não do termo no estilo.
            ''' </summary>
            ''' <param name="Nome">Termo a ser pesquisado.</param>
            ''' <value>TRUE caso exista ou FALSE caso não seja encontrado.</value>
            ''' <returns>TRUE caso exista ou FALSE caso não seja encontrado.</returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Exists(ByVal Nome As String) As Boolean
                Get
                    Return ArrayIndexFindByAtt(_atributosstr.ToArray, Nome, "Nome") <> -1
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Retorna posição do primeiro item no array pelo objeto ou atributo.
        ''' </summary>
        ''' <param name="LISTA">Array a ser pesquisado.</param>
        ''' <param name="Conteudo">Conteúdo que será procurado ou no índice do array ou em algum atributo.</param>
        ''' <param name="Atributo">Vazio para procurar na posição ou nome para pesquisa pela propriedade attribute.</param>
        ''' <param name="Inicio">Zero para procurar do início ou posição inicial do array.</param>
        ''' <returns>Retorna posição do item do array encontrado.</returns>
        ''' <remarks></remarks>
        Shared Function ArrayIndexFindByAtt(ByVal Lista As Array, ByVal Conteudo As String, Optional ByVal Atributo As String = "", Optional ByVal Inicio As Integer = 0) As Integer
            Dim z As Integer, item As Object = Nothing
            For z = 0 To Lista.Length - 1
                If Atributo = "" Then
                    item = Lista(z)
                Else
                    item = Lista(z).Attributes(Atributo)
                End If
                If Compare(item, Conteudo) Then
                    Exit For
                End If
            Next
            If z >= Lista.Length Then
                Return -1
            End If
            Return z
        End Function

        ''' <summary>
        ''' Classe que armazena detalhes sobre elemento de estilo do tipo "height:300px".
        ''' </summary>
        ''' <remarks></remarks>
        Public Class ElementoStr
            Private _nome As String = ""
            Private _conteudo As String = ""
            Private _separador As String
            Private _gex_valor_unid As String = "([-0-9.]+)(px|PX)?"
            Private _opera As ElementoStrOpera = ElementoStrOpera.Atribui

            ''' <summary>
            ''' Retorna string representando a forma de estilo.
            ''' </summary>
            ''' <returns>String representando a forma de estilo.</returns>
            ''' <remarks></remarks>
            Overrides Function ToString() As String
                Return _nome & _separador & _conteudo
            End Function

            ''' <summary>
            ''' Cria nova forma de estilo.
            ''' </summary>
            ''' <param name="AtributoStr">Atributo.</param>
            ''' <param name="Separador">Separador.</param>
            ''' <remarks></remarks>
            Sub New(ByVal AtributoStr As String, Optional ByVal Separador As String = ":")
                _separador = Separador
                AtributoStr = NZ(AtributoStr, "")
                Dim pos As Integer = AtributoStr.IndexOf(_separador)
                If pos = -1 Then
                    Conteudo = AtributoStr.Trim
                Else
                    Nome = AtributoStr.Substring(0, pos).Trim
                    Conteudo = AtributoStr.Substring(pos + 1).Trim
                End If
            End Sub

            ''' <summary>
            ''' Nome do atributo, parte à esquerda na definição.
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property Nome() As String
                Get
                    Return _nome
                End Get
                Set(ByVal value As String)
                    If value.StartsWith("+") Then
                        Operador = ElementoStrOpera.Aumenta
                        _nome = value.Substring(1)
                    ElseIf value.StartsWith("-") Then
                        Operador = ElementoStrOpera.Diminui
                        _nome = value.Substring(1)
                    Else
                        Operador = ElementoStrOpera.Atribui
                        _nome = value
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Conteúdo da definição, parte direita no termo.
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property Conteudo() As String
                Get
                    Return _conteudo
                End Get
                Set(ByVal value As String)
                    _conteudo = value
                End Set
            End Property

            ''' <summary>
            ''' Extração de valor do conteúdo, quando for acompanhado de termos como "PX" (pixels).
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property ConteudoValor() As String
                Get
                    Return RegexGroup(Conteudo, _gex_valor_unid, 1).Value
                End Get
                Set(ByVal value As String)
                    Conteudo = RegexGroupReplace(Conteudo, _gex_valor_unid, value, 1)
                End Set
            End Property

            ''' <summary>
            ''' Unidade do conteúdo, como "PX" (pixels).
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property ConteudoUnidade() As String
                Get
                    Return RegexGroup(Conteudo, _gex_valor_unid, 2).Value
                End Get
                Set(ByVal value As String)
                    Conteudo = RegexGroupReplace(Conteudo, _gex_valor_unid, value, 2)
                End Set
            End Property

            ''' <summary>
            ''' Acesso ao atributo definido para a classe.
            ''' </summary>
            ''' <param name="Nome">Nome do atributo.</param>
            ''' <value>Valor do atributo.</value>
            ''' <returns>Valor do atributo.</returns>
            ''' <remarks></remarks>
            Default Property Attributes(ByVal Nome As String) As String
                Get
                    If Compare(Nome, "Nome") Then
                        Return Me.Nome
                    ElseIf Compare(Nome, "Conteudo") Then
                        Return Conteudo
                    ElseIf Compare(Nome, "ConteudoValor") Then
                        Return ConteudoValor
                    ElseIf Compare(Nome, "ConteudoUnidade") Then
                        Return ConteudoUnidade
                    Else
                        Err.Raise(20000, MyBase.GetType.ToString, "Atributo '" & Nome & "' inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                    Return Nothing
                End Get

                Set(ByVal value As String)
                    If Compare(Nome, "Nome") Then
                        Nome = value
                    ElseIf Compare(Nome, "Conteudo") Then
                        Conteudo = value
                    ElseIf Compare(Nome, "ConteudoValor") Then
                        ConteudoValor = value
                    ElseIf Compare(Nome, "ConteudoUnidade") Then
                        ConteudoUnidade = value
                    Else
                        Err.Raise(20000, MyBase.GetType.ToString, "Atributo " & value & " inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Método de operação entre classes, para soma ou exclusão.
            ''' </summary>
            ''' <value>Tipo de operação.</value>
            ''' <returns>Tipo de operação.</returns>
            ''' <remarks></remarks>
            Property Operador() As ElementoStrOpera
                Get
                    Return _opera
                End Get
                Set(ByVal value As ElementoStrOpera)
                    _opera = value
                End Set
            End Property
        End Class

        ''' <summary>
        ''' Opções para interpretação do elemento quando adicionado ao conjunto elementosstr.
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum ElementoStrOpera
            Atribui
            Aumenta
            Diminui
        End Enum


    End Class

End Namespace