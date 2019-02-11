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

    End Class
End Namespace