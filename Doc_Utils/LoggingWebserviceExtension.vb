Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Web.Services.Protocols
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
''' <summary>
''' LoggingWebserviceExtension
''' 
''' Roughly based on the discussion in http://msdn.microsoft.com/en-us/magazine/cc188761.aspx
''' and http://msdn.microsoft.com/en-us/magazine/cc164007.aspx
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' NOT FOR PRODUCTION
''' 
''' [MG 18/06/2012]
''' </summary>
Public Class LoggingWebserviceExtension
    Inherits SoapExtension
    Private oldStream As Stream
    Private newStream As Stream

    ' Save the Stream representing the SOAP request or SOAP response into
    ' a local memory buffer.
    Public Overrides Function ChainStream(ByVal stream As Stream) As Stream
        oldStream = stream
        newStream = New MemoryStream()
        Return newStream
    End Function
    ''' <summary>
    ''' root path for logging
    ''' </summary>
    Private ReadOnly _loggingDirectory As String
    Public Sub New()
        'hardcoded for simplicity, obviously, this variable could be moved to a web.config key/value config pair

        _loggingDirectory = "e:\wslog\"
        

    End Sub

    Public Overrides Function GetInitializer(ByVal serviceType As Type) As Object
        'throw new Exception("The method or operation is not implemented.");
        Return Nothing
    End Function

    Public Overrides Function GetInitializer(ByVal methodInfo As LogicalMethodInfo, ByVal attribute As SoapExtensionAttribute) As Object
        'throw new Exception("The method or operation is not implemented.");
        Return Nothing
    End Function

    Public Overrides Sub Initialize(ByVal initializer As Object)
        'throw new Exception("The method or operation is not implemented.");
        Return
    End Sub

    Private Function GetFilePath(ByVal msg As SoapMessage) As String
        Dim stage As String = "unknownStage"
        Dim SearchString As String = Nothing
        Dim action As String = msg.MethodInfo.Name
        Select Case msg.Stage
            Case SoapMessageStage.BeforeDeserialize
                stage = "rq"
                Select Case action
                    Case "QuoteDocRequest"
                        SearchString = "Service_BusinessRef"
                    Case "QuoteDocRequestXml"
                        SearchString = "Service_BusinessRef"
                    Case "GetQuote"
                        SearchString = "Header_TransRef"
                    Case "ProcessMTA"
                        SearchString = "Header_PolRef"
                    Case Else
                        SearchString = "Header_QuoteRef"
                End Select

                Exit Select
            Case SoapMessageStage.AfterSerialize
                If action = "QuoteDocRequest" Then
                    SearchString = "Content-ID"
                Else
                    SearchString = "id"
                End If
                stage = "rs"
                Exit Select
        End Select

        msg.Stream.Position = 0

        Dim tmpreader As StreamReader = New StreamReader(msg.Stream)
        Dim strTest As String = tmpreader.ReadToEnd
        Dim intEndofXml As Integer = 0
        Dim strPart1 As String = Nothing
        Dim strTransRef As String = Nothing
        Try
            intEndofXml = strTest.IndexOf(SearchString, 1)
            If intEndofXml > -1 Then
                strPart1 = Right(strTest, (strTest.Length - intEndofXml))
            Else
                strPart1 = "notfound"
            End If
        Catch ex As Exception
            strPart1 = ex.Message
        End Try
        Dim strXmlArray() As String
        If stage = "rq" Then
            strXmlArray = strPart1.Split(Chr(34))
        Else
            Dim strDelim As String = Nothing
            If action = "QuoteDocRequest" Then
                strXmlArray = strPart1.Split(";")
            Else
                strXmlArray = strPart1.Split(">")
            End If
            strPart1 = strXmlArray(1)
            strXmlArray = strPart1.Split("_")
        End If
        If stage = "rq" Then
            strTransRef = strXmlArray(1)
        Else
            strTransRef = strXmlArray(0)
            strTransRef = strTransRef.Replace("<", Nothing)
            strTransRef = strTransRef.Replace(">", Nothing)
            'strTransRef = SearchString & "_arrlen" & strXmlArray.Length.ToString
        End If

        msg.Stream.Position = 0
        Dim fileName As String = action & "_" & stage & "_" & strTransRef & "_" & Convert.ToString(Guid.NewGuid()) & ".dump"
        Console.WriteLine(fileName)
        Return System.IO.Path.Combine(_loggingDirectory, fileName)
    End Function

    ' ProcessMessage is called to process SOAP messages
    ' after inbound messages are deserialized to input
    ' parameters and output parameters are serialized to
    ' outbound messages
    Public Overrides Sub ProcessMessage(ByVal message As SoapMessage)
        Select Case message.Stage
            Case SoapMessageStage.BeforeDeserialize
                WriteInput(message)
                Exit Select
            Case SoapMessageStage.AfterSerialize
                WriteOutput(message)
                Exit Select
        End Select

    End Sub

    Private Sub WriteLog(ByVal message As SoapMessage)
        Dim filepath As String = GetFilePath(message)
        Dim outFs As New StreamWriter(filepath, True)
        message.Stream.Position = 0
        Dim sr As New StreamReader(message.Stream)
        outFs.Write(sr.ReadToEnd())
        outFs.Flush()
        message.Stream.Position = 0
    End Sub
    Public Sub WriteOutput(ByVal message As SoapMessage)
        newStream.Position = 0
        Dim fs As New FileStream(GetFilePath(message), FileMode.Append, FileAccess.Write)
        Dim w As New StreamWriter(fs)

        Dim soapString As String = If((TypeOf message Is SoapServerMessage), "SoapResponse", "SoapRequest")
        w.WriteLine("-----" & soapString & " at " & DateTime.Now)
        w.Flush()
        Copy(newStream, fs)
        w.Close()
        newStream.Position = 0
        Copy(newStream, oldStream)
    End Sub
    Public Sub WriteInput(ByVal message As SoapMessage)
        Copy(oldStream, newStream)
        Dim fs As New FileStream(GetFilePath(message), FileMode.Append, FileAccess.Write)
        Dim w As New StreamWriter(fs)

        Dim soapString As String = If((TypeOf message Is SoapServerMessage), "SoapRequest", "SoapResponse")
        w.WriteLine("-----" & soapString & " at " & DateTime.Now)
        w.Flush()
        newStream.Position = 0
        Copy(newStream, fs)
        w.Close()
        newStream.Position = 0
    End Sub

    ' Copy copies the contents of a source stream 
    ' to a destination stream
    Private Sub Copy(ByVal src As Stream, ByVal dest As Stream)
        Dim reader As New StreamReader(src)
        Dim writer As New StreamWriter(dest)
        writer.Write(reader.ReadToEnd())
        writer.Flush()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

''' <summary>
''' Simply provides an attribute wrapper for the logging extension
''' </summary>
Public Class LoggingWebserviceExtensionAttribute
    Inherits SoapExtensionAttribute
    Private _extPriority As Integer

    Public Sub New()
        _extPriority = 1
    End Sub

    Public Overrides ReadOnly Property ExtensionType() As Type
        Get
            Return GetType(LoggingWebserviceExtension)
        End Get
    End Property

    Public Overrides Property Priority() As Integer
        Get
            Return _extPriority
        End Get
        Set(ByVal value As Integer)
            _extPriority = value
        End Set
    End Property
End Class
