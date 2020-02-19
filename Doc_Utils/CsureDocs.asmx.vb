Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Imports System.Security.Principal

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://dmatters.co.uk/webservices")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class CsureDocs
    '
    '----------------------------
    ' Initialise Global Variables
    '----------------------------
    '
    Inherits System.Web.Services.WebService
    Public strProgName As String = Nothing
    Public blnDebugMode As Boolean = False
    Public strDocRef As String = Nothing
    Public strArguments As String = Nothing
    Public strServerPool As String = Nothing
    Public strResult As String = Nothing
    Public strSrvTask As String = Nothing
    Public strDocFormat As String = Nothing
    Dim strUserName As String = Nothing
    Dim strUserPwd As String = Nothing
    Dim strSrvType As String = Nothing
    Public strSettingsPath As String = "wslog\Doc_Utils_Diagnostics.txt"
    ' Public strDocIPaddr As String = "194.200.67.19"
    Public strDocIPaddr As String = "wendy.datamatters.co.uk"
    Public strEnvPath As String = HttpContext.Current.Server.MapPath(".") & "\env.txt"
    Public strDrive As String = Nothing

    <WebMethod()> _
    Public Function GetDocuments_String(ByVal strXmlRequest As String) As String
        ' <LoggingWebserviceExtensionAttribute()> _
        Dim strDocResults As String = Nothing
        ' Dim strXmlstring As String = strXmlRequest.InnerXml
        Dim strXmlString As String = strXmlRequest
        If strXmlString = "debugging" Then blnDebugMode = True
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = Nothing
        Dim strQuoteType As String = Nothing
        Dim strPolRef As String = Nothing
        Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils

        '
        '------------------------------
        ' Check logon details valid etc
        '------------------------------
        '
        DoInitialChecks(strErrors, strXmlString, blnChecksOk)
        If blnChecksOk = True Then
            '
            '----------------------------
            ' Get policy and document URL
            '----------------------------
            '
            strSrvTask = ExtractFromXml("<Service_Function", strXmlString)
            strDocRef = ExtractFromXml("<Document_Url", strXmlString)
            strPolRef = ExtractFromXml("<Service_BusinessRef", strXmlString)
            If strSrvType = "Truck" Then
                strQuoteType = "cv"
            End If
            Dim strDocument As String = Nothing
            Dim strDocDesc As String = Nothing
            strDocResults = GetSingleDocument(strDocRef, strPolRef, strSrvTask, strDocument, strDocDesc)
        Else
            strDocResults = "<Error>Error:[S2.10]" & strErrors & "</Error>"
        End If
        Return strDocResults

    End Function
    <WebMethod()> _
    Public Function GetDocuments(ByVal strXmlRequest As System.Xml.XmlDocument) As String
        '
        ' <LoggingWebserviceExtensionAttribute()> _
        '-------------------------------------
        ' Web method to get specified Document
        '-------------------------------------
        '
        '---------------------------
        ' Initialise local variables
        '---------------------------
        '
        Dim strDocResults As String = Nothing
        Dim strXmlstring As String = strXmlRequest.InnerXml
        If strXmlstring = "debugging" Then blnDebugMode = True
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = Nothing
        Dim strQuoteType As String = Nothing
        Dim strPolRef As String = Nothing
        Dim strDocument As String = Nothing
        Dim strDocDesc As String = Nothing
        Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagson = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing
        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_GetDocuments_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)
      

        '
        '------------------------------
        ' Check logon details valid etc
        '------------------------------
        '
        DoInitialChecks(strErrors, strXmlstring, blnChecksOk)
        LogMessage("Running Checks", DiagnosticsOn, strLogPath)
        If blnChecksOk = True Then
            LogMessage("Checks ok", DiagnosticsOn, strLogPath)
            '
            '----------------------------
            ' Get policy and document URL
            '----------------------------
            '
            strSrvTask = ExtractFromXml("<Service_Function", strXmlstring)
            strDocRef = ExtractFromXml("<Document_Url", strXmlstring)
            strPolRef = ExtractFromXml("<Service_BusinessRef", strXmlstring)
            If strSrvType = "Truck" Then
                strQuoteType = "cv"
            End If
            LogMessage("Running GetSingleDocument with .. Docref=" & strDocRef & " - Polref=" & strPolRef & "document= " & strDocument & " Desc= " & strDocDesc, DiagnosticsOn, strLogPath)
            strDocResults = GetSingleDocument(strDocRef, strPolRef, strSrvTask, strDocument, strDocDesc)
        Else
            strDocResults = "<Error>Error:[S2.1]" & strErrors & "</Error>"
        End If
        Return strDocResults
    End Function
    <WebMethod()> _
    <LoggingWebserviceExtensionAttribute()> _
    Public Function GetMIMEDocuments(ByVal strXmlRequest As System.Xml.XmlDocument) As String
        '
        '<LoggingWebserviceExtensionAttribute()> _
        '-------------------------------------------------------------
        ' Web method to get specified Document and wrap in MIME header
        '-------------------------------------------------------------
        '
        '---------------------------
        ' Initialise local variables
        '---------------------------
        '
        Dim strDocResults As String = Nothing
        Dim strXmlstring As String = strXmlRequest.InnerXml
        If strXmlstring = "debugging" Then blnDebugMode = True
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = Nothing
        Dim strQuoteType As String = Nothing
        Dim strPolRef As String = Nothing
        Dim strMIME As String = Nothing
        Dim strDocument As String = Nothing
        Dim strDocDesc As String = Nothing

        Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils
        '
        '------------------------------
        ' Check logon details valid etc
        '------------------------------
        '
        DoInitialChecks(strErrors, strXmlstring, blnChecksOk)
        If blnChecksOk = True Then
            '
            '----------------------------
            ' Get policy and document URL
            '----------------------------
            '
            strSrvTask = ExtractFromXml("<Service_Function", strXmlstring)
            strDocRef = ExtractFromXml("<Document_Url", strXmlstring)
            strPolRef = ExtractFromXml("<Service_BusinessRef", strXmlstring)
            If strSrvType = "Truck" Then
                strQuoteType = "cv"
            End If
            strDocResults = GetSingleDocument(strDocRef, strPolRef, strSrvTask, strDocument, strDocDesc)
            '
            '--------------------
            ' Wrap in MIME header
            '--------------------
            '
            '     strMIME = "MIME-Version: 1.0 Content-Type: multipart/mixed; boundary=MIME_boundary"
            strMIME &= "--MIME_boundary"
            strMIME &= vbCrLf & strDocResults
            strMIME &= "--MIME_boundary--"
            strDocResults = strMIME
        Else
            strDocResults = "<Error>Error:[S2.2]" & strErrors & "</Error>"
        End If
        Return strDocResults
    End Function
    <WebMethod()> _
    Public Function DocListingRequest(ByVal strXmlRequest As System.Xml.XmlDocument) As String
        '  <LoggingWebserviceExtensionAttribute()> _
        '
        '---------------------------------------------------------------------
        ' Method to return a list of available documents for a specific Policy
        '---------------------------------------------------------------------
        '
        Dim strDocResults As String = Nothing
        Dim strXmlstring As String = strXmlRequest.InnerXml
        If strXmlstring = "debugging" Then blnDebugMode = True
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = Nothing
        Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils
        Dim strQuoteType As String = Nothing
        '
        '----------------------
        ' Check Login valid etc
        '----------------------
        '
        DoInitialChecks(strErrors, strXmlstring, blnChecksOk)
        If blnChecksOk = True Then
            '
            '---------------
            ' Extract Policy
            '---------------
            '
            strDocRef = ExtractFromXml("<Service_BusinessRef", strXmlstring)
            If strSrvType = "Truck" Then
                strQuoteType = "cv"
            End If
            strDocResults = DocumentationUtilities.DocsAvailable(strServerPool, strUserName, strQuoteType, strDocRef)
        Else
            strDocResults = "<Error>Error:[S2.3]" & strErrors & "</Error>"
        End If
        Return strDocResults
    End Function
    <WebMethod()> _
    Public Function QuoteDocRequest(ByVal strXmlRequest As System.Xml.XmlDocument) As String
        ' <LoggingWebserviceExtensionAttribute()> _
        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagsOn = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing
        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_QuoteDocRequest_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)
        Dim strDocResults As String = Nothing
        Dim strAllDocResults As String = Nothing
        strDocFormat = "string"

        Dim strXmlstring As String = strXmlRequest.InnerXml

        'Dim strXmlString As String = Nothing
        'If strXmlRequest = "debugging" Then
        '    blnDebugMode = True
        'End If
        '
        '------------------------------
        ' Check logon details valid etc
        '------------------------------
        '
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = ""

        DoInitialChecks(strErrors, strXmlstring, blnChecksOk)
        If blnChecksOk = True Then
            Dim strInsurer As String = ExtractFromXml("<Document_Insurer", strXmlstring)
            Dim strPremium As String = ExtractFromXml("<Document_Premium", strXmlstring)
            '    Dim strReference As String = ExtractFromXml("<Service_BusinessRef", strXmlString)
            Dim strDocType As String = ExtractFromXml("<Document_Type", strXmlstring)
            Dim strPolRef = ExtractFromXml("<Service_BusinessRef", strXmlstring)

            Dim strMode As String = "quote"
            Dim strQuoteType As String = "cv"
            Dim strPropNo As String = ""
            Dim strDocPath As String = Nothing
            Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils

            If strDocType = "sof" Or strDocType = "all" Then
                strDocResults = DocumentationUtilities.QuoteDocsAvailable(strServerPool, strUserName, strQuoteType, strPolRef, strPremium, strInsurer)
                strDocPath = ExtractFromXml("<Document_Url", strDocResults)
                Dim strdocarr() As String = strDocPath.Split("/")
                strDocRef = strdocarr(1)
            End If
            Dim strTempDocPath As String = strDocPath
            Dim strDocument As String = Nothing
            Dim strDocDesc As String = Nothing
            Dim blnDoingAllDocs As Boolean = False

            Dim strMIME As String = "MIME-Version: 1.0 Content-Type: multipart/mixed; boundary=MIME_boundary"

            If strDocType = "all" Then
                blnDoingAllDocs = True
                If strDocType = "all" Then
                    strDocType = "kf"
                End If
            End If
DocLoop:
            If strDocType = "kf" Or strDocType = "pw" Then
                strDocRef = GetDocName(strDocType, strInsurer)
            End If

            LogMessage("Document = " & strDocRef, DiagnosticsOn, strLogPath)
            Select Case strDocType
                Case "kf"
                    ' strDocRef = strInsurer & ".pdf"
                    strDocument = "\\" & strDocIPaddr & "\e\web\policyfast\public_docs\kf\cvt\" & strDocRef
                    strDocDesc = "KeyFacts"
                    strDocPath = "kf\cvt\" & strDocRef
                Case "pw"
                    ' strDocRef = strQuoteType & "_" & strInsurer & ".pdf"
                    strDocument = "\\" & strDocIPaddr & "\e\web\policyfast\public_docs\pw\cvt\" & strDocRef
                    strDocDesc = "PolicyWording"
                    strDocPath = "pw\cvt\" & strDocRef
                Case "sof"
                    If blnDoingAllDocs = True Then strDocPath = strTempDocPath
                    strDocument = "\\" & strDocIPaddr & "\e\web"

                    If strServerPool = "doris" Then
                        strDocument &= "\dm_intra\web_test"
                    End If
                    strDocument &= "\policyfast\docs\" & strDocPath

                    strDocDesc = "StatementOfFacts"
                    strDocument = strDocument.Replace("/", "\")
                    strDocPath = strDocPath.Replace("/", "\")
            End Select
            LogMessage("Document after getdocname = " & strDocRef & "path=" & strDocument, DiagnosticsOn, strLogPath)

            strDocResults = GetSingleDocument(strDocPath, strPolRef, strSrvTask, strDocument, strDocDesc)

            If blnDoingAllDocs = True Then
                Select Case strDocType
                    Case "kf"
                        strDocType = "pw"
                    Case "pw"
                        strDocType = "sof"
                    Case "sof"
                        strDocType = "finished"
                End Select
                strMIME &= "--MIME_boundary"
                strMIME &= vbCrLf & strDocResults
                strMIME &= "--MIME_boundary--"
                If strDocType <> "finished" Then GoTo DocLoop
                strDocResults = strMIME
            Else
                '
                '--------------------
                ' Wrap in MIME header
                '--------------------
                '
                strMIME &= "--MIME_boundary"
                strMIME &= vbCrLf & strDocResults
                strMIME &= "--MIME_boundary--"
                strDocResults = strMIME
            End If
        End If

        Return strDocResults
    End Function
    <WebMethod()> _
    Public Function QuoteDocRequestXml(ByVal strXmlRequest As XmlDocument) As System.Xml.XmlDocument
        '  <LoggingWebserviceExtensionAttribute()> _
        Dim strDocResults As String = Nothing
        Dim strAllDocResults As String = Nothing
        strDocFormat = "xml"

        Dim strXmlstring As String = strXmlRequest.InnerXml
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        Dim strFullyFormattedResponse As String = Nothing

        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagson = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing
        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_QuoteDocRequestXml_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)

        'Dim strXmlString As String = Nothing
        'If strXmlRequest = "debugging" Then
        '    blnDebugMode = True
        'End If
        '
        '------------------------------
        ' Check logon details valid etc
        '------------------------------
        '
        Dim blnChecksOk As Boolean = False
        Dim strErrors As String = ""

        DoInitialChecks(strErrors, strXmlstring, blnChecksOk)
        LogMessage("blnChecksOk = " & blnChecksOk, DiagnosticsOn, strLogPath)
        If blnChecksOk = True Then
            Dim strInsurer As String = ExtractFromXml("<Document_Insurer", strXmlstring)
            Dim strPremium As String = ExtractFromXml("<Document_Premium", strXmlstring)
            Dim strDocType As String = ExtractFromXml("<Document_Type", strXmlstring)
            Dim strPolRef = ExtractFromXml("<Service_BusinessRef", strXmlstring)

            Dim strMode As String = "quote"
            Dim strQuoteType As String = "cv"
            Dim strPropNo As String = ""
            Dim strDocPath As String = Nothing
            Dim DocumentationUtilities As DocUtils.DocUtils = New DocUtils.DocUtils

            If strDocType = "sof" Or strDocType = "all" Then
                strDocResults = DocumentationUtilities.QuoteDocsAvailable(strServerPool, strUserName, strQuoteType, strPolRef, strPremium, strInsurer)
                strDocPath = ExtractFromXml("<Document_Url", strDocResults)
                Dim strdocarr() As String = strDocPath.Split("/")
                strDocRef = strdocarr(1)
                LogMessage("strdocresults=" & strDocResults, DiagnosticsOn, strLogPath)
            End If
            Dim strTempDocPath As String = strDocPath
            Dim strDocument As String = Nothing
            Dim strDocDesc As String = Nothing
            Dim blnDoingAllDocs As Boolean = False


            Dim strMIME As String = "MIME-Version: 1.0 Content-Type: multipart/mixed; boundary=MIME_boundary"

            If strDocType = "all" Then
                blnDoingAllDocs = True
                If strDocType = "all" Then
                    strDocType = "kf"
                End If
            End If
            strFullyFormattedResponse = "<documents>"
DocLoop:
            LogMessage("doctype = " & strDocType, DiagnosticsOn, strLogPath)
            If strDocType = "kf" Or strDocType = "pw" Then
                strDocRef = GetDocName(strDocType, strInsurer)
            End If
            LogMessage("Document = " & strDocRef, DiagnosticsOn, strLogPath)

            Select Case strDocType
                Case "kf"
                    ' strDocRef = strInsurer & ".pdf"
                    strDocument = "\\" & strDocIPaddr & "\e\web\policyfast\public_docs\kf\cvt\" & strDocRef
                    strDocDesc = "KeyFacts"
                    strDocPath = "kf\cvt\" & strDocRef
                Case "pw"
                    ' strDocRef = strQuoteType & "_" & strInsurer & ".pdf"
                    strDocument = "\\" & strDocIPaddr & "\e\web\policyfast\public_docs\pw\cvt\" & strDocRef
                    strDocDesc = "PolicyWording"
                    strDocPath = "pw\cvt\" & strDocRef
                Case "sof"
                    If blnDoingAllDocs = True Then strDocPath = strTempDocPath
                    strDocument = "\\" & strDocIPaddr & "\e\web"

                    If strServerPool = "doris" Then
                        strDocument &= "\dm_intra\web_test"
                    End If
                    strDocument &= "\policyfast\docs\" & strDocPath

                    strDocDesc = "StatementOfFacts"
                    strDocument = strDocument.Replace("/", "\")
                    strDocPath = strDocPath.Replace("/", "\")
            End Select
            LogMessage("path after getdocname = " & strDocPath, DiagnosticsOn, strLogPath)
            strDocResults = GetSingleDocument(strDocPath, strPolRef, strSrvTask, strDocument, strDocDesc)
            If blnDoingAllDocs = True Then
                Select Case strDocType
                    Case "kf"
                        strDocType = "pw"
                    Case "pw"
                        strDocType = "sof"
                    Case "sof"
                        strDocType = "finished"
                End Select
                strMIME = DocumentationUtilities.GetFormattedDocument(strDocFormat, strDocResults)
                strFullyFormattedResponse &= strMIME

                If strDocType <> "finished" Then GoTo DocLoop
                strFullyFormattedResponse &= "</documents>"
            Else
                strMIME = DocumentationUtilities.GetFormattedDocument(strDocFormat, strDocResults)
                strFullyFormattedResponse &= strMIME
                strFullyFormattedResponse &= "</documents>"
            End If
        End If
EndOfFunction:
        Dim XmlDocResults As XmlDocument = New XmlDocument
        Try
            XmlDocResults.LoadXml(strFullyFormattedResponse)
        Catch ex As Exception
            XmlDocResults.LoadXml("<error> trying to load xml for " & ex.Message & "</error>")
        End Try

        Return XmlDocResults
    End Function
    Private Function ExtractFromXml(ByVal strElement As String, ByVal strXmlRequest As String) As String
        '
        '--------------------------------------
        ' Utility to extract data from XML file
        '--------------------------------------
        '
        Dim strXmlField As String = Nothing
        Dim IntTagStart As Integer = 0
        Dim strTemp As String = Nothing
        Dim strTempArray As String()
        Dim intExtractLen As Integer = 0
        IntTagStart = strXmlRequest.IndexOf(strElement)
        If IntTagStart < 0 Then
            strXmlField = "Error:[S2.4]Tag not found"
        Else
            intExtractLen = strXmlRequest.Length - IntTagStart
            If IntTagStart > 0 Then
                strTemp = Mid(strXmlRequest, IntTagStart, intExtractLen)
            Else
                strTemp = Left(strXmlRequest, intExtractLen)
            End If

            strTempArray = strTemp.Split(Chr(34))
            strXmlField = strTempArray(1)
        End If
        Return strXmlField
    End Function
    Private Sub DoInitialChecks(ByRef strErrorMsg As String, ByRef strXmlRequest As String, ByRef blnChecksOk As Boolean)
        '
        '-------------------------------------------------------
        ' Basic checks of login details and validity of XML file
        '-------------------------------------------------------
        '
        Dim logonOk As Boolean = True
        Dim strEnv As String = Nothing
        Dim strHeaderDets As String = Nothing
        Dim CsureDets As csdets.UserDetails = New csdets.UserDetails
        Dim xmldets As XmlTools.Xml = New XmlTools.Xml
        Dim cstools As Csure_Tools.CsureTools = New Csure_Tools.CsureTools
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagson = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing
        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_DoInitialChecks_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)

        If blnDebugMode = True Then
            strXmlRequest = Nothing
            ImportFile("d:\Truck_Webservice\xmlfiles\ServiceRequests\AllQuoteDocs.xml", strXmlRequest)
        End If
        '
        '----------------------
        ' Check the Environment
        '----------------------
        '
        strUserName = ExtractFromXml("<Header_UserName", strXmlRequest)
        LogMessage("strusername=" & strUserName, DiagnosticsOn, strLogPath)
        '
        ' Take environment from header for insurecom, auto-assign for 
        ' Acturis
        '
        If strUserName = "insurecom" Or Mid(strUserName, 1, 4) = "ICL-" Then
            strEnv = ExtractFromXml("Header_Environment", strXmlRequest)
        Else
            LogMessage("strEnvPath = " & strEnvPath, DiagnosticsOn, strLogPath)
            ImportFile(strEnvPath, strEnv)
        End If

        If Left(strEnv, 6) = "Error" Then
            strErrorMsg = "Error:[S2.5]No Environment specified"
            GoTo ExitPoint
        End If
        LogMessage("strEnv = " & strEnv, DiagnosticsOn, strLogPath)

        Select Case strEnv
            Case "T"
                strServerPool = "doris"
            Case "L"
                strServerPool = "wendy"
            Case Else
                strErrorMsg = "Error:[S2.5]No Environment specified"
                GoTo ExitPoint
        End Select
        '
        '-------------------------
        ' Get webservice settings
        '-------------------------
        '
        cstools.GetUserParams(strUserName, strServerPool)

        '
        '--------------------
        ' Check logon details
        '--------------------
        '        

        strUserPwd = ExtractFromXml("<Header_Password", strXmlRequest)
        If strUserName = Nothing Or strUserPwd = Nothing Then
            strErrorMsg = "Error:[S2.6]Invalid Credentials Supplied"
        Else
            logonOk = CsureDets.DetailsOk(strUserName, strUserPwd, "", "", strServerPool, cstools.TrustedUser)
            If logonOk = True Then
                '
                '------------------------------------------------
                ' Check a correct service task has been specified
                '------------------------------------------------
                '
                strSrvTask = ExtractFromXml("<Service_Function", strXmlRequest)
                Dim strSelCriteria As String = ExtractFromXml("<Document_Url", strXmlRequest)
                strSrvType = ExtractFromXml("<Service_BusinessType", strXmlRequest)
                Select Case strSrvTask
                    Case "DocumentListing"
                        If strSelCriteria = "All" Then
                            If strSrvType = "Truck" Then
                                blnChecksOk = True
                            Else
                                strErrorMsg = "<Error>Error:[S2.7]Invalid Service Type Requested</Error>"
                            End If
                        Else
                            strErrorMsg = "<Error>Error:[S2.8]Invalid Selection Criteria for this service type</Error>"
                        End If
                    Case "Request"
                        If strSrvType = "Truck" Then
                            blnChecksOk = True
                        End If
                    Case "QuoteDocsRequest"
                        If strSrvType = "Truck" Then
                            blnChecksOk = True
                        End If
                    Case Else
                        strErrorMsg = "<Error>Error:[S2.9]Invalid Service Function Requested</Error>"
                End Select

            Else
                strErrorMsg = CsureDets.StatusMessage
            End If
        End If
ExitPoint:
    End Sub
    Private Sub ImportFile(ByVal InputFilename As String, ByRef DataStream As String)
        '
        '-----------------------------------------------------
        ' Utility to import file, path supplied, file returned
        '-----------------------------------------------------
        '
        Dim UserDets As Microsoft.VisualBasic.ApplicationServices.User = New Microsoft.VisualBasic.ApplicationServices.User
        Dim strUser As String = UserDets.Name
        Dim ss As System.Web.Services.WebService = New System.Web.Services.WebService
        Dim stru As String = ss.User.Identity.Name
        Dim InputByte As String

        Try
            Dim fs As New IO.FileStream(InputFilename, IO.FileMode.Open)
            Dim sr As IO.BinaryReader = New IO.BinaryReader(fs)
            Dim sb As New System.Text.StringBuilder
            Do
                InputByte = Chr(sr.ReadByte)
                sb.Append(InputByte)
            Loop Until fs.Position = fs.Length
            '
            '-------------------------------------------------------------------------
            ' Append the string builder variable to the DataStream variable for return
            '-------------------------------------------------------------------------
            '
            DataStream = DataStream & sb.ToString()
            sr.Close()
        Catch ex As Exception
            DataStream = "<Error>Error:[S2.10]" & ex.Message & " for " & stru
        End Try

    End Sub

    Private Function GetSingleDocument(ByVal strDocRef As String, ByVal strPolRef As String, ByVal strSrvTask As String, ByVal strDocument As String, ByVal strDocDesc As String) As String
        '
        '----------------------------------------------------
        ' Function to get single Document as encoded PDF file
        '----------------------------------------------------
        '
        Dim strFullPath As String = Nothing
        Dim strPDF As String = Nothing
        Dim strBinaryPdf As String = Nothing
        Dim strFinalResult As String = Nothing
        Dim bytePDF As Byte() = Nothing
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagson = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing

        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_GetSingleDocument_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)
        LogMessage("Strdocument = " & strDocument & " Strdocdesc = " & strDocDesc & " strDocRef= " & strDocRef & " strsrvtask = " & strSrvTask, DiagnosticsOn, strLogPath)

        If strDocument = Nothing Then
            '
            '-------------------------
            ' First, get full pathname
            '-------------------------
            '
            LogMessage("calling GetFullDocPath: docname=" & strDocRef & "&reference=" & strPolRef, DiagnosticsOn, strLogPath)
            strProgName = "GetFullDocPath"
            strArguments = "&docname=" & strDocRef & "&reference=" & strPolRef & "&path=mds,pf-fleet,&formatreqd=xml"
            ConnectToD3()
            strDocument = ExtractFromXml("<FullPath", strResult)
            strDocDesc = ExtractFromXml("<FullDesc", strResult)
        End If
        LogMessage("strResult=" & strResult, DiagnosticsOn, strLogPath)
        '
        '-----------------------------------------
        ' Read in PDF file using returned pathname
        '-----------------------------------------
        '
        strDocument = strDocument.Replace("\", "/")
        LogMessage("document= " & strDocument, DiagnosticsOn, strLogPath)
        For Each c In Path.GetInvalidPathChars()
            If strDocument.Contains(c) Then
                strDocument = strDocument.Replace(c.ToString, "")
            End If
        Next
        Try
            ReadBinaryData(strDocument, bytePDF)
        Catch ex As Exception
            LogMessage("ex.message = " & ex.Message & "strdocument= " & strDocument, DiagnosticsOn, strLogPath)
            strResult = ex.Message
            GoTo EndOfDocService
        End Try
        LogMessage("Binary Result = " & bytePDF.Length.ToString, DiagnosticsOn, strLogPath)
        '
        ' ImportFile(strDocument, strPDF)
        ' If Left(strPDF, 7) <> "<Error>" Then
        '
        '-----------------
        ' Encode to Base64
        '-----------------
        '
        'bytePDF = Encoding.UTF8.GetBytes(strPDF)
        strBinaryPdf = System.Convert.ToBase64String(bytePDF)
        strResult = vbCrLf & "Content-Type: application/pdf"
        strResult &= vbCrLf & "Content-Transfer-Encoding : base64"
        strResult &= vbCrLf & "Content-ID : <" & strPolRef & "_" & strDocDesc & "_" & strDocRef & ">"
        strResult &= vbCrLf & "Content-Disposition: attachment; filename=" & strDocDesc & ".pdf"
        strResult &= vbCrLf & vbCrLf & strBinaryPdf
        strResult &= vbCrLf & vbCrLf
        'Else
        'Dim struser As String = Environment.UserName
        'strResult = strPDF
        'End If
EndOfDocService:
        Return strResult
    End Function
    Private Sub ConnectToD3()
        '
        '----------------------------------------------------------------
        ' Utility to connect to D3 and run specified FlashConnect routine
        '----------------------------------------------------------------
        '
        Dim strArg As String = Nothing
        Dim strMethod As String = "POST"
        '      Dim strServerPool As String = "doris"
        Dim Http As New MSXML.XMLHTTPRequest
        Dim RandomNumber As Random = New Random
        Dim strRandomString As String = "&random=" & RandomNumber.Next.ToString
        Dim strDomain As String = Nothing
        Select Case strServerPool
            Case "doris"
                strDomain = "http://www.datamatters.info"
            Case "wendy"
                strDomain = "http://www.coversure.co.uk"
        End Select
        Dim ConnectionString As String = strDomain & "/cgi-bin/fccgi.exe?w3exec=" + strProgName
        If strServerPool <> Nothing Then
            strArg = "&w3serverpool=" & strServerPool & strArguments & strRandomString
        Else
            strArg = strArguments & strRandomString
        End If

        If strMethod = Nothing Then
            strMethod = "GET"
        End If

        Try
            Http.open(strMethod, ConnectionString, False)
            Http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            Http.setRequestHeader("Timeout", "600000")
            Http.send(strArg)
            strResult = Http.responseText
        Catch ex As Exception
            strResult = strProgName & " - " & ex.Message
        End Try
    End Sub

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Protected Sub ReadBinaryData(ByVal strPath As String, ByRef strRecord As Byte())
        '
        '----------------------------------------------------
        ' Read data from file, path selected, record returned
        '----------------------------------------------------
        '
        Dim input As New System.IO.FileStream(strPath, System.IO.FileMode.Open)
        Dim reader As New System.IO.BinaryReader(input)
        strRecord = reader.ReadBytes(CInt(input.Length))
    End Sub
    Private Sub LogMessage(ByVal ErrorMsg As String, ByVal DiagnosticsOn As Boolean, ByVal strLogPath As String)
        If DiagnosticsOn = True Then
            Try
                FileOpen(1, strLogPath, OpenMode.Append)
                Print(1, vbCrLf & "=================================" & vbCrLf & ErrorMsg & vbCrLf & "=================================" & vbCrLf & vbCrLf)
                FileClose(1)
            Catch ex As Exception
                strResult = "error:" & ex.Message & "drive = " & strDrive & "PATH=" & strSettingsPath
                FileOpen(1, "e:\wslog\truck\general.txt", OpenMode.Append)
                Print(1, strResult)
                FileClose(1)
            End Try
        End If
    End Sub
    Private Function GetDocName(ByVal strDocType As String, ByVal strInsurer As String) As String
        '
        '--------------------------------------
        ' Utility to retrieve latest document
        '--------------------------------------
        '
        strProgName = "getlatestdoc"
        strArguments = "&insurer=" & strInsurer & "&doctype=" & strDocType & "&path=mds,pf-fleet,"
        ConnectToD3()
        Return strResult
    End Function
    Public Function FilenameIsOK(ByVal fileNameAndPath As String) As Boolean
        Dim LogDetails As String = Nothing
        Dim strDiagsOn As String = Nothing
        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagsOn = "OFF"
        End Try
        Dim DiagnosticsOn As Boolean = False
        Dim strLogPath As String = Nothing

        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_GetSingleDocument_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If

        LogMessage("Opening Log to " & strLogPath, DiagnosticsOn, strLogPath)
        For Each c In Path.GetInvalidFileNameChars()
            If fileNameAndPath.Contains(c) Then
                LogMessage("invalid file character of " & c.ToString, DiagnosticsOn, strLogPath)
            End If
        Next
        For Each c In Path.GetInvalidPathChars()
            If fileNameAndPath.Contains(c) Then
                LogMessage("invalid path character of " & c.ToString, DiagnosticsOn, strLogPath)
            End If
        Next
        Return True
    End Function
End Class
