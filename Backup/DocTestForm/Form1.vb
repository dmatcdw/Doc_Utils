Imports System.Xml
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim xmlRequest As XmlDocument = New XmlDocument
        Dim xmlResults As XmlNode = Nothing
        xmlrequest.LoadXml(TextBox1.Text)
        Dim svcDocUtils As DocUtilsSvc.CsureDocs = New DocUtilsSvc.CsureDocs
        xmlResults = svcDocUtils.QuoteDocRequestXml(xmlRequest)
        TextBox2.Text = xmlResults.InnerXml
    End Sub
End Class
