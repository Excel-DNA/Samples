Imports System.Xml
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports System.IO

''' <summary>Dialog used to display and edit the CustomXMLPart utilized for storing the DBModif definitions</summary>
Public Class EditDBModifDef
    ''' <summary>the edited CustomXmlParts for the DBModif definitions</summary>
    Private CustomXmlParts As Object

    ''' <summary>put the custom xml definition in the edit box for display/editing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditDBModifDef_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        ' Make a StringWriter to hold the result.
        Using sw As New System.IO.StringWriter()
            ' Make a XmlTextWriter to format the XML.
            Using xml_writer As New XmlTextWriter(sw)
                xml_writer.Formatting = Formatting.Indented
                Dim doc As New XmlDocument()
                doc.LoadXml(CustomXmlParts(1).XML)
                doc.WriteTo(xml_writer)
                xml_writer.Flush()
                ' Display the result.
                Me.EditBox.Text = sw.ToString
            End Using
        End Using
    End Sub

    ''' <summary>store the displayed/edited textbox content back into the custom xml definition, indluding validation feedback</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click

        ' Make a StringWriter to reformat the indented XML.
        Using sw As New System.IO.StringWriter()
            ' Make a XmlTextWriter to (un)format the XML.
            Using xml_writer As New XmlTextWriter(sw)
                ' revert indentation...
                xml_writer.Formatting = Formatting.None
                Dim doc As New XmlDocument()
                Try
                    ' validate definition XML
                    Dim schemaString As String = My.Resources.SchemaFiles.DBModifDef
                    Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
                    doc.Schemas.Add("DBModifDef", schemadoc)
                    Dim eventHandler As Schema.ValidationEventHandler = New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
                    doc.LoadXml(Me.EditBox.Text)
                    doc.Validate(eventHandler)
                Catch ex As Exception
                    DBModifs.ErrorMsg("Problems with parsing changed definition: " + ex.Message, "Edit DB Modifier Definitions XML")
                    Exit Sub
                End Try
                doc.WriteTo(xml_writer)
                xml_writer.Flush()
                ' store the result in CustomXmlParts
                CustomXmlParts(1).Delete
                CustomXmlParts.Add(sw.ToString)
            End Using
        End Using

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>validation handler for XML schema (DBModifDef) checking</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub myValidationEventHandler(sender As Object, e As Schema.ValidationEventArgs)
        ' simply pass back Errors and Warnings as an exception
        If e.Severity = Schema.XmlSeverityType.Error Or e.Severity = Schema.XmlSeverityType.Warning Then Throw New Exception(e.Message)
    End Sub

    ''' <summary>no change was made to definition</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>show the current line and column for easier detection of problems in xml document</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_SelectionChanged(sender As Object, e As EventArgs) Handles EditBox.SelectionChanged
        Me.PosIndex.Text = "Line: " + (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString + ", Column: " + (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString
    End Sub
End Class