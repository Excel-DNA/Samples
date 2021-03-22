# CustomXMLParts

This sample shows CustomXMLParts handling techniques from my [DB-Addin](https://github.com/rkapl123/DBAddin) as well as XML schema validation.

Here, I use CustomXMLParts as an advanced storage for Database modifier definitions (for writing Excel Data to a Database table (DBMapper), doing DML Statements such as insert/update/delete (DBAction) and executing sequences of DBMappers and DBActions (DB Sequence)).

The Sample only contains creating and viewing/editing DBMapper Definitions in a stripped down class (all in one) without any other meaningful action code to concentrate on the usage of CustomXMLParts. Besides that, I demonstrate the usage of validating the XML against an existing schema.

## CustomXMLParts Usage

All CustomXMLParts Objects are found in the namespace `Microsoft.Office.Core`, referenced by the Office.dll within Exceldna.Interop.  

Fetching an existing CustomXMLParts XML document is done by selecting the required namespace:

```VB
        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
```

Adding the XML document if the namespace doesn't yet exist (in a new workbook) is done by adding the root element:
```VB
        If CustomXmlParts.Count = 0 Then
            ' in case no CustomXmlPart in Namespace DBModifDef exists in the workbook, add one
            ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        End If
```

### Adding sub elements

Sub elements are added with the Methond `AppendChildNode` of the selected node:

```VB
        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType, NamespaceURI:="DBModifDef")
```

The appended child element is placed last, to append further child elements, you need to call `LastChild`
```VB
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
        ' append the detailed settings to the definition element
        dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:=createdDBModifType + Guid.NewGuid().ToString())
        dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
```

### Retrieving elements

When retrieving element values, it's a good idea to check for the count of nodes contained to avoid exceptions:
```VB
        Dim nodeCount As Integer = definitionXML.SelectNodes("ns0:" + nodeName).Count
        If nodeCount = 0 Then
            getParamFromXML = "" ' optional nodes become empty strings
        Else
            getParamFromXML = definitionXML.SelectSingleNode("ns0:" + nodeName).Text
        End If
```

### Iterating through nodes

When iterating through nodes you take the `ChildNodes` method of the (root) node object and us `BaseName` of the iterator variable (node object) to get it's element name. 
Here the name is usually in the (one and only) attribute "name" of the element, so if that exists, it is taken as the nodes name.

```VB
	For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
		Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
		If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
			Dim nodeName As String
			If customXMLNodeDef.Attributes.Count > 0 Then
				nodeName = customXMLNodeDef.Attributes(1).Text
			Else
				nodeName = customXMLNodeDef.BaseName + "unknown"
			End If

			' finally create the DBModif Object and fill parameters into CustomXMLPart:
			Dim newDBModif As DBModif = New DBModif(customXMLNodeDef, DBModiftype)
...
```

## SchemaFile Usage

To utilize validated XML files, you need to create a [XML schema definition, short XSD](https://www.w3.org/TR/xmlschema11-1/) document, that you add to a [XML resource component](https://docs.microsoft.com/en-us/dotnet/framework/resources/creating-resource-files-for-desktop-apps).
Of course you can add as many xsd schema definitions as you want to the XML resource component.

After adding the XSD file to the XML resource component, you can fetch the schema in your code with  
`Dim schemaString As String = My.Resources.<XML resource Name>.<XSD Name>`.


Validation of an existing XML string is then done with
```VB
...
	Dim schemaString As String = My.Resources.<XML resource Name>.<XSD Name>
	Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
	doc.Schemas.Add("<Schema Name>", schemadoc)
	Dim eventHandler As Schema.ValidationEventHandler = New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
	doc.LoadXml(<your XML String, e.g. taken from Me.EditBox.Text>)
	doc.Validate(eventHandler)
...
	Sub myValidationEventHandler(sender As Object, e As Schema.ValidationEventArgs)
		' simply pass back Errors and Warnings as an exception
		If e.Severity = Schema.XmlSeverityType.Error Or e.Severity = Schema.XmlSeverityType.Warning Then Throw New Exception(e.Message)
	End Sub
```

## Testing Sample

After building the sample, simply activate the Addin by opening CustomXMLParts-AddIn64.xll or CustomXMLParts-AddIn.xll (32 bit) in bin/Debug or bin/Release and select the new CustomXMLParts Ribbon.