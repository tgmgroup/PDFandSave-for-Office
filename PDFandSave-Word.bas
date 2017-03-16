Attribute VB_Name = "PDFandSave"
' Create New Function 

PDFandSavePublic Sub PDFandSaveWord() 

	ActiveDocument.Save
	SaveActiveDocumentAsPdfWord
	
	End SubSub 

SaveActiveDocumentAsPdfWord()

	Dim strPath As String    
	On Error GoTo Errhandler    
	
	If InStrRev(ActiveDocument.FullName, ".") <> 0 Then        
		strPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1) & ".pdf"        
		ActiveDocument.SaveAs FileName:=strPath, FileFormat:=wdFormatPDF    
		End If    
		
	On Error GoTo 0    
	
	Exit Sub

Errhandler:

	MsgBox "There was an error saving a copy of this document as PDF. " & _    
	"Ensure that the PDF is not open for viewing and that the destination path is writable. Error code: " & Err
	
	End Sub
