Attribute VB_Name = "PDFandSaveExcel"
' Creates New Functions PDFandSaveExcelSheet(active sheet) and PDFandSaveExcelWB(all sheets in workbook)
Public Sub PDFandSaveExcelSheet()

	ActiveWorkbook.Save
	SaveActiveDocumentAsPdfExcelSheet
	
	End Sub

Public Sub PDFandSaveExcelWB()
	
	ActiveWorkbook.Save
	SaveActiveDocumentAsPdfExcelWB
	
	End Sub
	
Sub SaveActiveDocumentAsPdfExcelSheet()
	
	Dim strPath As String
	
	On Error GoTo Errhandler
	
	If InStrRev(ActiveWorkbook.FullName, ".") <> 0 Then
	
		strPath = Left(ActiveWorkbook.FullName, InStr(ActiveWorkbook.FullName, ".") - 1) & ".pdf"
		ActiveSheet.ExportAsFixedFormat _
		Type:=xlTypePDF, _
		Filename:=strPath, _
		Quality:=xlQualityStandard, _
		IncludeDocProperties:=True, _
		IgnorePrintAreas:=False, _ 
		OpenAfterPublish:=True
		End If
		
	On Error GoTo 0
		
	Exit Sub

Errhandler:
	
	MsgBox "There was an error saving a copy of this document as PDF. " & _    
	"Ensure that the PDF is not open for viewing and that the destination path is writable. Error code: " & Err

	End Sub

Sub SaveActiveDocumentAsPdfExcelWB()

	Dim strPath As String
	Dim sheetName As String
	Dim workSheet As workSheet    
	
	On Error GoTo Errhandler    
	
	If InStrRev(ActiveWorkbook.FullName, ".") <> 0 Then        
		
		strPath = Left(ActiveWorkbook.FullName, InStr(ActiveWorkbook.FullName, ".") - 1)    
		
		For Each workSheet In Worksheets        
		
			workSheet.Select        
			sheetName = workSheet.Name        
			ActiveSheet.ExportAsFixedFormat _        
			Type:=xlTypePDF, _        
			Filename:=strPath & " - " & sheetName & ".pdf", _        
			Quality:=xlQualityStandard, _        
			IncludeDocProperties:=True, _        
			IgnorePrintAreas:=False, _        
			OpenAfterPublish:=True    
			
			Next workSheet    
		
		End If       
	
	On Error GoTo 0    
	
Exit Sub

Errhandler:    
	
	MsgBox "There was an error saving a copy of this document as PDF. " & _    
	"Ensure that the PDF is not open for viewing and that the destination path is writable. Error code: " & Err
	
	End Sub
