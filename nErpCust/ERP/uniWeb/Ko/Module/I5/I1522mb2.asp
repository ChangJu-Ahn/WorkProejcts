<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 검사결과등록 
'*  3. Program ID           : I1522MB2
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
    
    On Error Resume Next                                                         
    Err.Clear                                                                    
    
    Call HideStatusWnd                                                           

	Dim iPI5G130

	Dim txtSpread
	Dim iErrorPosition

	txtSpread = Request("txtSpread")

    Set iPI5G130 = Server.CreateObject("PI5G130.cIVMIInspectionResult")
    
    If CheckSYSTEMError(Err,True) = True Then
    	Response.End	
    End If

	Call iPI5G130.I_VMI_INSPECTION_RESULT(gStrGlobalCollection, _
										txtSpread, _
										iErrorPosition)  

    If CheckSYSTEMError(Err,True) = True Then
		Set iPI5G130 = Nothing
		If iErrorPosition <> 0 Then
			Call SheetFocus(iErrorPosition, 1)
		End If
    	Response.End	
    End If
   
    Set iPI5G130 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCrlf
	Response.Write " Parent.DbSaveOk " & vbCrlf
	Response.Write " </Script>" & vbCrlf
	Response.End

'===============================================================    
Sub SheetFocus(ByVal lRow, ByVal lCol)
	Response.Write " <Script Language=VBScript> "                    & vbCrLF
	Response.Write " With parent.frm1 "                              & vbCrlf
	Response.Write "	.vspdData.focus "                           & vbCrlf
	Response.Write "	.vspdData.Row = " & lRow                    & vbCrlf
	Response.Write "	.vspdData.Col = " & lCol                    & vbCrlf
	Response.Write "	.vspdData.Action = 0 "                      & vbCrlf
	Response.Write "	.vspdData.SelStart = 0 "                    & vbCrlf
	Response.Write "	.vspdData.SelLength = len(.vspdData.Text) " & vbCrlf
	Response.Write " End With"                                       & vbCrlf
	Response.Write " </Script>"                                      & vbCrLF
End Sub
%>


