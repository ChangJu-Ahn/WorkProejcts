<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>

<% Server.ScriptTimeout=9600 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

'**********************************************************************************************
'*  1. Module Name          : Product
'*  2. Function Name        : 
'*  3. Program ID           : p4611mb5_ko441
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PP4G214_KO226
'*  7. Modified date(First) : 2010/08/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Jin Ha
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :
'**********************************************************************************************

On Error Resume Next
Err.Clear
	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)

Call SubBizSaveMulti()

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
 Sub subBizSaveMulti()															'☜: 저장 요청을 받음 
 
    Dim oPBAT101
    Dim iErrorPosition

    On Error Resume Next 		
    Err.Clear																		'☜: Protect system from crashing

    Set oPBAT101 = Server.CreateObject("PXI3G122_KO441.cUploadErpProdRslt")    
    
    If CheckSYSTEMError(Err,True) = True Then
	Exit Sub
    End If
	
    Call oPBAT101.UPLOAD_ERP_PROD_RSLT_MAIN(gStrGlobalCollection, "") 


    Select Case Trim(Cstr(Err.Description))
	Case "B_MESSAGE" & Chr(11) & "990000"
		Call DisplayMsgBox("990000", VbOKOnly, "", "", I_MKSCRIPT)

	Case "B_MESSAGE" & Chr(11) & "800161"
		Call DisplayMsgBox("800161", VbOKOnly, "", "", I_MKSCRIPT)

	Case Else
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행 : ", "", "", "", "") = True Then
			Set oPBAT101 = Nothing

			Response.Write "<Script language=vbs> " & vbCr 
			Response.Write " parent.DBSaveError2 "   & vbCr	
			Response.Write "</Script> "  
	
			Exit Sub
		End If
    End Select

    
    Set oPBAT101 = Nothing                                                   '☜: Unload Comproxy  
    
    Response.Write "<Script language=vbs> " & vbCr 
    Response.Write " parent.DbSaveOk2 "      & vbCr						
    Response.Write "</Script> "    
    
End Sub	

 Sub subBizSaveMulti_PXI3G136_KO441()															'☜: 저장 요청을 받음 
 
    Dim oPBAT101
    Dim iErrorPosition

    On Error Resume Next 		
    Err.Clear																		'☜: Protect system from crashing

    Set oPBAT101 = Server.CreateObject("PXI3G136_KO441.cUploadErpProdRslt")    
    
    If CheckSYSTEMError(Err,True) = True Then
	Exit Sub
    End If
	
    Call oPBAT101.UPLOAD_ERP_PROD_RSLT_MAIN(gStrGlobalCollection, "") 


    Select Case Trim(Cstr(Err.Description))
	Case "B_MESSAGE" & Chr(11) & "990000"
		Call DisplayMsgBox("990000", VbOKOnly, "", "", I_MKSCRIPT)

	Case "B_MESSAGE" & Chr(11) & "800161"
		Call DisplayMsgBox("800161", VbOKOnly, "", "", I_MKSCRIPT)

	Case Else
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행 : ", "", "", "", "") = True Then
			Set oPBAT101 = Nothing

			Response.Write "<Script language=vbs> " & vbCr 
			Response.Write " parent.DBSaveError2 "   & vbCr	
			Response.Write "</Script> "  
	
			Exit Sub
		End If
    End Select

    
    Set oPBAT101 = Nothing                                                   '☜: Unload Comproxy  
    
    Response.Write "<Script language=vbs> " & vbCr 
    Response.Write " parent.DbSaveOk2 "      & vbCr						
    Response.Write "</Script> "    
    
End Sub	
'============================================================================================================
' Name : RemovedivTextArea
' Desc : 
'============================================================================================================
Sub RemovedivTextArea()
    On Error Resume Next                                                             
    Err.Clear                                                                        
	
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
End Sub
%>
