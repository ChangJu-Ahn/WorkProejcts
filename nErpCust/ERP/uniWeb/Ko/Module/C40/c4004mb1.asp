<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Cost Accounting
'*  2. Function Name        : Distribution Factor
'*  3. Program ID           :
'*  4. Program Name         : 오더배부규칙 정보 등록 
'*  5. Program Desc         : 오더배부규칙 정보 등록  
'*  6. Modified date(First) : 2005/09/14
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call LoadBasisGlobalInf()
	call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
	 
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow
    Dim lgStrPrevKey,iStrPrevKey
    	
	On Error Resume Next								'☜: 
	Err.Clear

	Call HideStatusWnd

'---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)


'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD

        Case CStr(UID_M0001)                                                         '☜: Query     
			'Call SubBizQueryMulti()		
        Case CStr(UID_M0002)                                                  '☜: Save,Update             
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim iPC4G004Q
	
	Set iPC4G004Q = Server.CreateObject("PC4G004.cMfCDstbRuleDel")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If    
	
	Call iPC4G004Q.C_MFC_DSTB_DEL(gStrGloBalCollection, Trim(Request("txtVER_CD")))	
    
    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus  
       Set iPC4G004Q = Nothing
       Exit Sub       
    End If    
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbDeleteOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
    
    Set iPC4G004Q = Nothing

	    
End Sub


 	
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
  
    Dim PC4G004Data
    Dim sSQLI1, strCon
    Dim iErrorPosition 
    
    sSQLI1    = Trim(Request("txtSpread"))
    strCon = ucase(trim(request("txtVER_CD"))	)
   
    Set PC4G004Data = Server.CreateObject("PC4G004.cMFCDstbRuleSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    
    
    Call PC4G004Data.C_MFC_DSTB_RULE_SVR(gStrGlobalCollection, strCon, sSQLI1, iErrorPosition)			
	
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
		Set PC4G004Data= Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",2)" & vbCrLF
		Response.Write "</Script>" & vbCrLF	
		Response.End
	End If

	Set PC4G004Data = Nothing	

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End  

   
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
    Response.end
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

