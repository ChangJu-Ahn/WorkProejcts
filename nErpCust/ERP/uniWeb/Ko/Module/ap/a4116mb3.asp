
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account 
'*  2. Function Name        : 
'*  3. Program ID           : a4116mb3.adp
'*  4. Program Name         : (-)채권/출금반제 삭제 Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : YOU SO EUN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
														'☜ : ASP가 캐쉬되지 않도록 한다.
														'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

On Error Resume Next															'☜: Protect system from crashing
Err.Clear																		'☜: Clear Error status

Call HideStatusWnd																'☜: Hide Processing message
Call LoadBasisGlobalInf()														

Call SubBizDelete()																'☜: Single --> Delete

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim iPAPG080
    Dim iCommandSent
    Dim I5_a_allc_paym
    
    Const A356_I5_paym_no = 0 
    
    redim I5_a_allc_paym(18)

	Dim I9_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A356_I9_a_data_auth_data_BizAreaCd = 0
	Const A356_I9_a_data_auth_data_internal_cd = 1
	Const A356_I9_a_data_auth_data_sub_internal_cd = 2
	Const A356_I9_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A356_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

    iCommandSent = "DELETE"
    I5_a_allc_paym(A356_I5_paym_no) = Trim(Request("txtAllcNo"))
    
    Set iPAPG080 = Server.CreateObject ("PAPG080.cAMntAllcPayByArSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	
	Call iPAPG080.A_MAINT_ALLC_PAYM_BY_AR_SVR(gStrGlobalCollection,iCommandSent,,,,,,I5_a_allc_paym,,,,,,I9_a_data_auth)
 
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG080 = Nothing
		Exit Sub
	End If

    Set iPAPG080 = Nothing
    
    Response.Write "<Script Language=VBScript>	" & vbCr         
    Response.Write " parent.DbDeleteOk()		" & vbCr  
    Response.Write "</Script>					" & vbCr
        
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   		                                                                    
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------     
End Sub    

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>
