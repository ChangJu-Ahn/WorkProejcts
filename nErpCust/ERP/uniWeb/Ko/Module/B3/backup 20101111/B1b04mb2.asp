<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : B1b04mb2
'*  4. Program Name         : HS부호등록 
'*  5. Program Desc         : HS부호등록 
'*  6. Component List       : PB1GB41.cBMaintHsCodeS
'*  7. Modified date(First) : 2000/03/27	
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Sun Jung	
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/incSvrMain.asp" -->
<%
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Dim lgOpModeCRUD
	Dim iPB1GB41																	'☆ : 입력/수정용 ComPlus Dll 사용 변수 
	Dim istrData
	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd																 '☜: Hide Processing message

	lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
	Select Case lgOpModeCRUD	         
	    Case CStr(UID_M0002), CStr(UID_M0005), Cstr(UID_M0005)                       '☜: Save,Update
	         Call SubBizSaveMulti()

	End Select
	
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim E2_ief_Supplied																		'문제발생시 문제를 일으킨 레코드 숫자를 반환한다.
    Dim i_user_id
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim iDCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")
    
	Set iPB1GB41 = Server.CreateObject("PB1GB41.cBMaintHsCodeS")
	
	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End If
	
	Call iPB1GB41.B_MAINT_HS_CODE_SVR(gStrGlobalCollection,i_user_id,itxtSpread,E2_ief_Supplied)

	If CheckSYSTEMError2(Err,True,E2_ief_Supplied & "행:" ,"","","","") = true then 		
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
		Set iPB1GB41 = Nothing															'☜: ComProxy Unload
		Exit Sub
	End If
	
    Set iPB1GB41 = Nothing    
	
	Response.Write "<Script Language=vbscript>"												& vbCr
	Response.Write "With Parent "															& vbCr
	Response.Write " .DBSaveOK "           & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr 
End Sub    

	
%>

