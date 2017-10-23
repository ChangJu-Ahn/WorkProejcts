
<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4213mb2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관란 내역등록 Save Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/03/27																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/27 : Coding Start												*
'********************************************************************************************************
	Dim lgOpModeCRUD
	Dim OBJ_PM6G131											'☆ : 매출내역등록입력/수정/삭제용 ComProxy Dll 사용 변수 
	
	On Error Resume Next									'☜: Protect system from crashing
	Err.Clear 												'☜: Clear Error status
				
	Call LoadBasisGlobalInf()	
	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)
	
	Select Case lgOpModeCRUD
	        Case CStr(UID_M0002)
	             Call SubBizSaveMulti()
	End Select

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSaveMulti()
	On Error Resume Next			
    Err.Clear
    
	Dim iErrorPosition
	Dim txtCcNo
	
	'-------------------
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
             
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
    '---------------------    

	Set OBJ_PM6G131 = Server.CreateObject("PM6G131.cMMaintImportCcLanS")

	If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G131 = Nothing
		Exit Sub
	End If
	
	txtCcNo = Trim(Request("txtCCNo"))
	
	Call OBJ_PM6G131.M_MAINT_IMPORT_CC_LAN_SVR(gStrGlobalCollection, , txtCcNo, itxtSpread, iErrorPosition)
	
	If CheckSYSTEMError2(Err,True, iErrorPosition(0) & "행:" ,"","","","") = True then 		
		Set OBJ_PM6G131 = Nothing
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr
		Response.Write "</Script> "
		Exit Sub
	End If

	Set OBJ_PM6G131 = Nothing
	

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr																
	Response.Write ".DbSaveOk  " & vbCr	
	Response.Write "End With   " & vbCr
	Response.Write "</Script>	   " & vbCr		

End Sub


%>