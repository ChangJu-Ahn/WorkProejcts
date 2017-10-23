<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4212mb2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관 내역등록 Save Transaction 처리용 ASP								*
'*  7. Modified date(First) : 2000/06/21																*
'*  8. Modified date(Last)  : 2002/07/2			     													*
'*  9. Modifier (First)     : Kim Jin Ha        														*
'* 10. Modifier (Last)      : Kim Jin Ha			    												*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/27 : Coding Start												*
'*			                : 2. 2002/06/21 : Com Plus Conv.
'********************************************************************************************************
	'Dim lgOpModeCRUD

	On Error Resume Next																	'☜: Protect system from crashing
	Err.Clear 																				'☜: Clear Error status
	
	Call LoadBasisGlobalInf()			
	Call HideStatusWnd
	
	Call SubBizSaveMulti()
	

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSaveMulti()
	On Error Resume Next																	'☜: Protect system from crashing
	Err.Clear

	Dim OBJ_PM6G121
	Dim iErrorPosition 												
	Dim iStrCcNo
	
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

	Set OBJ_PM6G121 = Server.CreateObject("PM6G121.cMMaintImportCcDtlS")

    If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G121 = Nothing
		Exit Sub
	End If
   
    iStrCcNo = Trim(Request("txtCCNo"))
    	
	Call OBJ_PM6G121.M_MAINT_IMPORT_CC_DTL_SVR(gStrglobalcollection, iStrCcNo, itxtSpread, CLng(iErrorPosition))
												      
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행:","","","","") = True Then
		Set OBJ_PM6G121 = Nothing	
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr
		Response.Write "</Script> "
		Exit Sub
	End If
		
	Set OBJ_PM6G121 = Nothing																'☜: Unload Comproxy

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "parent.DbSaveOk	" & vbCr
	Response.Write "</Script>" & vbCr			


End Sub
%>
