<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%	
	call LoadBasisGlobalInf()

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212mb2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C 내역등록 Save Transaction 처리용 ASP								*
'*  7. Modified date(First) : 2000/03/27																*
'*  8. Modified date(Last)  : 2000/03/27																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/27 : Coding Start												*
'********************************************************************************************************

	Dim lgOpModeCRUD
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD = Request("txtMode")		

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0002), CStr(UID_M0005), Cstr(UID_M0005)                       '☜: Save,Update
	             Call SubBizSaveMulti()

	End Select

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next					'☜: Protect system from crashing
	Err.Clear 
		
	Dim iPM4G121
	Dim iErrorPosition
	Dim str_LCNO
	
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
    
	' Master L/C Detail Save용 Object
	Set iPM4G121 = Server.CreateObject("PM4G121.cMMaintLcDtlS")

	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM4G121 = Nothing
		Exit Sub
	End If
								
	str_LCNO = Trim(Request("txtLCNo"))	
	
	Call iPM4G121.M_MAINT_LC_DTL_SVR(gStrGlobalCollection, str_LCNO, , itxtSpread, , , iErrorPosition)
		
	If CheckSYSTEMError2(Err,True, iErrorPosition & "행:" ,"","","","") = True then 		
		Set iPM4G121 = Nothing
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "
		Exit Sub
	End If

	Set iPM4G121 = Nothing
		
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent"				& vbCr						'☜: 화면 처리 ASP 를 지칭함 
	Response.Write ".DbSaveOk"					& vbCr
	Response.Write "End With"					& vbCr
	Response.Write "</Script>"					& vbCr

End Sub
%>