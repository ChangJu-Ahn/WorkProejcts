<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf()
 '*******************************************************************************************************
 '*  1. Module Name          : Procuremant																*
 '*  2. Function Name        :																			*
 '*  3. Program ID           : M5212mb2.asp																*
 '*  4. Program Name         :																			*
 '*  5. Program Desc         : 수입B/L 내역등록 Save Transaction 처리용 ASP								*
 '*  7. Modified date(First) : 2000/03/27																*
 '*  8. Modified date(Last)  : 2003/06/13																			*
 '*  9. Modifier (First)     : Sun-jung Lee																*
 '* 10. Modifier (Last)      : Jin-hyun Shin															*
 '* 11. Comment              :																			*
 '* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
 '*                            this mark(⊙) Means that "may  change"									*
 '*                            this mark(☆) Means that "must change"									*
 '* 13. History              : 1. 2000/03/27 : Coding Start												*
 '*******************************************************************************************************

 Dim lgOpModeCRUD

 On Error Resume Next
 Err.Clear
    
 Call HideStatusWnd
 lgOpModeCRUD = Request("txtMode")
    

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
  
	Dim iM52121
	Dim iErrorPosition
	Dim str_txtblno
	
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
		
	If Request("txtBLNo") = "" Then           '⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Exit Sub
	End If

	Set iM52121 = Server.CreateObject("PM5G121.cMMaintImportBlDtlS")

	If CheckSYSTEMError(Err,True) = True Then
		Set iM52121 = Nothing
		Exit Sub
	End If
	
	str_txtblno   =  Trim(Request("txtBLNo"))
  
	Call iM52121.M_MAINT_IMPORT_BL_DTL_SVR(gStrGlobalCollection, str_txtblno, itxtSpread,iErrorPosition)
 
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행:","","","","") = True Then
		Set iM52121 = Nothing	
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr
		Response.Write "</Script> "
		Exit Sub
	End If
	
	Set iM52121 = Nothing 

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent"   & vbCr               '☜: 화면 처리 ASP 를 지칭함 
	Response.Write "	.DbSaveOk"    & vbCr 
	Response.Write "End With"    & vbCr
	Response.Write "</Script>"     & vbCr
   

End Sub


%>
