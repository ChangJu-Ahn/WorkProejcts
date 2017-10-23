<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%	
call LoadBasisGlobalInf()

	Dim lgOpModeCRUD
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")

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
																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim iPM4G221																' Open L/C Detail Save용 Object
	Dim CommandSent
	Dim  I2_m_lc_amend_hdr
    Const M422_I2_lc_amd_no = 0    '  View Name : import m_lc_amend_hdr
    Const M422_I2_lc_kind = 1
    Const M422_E2_count = 0    '  View Name : export_group_count ief_supplied

	Dim I1_s_wks_user_user_id   
	Dim prErrorPosition
    
    '-------------------
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
    
    Dim strPreMessage 'Prefix message string that includes row info. '200310
             
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

	Set iPM4G221 = Server.CreateObject("PM4G221.cMMaintLcAmendDtlS")
		
	ReDim I2_m_lc_amend_hdr(1)
	I2_m_lc_amend_hdr(M422_I2_lc_amd_no) =  Request("txtLCAmdNo")
	I2_m_lc_amend_hdr(M422_I2_lc_kind) = "M"

	If CheckSYSTEMError(Err,True) = True Then
		Set iPM4G221 = Nothing
		Exit Sub
	End If
		
	CommandSent = "save"
		
    Call iPM4G221.MAINT_LC_AMEND_DTL_SVR(gStrGlobalCollection, CommandSent, I1_s_wks_user_user_id, I2_m_lc_amend_hdr, itxtSpread, prErrorPosition)
	
	'200310 행번호가 리턴되어야 행정보를 메세지에 추가한다.
	If Trim(prErrorPosition) = "" then 		
		strPreMessage = ""
	Else
		strPreMessage =  prErrorPosition & "행:"
	End if	

	If CheckSYSTEMError2(Err,True, strPreMessage ,"","","","") = True then 		
		Set iPM4G221 = Nothing
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "
		Exit Sub
	End If
       
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write " .frm1.txtLCAmdNo.value = """ & ConvSPChars(I2_m_lc_amend_hdr(M422_I2_lc_amd_no)) & """" & vbCr
	Response.Write " .DbSaveOk" & vbCr
	Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr

	Set iPM4G221 = Nothing			
 
End Sub


%>
