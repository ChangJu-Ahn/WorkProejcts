

<%@ LANGUAGE=VBSCript%>

<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp" -->

<%Call LoadBasisGlobalInf()%>

<%
Dim lgOpModeCRUD
        
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status
	    		
Call HideStatusWnd                                                               '☜: Hide Processing message
	
lgOpModeCRUD  = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call  SubBizQueryMulti()
    Case CStr(UID_M0002)
         Call SubBizSaveMulti()
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim TmpBuffer
    Dim iMax
    Dim iIntLoopCount
    Dim iTotalStr
    
    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrPrevKey
    Dim istrData
    Dim istrTemp
    
    Dim iStrNextKey  	
    Dim iarrValue
    Const C_SHEETMAXROWS_D  = 100
    
    Dim iPM1G418
	Dim imp_next_m_config_process
	Dim exp_cond_m_config_process
	

	Dim imp_m_config_process
	Const C_i_m_config_process_po_type_cd = 0
	Const C_i_m_config_process_usage_flg = 1

    Dim exp_group
    Const IG1_I2_m_config_process_po_type_cd = 0
    Const IG1_I2_m_config_process_po_type_nm = 1
    '법인간거래여부 추가 
    Const IG1_I2_m_config_process_intercom_flg = 2
    
    Const IG1_I2_m_config_process_sto_flg = 3			' added for STO
    Const IG1_I2_m_config_process_import_flg = 4
    Const IG1_I2_m_config_process_bl_flg = 5
    Const IG1_I2_m_config_process_cc_flg = 6
    Const IG1_I2_m_config_process_rcpt_flg = 7
    Const IG1_I2_m_config_process_iv_flg = 8
    Const IG1_I2_m_config_process_ret_flg = 9
    Const IG1_I2_m_config_process_subcontra_flg = 10
    Const IG1_I2_m_config_process_rcpt_type = 11
    Const IG1_I2_m_config_process_rcpt_type_nm = 12
    Const IG1_I2_m_config_process_issue_type = 13
    Const IG1_I2_m_config_process_issue_type_nm = 14
    Const IG1_I2_m_config_process_iv_type = 15
    Const IG1_I2_m_config_process_iv_type_nm = 16
    Const IG1_I2_m_config_process_so_type = 17		' added for STO
    Const IG1_I2_m_config_process_so_type_nm = 18	' added for STO
    Const IG1_I2_m_config_process_usage_flg = 19

	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
  	
  	If iStrPrevKey <> "" then					
		iarrValue = Split(iStrPrevKey, gColSep)
		imp_next_m_config_process = Trim(iarrValue(0))										
	else			
		imp_next_m_config_process = ""
	End If 
	

    redim imp_m_config_process(C_i_m_config_process_usage_flg)
    
    imp_m_config_process(C_i_m_config_process_po_type_cd)= Trim(UCase(Request("txtPotypeCd")))
	imp_m_config_process(C_i_m_config_process_usage_flg)= UCase(Request("txtUseflg"))
   
    Set iPM1G418 = Server.CreateObject("PM1G418.cMLstConfigProcessS")    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM1G418 = Nothing
        Exit Sub
	End if

    call iPM1G418.M_LIST_CONFIG_PROCESS_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, imp_m_config_process, _
     imp_next_m_config_process, exp_group, exp_cond_m_config_process)
  
    
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPM1G418 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if
	
	Set iPM1G418 = Nothing												'☜: ComProxy Unload
	
	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write " Parent.frm1.txtPotypeNm.value   = """ & ConvSPChars(exp_cond_m_config_process(0)) & """" & vbCr
    Response.Write "</Script> "		
	
	iLngMaxRow = CLng(Request("txtMaxRows"))
	
	iIntLoopCount = 0
	iMax = UBound(exp_group,2)
	ReDim TmpBuffer(iMax)
	
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(exp_group(IG1_I2_m_config_process_po_type_cd,iLngRow)) 
           Exit For
        End If  	
		
		istrData = ""
		istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_po_type_cd, iLngRow))
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_po_type_nm, iLngRow))
        '법인간거래여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_intercom_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if
        'STO여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_sto_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if
        '수입여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_import_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if
        '선적여부 
       	if ConvSPChars(exp_group(IG1_I2_m_config_process_bl_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        end if
        '통관여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_cc_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if
        '입고여부 
       	if ConvSPChars(exp_group(IG1_I2_m_config_process_rcpt_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        end if
        '매입여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_iv_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if
        '반품여부 
       	if ConvSPChars(exp_group(IG1_I2_m_config_process_ret_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        end if
        '사급여부 
        if ConvSPChars(exp_group(IG1_I2_m_config_process_subcontra_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        End if

        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_rcpt_type, iLngRow))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_rcpt_type_nm, iLngRow))
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_issue_type, iLngRow))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_issue_type_nm, iLngRow))
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_iv_type, iLngRow))
        istrData = istrData & Chr(11) & " "      'PopUp
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_iv_type_nm, iLngRow))
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_so_type, iLngRow)) 'added for STO
        istrData = istrData & Chr(11) & " "      'PopUp  'added for STO
        istrData = istrData & Chr(11) & ConvSPChars(exp_group(IG1_I2_m_config_process_so_type_nm, iLngRow))  'added for STO
        
        '사용여부 
       	if ConvSPChars(exp_group(IG1_I2_m_config_process_usage_flg, iLngRow)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        end if
        
		istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next	
    
    iTotalStr = Join(TmpBuffer, "")
    
    Response.Write "<Script language=vbs> " & vbCr   
    
    if UCase(Request("txtPotypeCd")) = "" then
		Response.Write " Parent.frm1.txtPotypeNm.value   = """ & " " & """" & vbCr     
    end if
       
    Response.Write " Parent.frm1.hdnUseflg.value = """ & UCase(Request("hdnUseflg")) & """		"	& vbCr    
    Response.Write " Parent.SetSpreadColor -1,-1												"   & vbCr
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData					"	& vbCr
    Response.Write " Parent.lgQuery = True"															& vbCr
    Response.Write " Parent.ggoSpread.SSShowData        """ & iTotalStr		& """				"	& vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey	& """				"	& vbCr  
    Response.Write " Parent.DbQueryOk "					& vbCr  
    Response.Write " Parent.frm1.vspdData.focus "		& vbCr 																		    	& vbCr   
    Response.Write "</Script> "		
    
End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    
    Dim iPM1G411
    Dim iErrorPosition
    
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

    
    Call RemovedivTextArea()
    
	Set iPM1G411 = Server.CreateObject("PM1G411.cMMaintConfigProcS")    

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End If
	
    Call iPM1G411.M_MAINT_CONFIG_PROCESS_SVR(gStrGlobalCollection, itxtSpread, iErrorPosition)

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPM1G411 = Nothing
       Exit Sub
	End If

    Set iPM1G411 = Nothing    
                 
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write "Parent.DBSaveOk "      & vbCr   
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

