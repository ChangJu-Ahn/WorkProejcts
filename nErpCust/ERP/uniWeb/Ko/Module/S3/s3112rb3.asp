<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
    Dim lgOpModeCRUD
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
            Call SubBizQuery()      
        Case "COMPLETE"														         '☜: 현재 조회/Prev/Next 요청을 받음 
       		Call SubBizCOMPLETE()
       	Case "PARTIAL"														         '☜: 현재 조회/Prev/Next 요청을 받음 
			Call SubBizPARTIAL()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iPS3G120
    Dim iLngMaxRow
    Dim iStrData
    Dim I2_s_so_dtl
    Dim I3_s_so_hdr 
    Dim E1_b_item 
    Dim E2_b_plant 
    Dim E3_s_so_dtl 
    Dim E4_s_so_schd 
    Dim EG1_exp_grp

	Const S337_E1_item_cd      = 0    '[CONVERSION INFORMATION]  View Name : exp b_item
	Const S337_E1_item_nm      = 1
	
	Const S337_E2_plant_cd     = 0    '[CONVERSION INFORMATION]  View Name : exp b_plant
	Const S337_E2_plant_nm     = 1

	Const S337_E3_tracking_no  = 0    '[CONVERSION INFORMATION]  View Name : exp s_so_dtl
	Const S337_E3_dlvy_dt      = 1
	Const S337_E3_so_qty       = 2
	Const S337_E3_bonus_qty    = 3
	Const S337_E3_so_unit      = 4

	Const S337_E4_dlvy_dt              = 0    '[CONVERSION INFORMATION]  View Name : exp s_so_schd
	Const S337_E4_promise_dt           = 1
	Const S337_E4_cfm_qty              = 2
	Const S337_E4_cfm_bonus_qty        = 3
	Const S337_E4_cfm_base_qty         = 4
	Const S337_E4_cfm_bns_base_qty     = 5
	Const S337_E4_atp_flag             = 6

	Const S337_EG1_E1_dlvy_dt          = 0    '[CONVERSION INFORMATION]  View Name : exp_item s_so_schd
	Const S337_EG1_E1_promise_dt       = 1
	Const S337_EG1_E1_cfm_qty          = 2
	Const S337_EG1_E1_cfm_bonus_qty    = 3
	Const S337_EG1_E1_cfm_base_qty     = 4
	Const S337_EG1_E1_cfm_bns_base_qty = 5
	Const S337_EG1_E1_atp_flag = 6
   

    On Error Resume Next                                                              
    Err.Clear         
   	I3_s_so_hdr = Trim(Request("txtSONo"))
	I2_s_so_dtl = Trim(Request("txtSOSeq"))
	                                                                '☜: Clear Error status

    Set iPS3G120 = Server.CreateObject("PS3G120.cSCheckAtpSvr")    

	If CheckSYSTEMError(Err,True) = True Then

       Exit Sub
    End If  
                                                                 '☜: Clear Error status
    Call iPS3G120.S_CHECK_ATP_SVR(gStrGlobalCollection,"", I2_s_so_dtl ,I3_s_so_hdr , _
                            E1_b_item , E2_b_plant , E3_s_so_dtl , E4_s_so_schd , EG1_exp_grp)
   

	If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G120 = Nothing
       Exit Sub
    End If  

    Set iPS3G120 = Nothing
    
       
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write "With  parent.frm1	           " & vbCr
	Response.Write ".txtItem.value        = """ & ConvSPChars(E1_b_item(S337_E1_item_cd))       & """" & vbCr
	Response.Write ".txtItemNm.value      = """ & ConvSPChars(E1_b_item(S337_E1_item_nm))       & """" & vbCr
	Response.Write ".txtPlant.value       = """ & ConvSPChars(E2_b_plant(S337_E2_plant_cd))     & """" & vbCr
	Response.Write ".txtPlantNm.value     = """ & ConvSPChars(E2_b_plant(S337_E2_plant_nm))     & """" & vbCr
	Response.Write ".txtTrackingNo.value  = """ & ConvSPChars(E3_s_so_dtl(S337_E3_tracking_no)) & """" & vbCr
	Response.Write ".txtSOUnit.value      = """ & ConvSPChars(E3_s_so_dtl(S337_E3_so_unit))     & """" & vbCr

	Response.Write ".txtSchdDlvyDt.text   = """ & UNIDateClientFormat(E3_s_so_dtl(S337_E3_dlvy_dt))  & """" & vbCr

	Response.Write ".txtSOQty.text        = """ & UNINumClientFormat(E3_s_so_dtl(S337_E3_so_qty), ggQty.DecPoint, 0)       & """" & vbCr
	Response.Write ".txtBonusQty.text     = """ & UNINumClientFormat(E3_s_so_dtl(S337_E3_bonus_qty), ggQty.DecPoint, 0)    & """" & vbCr
	Response.Write ".txtHBaseQty.value    = """ & UNINumClientFormat(E4_s_so_schd(S337_E4_cfm_base_qty), ggQty.DecPoint, 0)       & """" & vbCr
	Response.Write ".txtHBonusBaseQty.value = """ & UNINumClientFormat(E4_s_so_schd(S337_E4_cfm_bns_base_qty), ggQty.DecPoint, 0)  & """" & vbCr
	Response.Write ".txtHAtpFlag.value    = """ & E4_s_so_schd(S337_E4_atp_flag)  & """" & vbCr
	If E4_s_so_schd(S337_E4_atp_flag) = "L" Then
		Response.Write ".chkPPFlg.checked = True  " & vbCr
	Else
		Response.Write ".chkPPFlg.checked = False " & vbCr
	End If

	Response.Write ".txtAvalSchdDlvyDt.text = """ & UNIDateClientFormat(E4_s_so_schd(S337_E4_dlvy_dt))     & """" & vbCr

	Response.Write ".txtAvalGIDt.text       = """ & UNIDateClientFormat(E4_s_so_schd(S337_E4_promise_dt))  & """" & vbCr	
    Response.Write "End With           " & vbCr															    	
    Response.Write "</Script>          " & vbCr 
    Response.Write "<Script language=vbs>  " & vbCr      
	Response.Write "With parent.frm1 " & vbCr  
    iLngMaxRow  = CLng(Request("txtMaxRows"))											'Save previous Maxrow
    iStrData = ""
    
    Dim iLngRow
     
	If UBound(EG1_exp_grp,2) > 0 Then
		For iLngRow = 0 To UBound(EG1_exp_grp,2)

			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(S337_EG1_E1_dlvy_dt,iLngRow))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(S337_EG1_E1_promise_dt,iLngRow))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_qty,iLngRow), ggQty.DecPoint, 0)										'3
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_bonus_qty,iLngRow), ggQty.DecPoint, 0)										'3
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_base_qty,iLngRow), ggQty.DecPoint, 0)										'3
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_bns_base_qty,iLngRow), ggQty.DecPoint, 0)										'3
			If EG1_exp_grp(S337_EG1_E1_atp_flag,iLngRow) = "F" then
				iStrData = iStrData & Chr(11) & "N"
			Else
				iStrData = iStrData & Chr(11) & "Y"
			End If
        
			iStrData = iStrData & Chr(11) & EG1_exp_grp(S337_EG1_E1_atp_flag,iLngRow)

			iStrData = iStrData & Chr(11) & Chr(12)

		Next

	Else
		
		Response.Write ".txtHBaseQty.value = """ & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_base_qty,0), ggQty.DecPoint, 0) & """" & vbCr
		Response.Write ".txtHBonusBaseQty.value = """ & UNINumClientFormat(EG1_exp_grp(S337_EG1_E1_cfm_bns_base_qty,0), ggQty.DecPoint, 0) & """" & vbCr
		Response.Write ".txtHAtpFlag.value = """ & EG1_exp_grp(S337_EG1_E1_atp_flag,0) & """" & vbCr
		If Trim(EG1_exp_grp(S337_EG1_E1_atp_flag,0)) = "L" Then
			Response.Write ".chkPPFlg.checked = True  " & vbCr
		Else
			Response.Write ".chkPPFlg.checked = False " & vbCr
		End If

		Response.Write ".txtAvalSchdDlvyDt.text = """ & UNIDateClientFormat(EG1_exp_grp(S337_EG1_E1_dlvy_dt,0)) & """" & vbCr

		Response.Write ".txtAvalGIDt.text = """ & UNIDateClientFormat(EG1_exp_grp(S337_EG1_E1_promise_dt,0)) & """" & vbCr
	

	End If

	Response.Write "parent.ggoSpread.Source = .vspdData    " & vbCr 
    Response.Write "parent.ggoSpread.SSShowDataByClip        """ & istrData	       & """" & vbCr 
    Response.Write "parent.lgStrPrevKey              = """ & iStrNextKey	   & """" & vbCr  
	Response.Write "parent.DbqueryOk()                " & vbCr  
	Response.Write ".vspdData.focus                   " & vbCr 
	Response.Write ".vspdData.Row = 1                 " & vbCr
	Response.Write ".vspdData.SelModeSelected = True  " & vbCr
	Response.Write "End With  " & vbCr 
    Response.Write "</Script> " & vbCr 


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    

Sub SubBizCOMPLETE()
	Dim iPS3G131
    Dim IG1_imp_grp 
    Dim I1_s_so_dtl 
    Dim I2_s_so_hdr 
    Dim iErrorPosition
    
	Const S339_IG1_I1_row_num          = 0    '[CONVERSION INFORMATION]  View Name : imp_item s_wks_client
	Const S339_IG1_I2_dlvy_dt          = 1    '[CONVERSION INFORMATION]  View Name : imp_item s_so_schd
	Const S339_IG1_I2_promise_dt       = 2
	Const S339_IG1_I2_cfm_qty          = 3
	Const S339_IG1_I2_cfm_bonus_qty    = 4
	Const S339_IG1_I2_cfm_base_qty     = 5
	Const S339_IG1_I2_cfm_bns_base_qty = 6
	Const S339_IG1_I2_atp_flag         = 7
			
    On Error Resume Next                                                              
    Err.Clear                                                                         '☜: Clear Error status

	I2_s_so_hdr = Trim(Request("txtSONo"))
	
	I1_s_so_dtl = UNICDbl(Request("txtSOSeq"), 0)

	Redim  IG1_imp_grp(0,S339_IG1_I2_atp_flag)

	IG1_imp_grp(0,S339_IG1_I2_dlvy_dt)       = UNIConvDate(Request("txtSchdDlvyDt"))
	IG1_imp_grp(0,S339_IG1_I2_promise_dt)    = UNIConvDate(Request("txtPromiseGIDt"))
	IG1_imp_grp(0,S339_IG1_I2_cfm_qty)       = UNIConvNum(Request("C_ComfirmSOQty"),0)
	IG1_imp_grp(0,S339_IG1_I2_cfm_bonus_qty) = UNIConvNum(Request("C_ComfirmBonusQty"),0)
	IG1_imp_grp(0,S339_IG1_I2_cfm_base_qty)  = UNIConvNum(Request("C_ComfirmBaseQty"),0)
	IG1_imp_grp(0,S339_IG1_I2_cfm_bns_base_qty) = UNIConvNum(Request("C_ComfirmBaseBonusQty"),0)
	IG1_imp_grp(0,S339_IG1_I2_atp_flag)      = Trim(Request("C_ATPFlg"))
    IG1_imp_grp(0,S339_IG1_I1_row_num)       = 1

    Set iPS3G131 = Server.CreateObject("PS3G131.cSRegenSchdByAtpSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  

    iErrorPosition = iPS3G131.S_REGEN_SO_SCHD_BY_ATP_SVR(gStrGlobalCollection ,IG1_imp_grp , I1_s_so_dtl ,  I2_s_so_hdr ,  "")
              
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS3G131 = Nothing
       Exit Sub
	End If
    
    Set iPS3G131 = Nothing
    Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write "With parent		            " & vbCr																	
	Response.Write ".ATPOk                      " & vbCr
	Response.Write "End With                    " & vbCr
    Response.Write "</Script>                   " & vbCr
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub  
Sub SubBizPARTIAL()

	Dim arrVal
	Dim arrTemp	
	Dim LngRow
    Dim IG1_imp_grp 
    Dim iPS3G131
    Dim I1_s_so_dtl 
    Dim I2_s_so_hdr 
    Dim LngMaxRow
    Dim iErrorPosition
	Const S339_IG1_I1_row_num          = 0    '[CONVERSION INFORMATION]  View Name : imp_item s_wks_client
	Const S339_IG1_I2_dlvy_dt          = 1    '[CONVERSION INFORMATION]  View Name : imp_item s_so_schd
	Const S339_IG1_I2_promise_dt       = 2
	Const S339_IG1_I2_cfm_qty          = 3
	Const S339_IG1_I2_cfm_bonus_qty    = 4
	Const S339_IG1_I2_cfm_base_qty     = 5
	Const S339_IG1_I2_cfm_bns_base_qty = 6
	Const S339_IG1_I2_atp_flag         = 7

    On Error Resume Next                                                              
    Err.Clear       
    LngMaxRow = CInt(Request("txtMaxRows"))	                                                                  '☜: Clear Error status
	I2_s_so_hdr = Trim(Request("txtSONo"))
	I1_s_so_dtl = UNICDbl(Request("txtSOSeq"), 0)
	
	arrTemp = Split(Request("txtspread"), gRowSep)
		
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
    Redim IG1_imp_grp(LngMaxRow - 1, S339_IG1_I2_atp_flag) 
    
	For LngRow = 1 To LngMaxRow
		arrVal = Split(arrTemp(LngRow-1), gColSep)
		IG1_imp_grp(LngRow - 1,S339_IG1_I1_row_num)		    	= UNIConvNum(arrVal(0),0)	
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_dlvy_dt)			    = UNIConvDate(arrVal(1))
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_promise_dt)			= UNIConvDate(arrVal(2))
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_cfm_qty)			    = UNIConvNum(arrVal(3),0)	
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_cfm_bonus_qty)		= UNIConvNum(arrVal(4),0)
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_cfm_base_qty)		= UNIConvNum(arrVal(5),0) 	
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_cfm_bns_base_qty)	= UNIConvNum(arrVal(6),0)	
		IG1_imp_grp(LngRow - 1,S339_IG1_I2_atp_flag)			= Trim(arrVal(7))
	
	Next

    Set iPS3G131 = Server.CreateObject("PS3G131.cSRegenSchdByAtpSvr")     

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  

    iErrorPosition = iPS3G131.S_REGEN_SO_SCHD_BY_ATP_SVR( gStrGlobalCollection ,IG1_imp_grp , I1_s_so_dtl , _
                   I2_s_so_hdr ,  "")
                    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS3G131 = Nothing
       Exit Sub
	End If 
    
    Set iPS3G131 = Nothing
    
    Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write "With parent		            " & vbCr																	
	Response.Write ".ATPOk                      " & vbCr
	Response.Write "End With                    " & vbCr
    Response.Write "</Script>                   " & vbCr
    
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub  
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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