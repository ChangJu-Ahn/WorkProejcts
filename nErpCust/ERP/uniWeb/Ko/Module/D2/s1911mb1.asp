<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s1911mb1.asp																*
'*  4. Program Name         : 수주형태등록																*
'*  5. Program Desc         : 수주형태등록																*
'*  6. Comproxy List        :  																			*
'*  7. Modified date(First) : 2000/08/25																*
'*  8. Modified date(Last)  : 2005/01/24																*
'*  9. Modifier (First)     : Juvenile	 																*
'* 10. Modifier (Last)      : Sim Hae Young																*
'* 11. Comment              : 납기일수,수주관리여부 column추가 
'			    : 수주형태명 뒤에 멀티컴퍼니거래여부 칼럼추가(2005/01/24)										*
'********************************************************************************************************

    Dim lgOpModeCRUD
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    End Select

'============================================================================================================
Sub SubBizQueryMulti()

	Dim iLngRow
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey

    Dim iPD2GS64

    Dim I1_Next_So_Type_Cfg
    Dim I2_So_Type_Cfg
    Dim E1_So_Type_Cfg
    Dim EG1_EXP_GRP
    Dim E2_Next_So_Type_Cfg

    Dim intGroupCount
    Dim iStrNextKey
    Dim iarrValue

    Const C_SHEETMAXROWS_D  = 100

    'import view에 대한 상수 
    Const I2_So_Type_Cfg_So_Type = 0
    Const I2_So_Type_Cfg_Usg_Flg = 1

    'exp view에 대한 상수(exp_s_so_type_config)
    Const E1_So_Type_Cfg_So_Type = 0
    Const E1_So_Type_Cfg_So_Type_Nm = 1

    ' exp_grp 저장 
    Const EG1_EXP_GRP_so_type_cfg_so_type = 0
    Const EG1_EXP_GRP_so_type_cfg_so_type_nm = 1
    '2002-12-04 추가 2003-01-15 위치이동 

    '멀티컴퍼니거래여부 
    Const EG1_EXP_GRP_so_type_cfg_intercom_flg = 2


    Const EG1_EXP_GRP_so_type_cfg_sto_flag = 3

    Const EG1_EXP_GRP_so_type_cfg_rel_dn_flag = 4
    Const EG1_EXP_GRP_so_type_cfg_rel_bill_flag = 5
    Const EG1_EXP_GRP_so_type_cfg_ret_item_flag = 6
    Const EG1_EXP_GRP_so_type_cfg_sp_stk_flag = 7
    Const EG1_EXP_GRP_so_type_cfg_ci_flag = 8
    Const EG1_EXP_GRP_so_type_cfg_export_flag = 9
    Const EG1_EXP_GRP_so_type_cfg_auto_dn_flag = 10
    Const EG1_EXP_GRP_so_type_cfg_mov_type = 11
    Const EG1_EXP_GRP_so_type_cfg_mov_type_nm = 12
    Const EG1_EXP_GRP_so_type_cfg_trans_type = 13
    Const EG1_EXP_GRP_so_type_cfg_trans_type_nm = 14
    Const EG1_EXP_GRP_so_type_cfg_dlvy_lt = 15
    Const EG1_EXP_GRP_so_type_cfg_so_mgmt_flag = 16
    Const EG1_EXP_GRP_so_type_cfg_credit_chk_flag = 17
    Const EG1_EXP_GRP_so_type_cfg_deposit_chk_flag = 18

    Const EG1_EXP_GRP_so_type_cfg_usage_flag = 19
    Const EG1_EXP_GRP_so_type_cfg_ext1_qty = 20
    Const EG1_EXP_GRP_so_type_cfg_ext2_qty = 21
    Const EG1_EXP_GRP_so_type_cfg_ext1_amt = 22
    Const EG1_EXP_GRP_so_type_cfg_ext2_amt = 23
    Const EG1_EXP_GRP_so_type_cfg_ext1_cd = 24
    Const EG1_EXP_GRP_so_type_cfg_ext2_cd = 25

    Dim strSOType
    Redim I2_So_Type_Cfg(I2_So_Type_Cfg_Usg_Flg)

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status

	I2_So_Type_Cfg(I2_So_Type_Cfg_So_Type) = Trim(Request("txtSOType"))
	I2_So_Type_Cfg(I2_So_Type_Cfg_Usg_Flg) = Trim(Request("rdoUsageFlg"))
	strSOType = Request("txtSOType")
	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key

	If iStrPrevKey <> "" then
		iarrValue = Split(iStrPrevKey, gColSep)
		I1_Next_So_Type_Cfg = Trim(iarrValue(0))
	else
		I1_Next_So_Type_Cfg = ""
	End If

	Set iPD2GS64 = Server.CreateObject("PD2GS64.CListSoTypeConfigSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

	Call iPD2GS64.ListSoTypeConfigSvr(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_Next_So_Type_Cfg, I2_So_Type_Cfg, _
											   E1_So_Type_Cfg, EG1_EXP_GRP)

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs> " & vbCr
		Response.Write " Parent.frm1.txtSOTypeNm.value   = """"" & vbCr
		Response.Write "</Script>"
       Set iPD2GS64 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If

    Set iPD2GS64 = Nothing

    iLngMaxRow  = CLng(Request("txtMaxRows"))

	For iLngRow = 0 To UBound(EG1_EXP_GRP,1)

		If  iLngRow >= C_SHEETMAXROWS_D  Then
		   iStrNextKey = ConvSPChars(EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_so_type))
           Exit For
        End If

		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_so_type))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_so_type_nm))

		'멀티컴퍼니거래여부 
		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_intercom_flg) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if

		'2002-12-04 추가 2003-01-15 위치이동 
		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_sto_flag) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_export_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_ret_item_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_ci_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_rel_dn_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_auto_dn_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_rel_bill_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_sp_stk_flag) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_mov_type))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_mov_type_nm))

		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_trans_type))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_EXP_GRP(iLngRow,EG1_EXP_GRP_so_type_cfg_trans_type_nm))


		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_dlvy_lt),0, 0)

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_so_mgmt_flag) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_credit_chk_flag) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_deposit_chk_flag) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_usage_flag) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End if
			
		istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
		istrData = istrData & Chr(11) & Chr(12)

	Next
	
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.frm1.txtSOTypeNm.value   = """ & ConvSPChars(E1_So_Type_Cfg(E1_So_Type_Cfg_So_Type_Nm)) & """" & vbCr
    Response.Write " Parent.frm1.txtHSOType.value =  """ & strSOType										& """" & vbCr
    Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData									      " & vbCr
    Response.Write " Parent.ggoSpread.SSShowData        """ & istrData										     & """" & vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey										& """" & vbCr
    Response.Write " Parent.frm1.vspdData.ReDraw = False															" & vbCr
    Response.Write " Parent.SetSpreadColor -1 , -1																     	  " & vbCr
    Response.Write "</Script> "																							& vbCr

	For iLngRow = 0 To UBound(EG1_EXP_GRP,1)

		'멀티컴퍼니거래여부'
		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_intercom_flg) = "Y" Then
		End If


		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_sto_flag) = "Y" Then
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_ExportFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_CiFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_RelDnFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_RelBillFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_SoMgmtFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_CreditChkFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_DepositChkFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write "</Script>"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_export_flag) = "Y" Then
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_RetItemFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_RelDnFlg," & iLngRow + 1 & "," & iLngRow + 1 & vbCr
			Response.Write "</Script>"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_rel_dn_flag) <> "Y" Then
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_AutoDnFlg," & iLngRow + 1 & "," & iLngRow + 1  & vbCr
			Response.Write "</Script>"
		End if

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_rel_dn_flag) = "N" Then

			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Call Parent.OnChgRelBillFlg(  """ & "DnUnCheck" & """  , " & iLngRow + 1 & ") " & vbCr
			Response.Write "</Script>"
		Else
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Call Parent.OnChgRelBillFlg(  """ & "DnCheck" & """  , " & iLngRow + 1 & ") " & vbCr
			Response.Write "</Script>"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_rel_bill_flag) = "N" Then
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Call Parent.OnChgRelBillFlg(  """ & "BillUnCheck" & """  ," & iLngRow + 1 & ") " & vbCr
			Response.Write "</Script>"
		Else
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Call Parent.OnChgRelBillFlg(  """ & "BillCheck" & """  ," & iLngRow + 1 & ") " & vbCr
			Response.Write "</Script>"
		End If

		If EG1_EXP_GRP(iLngRow, EG1_EXP_GRP_so_type_cfg_ci_flag) = "Y" Then
			Response.Write "<Script language = vbs> " & vbCr
			Response.Write " Parent.ggoSpread.SSSetProtected	Parent.S_RelDnFlg," & iLngRow + 1 & "," & iLngRow + 1  & vbCr
			Response.Write "</Script>"
		End If

	Next

	Response.Write "<Script language = vbs> " & vbCr
	Response.Write " Parent.frm1.vspdData.ReDraw = True " & vbCr
	Response.Write " Parent.DbQueryOk " & vbCr
	Response.Write "</Script>"

End Sub


'============================================================================================================
Sub SubBizSaveMulti()

	Dim iPD2GS63
	Dim iErrorPosition
	Dim iLngMaxRow
	Dim strtxtSpread

    On Error Resume Next
    Err.Clear																			 '☜: Clear Error status


    If Request("txtMaxRows") = "" Then
		Call ServerMesgBox("MaxRows 조건값이 비어있습니다!",vbInformation, I_MKSCRIPT)
		Response.End
	End If

	strtxtSpread = Trim(Request("txtSpread"))
	Set iPD2GS63 = Server.CreateObject("PD2GS63.CMaintSoTypeConfigSvr")

    If CheckSYSTEMError(Err,True) = True Then
      Exit Sub
    End If

    Call iPD2GS63.MaintSoTypeConfigSvr(gStrGlobalCollection, strtxtSpread, iErrorPosition)

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPD2GS63 = Nothing
       Exit Sub
	End If

    Set iPD2GS63 = Nothing

    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.DBSaveOk "      & vbCr
    Response.Write "</Script> "

End Sub

'============================================================================================================
Sub SetErrorStatus()
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
End Sub

%>
