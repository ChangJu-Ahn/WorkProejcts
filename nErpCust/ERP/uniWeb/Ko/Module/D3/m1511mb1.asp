<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->

<%Call LoadBasisGlobalInf()%>
<%

Dim lgOpModeCRUD
 
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
	
lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
        
End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Dim TmpBuffer
    Dim iMax
    Dim iIntLoopCount
    Dim iTotalStr
    
    Dim iLngMaxRow
    Dim iLngRow
    Dim OBJ_PM1G438
    Dim istrData
    
    Dim lgStrPrevKey
    Dim iStrPrevKey
    Dim iGroupCount
    Dim StrNextKey  	
    Dim arrValue
    Dim I2_m_mvmt_type_next_cd
    
    Const C_SHEETMAXROWS_D  = 100
    
    Dim I1_m_mvmt_type
    Const C_imp_mvmt_io_type_cd		= 0
    Const C_imp_mvmt_usage_flg		= 1
    
    Dim exp_group
    
    Const C_exp_mvmt_cd_minor_nm	= 0
    Const C_exp_io_type_cd			= 1
    Const C_exp_io_type_nm			= 2
    Const C_exp_mvmt_cd				= 3
    Const C_exp_rcpt_flg			= 4
    Const C_exp_import_flg			= 5
    Const C_exp_ret_flg				= 6
    Const C_exp_subcontra_flg		= 7
    Const C_exp_usage_flg			= 8
    Const C_exp_ext1_cd				= 9
    Const C_exp_ext2_cd				= 10
    Const C_exp_ext3_cd				= 11
    Const C_exp_ext4_cd				= 12
    Const C_exp_child_settle_flg	= 13	'자품목 처리여부 
    Const C_exp_subcontra2_flg		= 14	'외주가공여부 
    
    'KSJ추가 2007.06.14 매입일괄등록 
    Const C_exp_except_iv_type_cd	= 15	'일괄매입형태 
    Const C_exp_except_iv_type_nm	= 16	'일괄매입형태명 
    
    Dim E1_m_mvmt_type
	Const C_exp_mvmt_io_type_cd		= 0
    Const C_exp_mvmt_io_type_nm		= 1
    
	
	Dim E2_m_mvmt_type_next_cd
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Redim I1_m_mvmt_type(C_imp_mvmt_usage_flg)
    I1_m_mvmt_type(C_imp_mvmt_io_type_cd) = UCase(Trim(Request("txtGmTypeCd")))
    I1_m_mvmt_type(C_imp_mvmt_usage_flg)  = Trim(Request("txtUseflg"))

    If Trim(Request("lgStrPrevKey")) <> ""  Then
	   I2_m_mvmt_type_next_cd = Trim(Request("lgStrPrevKey"))
	End If    
    
    Set OBJ_PM1G438 = Server.CreateObject("PM1G438.cMListMvmtTypeS")
    
    If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if
    
    Call OBJ_PM1G438.M_LIST_MVMT_TYPE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_mvmt_type, _
								I2_m_mvmt_type_next_cd, E1_m_mvmt_type, exp_group, E2_m_mvmt_type_next_cd)    
    
    If CheckSYSTEMError(Err,True) = true then 	 		
		Set OBJ_PM1G438 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End If 
	Set OBJ_PM1G438 = Nothing
	
    iLngMaxRow = CLng(Request("txtMaxRows"))
    iIntLoopCount = 0
	iMax = UBound(exp_group,1)
	ReDim TmpBuffer(iMax)
	
    If Not IsEmpty(exp_group) then
		For iLngRow = 0 To iMax
		    istrData = ""
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_io_type_cd))
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_io_type_nm))
		    
		    If exp_group(iLngRow,C_exp_rcpt_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else  istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_ret_flg) = "Y" then istrData = istrData & Chr(11) & "1"  Else  istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_subcontra_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_child_settle_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_subcontra2_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_import_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else istrData = istrData & Chr(11) & "0"
		    If exp_group(iLngRow,C_exp_usage_flg) = "Y" then istrData = istrData & Chr(11) & "1" Else istrData = istrData & Chr(11) & "0"
		    
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_mvmt_cd))
		    istrData = istrData & Chr(11) & ""
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_mvmt_cd_minor_nm))
		    
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_except_iv_type_cd))
		    istrData = istrData & Chr(11) & ""
		    istrData = istrData & Chr(11) & ConvSPChars(exp_group(iLngRow,C_exp_except_iv_type_nm))
		    
		    istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
		    istrData = istrData & Chr(11) & Chr(12)
		    
		    TmpBuffer(iIntLoopCount) = istrData
			iIntLoopCount = iIntLoopCount + 1
		Next
		iTotalStr = Join(TmpBuffer, "")
		
	End If
	 
    If  iStrPrevKey = E2_m_mvmt_type_next_cd  Then
		iStrPrevKey = ""
	Else
		StrNextKey = E2_m_mvmt_type_next_cd
	End If
        
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With Parent "               & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData"                                  & vbCr
	Response.Write " .ggoSpread.SSShowData     """ & iTotalStr                     & """" & vbCr
	Response.Write " .lgStrPrevKey           = """ & StrNextKey                    & """" & vbCr
	Response.Write " .frm1.hdnGmType.value   = """ & UCase(Request("txtGmTypeCd")) & """" & vbCr
	Response.Write " .frm1.hdnUseflg.value   = """ & UCase(Request("txtUseflg"))   & """" & vbCr
	Response.Write " .frm1.txtGmTypeNm.value = """ & ConvSPChars(E1_m_mvmt_type(C_exp_mvmt_io_type_nm)) & """" & vbCr
    Response.Write " .DbQueryOk "		    	& vbCr 
    Response.Write " .frm1.vspdData.focus "		& vbCr 			
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr    
	'Response.End
End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim OBJ_PM1G431
    Dim iErrorPosition
    Dim txtSpread
    
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Set OBJ_PM1G431 = Server.CreateObject("PM1G431.cMMaintMvmtTypeS")
	
	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End If
	txtSpread = Trim(Request("txtSpread"))
	
	Call OBJ_PM1G431.M_MAINT_MVMT_TYPE_SVR(gStrGlobalCollection, txtSpread, iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set OBJ_PM1G431 = Nothing
       Exit Sub
	End If
	Set OBJ_PM1G431 = Nothing    
       
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 
    'Response.End   
End Sub    


%>

