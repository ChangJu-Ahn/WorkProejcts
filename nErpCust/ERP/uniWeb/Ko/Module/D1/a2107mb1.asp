<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")


	Call HideStatusWnd

    Err.Clear
    On Error Resume Next

    Dim lgOpModeCRUD

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    End Select


Sub SubBizQueryMulti()
    On Error Resume Next
    Err.Clear

    Dim pPD1G035

    Dim iStrData

    Dim StrNextKey1
    Dim StrNextKeyOne_Seq
    Dim StrNextKeyJnl1			' 거래항목 다음 값 
    Dim StrNextKeyDrCrFg1		' 차대구분 다음 값 
    Dim StrNextKeyAcct1			' 계정 다음 값 
    Dim StrNextKeyBizArea1		' 사업장 다음 값 
    Dim lgStrPrevKeyOne_Seq
    Dim iIntQueryCount
    Dim LngMaxRow
    DIm iStrPrevKey
    'Dim lgStrPrevKey1
    'Dim lgStrPrevKeyJnl1		' 거래항목 이전 값 
    'Dim lgStrPrevKeyDrCrFg1	' 차대구분 이전 값 
    'Dim lgStrPrevKeyAcct1		' 계정 이전 값 
    'Dim lgStrPrevKeyBizArea1	' 사업장 이전 값 

    Const C_SHEETMAXROWS_D	= 100
    const C_AJnlFormSeq		= 0
    const C_MaxQueryReCord	= 1
    const C_AJnlFormKey		= 2

    Dim LngRow
    Dim GroupCount



    Dim iIntLoopCount
    Dim iLngRow,iLngCol



    Dim importArray
    Dim E1_a_trans_nm
    Dim EG1_export_group
    Dim EG2_export_group  'Not Used

    Const A248_EG1_E2_jnl_cd	= 0
    Const A248_EG1_E2_jnl_nm	= 1
    Const A248_EG1_E3_seq		= 2
    Const A248_EG1_E3_dr_cr_fg	= 3
    Const A248_EG1_E3_event_cd	= 4
    Const A248_EG1_E1_jnl_nm	= 5
    Const A248_EG1_E4_acct_cd	= 6
    Const A248_EG1_E4_acct_nm	= 7

    iIntQueryCount      = Request("lgPageNo")
    StrNextKeyOne_Seq   = Request("lgStrPrevKeyOne_Seq")
    LngMaxRow           = Request("txtMaxRows_One")  '☜: 최대 업데이트된 갯수 

    ReDim importArray(2)
    importArray(C_AJnlFormSeq)      = UNIConvNum(StrNextKeyOne_Seq,0)
    importArray(C_MaxQueryReCord)   = C_SHEETMAXROWS_D
    importArray(C_AJnlFormKey)      = Request("txtTransType")

    ReDim E1_a_trans_nm(1)
    Const EA_a_acct_trans_type_trans_type1 = 0
    Const EA_a_acct_trans_type_trans_nm1 = 1

    If CheckSYSTEMError(Err, True) = True Then
        Set pPD1G035  = Nothing
        Exit Sub
    End If

    Set pPD1G035  = Server.CreateObject("PD1G035.cAListJnlFormSvr")

    Call pPD1G035.A_LIST_JNL_FORM_SVR(gStrGlobalCollection, importArray, E1_a_trans_nm, EG1_export_group, EG2_export_group)

    If CheckSYSTEMError(Err, True) = True Then
        Set pPD1G035  = Nothing
        Exit Sub
    End If


    Set pPD1G035  = Nothing

    iStrData = ""
    iIntLoopCount = 0 

	If isEmpty(EG1_export_group) = False then

	   For iLngRow = 0 To UBound(EG1_export_group, 1)
	      iIntLoopCount = iIntLoopCount + 1

	       If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then 

	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E2_jnl_cd))
	          iStrData = iStrData & Chr(11) & ""
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E2_jnl_nm))
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E3_seq))
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E3_dr_cr_fg))
	          iStrData = iStrData & Chr(11) & ""
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E3_event_cd))
	          iStrData = iStrData & Chr(11) & ""
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E1_jnl_nm))
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E4_acct_cd))
	          iStrData = iStrData & Chr(11) & ""
	          iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A248_EG1_E4_acct_nm))
	          iStrData = iStrData & Chr(11) & "N" '"" 2002.09.10 jsk
	          iStrData = iStrData & Chr(11) & ""
	          iStrData = iStrData & Chr(11) & ""  '"N" 2002.09.10 jsk
	          iStrData = iStrData & Chr(11) & Cstr(iLngRow + 1 + LngMaxRow) & Chr(11) & Chr(12)

	       Else
	          iStrPrevKey   = EG1_export_group(UBound(EG1_export_group, 1), A248_EG1_E2_jnl_cd)
	          StrNextKeyOne_Seq = EG1_export_group(iLngRow, A248_EG1_E3_seq)
	          iIntQueryCount  = iIntQueryCount + 1
	     Exit For
	    End If
	   Next
	End If

 If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
    iStrPrevKey = ""
    iIntQueryCount = ""
 End If

    Response.Write " <Script Language=vbscript>                         " & vbCr
    Response.Write " With parent                                        " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write " .ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write " .lgStrPrevKeyOne_Seq   = """ & StrNextKeyOne_Seq    & """" & vbCr
    Response.Write " .frm1.txtTransNM.value = """ & ConvSPChars(E1_a_trans_nm(EA_a_acct_trans_type_trans_nm1))    & """" & vbCr
    Response.Write " .frm1.hTransType.value = """ & ConvSPChars(E1_a_trans_nm(EA_a_acct_trans_type_trans_type1))     & """" & vbCr
    Response.Write " .lgPageNo = """ & iIntQueryCount    & """" & vbCr
    Response.Write " .lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)    & """" & vbCr
    Response.Write " .DbQuery_OneOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
 End Sub
'--------------------------------------------------------------------------------------------------------
'                                   SAVE
'--------------------------------------------------------------------------------------------------------
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear

	Dim pPD1G035

	Set pPD1G035 = Server.CreateObject("PD1G035.cAMngJnlFormSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Set pPD1G035 = Nothing
		Exit Sub
	End If


	Call pPD1G035.A_MANAGE_JNL_FORM_SVR(gStrGlobalCollection, Request("txtSpread1"), Request("txtSpread2"))


	If CheckSYSTEMError(Err, True) = True Then
		Set pPD1G035 = Nothing
		Exit Sub
	End If

	Set pPD1G035 = Nothing


%>
 <Script Language=vbscript>
  Parent.DbSave_OneOk("<%=Request("txtTransType")%>")    '☜: 화면 처리 ASP 를 지칭함 
  'Parent.DbSave_Two
 </Script>
<%
End Sub
%>

