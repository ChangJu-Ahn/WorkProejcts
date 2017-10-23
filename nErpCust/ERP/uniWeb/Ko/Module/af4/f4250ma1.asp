
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : F4250MA1
*  4. Program Name         : 차입금상환등록 
*  5. Program Desc         : Single Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/30
*  8. Modified date(Last)  : 2003/05/19
*  9. Modifier (First)     : Ahn, do hyun
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<Script Language="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID	= "F4250MB1.asp"						           '☆: Biz Logic ASP Name
'Const BIZ_PGM_ID2	= "F4250MB2.asp"

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgNm_of_Cdfg
dim strIntClsPlanAmt
dim strIntClsPlanLocAmt
Dim txtXchRate

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgNm_of_Cdfg	  = ""
	IsOpenPop = False
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    Dim strYear                                                 '⊙: User-defined variables for year  (You can delete this line if you need not)
    Dim strMonth                                                '⊙: User-defined variables for month (You can delete this line if you need not)
    Dim strDay                                                  '⊙: User-defined variables for day   (You can delete this line if you need not)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "LOOKUP"
                  lgKeyStream = Trim(Frm1.txtPayNo.Value)       'You Must append one character(Parent.gColSep)
       Case "D"
                  lgKeyStream = Trim(Frm1.hPayNo.Value) & Parent.gColSep
                  lgKeyStream = lgKeyStream & frm1.txtLoanNo.value & Parent.gColSep
       Case "NEXT"
                  lgKeyStream = Trim(Frm1.hPayNo.Value)       'You Must append one character(Parent.gColSep)
       Case "PREV"
                  lgKeyStream = Trim(Frm1.hPayNo.Value)       'You Must append one character(Parent.gColSep)
   End Select
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
                   'ComboObject Name      Name   Value  Separator
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

	Call InitVariables
	Call InitComboBox()

	frm1.txtPayNo.focus
	Call SetToolbar("1110000000001111")                                              '☆: Developer must customize

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

'	Call ggoOper.SetReqAttr(frm1.txtEtcPay, "Q")
'	Call ggoOper.SetReqAttr(frm1.txtEtcPayLoc, "Q")

    Call SetDefaultVal()
    
	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing    
	
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This check required field
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    If DbQuery("LOOKUP") = False Then                                                       '☜: Query db data
       Exit Function
    End If
   
    frm1.txtPayNo.focus
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                        '☜: Clear Condition Field

	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call txtRcptTypeCd_OnChange()    
    Call SetToolbar("1110100000001111")
	Call InitVariables

	frm1.txtPayNo.focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncNew = True															      '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                             '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")                         '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If frm1.hConfFg.value = "C" Then
        Call DisplayMsgBox("114114","x","x","x")
		Exit Function
	End If	

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbDelete = False Then                                                      '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	If frm1.hConfFg.value = "C" Then
        Call DisplayMsgBox("114114","x","x","x")
		Exit Function
	End If	

'	If UNICDbl(frm1.txtXchRate.Text) = "0" Then
'        Call DisplayMsgBox("173125","x","x","x")
'        Exit Function
'	End If

	If Trim(frm1.txtRcptTypeCd.value) = "DP" Then
		Call CommonQueryRs("C.TRANS_STS, DOC_CUR "," F_DPST C","C.BANK_CD = " & FilterVar(frm1.txtBankCd.value, "''", "S") & _
			" AND C.BANK_ACCT_NO = " & FilterVar(frm1.txtBankAcctNo.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If lgF0 = "CN" & chr(11) or lgF0 = "MT" & chr(11) Then
			Call DisplayMsgBox("140629","x","x","x")
			Exit Function
		End If

		If UCase(lgF1) <> UCase(Trim(frm1.txtDocCur.value)) & chr(11) Then
			Call DisplayMsgBox("140721","x","x","x")
			Exit Function
		End If
	End If

    Call Nm_of_Cd_Call

	Select Case lgNm_of_Cdfg
		Case "txtRcptTypeCd"
			frm1.txtRcptTypeCd.focus
			Exit Function
		Case "txtDeptCd"
			frm1.txtDeptCd.focus
			Exit Function
		Case "txtBankCd"
			frm1.txtBankCd.focus
			Exit Function
	End Select
	If Trim(frm1.txtRcptAcctCd.value) <> "" Then
		If CommonQueryRs("A.ACCT_CD","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.txtRcptTypeCd.value, "''", "S") & _
				" AND A.ACCT_CD = " & FilterVar(frm1.txtRcptAcctCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("141167","x",frm1.txtRcptAcctCd.Alt,"x")
			Exit Function
		End If
	End If
	If Trim(frm1.txtIntPayAcctCd.value) <> "" Then
		If CommonQueryRs("A.ACCT_CD","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("PI", "''", "S") & " " & _
				" AND A.ACCT_CD = " & FilterVar(frm1.txtIntPayAcctCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("141167","x",frm1.txtIntPayAcctCd.Alt,"x")
			Exit Function
		End If
	End If

	If Trim(frm1.txtChargeAcctCd.value) <> "" Then
		If CommonQueryRs("A.ACCT_CD","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("BC", "''", "S") & " " & _
				" AND A.ACCT_CD = " & FilterVar(frm1.txtChargeAcctCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("141167","x",frm1.txtChargeAcctCd.Alt,"x")
			Exit Function
		End If
	End If

	If Trim(frm1.txtEtcBPAcctCd.value) <> "" Then
		If CommonQueryRs("A.ACCT_CD","A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C","A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("BP", "''", "S") & " " & _
				" AND A.ACCT_CD = " & FilterVar(frm1.txtEtcBPAcctCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("141167","x",frm1.txtEtcBPAcctCd.Alt,"x")
			Exit Function
		End If
	End If

	If UNICDbl(frm1.txtPlanAmtPR.Text) = 0 and UNICDbl(frm1.txtPlanAmtPI.Text) = 0 then
        Call DisplayMsgBox("141138","x","x","x")'두개다입력 
		Exit Function
	End If

	If UNICDbl(frm1.txtPlanAmtPR.Text) < 0 then
        Call DisplayMsgBox("141139","x","","x")'양수 
		Exit Function
	End If

	If (UNICDbl(frm1.txtIntPayAmt.Text) + UNICDbl(frm1.txtPlanAmtPI.Text) < 0 ) and  _
		Trim(frm1.hIntPayStnd.value) = "DI" then
		'원금상환총액과 원금상환액의 합은 0보다 크거나 같아야합니다.
        Call DisplayMsgBox("141170","x","x","x")
		Exit Function
	End If

    If UNICDbl(frm1.txtLoanBalAmt.Text) + UNICDbl(frm1.hPlanAmtPR.value) < UNICDbl(frm1.txtPlanAmtPR.Text)  Then
		Call DisplayMsgBox("141141","x","x","x")
		Exit Function
	End if
	
	If Trim(frm1.hIntPayStnd.value) = "DI" and UNICDbl(frm1.txtDfrIntPay.Text) > UNICDbl(frm1.txtPlanAmtPI.Text) Then
		Call DisplayMsgBox("141142","x","x","x")
		Exit Function
	End if

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If
    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									     '⊙: This lock the suitable field
    Call CancelRestoreToolBar()                                                  '⊙: If you do not follow common toolbar operation ....
    Call SetToolbar("1110100000001111")
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncCancel() 
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
' Keep : This Function is related to multi but single
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    Dim IntRetCD
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
     
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
   	
    Call InitVariables														     '⊙: Initializes local global variables

    If DbQuery("PREV") = False Then                                                       '☜: Query db data
       Exit Function
    End If


    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim IntRetCD
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
     
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Call InitVariables														     '⊙: Initializes local global variables

    If DbQuery("NEXT") = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_SINGLE)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_SINGLE, True)

    FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")	                 '⊙: Data is changed.  Do you want to exit? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    FncExit = True                                                               '☜: Processing is OK

End Function


'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Call MakeKeyStream(pDirect)

    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
   	
    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtPrevNext="      & pDirect                         '☜: Direction
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인    
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    DbQuery = True                                                               '☜: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strYear,strMonth,strDay 
	Dim strYYYYMMDD
    Err.Clear                                                                    '☜: Clear err status
    
	DbSave = False														         '☜: Processing is NG

	Call LayerShowHide(1)

	Call ExtractDateFrom(frm1.txtPayPlanDt,frm1.txtPayPlanDt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMMDD = strYear & strMonth & strDay				' ?에 Pay_dt를 넣어주세요.   	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode                                     ' U:Update mode, C:Create Mode
		.htxtPayPlanDt.value  = strYYYYMMDD
		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end		
	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This Sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)

    Call MakeKeyStream("D")
		
    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="     & "D"                     '☜: Query Key

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Call ggoOper.LockField(Document, "Q")
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call txtRcptTypeCd_userChange

	If UNICDbl(frm1.txtPlanAmtPR.Text) = 0 Then
		Call ggoOper.SetReqAttr(frm1.txtPlanAmtPR, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPlanLocAmtPR, "Q")
'		Call ggoOper.SetReqAttr(frm1.txtEtcPay, "Q")
'		Call ggoOper.SetReqAttr(frm1.txtEtcPayLoc, "Q")
	Else
'		Call ggoOper.SetReqAttr(frm1.txtEtcPay, "D")
'		Call ggoOper.SetReqAttr(frm1.txtEtcPayLoc, "D")

	End If

	If UNICDbl(frm1.txtEtcPay.Text) <> 0 Then
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "N")
	End If
	If UNICDbl(frm1.txtEtcBPPay.Text) <> 0 Then
		Call ggoOper.SetReqAttr(frm1.txtEtcBPAcctCd, "N")
	End If

	If UNICDbl(frm1.txtPlanAmtPI.Text) = 0 Then
		Call ggoOper.SetReqAttr(frm1.txtPlanAmtPI, "Q")
		Call ggoOper.SetReqAttr(frm1.txtPlanLocAmtPI, "Q")
	ElseIf Trim(frm1.hIntPayStnd.value) = "DI" AND UNICDbl(frm1.txtPlanAmtPI.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtIntPayAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtIntPayAcctCd, "Q")
	End If

	Call CurFormatNumericOCX()
    Frm1.txtPayNo.focus 

    Call CancelRestoreToolBar()                                                  '⊙: If you do not follow common toolbar operation ....
	If Frm1.hConfFg.Value = "C" Then
		Call SetToolbar("111000001101111")
	Else
		Call SetToolbar("111110001101111")
	End If

	lgBlnFlgChgValue	= false
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   

End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk(byval pDirect)
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    Select Case lgIntFlgMode
		Case Parent.OPMD_CMODE
			frm1.txtPayNo.value = pDirect
    End Select 
	
    Call InitVariables

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call MainQuery()
End Sub
	
'========================================================================================================
' Name : DbDeleteOk
' Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call FncNew()
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

		ggoOper.FormatFieldByObjectOfCur .txtLoanAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtRdpAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtLoanBalAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtIntPayAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		ggoOper.FormatFieldByObjectOfCur .txtPlanAmtPR,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtPlanAmtPI,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
'		ggoOper.FormatFieldByObjectOfCur .txtTotAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDfrIntPay,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtEtcPay,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtEtcBPPay,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub


'============================================================================================================
' Name : Nm_of_Cd_Call()
' Desc : Delete DB data
'============================================================================================================
Sub Nm_of_Cd_Call()
    On Error Resume Next                                                               '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
	Call CommonQueryRs(" A.MINOR_NM ","B_MINOR A, B_CONFIGURATION B","A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 4 AND B.REFERENCE IN (" & FilterVar("FO", "''", "S") & " ," & FilterVar("DP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " ," & FilterVar("CK", "''", "S") & " ) AND a.minor_cd = " & FilterVar(Trim(frm1.txtRcptTypeCd.value)	, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IsNull(lgF0) or lgF0 = "" Then
		Call DisplayMsgBox("141140","X","X","X")            '☜ : No data is found.
		lgNm_of_Cdfg = "txtRcptTypeCd"
		Exit Sub
	Else 
		frm1.txtRcptTypeNm.value = Replace(lgf0,chr(11),"")

		Call CommonQueryRs(" A.DEPT_NM ","B_ACCT_DEPT A","A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & " AND a.dept_cd = " & FilterVar(Trim(frm1.txtDeptCd.value)	, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If IsNull(lgF0) or lgF0 = "" Then
			Call DisplayMsgBox("800062","X","X","X")            '☜ : No data is found. 
			lgNm_of_Cdfg = "txtDeptCd"
			Exit Sub
		Else
			frm1.txtDeptNm.value = Replace(lgf0,chr(11),"")

			If Trim(frm1.txtBankCd.value) <> "" Then
				Call CommonQueryRs(" A.BANK_NM ","B_BANK A, B_BANK_ACCT B","A.BANK_CD = B.BANK_CD AND A.BANK_CD = " & FilterVar(frm1.txtBankCd.value, "''", "S") & " and bank_acct_no = " & FilterVar(frm1.txtBankAcctNo.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IsNull(lgF0) or lgF0 = "" Then
					Call DisplayMsgBox("200033","X","X","X")            '☜ : No data is found. 
					lgNm_of_Cdfg = "txtBankCd"
					Exit Sub
				Else
					frm1.txtBankNm.value = Replace(lgf0,chr(11),"")
				End If
			End If
		End If
	End If

	lgNm_of_Cdfg = ""
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : sumPRPI()
' Desc : developer describe this line 
'========================================================================================================
Function sumPRPI()

'	frm1.txtTotAmt.Text = UNIFormatNumber(UNICDbl(frm1.txtPlanAmtPR.Text) + Parent.UNICDbl(frm1.txtPlanAmtPI.Text), ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
'	frm1.txtTotLocAmt.Text = UNIFormatNumber(UNICDbl(frm1.txtPlanLocAmtPR.Text) + Parent.UNICDbl(frm1.txtPlanLocAmtPI.Text), ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
    lgBlnFlgChgValue = True

End Function

'========================================================================================================
' Name : OpenPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopup(Byval strCode,Byval iWhere)
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
	Case 2
		If frm1.txtBankCd.className = Parent.UCN_PROTECTED Then Exit Function		
		
		arrParam(0) = "은행팝업"									' 팝업 명칭 
		arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE 명칭 
		arrParam(2) = strCode													' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition			
		arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
		arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "		
		arrParam(5) = frm1.txtBankCd.Alt									' 조건필드의 라벨 명칭 

		arrField(0) = "A.BANK_CD"						' Field명(0)
		arrField(1) = "A.BANK_NM"						' Field명(1)
		arrField(2) = "B.BANK_ACCT_NO"					' Field명(2)
   
		arrHeader(0) = frm1.txtBankCd.Alt					' Header명(0)
		arrHeader(1) = frm1.txtBankNm.Alt					' Header명(1)
		arrHeader(2) = frm1.txtBankAcctNo.Alt				' Header명(2)					

	Case 3
		arrParam(0) = strCode		            '  Code Condition
		arrParam(1) = frm1.txtPayPlanDt.Text
		arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
		arrParam(3) = "F"									' 결의일자 상태 Condition  
'		If lgInternalCd <> "" Then
'			arrParam(4) = " INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
'		Else
'			arrParam(4) = ""
'		End If
'
'		If lgSubInternalCd <> "" Then
'			arrParam(4) = " INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
'		Else
'			arrParam(4) = ""
'		End If		
		
		' 권한관리 추가 
		arrParam(5) = lgAuthBizAreaCd
		arrParam(6) = lgInternalCd
		arrParam(7) = lgSubInternalCd
		arrParam(8) = lgAuthUsrID		
    
	Case 4		'출금유형 
		arrParam(0) = "출금유형"
		arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
					& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD AND B.SEQ_NO = 4 AND B.REFERENCE IN (" & FilterVar("FO", "''", "S") & " ," & FilterVar("DP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " ," & FilterVar("CK", "''", "S") & " ) "
		arrParam(5) = frm1.txtRcptTypeCd.Alt
	
		arrField(0) = "A.MINOR_CD"
		arrField(1) = "A.MINOR_NM"
			    
		arrHeader(0) = frm1.txtRcptTypeCd.Alt
		arrHeader(1) = frm1.txtRcptTypeNm.Alt
	Case 5
		If frm1.txtBankAcctNo.className = Parent.UCN_PROTECTED Then Exit Function		
		
		arrParam(0) = "계좌번호팝업"								' 팝업 명칭 
		arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition			
		arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
		arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "		
		arrParam(5) = frm1.txtBankAcctNo.Alt								' 조건필드의 라벨 명칭 

		arrField(0) = "B.BANK_ACCT_NO"					' Field명(0)
		arrField(1) = "A.BANK_CD"						' Field명(0)
		arrField(2) = "A.BANK_NM"						' Field명(0)
    
		arrHeader(0) = frm1.txtBankAcctNo.Alt					' Header명(0)
		arrHeader(1) = frm1.txtBankCd.Alt					' Header명(0)
		arrHeader(2) = frm1.txtBankNm.Alt						' Header명(0)				

	Case 6
		If frm1.txtIntPayAcctCd.className = "protected" Then Exit Function    
			
		arrParam(0) = "이자비용계정팝업"								' 팝업 명칭 
		arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			" and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("PI", "''", "S") & " " '& FilterVar(Trim(frm1.cboIntPayStnd.Value),"''","S")	' Where Condition
		arrParam(5) = frm1.txtIntPayAcctCd.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "A.Acct_CD"									' Field명(0)
		arrField(1) = "A.Acct_NM"									' Field명(1)
		arrField(2) = "B.GP_CD"										' Field명(2)
		arrField(3) = "B.GP_NM"										' Field명(3)
			
		arrHeader(0) = frm1.txtIntPayAcctCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtIntPayAcctNm.Alt								' Header명(1)
		arrHeader(2) = "그룹코드"									' Header명(2)
		arrHeader(3) = "그룹명"										' Header명(3)

	Case 7
		If frm1.txtRcptAcctCd.className = "protected" Then Exit Function    
				
		arrParam(0) = "출금계정팝업"								' 팝업 명칭 
		arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			" and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar(frm1.txtRcptTypeCd.Value, "''", "S")	' Where Condition
		arrParam(5) = frm1.txtRcptAcctCd.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "A.Acct_CD"									' Field명(0)
		arrField(1) = "A.Acct_NM"									' Field명(1)
		arrField(2) = "B.GP_CD"										' Field명(2)
		arrField(3) = "B.GP_NM"										' Field명(3)
				
		arrHeader(0) = frm1.txtRcptAcctCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtRcptAcctNm.Alt								' Header명(1)
		arrHeader(2) = "그룹코드"									' Header명(2)
		arrHeader(3) = "그룹명"										' Header명(3)						

	Case 8
		If frm1.txtChargeAcctCd.className = "protected" Then Exit Function    
				
		arrParam(0) = "부대비용계정팝업"								' 팝업 명칭 
		arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			" and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("BC", "''", "S") & " " ' & FilterVar(Trim(frm1.txtRcptType.Value),"''","S")	' Where Condition
		arrParam(5) = frm1.txtChargeAcctCd.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "A.Acct_CD"									' Field명(0)
		arrField(1) = "A.Acct_NM"									' Field명(1)
		arrField(2) = "B.GP_CD"										' Field명(2)
		arrField(3) = "B.GP_NM"										' Field명(3)
				
		arrHeader(0) = frm1.txtChargeAcctCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtChargeAcctNm.Alt								' Header명(1)
		arrHeader(2) = "그룹코드"									' Header명(2)
		arrHeader(3) = "그룹명"										' Header명(3)						
	Case 9
		If frm1.txtEtcBPAcctCd.className = "protected" Then Exit Function    
				
		arrParam(0) = "수수료계정팝업"								' 팝업 명칭 
		arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD" & _
			" and C.trans_type = " & FilterVar("FI002", "''", "S") & "  and C.jnl_cd = " & FilterVar("BP", "''", "S") & " " ' & FilterVar(Trim(frm1.txtRcptType.Value),"''","S")	' Where Condition
		arrParam(5) = frm1.txtEtcBPAcctCd.Alt							' 조건필드의 라벨 명칭 

		arrField(0) = "A.Acct_CD"									' Field명(0)
		arrField(1) = "A.Acct_NM"									' Field명(1)
		arrField(2) = "B.GP_CD"										' Field명(2)
		arrField(3) = "B.GP_NM"										' Field명(3)
				
		arrHeader(0) = frm1.txtEtcBPAcctCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtEtcBPAcctNm.Alt								' Header명(1)
		arrHeader(2) = "그룹코드"									' Header명(2)
		arrHeader(3) = "그룹명"										' Header명(3)						


    Case Else
		Exit Function
    End Select    
	
	IsOpenPop = True
	
	Select Case iWhere
		Case 3
			arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		Case 2, 5
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=680px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		Case Else
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
					 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select


	IsOpenPop = False
	
	If arrRet(0) = "" Then
			Select Case iWhere
			Case 1
				frm1.txtPayNo.focus
			Case 2
				frm1.txtBankCd.focus
			Case 3
				frm1.txtPayPlanDt.focus
			Case 4		'출금유형 
				frm1.txtRcptTypeCd.focus
			Case 5
				Frm1.txtBankAcctNo.focus
			Case 6
				frm1.txtIntPayAcctCd.focus
			Case 7
				frm1.txtRcptAcctCd.focus
			Case 8
				frm1.txtChargeAcctCd.focus
			Case 9
				frm1.txtEtcBPAcctCd.focus
			End Select          
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'======================================================================================================
'	Name : SubSetSchoolInf()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetPopUp(arrRet,iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtPayNo.value = arrRet(0)
				.txtPayNo.focus
			Case 2
				.txtBankCd.value = arrRet(0)
				.txtBankNm.value = arrRet(1)
				.txtBankAcctNo.value = arrRet(2)
				.txtBankCd.focus
				lgBlnFlgChgValue = True
			Case 3
				.txtPayPlanDt.text = arrRet(3)
			    .txtDeptCd.value = arrRet(0)
			    .txtDeptNm.value = arrRet(1)
			    .txtDeptCd.focus
				call txtDeptCd_OnChange()  
				lgBlnFlgChgValue = True
			Case 4		'출금유형 
				.txtRcptTypeCd.value = arrRet(0)
				.txtRcptTypeNm.value = arrRet(1)		
				Call txtRcptTypeCd_OnChange
				.txtRcptTypeCd.focus
				lgBlnFlgChgValue = True
			Case 5
				.txtBankAcctNo.value = arrRet(0)
				.txtBankCd.value = arrRet(1)
				.txtBankNm.value = arrRet(2)
				.txtBankAcctNo.focus
				lgBlnFlgChgValue = True
			Case 6
			    .txtIntPayAcctCd.value = arrRet(0)
			    .txtIntPayAcctNm.value = arrRet(1)
			    .txtIntPayAcctCd.focus
				lgBlnFlgChgValue = True
			Case 7
			    .txtRcptAcctCd.value = arrRet(0)
			    .txtRcptAcctNm.value = arrRet(1)
			    .txtRcptAcctCd.focus
				lgBlnFlgChgValue = True
			Case 8
			    .txtChargeAcctCd.value = arrRet(0)
			    .txtChargeAcctNm.value = arrRet(1)
			    .txtChargeAcctCd.focus
				lgBlnFlgChgValue = True
			Case 9
			    .txtEtcBPAcctCd.value = arrRet(0)
			    .txtEtcBPAcctNm.value = arrRet(1)
			    .txtEtcBPAcctCd.focus
				lgBlnFlgChgValue = True

		End Select
	End With
End Sub

'=======================================================================================================
'   Event Desc : 입금유형별 Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptTypeCd_OnChange()
	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    strval = frm1.txtRcptTypeCd.value
            
    If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		Select Case UCase(lgF0)
			Case "CS" & Chr(11)
				frm1.txtBankCd.value = ""
				frm1.txtBankNm.value = ""
				frm1.txtBankAcctNo.value = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
											
			Case "DP" & Chr(11)			' 예적금 
				Call ggoOper.SetReqAttr(frm1.txtBankCd, "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
			Case "NO" & Chr(11)
				frm1.txtBankCd.value = ""
				frm1.txtBankNm.value = ""
				frm1.txtBankAcctNo.value = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
			Case Else
				frm1.txtBankCd.value = ""
				frm1.txtBankNm.value = ""
				frm1.txtBankAcctNo.value = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		End Select
	Else
		frm1.txtBankCd.value = ""
		frm1.txtBankNm.value = ""
		frm1.txtBankAcctNo.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
	End If
	
	frm1.txtRcptAcctCd.value = ""
	frm1.txtRcptAcctNm.value = ""
End Sub

'=======================================================================================================
'   Event Desc : 입금유형별 Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptTypeCd_userChange()
	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    strval = frm1.txtRcptTypeCd.value
            
    If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		Select Case UCase(lgF0)
		Case "CS" & Chr(11)
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")									
										
		Case "DP" & Chr(11)			' 예적금 
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "N")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "N")
		Case "NO" & Chr(11)
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		Case Else
			frm1.txtBankCd.value = ""
			frm1.txtBankNm.value = ""
			frm1.txtBankAcctNo.value = ""
			Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
		End Select
	Else
		frm1.txtBankCd.value = ""
		frm1.txtBankNm.value = ""
		frm1.txtBankAcctNo.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBankCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcctNo, "Q")
	End If
End Sub

'================================================================
'상환번호 참조 팝업 
'================================================================
Function OpenPopupPay()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("F4250RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "F4250RA2", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then
		Exit Function
	Else
	    Call ggoOper.ClearField(Document, "1")											'☜: Clear Contents  Field
		Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

		frm1.txtPayNo.value = arrRet(0)
	End If
End Function

'================================================================
'차입금참조 팝업 
'================================================================
Function OpenPopupLoan()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("F4250RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "F4250RA1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False

	If arrRet(0) = ""  Then
		frm1.txtLoanNo.focus
		Exit Function
	Else
		Call ggoOper.ClearField(Document, "A")											'☜: Clear Contents  Field
		Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

		frm1.txtLoanBalAmt.Text = arrRet(12)
		frm1.txtLoanPlcCd.value = arrRet(14)
		frm1.txtLoanPlcNm.value = arrRet(15)
		frm1.txtLoanDt.Text = arrRet(16)
		frm1.txtDueDt.Text = arrRet(17)
		frm1.cboLoanFg.value = arrRet(3)
		frm1.txtLoanType.value = arrRet(5)
		frm1.txtLoanTypeNm.value = arrRet(6)
		frm1.txtLoanAmt.Text = arrRet(18)
		frm1.txtRdpAmt.Text = arrRet(20)												'원금상환총액 
		frm1.txtIntPayAmt.Text = arrRet(22)												'이자지급총액 
		frm1.txtLoanIntRate.Text = UNIFormatNumber(arrRet(25),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit)	'이자율 
		
		frm1.txtDocCur.value = arrRet(7)
		frm1.txtPayPlanDt.Text = arrRet(2)

		IF arrRet(26) = "DI" Then			'미지급이자 
			Call CommonQueryRs("sum(int_cls_amt),sum(int_cls_loc_amt), xch_rate","f_ln_mon_dfr_int"," loan_no = " & FilterVar(arrRet(0), "''", "S") & _
				" and int_pay_plan_dt =  " & FilterVar(UNIConvDate(arrRet(2)), "''", "S") & " and CLS_FG = " & FilterVar("Y", "''", "S") & "  group by xch_rate"  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			strIntClsPlanAmt = Replace(lgf0,chr(11),"")
			strIntClsPlanLocAmt = Replace(lgf1,chr(11),"")
			frm1.txtDfrIntPay.Text = strIntClsPlanAmt 
			'frm1.txtDfrIntPay.Text = UNIFormatNumber(strIntClsPlanAmt,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			frm1.txtDfrIntPayLoc.Text = UNIFormatNumber(strIntClsPlanLocAmt,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
'			frm1.hDfrXchRate.value = Replace(lgf2,chr(11),"")
		Else
			frm1.txtDfrIntPay.Text = strIntClsPlanAmt 
			'frm1.txtDfrIntPay.Text = UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			frm1.txtDfrIntPayLoc.Text = UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
'			frm1.hDfrXchRate.value = 0
		End If
		
		frm1.txtLoanNo.value = arrRet(0)
		
		frm1.txtLoanNm.value = arrRet(1)

		If UNICDbl(arrRet(8)) = 0 Then
			Call ggoOper.SetReqAttr(frm1.txtPlanAmtPR, "Q")
			Call ggoOper.SetReqAttr(frm1.txtPlanLocAmtPR, "Q")
'			Call ggoOper.SetReqAttr(frm1.txtEtcPay, "Q")
'			Call ggoOper.SetReqAttr(frm1.txtEtcPayLoc, "Q")
		Else
'			Call ggoOper.SetReqAttr(frm1.txtEtcPay, "D")
'			Call ggoOper.SetReqAttr(frm1.txtEtcPayLoc, "D")
		End If
		If UNICDbl(arrRet(10)) = 0 Then
			Call ggoOper.SetReqAttr(frm1.txtPlanAmtPI, "Q")
			Call ggoOper.SetReqAttr(frm1.txtPlanLocAmtPI, "Q")
		End If

		If arrRet(26) = "AI" Then
			frm1.txtIntPayStnd1.checked = true
			frm1.txtIntPayStnd2.checked = false
			frm1.hIntPayStnd.value = "AI"
		ElseIf arrRet(26) = "DI" Then
			frm1.txtIntPayStnd1.checked = false
			frm1.txtIntPayStnd2.checked = true
			frm1.hIntPayStnd.value = "DI"
		End If
		
		frm1.txtPlanAmtPR.Text = arrRet(8)
		frm1.txtPlanAmtPI.Text = arrRet(10)
		frm1.hPlanAmtPR.value = arrRet(8)
		frm1.hPlanAmtPI.value = arrRet(10)
		frm1.hPayPlanDt.value = arrRet(2)

		frm1.hXchRate.value = UNIFormatNumber(arrRet(24),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit)		'환율 
'		frm1.txtPlanLocAmtPR.Text = arrRet(9)
'		frm1.txtPlanLocAmtPI.Text = arrRet(11)
'		frm1.hPlanLocAmtPR.value = arrRet(9)
'		frm1.hPlanLocAmtPI.value = arrRet(11)
'		frm1.txtTotLocAmt.Text = UNICDbl(arrRet(9)) + UNICDbl(arrRet(11))
		frm1.txtLoanBalLocAmt.Text = arrRet(13)
		frm1.txtLoanLocAmt.Text = arrRet(19)
		frm1.txtRdpLocAmt.Text = arrRet(21)
		frm1.txtIntPayLocAmt.Text = arrRet(23)

		frm1.hInt_Pay_Perd.value = arrRet(28)
		frm1.hInt_Pay_Perd_Base.value = arrRet(29)
		If arrRet(27) = "BK" Then
			frm1.txtLoanPlcType1.checked = true
			frm1.txtLoanPlcType2.checked = false
		ElseIf arrRet(27) = "BP" Then
			frm1.txtLoanPlcType1.checked = false
			frm1.txtLoanPlcType2.checked = true
		End If
		
		frm1.hDay_Mthd.value = arrRet(30)
		frm1.txtIntAcctCd.value = arrRet(31)
	End If

	Call CurFormatNumericOCX()
	Call SetToolbar("1110100000001111")                                              '☆: Developer must customize

	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue = True
	frm1.txtLoanNo.focus
End Function

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim arrRet 
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtLoanNo.focus
End Function

'============================================================
'결의전표 팝업 
'============================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	frm1.txtLoanNo.focus
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'---------------->  from here , Condition area event of 3rd party control 
'=======================================================================================================
'   Event Name : _DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPayPlanDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPayPlanDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPayPlanDt.Focus
    End If
End Sub

'========================================================================================================
' Name : onChange
' Desc : developer describe this line
'========================================================================================================
Sub txtPayPlanDt_Change()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtPayPlanDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtPayPlanDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
						
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
			
		End If
	End If

    lgBlnFlgChgValue = True
End Sub

Sub txtXchRate_Change()
	If Trim(frm1.txtDocCur.value) <> "" Then
		frm1.txtPlanLocAmtPR.Text = "0"
		frm1.txtPlanLocAmtPI.Text = "0"
		frm1.txtEtcPayLoc.Text = "0"
		frm1.txtEtcBPPayLoc.value = "0"
	End If
	
	lgBlnFlgChgValue = True
End Sub

Sub txtBankAcctNo_onChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtDeptCd_onChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtDeptCd.value = "") Then		Exit sub
	If Trim(frm1.txtPayPlanDt.Text = "") Then	Exit sub
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtPayPlanDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
End Sub

Sub txtBankCd_onChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtPlanAmtPR_Change()

	If frm1.txtDocCur.value <> "" Then
		frm1.txtPlanLocAmtPR.Text = "0"
	End If

	Call sumPRPI()
End Sub

Sub txtPlanAmtPI_Change()
	If Trim(frm1.txtDocCur.value) <> "" Then
		frm1.txtPlanLocAmtPI.Text = "0"
	End If

	If Trim(frm1.hIntPayStnd.value) = "DI" AND UNICDbl(frm1.txtPlanAmtPI.Text) - UNICDbl(frm1.txtDfrIntPay.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtIntPayAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtIntPayAcctCd, "Q")
		frm1.txtIntPayAcctCd.value = ""
		frm1.txtIntPayAcctNm.value = ""
	End If

	Call sumPRPI()
End Sub

Sub txtPlanLocAmtPR_Change()
	Call sumPRPI()
End Sub

Sub txtPlanLocAmtPI_Change()
	Call sumPRPI()
End Sub

Sub txtRepayDesc_onChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtEtcPay_Change()
	If Trim(frm1.txtDocCur.value <> "") Then
		frm1.txtEtcPayLoc.Text = "0"
	End If

	If UNICDbl(frm1.txtEtcPay.Text) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")
		frm1.txtChargeAcctCd.value = ""
		frm1.txtChargeAcctNm.value = ""
	End If
    lgBlnFlgChgValue = True
End Sub

Sub txtEtcPayLoc_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtEtcBPPay_Change()
	If Trim(frm1.txtDocCur.value <> "") Then
		frm1.txtEtcBPPayLoc.Value = "0"
	End If

	If UNICDbl(frm1.txtEtcBPPay.value) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtEtcBPAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtEtcBPAcctCd, "Q")
		frm1.txtEtcBPAcctCd.value = ""
		frm1.txtEtcBPAcctNm.value = ""
	End If

    lgBlnFlgChgValue = True
End Sub

Sub txtEtcBPPayLoc_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtRcptAcctCd_onChange()
	frm1.txtRcptAcctNm.value = ""
    lgBlnFlgChgValue = True
End Sub

Sub txtIntPayAcctCd_onChange()
	frm1.txtIntPayAcctNm.value = ""
    lgBlnFlgChgValue = True
End Sub

Sub txtChargeAcctCd_onChange()
	frm1.txtChargeAcctNm.value = ""
    lgBlnFlgChgValue = True
End Sub

Sub txtEtcBPAcctCd_onChange()
	frm1.txtEtcBPAcctNm.value = ""
    lgBlnFlgChgValue = True
End Sub



'========================================================================================================
' Name : txtPayNo_KeyPress
' Desc : developer describe this line
'========================================================================================================
Sub txtPayNo_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then                                            ' 13 :Enter key
		Call MainQuery
	End if

End Sub

'---------------->  from here , Content area event of 3rd party control 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
				
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>
						<TABLE CELLSPACING=0 CELLPADDING=0 align=right>
							<TR>
								<td><A HREF="VBSCRIPT:OpenPopupLoan()">차입금참조</a> | 
									<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</a> |
									<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</a>						
								</td>
						    </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>상환번호</TD>
                                    <TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtPayNo"  SIZE=20 MAXLENGTH=18   TAG="12XXXU" ALT="상환번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="txtPayNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupPay()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
                                <TD CLASS=TD5 NOWRAP>차입금번호</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNo"  SIZE="18" TAG="24xxxU" ALT="차입금번호"></TD>
                                <TD CLASS=TD5 NOWRAP>차입내역</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanNm"  SIZE="30" TAG="24xxxU" ALT="차입내역"></OBJECT></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>차입일</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtLoanDt name=txtLoanDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24X" ALT="차입일"></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>상환만기일</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtDueDt name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24Z" ALT="상환만기일"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>차입처구분</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" tag="24xxxU" NAME=txtLoanPlcType ID=txtLoanPlcType1 Value = "BK"><LABEL FOR=Radio_Loan_fg0>은행</LABEL>&nbsp;
											 <INPUT TYPE="RADIO" CLASS="Radio" tag="24xxxU" NAME=txtLoanPlcType ID=txtLoanPlcType2 Value = "BP"><LABEL FOR=Radio_Loan_fg1>거래처</LABEL>&nbsp;</TD>
                                <TD CLASS=TD5 NOWRAP>차입처</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanPlcCd" SIZE="10" TAG="24xxxU" ALT="차입처"> <INPUT TYPE=TEXT NAME="txtLoanPlcNm"  SIZE=20 TAG="24xxxU" ALT="차입처명"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>장단기구분</TD>
                                <TD CLASS=TD6 NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" CLASS=FPDS140 TAG="24xxxU"><OPTION VALUE=""></OPTION></SELECT></TD>
                                <TD CLASS=TD5 NOWRAP>차입용도</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLoanType"  SIZE=10 TAG="24xxxU" ALT="차입용도"> <INPUT TYPE=TEXT NAME="txtLoanTypeNm"  SIZE=20 TAG="24xxxU" ALT="차입용도명"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>거래통화|환율</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur"  SIZE=10 TAG="24xxxZ" ALT="거래통화">&nbsp;
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hXchRate Name=hXchRate ALT="환율" align="top" CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X5Z"></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>이자율</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLoanIntRate Name=txtLoanIntRate ALT="이자율" CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X5Z"></OBJECT>');</SCRIPT>&nbsp;%</TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>차입금액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLoanAmt name=txtLoanAmt title=FPDOUBLESINGLE ALT="차입금액" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLoanLocAmt name=txtLoanLocAmt title=FPDOUBLESINGLE ALT="차입금액(자국)" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>원금상환총액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRdpAmt name=txtRdpAmt title=FPDOUBLESINGLE ALT="원금상환금액" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRdpLocAmt name=txtRdpLocAmt title=FPDOUBLESINGLE ALT="원금상환금액(자국)" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>차입잔액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLoanBalAmt name=txtLoanBalAmt title=FPDOUBLESINGLE ALT="차입잔액" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLoanBalLocAmt name=txtLoanBalLocAmt title=FPDOUBLESINGLE ALT="차입잔액(자국)" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>이자지급총액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIntPayAmt name=txtIntPayAmt title=FPDOUBLESINGLE ALT="이자지급총액" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIntPayLocAmt name=txtIntPayLocAmt title=FPDOUBLESINGLE ALT="이자지급총액(자국)" tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							</TR>
 							<TR>
                                <TD CLASS=TD5 NOWRAP>상환일자</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtPayPlanDt name=txtPayPlanDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22Z" ALT="상환일자"></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>환율</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtXchRate Name=txtXchRate ALT="환율" CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X5Z"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>부서</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT id=txtDeptCd NAME="txtDeptCd"  SIZE="10" MAXLENGTH="10"   TAG="22xxxU" ALT="부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="txtDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDeptCd.value,3)">
													<INPUT TYPE=TEXT NAME="txtDeptNm"  SIZE=17   TAG="24xxxU" ALT="부서명"></TD>
                                <TD CLASS=TD5 NOWRAP>출금유형</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRcptTypeCd"  SIZE="10" MAXLENGTH="10"   TAG="22xxxU" ALT="출금유형코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="txtRcptTypeCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptTypeCd.value,4)">
													<INPUT TYPE=TEXT NAME="txtRcptTypeNm"  SIZE=17 TAG="24xxxU" ALT="출금유형명"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>지급계좌번호</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBankAcctNo"  SIZE=20 MAXLENGTH=30   TAG="24xxxU" ALT="지급계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="txtDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.value,5)"></TD>
								<TD CLASS="TD5" NOWRAP>출금계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtRcptAcctCd" ALT="출금계정" SIZE="10" MAXLENGTH="20"  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptAcctCd.value, 7)">
													   <INPUT NAME="txtRcptAcctNm" ALT="출금계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>지급은행</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBankCd"  SIZE="10" MAXLENGTH="10"  tag="24XXXU" ALT="지급은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.value,2)">
                                					<INPUT TYPE=TEXT NAME="txtBankNm"  SIZE=20 TAG="24xxxU" ALT="지급은행명"></TD>
                                <TD CLASS=TD5 NOWRAP>이자지급형태</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" tag="24xxxU" NAME=txtIntPayStnd ID=txtIntPayStnd1 Value = "AI"><LABEL FOR=Radio_Loan_fg2>선급</LABEL>&nbsp;
											 <INPUT TYPE="RADIO" CLASS="Radio" tag="24xxxU" NAME=txtIntPayStnd ID=txtIntPayStnd2 Value = "DI"><LABEL FOR=Radio_Loan_fg3>후급</LABEL>&nbsp;</TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>원금상환액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPlanAmtPR name=txtPlanAmtPR ALT="원금상환액" title=FPDOUBLESINGLE tag="22X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPlanLocAmtPR name=txtPlanLocAmtPR ALT="원금상환액(자국)" title=FPDOUBLESINGLE tag="21X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
                                <TD CLASS=TD5 NOWRAP>이자지급액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPlanAmtPI name=txtPlanAmtPI ALT="이자지급액" title=FPDOUBLESINGLE tag="22X2" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPlanLocAmtPI name=txtPlanLocAmtPI ALT="이자지급액(자국)" title=FPDOUBLESINGLE tag="21X2" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>미지급이자액|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtDfrIntPay name=txtDfrIntPay ALT="미지급이자액" title=FPDOUBLESINGLE tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtDfrIntPayLoc name=txtDfrIntPayLoc ALT="미지급이자액(자국)" title=FPDOUBLESINGLE tag="24X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>이자비용계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtIntPayAcctCd" ALT="이자비용계정" SIZE="10" MAXLENGTH="20"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIntAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtIntPayAcctCd.value, 6)">
													   <INPUT NAME="txtIntPayAcctNm" ALT="이자비용계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>부대비용|자국</TD>
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEtcPay name=txtEtcPay ALT="부대비용" title=FPDOUBLESINGLE tag="21X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEtcPayLoc name=txtEtcPayLoc ALT="부대비용(자국)" title=FPDOUBLESINGLE tag="21X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>부대비용계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtChargeAcctCd" ALT="부대비용계정" SIZE="10" MAXLENGTH="20"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtChargeAcctCd.value, 8)">
													   <INPUT NAME="txtChargeAcctNm" ALT="부대비용계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>수수료|자국</TD> 
                                <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEtcBPPay name=txtEtcBPPay ALT="수수료" title=FPDOUBLESINGLE tag="21X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT>&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEtcBPPayLoc name=txtEtcBPPayLoc ALT="수수료(자국)" title=FPDOUBLESINGLE tag="21X2Z" CLASS=FPDS140></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>수수료계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEtcBPAcctCd" ALT="수수료계정" SIZE="10" MAXLENGTH="20"  tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEtcBPAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtEtcBPAcctCd.value, 9)">
													   <INPUT NAME="txtEtcBPAcctNm" ALT="수수료계정명" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사용자필드1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld1"  SIZE=40 MAXLENGTH="50" TAG="21xxx" ALT="사용자필드1"></TD>
								<TD CLASS=TD5 NOWRAP>사용자필드2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUserFld2"  SIZE=40 MAXLENGTH="50" TAG="21xxx" ALT="사용자필드2"></TD>
							</TR>
							<TR>
                                <TD CLASS=TD5 NOWRAP>비고</TD>
                                <TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRepayDesc"  SIZE=80 MAXLENGTH=128   TAG="21xxx" ALT="비고"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>

	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=hidden NAME="txtTempGlNo"  SIZE=20 TAG="24xxxU" ALT="결의전표번호">
<INPUT TYPE=hidden NAME="txtGlNo"  SIZE=20 TAG="24xxxU" ALT="회계전표번호">

<INPUT TYPE=HIDDEN NAME="txtIntAcctCd"			TAG="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"			tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hPayNo"				TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"				TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"			TAG="X4" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"			TAG="X4" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hPayPlanDt"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtPayPlanDt"			TAG="24" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hIntPayStnd"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hPlanAmtPR"			TAG="24" Tabindex="-1">	<!--원금 상환액(쿼리해온값)-->
<INPUT TYPE=HIDDEN NAME="hPlanAmtPI"			TAG="24" Tabindex="-1">	
<INPUT TYPE=HIDDEN NAME="hPlanLocAmtPR"			TAG="24" Tabindex="-1">	<!--원금 상환액(자국)-->
<INPUT TYPE=HIDDEN NAME="hPlanLocAmtPI"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hEtcPay"				TAG="24" Tabindex="-1">	<!--부대비용(쿼리해온값)-->
<INPUT TYPE=HIDDEN NAME="hEtcPayLoc"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hEtcBPPay"				TAG="24" Tabindex="-1">	<!--부대비용(쿼리해온값)-->
<INPUT TYPE=HIDDEN NAME="hEtcBPPayLoc"			TAG="24" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hlocPlanLocAmtPR"		TAG="24" Tabindex="-1">	<!--상환시 환율적용한 원금상환액 -->
<INPUT TYPE=HIDDEN NAME="hlocPlanLocAmtPI"		TAG="24" Tabindex="-1">	
<INPUT TYPE=HIDDEN NAME="hlocEtcPayLoc"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hlocEtcBPPayLoc"		TAG="24" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hInt_Pay_Perd"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hInt_Pay_Perd_Base"	TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hDay_Mthd"				TAG="24" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hInt_Pay_Dt"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hIntBaseMthd"			TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hPayObj"				TAG="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hConfFg"				TAG="24" Tabindex="-1">

<INPUT TYPE=HIDDEN NAME="hGlNo"					TAG="24" Tabindex="-1"><!--전표번호 -->
<INPUT TYPE=HIDDEN NAME="hTempGlNo"				TAG="24" Tabindex="-1"><!--결의전표번호 -->
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"			tag="24" Tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY
