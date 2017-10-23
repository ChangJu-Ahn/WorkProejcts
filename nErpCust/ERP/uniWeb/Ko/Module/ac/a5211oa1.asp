<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : A5211OA1
'*  4. Program Name         : 보조부출력 
'*  5. Program Desc         : Report of Subledger Detail
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/13
'*  8. Modified date(Last)  : 2004/01/12
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!-- '==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================= 
Dim IsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE											'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False													'Indicates that no value changed
End Sub

'========================================================================================
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

	EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtDateFr.Text = StartDate 
	frm1.txtDateTo.Text = EndDate 
End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "PA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()

End Sub

'========================================================================================
Sub SetSpreadLock()

End Sub

'========================================================================================
Sub SetSpreadColor(ByVal lRow)

End Sub

'========================================================================================
Function OpenPopUp(Byval param, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 6
			arrParam(0) = "사업장코드 팝업"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 													' TABLE 명칭 
			arrParam(2) = param															' Code Condition
			arrParam(3) = ""															' Name Cindition
			
			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"													' Field명(0)
			arrField(1) = "BIZ_AREA_NM"													' Field명(1)

			arrHeader(0) = "사업장코드"												' Header명(0)
			arrHeader(1) = "사업장명"												' Header명(1)
		Case 1, 2
			arrParam(0) = "계정코드 팝업"											' 팝업 명칭 
			arrParam(1) = " A_ACCT A, A_ACCT_GP B  "									' TABLE 명칭 
			arrParam(2) = Trim(param)													' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "ISNULL(A.SUBLEDGER_1,'') <> '' AND A.GP_CD=B.GP_CD"			' Where Condition
			
			If LTrim(RTrim(frm1.txtCtrlCd.value)) <> "" Then
				arrParam(4) = arrParam(4) & " AND ISNULL(A.SUBLEDGER_1,'') = " & FilterVar(frm1.txtCtrlCd.value, "''", "S")	' Where Condition
			End If

			arrParam(5) = "계정코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"													' Field명(0)
			arrField(1) = "A.ACCT_NM"													' Field명(1)
     		arrField(2) = "B.GP_CD"														' Field명(2)
			arrField(3) = "B.GP_NM"														' Field명(3)

			arrHeader(0) = "계정코드"												' Header명(0)
			arrHeader(1) = "계정명"													' Header명(1)
			arrHeader(2) = "그룹코드"												' Header명(2)
			arrHeader(3) = "그룹명"
		Case 3
			arrParam(0) = "보조부항목 팝업"											' 팝업 명칭 
			arrParam(1) = "A_CTRL_ITEM A"												' TABLE 명칭 
			arrParam(2) = Trim(param)													' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = "A.CTRL_CD in (select subledger_1 from A_Acct) "				' Where Condition
			arrParam(5) = "관리항목코드"											' 조건필드의 라벨 명칭 

			arrField(0) = "A.CTRL_CD"													' Field명(0)
			arrField(1) = "A.CTRL_NM"													' Field명(1)

			arrHeader(0) = "관리항목코드"											' Header명(0)
			arrHeader(1) = "관리항목명"												' Header명(1)
		Case 4
			arrParam(0) = Trim(frm1.txtCtrlNm.value)									' 팝업 명칭 
			arrParam(1) = Trim(frm1.hTblId.value) 
			arrParam(2) = ""															' Code Condition
			arrParam(3) = ""															' Name Condition

			arrParam(4) = Trim(frm1.hDataColmID.value) & _
					" in (select distinct CTRL_Val1 from A_SUBLEDGER_SUM where convert(datetime,fisc_yr+fisc_mnth+(case when fisc_dt in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ) then " & FilterVar("01", "''", "S") & "  else fisc_dt end),112) between '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"") & "' and '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"") & "'	 )"	' Where Condition

			arrParam(5) = Trim(frm1.txtCtrlNm.value)									' 조건필드의 라벨 명칭 

			arrField(0) = Trim(frm1.hDataColmID.value)									' Field명(0)
			arrField(1) = Trim(frm1.hDataColmNm.value)									' Field명(1)

			arrHeader(0) = Trim(frm1.hDataColmID.value)									' Header명(0)
			arrHeader(1) = Trim(frm1.hDataColmNm.value)									' Header명(1)
		Case 5
			arrParam(0) = Trim(frm1.txtCtrlNm.value)									' 팝업 명칭 
			arrParam(1) = "A_ACCT A,A_SUBLEDGER_SUM B"
			arrParam(2) = ""															' Code Condition
			arrParam(3) = ""															' Name Condition

			arrParam(4) = " A.SUBLEDGER_1 = " & FilterVar(frm1.txtCtrlCd.value, "''", "S")  & " and " & _
						" a.acct_cd = b.acct_cd and convert(datetime,b.fisc_yr+b.fisc_mnth+(case when b.fisc_dt in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ) then " & FilterVar("01", "''", "S") & "  else b.fisc_dt end),112) between '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"") & "' and '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"") & "'	 "	' Where Condition

			arrParam(5) = Trim(frm1.txtCtrlNm.value)									' 조건필드의 라벨 명칭 

			arrField(0) = "b.ctrl_val1"													' Field명(0)
			arrField(1) = ""

			arrHeader(0) = Trim(frm1.txtCtrlNm.value)									' Header명(0)
			arrHeader(1) = ""
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtBizAreaCd.focus
			Case 1
				frm1.txtAcctCdFr.focus
			Case 2
				frm1.txtAcctCdTo.focus
			Case 3
				frm1.txtCtrlCd.focus
			Case 4
				frm1.txtCtrlVal.focus
			Case 5
				frm1.txtCtrlVal.focus
			Case 6
				frm1.txtBizAreaCd1.focus
		End select
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If

End Function

'========================================================================================================= 
Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	Select Case iWhere
		Case 0	'사업장코드 
			frm1.txtBizAreaCd.value = focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 1	'시작계정코드 
			frm1.txtAcctCdFr.focus
			frm1.txtAcctCdFr.value = arrRet(0)
			frm1.txtAcctNmFr.value = arrRet(1)
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 2	'종료계정코드 
			frm1.txtAcctCdTo.focus
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 3
			frm1.txtCtrlCd.focus
			frm1.txtCtrlCd.value = arrRet(0)
			frm1.txtCtrlNm.value = arrRet(1)

			CtrlVal.innerHTML = frm1.txtCtrlNm.value 
			frm1.txtCtrlVal.value	= ""
			frm1.txtCtrlValNm.value	= ""
			
			Call ElementVisible(frm1.txtCtrlVal, 1)
			Call ElementVisible(frm1.txtCtrlValNm, 1)
			Call ElementVisible(frm1.btnCtrlVal, 1)
		Case 4
			frm1.txtCtrlVal.focus
			frm1.txtCtrlVal.value = arrRet(0)
			frm1.txtCtrlValNm.value = arrRet(1)	
		Case 5
			frm1.txtCtrlVal.focus
			frm1.txtCtrlVal.value = arrRet(0)
		Case 6	'사업장코드 
			frm1.txtBizAreaCd1.value = focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
		Case Else
			Exit Function
	End select	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function QueryCtrlVal()
    Dim ArrRet

    If frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	End If

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	ArrRet 	= Split(lgF0,Chr(11))

	If Trim(ArrRet(0)) <> "" Then
		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)

		Call OpenPopUp(0,4)
	Else
		Call OpenPopUp(0,5)
	End If
End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables
    Call SetDefaultVal
    Call SetToolBar("1000000000001111")
    
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


	frm1.txtCtrlCd.focus
	Call ElementVisible(frm1.txtCtrlVal, 0)
	Call ElementVisible(frm1.txtCtrlValNm, 0)
	Call ElementVisible(frm1.btnCtrlVal, 0)
End Sub

'========================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")	
        frm1.fpDateTime1.Focus
    End If
End Sub

'========================================================================================
Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
        Call SetFocusToDocument("M")	
        frm1.fpDateTime2.Focus
    End If
End Sub

'========================================================================================
Sub txtCtrlCd_OnBlur()
	Dim ArrRet
	Dim ArrParam(2)
  
	On Error Resume Next
	
    Call CommonQueryRs( "CTRL_CD,CTRL_NM" ,	"A_CTRL_ITEM", _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
	ArrRet 	= Split(lgF0,Chr(11))

	If ArrRet(0) = "" Then
		frm1.txtCtrlNm.value = ""
		CtrlVal.innerHTML = ""
		frm1.txtCtrlVal.value	= ""
		frm1.txtCtrlValNm.value	= ""		

		Call ElementVisible(frm1.txtCtrlVal, 0)
		Call ElementVisible(frm1.txtCtrlValNm, 0)
		Call ElementVisible(frm1.btnCtrlVal, 0)
		Exit Sub	
	End If

	ArrParam(0) = ArrRet(0)
	ArrRet 	= Split(lgF1,Chr(11))
	ArrParam(1) = ArrRet(0)

	frm1.txtCtrlCd.value = ArrParam(0)
	frm1.txtCtrlNm.value = ArrParam(1)

	CtrlVal.innerHTML = frm1.txtCtrlNm.value 
	frm1.txtCtrlVal.value	= ""
	frm1.txtCtrlValNm.value	= ""

	Call ElementVisible(frm1.txtCtrlVal, 1)
	Call ElementVisible(frm1.txtCtrlValNm, 1)
	Call ElementVisible(frm1.btnCtrlVal, 1)
End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd.value)
	If strCd = "" Then
		frm1.txtBizAreaNm.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm.value = ""
			IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm.value = ""
			else
				frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
	
End sub

'========================================================================================================
'   Event Name : txtBizAreaCd1_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd1_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd1.value)
	If strCd = "" Then
		frm1.txtBizAreaNm1.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm1.value = ""
			IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm1.value = ""
			else
				frm1.txtBizAreaNm1.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
 
End sub

'========================================================================================
Function SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
    Dim Fiscyyyy,Fiscmm,Fiscdd,DateFryyyy,DateFrmm,DateFrdd
	Dim IntRetCD
	
    Call ExtractDateFrom(Parent.gFiscStart,Parent.gServerDateFormat,Parent.gServerDateType,Fiscyyyy,Fiscmm,Fiscdd)
    Call ExtractDateFrom(frm1.txtDateFr.text,frm1.txtDateFr.UserDefinedFormat,Parent.gComDateType,DateFryyyy,DateFrmm,DateFrdd)

	If Fiscmm > DateFrmm Then
		Fiscyyyy = cstr(cint(DateFryyyy) - 1)
	Else
		Fiscyyyy	= DateFryyyy
	End If

	VarFiscDt = Fiscyyyy & Fiscmm & Fiscdd

	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"")
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"")

	VarCtrlCd	= "%"
	VarCtrlVal	= "%"

	VarAcctCdFr = "0"
	VarAcctCdTo = "zzzzzzzzzzzzzzzzzz"

	'VarBizAreaCd = "%"

	If Len(frm1.txtCtrlCd.value) > 0 Then 
		VarCtrlCd = Trim(frm1.txtCtrlCd.value)
	Else 
		frm1.txtCtrlNm.value = ""
	End If

	If Len(frm1.txtCtrlVal.value) > 0 Then 
		VarCtrlVal = Trim(frm1.txtCtrlVal.value)
	Else 
		frm1.txtCtrlValNm.value = ""
	End If

	If Len(frm1.txtAcctCdFr.value) > 0 Then 
		VarAcctCdFr = Trim(frm1.txtAcctCdFr.value)
	Else 
		frm1.txtAcctNmFr.value = ""
	End If

	If Len(frm1.txtAcctCdTo.value) > 0 Then 
		VarAcctCdTo = Trim(frm1.txtAcctCdTo.value)
	Else 
		frm1.txtAcctNmTo.value = ""
	End If

	If frm1.txtBizAreaCd.value = "" then 
		frm1.txtBizAreaNm.value = ""
		If lgAuthBizAreaCd <> "" Then			
			VarBizAreaCd  = lgAuthBizAreaCd
		Else
			VarBizAreaCd = "0"
		End If			
	Else 
		If lgAuthBizAreaCd <> "" Then
			VarBizAreaCd = Trim(FilterVar(frm1.txtBizAreaCD.value,"","SNM"))
			If UCASE(lgAuthBizAreaCd) <> UCASE(VarBizAreaCd) Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				frm1.txtBizAreaCD.focus()
				SetPrintCond =  False
				Exit Function
			End If
		Else
			VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
		End If
	End if

	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		If lgAuthBizAreaCd <> "" Then			
			VarBizAreaCd1 = lgAuthBizAreaCd
		Else
			VarBizAreaCd1 = "ZZZZZZZZZZ"
		End If			
	Else 
		If lgAuthBizAreaCd <> "" Then
			VarBizAreaCd1 = Trim(FilterVar(frm1.txtBizAreaCD1.value,"","SNM"))
			If UCASE(lgAuthBizAreaCd) <> UCASE(VarBizAreaCd1) Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				frm1.txtBizAreaCD1.focus()
				SetPrintCond =  False
				Exit Function
			End If
		Else
			VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
		End If
	End if
'msgbox VarBizAreaCd & "**" & VarBizAreaCd1

	StrEbrFile = "a5211ma2"

'	if frm1.PrintOpt1.checked = True and frm1.txtBizAreaCd.value <> "" then
'		StrEbrFile = "a5211ma1.ebr"
'	elseif  frm1.PrintOpt1.checked = True and frm1.txtBizAreaCd.value = "" then
'	StrEbrFile = "a5211ma2.ebr"
'	end if

	SetPrintCond =  True
	
End Function

'==========================================================================================
Function CompareAcctCdByDB(ByVal FromCd , ByVal ToCd)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareAcctCdByDB = False

    If FromCd.value <> "" And ToCd.value <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareAcctCdByDB = True
End Function

'========================================================================================
Function FncBtnPrint() 
    Dim StrUrl
    Dim StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal
    Dim IntRetCD
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtDateFr.Text,frm1.txtDateTo.Text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,Parent.gComDateType,True) = False Then		
		Exit Function
    End If

    If CompareAcctCdByDB(frm1.txtAcctCdFr,frm1.txtAcctCdTo) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
        frm1.txtAcctCdFr.focus
		Exit Function
	End If		
	
	'회계일자 조회기간 Check
'	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
'		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
'		frm1.txtDateFr.focus
'		Exit Function
'	End If
	
	'계정코드 조회조건 Check
'	If Trim(frm1.txtAcctCdFr.value) > Trim(frm1.txtAcctCdTo.value) Then
'		Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
'		frm1.txtAcctCdFr.focus
'		Exit Function
'	End If
	
	IntRetCD = 	SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
	If IntRetCD = False Then
	    Exit Function
 	End If
 	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr					' '|'-> ebr파일을 부를때 사용되는 구분자.(url에서 ?뒤에파라메터로 붙여주는 것이라고 보면 됨.
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
	StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	StrUrl = StrUrl & "|Currency|" & Parent.gCurrency
	StrUrl = StrUrl & "|AcctCdFr|" & VarAcctCdFr
	StrUrl = StrUrl & "|AcctCdTo|" & VarAcctCdTo
	
	StrUrl = StrUrl & "|CtrlCd|" & VarCtrlCd
	StrUrl = StrUrl & "|CtrlVal|" & VarCtrlVal

	Call FncEBRPrint(EBAction,ObjName,StrUrl)
End Function

'========================================================================================
Function FncBtnPreview() 
    Dim StrUrl
    Dim StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal
    Dim IntRetCD
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtDateFr.Text,frm1.txtDateTo.Text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
                        "970025",frm1.txtDateFr.UserDefinedFormat,Parent.gComDateType,True) = False Then		
		Exit Function
    End If

    If CompareAcctCdByDB(frm1.txtAcctCdFr,frm1.txtAcctCdTo) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
        frm1.txtAcctCdFr.focus
		Exit Function
	End If		

	'회계일자 조회기간 Check
'	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
'		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
'		frm1.txtDateFr.focus
'		Exit Function
'	End If
	
	'계정코드 조회조건 Check
'	If Trim(frm1.txtAcctCdFr.value) > Trim(frm1.txtAcctCdTo.value) Then
'		Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
'		frm1.txtAcctCdFr.focus
'		Exit Function
'	End If
	
	
	IntRetCD = 	SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
	If IntRetCD = False Then Exit Function

	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
	StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	StrUrl = StrUrl & "|Currency|" & Parent.gCurrency
	StrUrl = StrUrl & "|AcctCdFr|" & VarAcctCdFr
	StrUrl = StrUrl & "|AcctCdTo|" & VarAcctCdTo
	
	StrUrl = StrUrl & "|CtrlCd|" & VarCtrlCd
	StrUrl = StrUrl & "|CtrlVal|" & VarCtrlVal

	Call FncEBRPreview(ObjName,StrUrl)
End Function

'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
End Function

'========================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
    FncExit = True
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>

<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=시작회계일자 id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=종료회계일자 id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="CtrlCd" NOWRAP>보조부항목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="보조부항목" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCtrlCd.Value,3)"> <INPUT TYPE="Text" NAME="txtCtrlNm" SIZE=25 tag="14X" ALT="보조부항목명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="CtrlVal" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlVal" SIZE=10 MAXLENGTH=30 tag="11XXXU" ALT="" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlVal" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call QueryCtrlVal()"> <INPUT TYPE="Text" NAME="txtCtrlValNm" SIZE=25 tag="14X" ALT=""></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="FrAcct" NOWRAP>시작계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdFr" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="시작계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDFr.Value,1)"> <INPUT TYPE="Text" NAME="txtAcctNmFr" SIZE=25 tag="14X" ALT="시작계정명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="ToAcct" NOWRAP>종료계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdTo" SIZE=10 MAXLENGTH=20 tag="11XXUX" ALT="종료계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDTo.Value,2)"> <INPUT TYPE="Text" NAME="txtAcctNmTo" SIZE=25 tag="14X" ALT="종료계정명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,6)"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hTblId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmID" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmNm" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

