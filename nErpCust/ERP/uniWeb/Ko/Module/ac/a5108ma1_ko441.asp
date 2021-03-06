<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5108MA1
'*  4. Program Name         : 
'*  5. Program Desc         : Ado query Sample with DBAgent(Sort)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
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
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">				  </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'========================================================================================================

Const BIZ_PGM_ID 		= "a5108mb1_ko441.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
Const C_MaxKey          = 0					                          '☆: SpreadSheet의 키의 갯수 


Const C_ThisLeftAmt		= 3
Const C_ThisRightAmt	= 4

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim lgMaxFieldCount
Dim lgCookValue 

Dim lgFiscStart
Dim lgFromGlDt
Dim lgToGlDt

Dim lgSaveRow 

Dim strSp_Id

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================
Sub SetDefaultVal()
	
	Dim FromGlDate

	FromGlDate = UNIDateAdd("m", 1, Parent.gFiscStart,Parent.gServerDateFormat)
	FromGlDate = UNIDateAdd("d", -1, FromGlDate,Parent.gServerDateFormat)


	frm1.txtFromGlDt.Text		= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat)
	'frm1.txtToGlDt.Text			= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat) 
	'frm1.txtToGlDt.Text			= UniConvDateAToB(FromGlDate ,Parent.gServerDateFormat,Parent.gDateFormat) 
    frm1.txtToGlDt.Text			= UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
End Sub

'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>
    <% Call LoadBNumericFormatA("Q", "A", "COOKIE", "QA") %>
End Sub


'========================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, Parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)

       Call MainQuery()

       WriteCookie CookieSplit , ""

	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF

End Function

'========================================================================================
Sub InitComboBox()
End Sub

'========================================================================================
Sub InitSpreadSheet()
'msgbox "InitSpreadSheet"
    Call SetZAdoSpreadSheet("A5108MA1_GRD01_KO441", "S", "A", "V20021211", parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
    End With
End Sub



'========================================================================================
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")
    Call CookiePage(0)
    frm1.txtFromGlDt.focus
    
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

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================
Function FncQuery() 
	Dim IntRetCD

    FncQuery = False    
    Err.Clear

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtFromGlDt.Text,frm1.txtToGlDt.Text,frm1.txtFromGlDt.Alt,frm1.txtToGlDt.Alt, _
	 "970024", frm1.txtFromGlDt.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtToGlDt.Focus
		Exit Function
	End If


    
    If frm1.txtBizArea1Cd.value = "" Then frm1.txtBizAreaNm1.value = "" End If
    If frm1.txtBizArea2Cd.value = "" Then frm1.txtBizAreaNm2.value = "" End If
    If frm1.txtBizArea3Cd.value = "" Then frm1.txtBizAreaNm3.value = "" End If
    If frm1.txtBizArea4Cd.value = "" Then frm1.txtBizAreaNm4.value = "" End If
    If frm1.txtBizArea5Cd.value = "" Then frm1.txtBizAreaNm5.value = "" End If                
'    If (frm1.txtBizArea1Cd.value = "" ) And (frm1.txtBizArea2Cd.value = "") And (frm1.txtBizArea3Cd.value = "") And (frm1.txtBizArea4Cd.value = "") And (frm1.txtBizArea5Cd.value = "") Then
'	    Call DisplayMsgBox("169803","X","X","X")
'		frm1.txtBizArea1Cd.Focus
'		Exit Function                
'	End If
	
	
    If frm1.txtClassType.value <> "" Then
   		IntRetCD = CommonQueryRs(" CLASS_TYPE_NM, CLASS_TYPE"," A_ACCT_CLASS_TYPE ","  CLASS_TYPE = " & FilterVar(frm1.txtClassType.Value, "''", "S") & " and CLASS_TYPE Like " & FilterVar("BS%", "''", "S") & "  " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		If IntRetCD = False  Then
		    Call DisplayMsgBox("110500","X","X","X")
			frm1.txtClassType.Focus
			Exit Function
		End If
	End If

    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True
End Function



'========================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear

	Call Parent.FncPrint()
    FncPrint = True
End Function


'========================================================================================
Function FncExcel()
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(Parent.C_MULTI)
    FncExcel = True
End Function

'========================================================================================
Function FncFind()
    FncFind = False 
    Err.Clear
	Call Parent.FncFind(Parent.C_MULTI, True)
    FncFind = True 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub


'========================================================================================
Function FncExit()
    FncExit = False
    Err.Clear
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal, strZeroFg
    Dim strYYYY, strMM, strDD    

    Err.Clear
'msgbox "DbQuery"
    DbQuery = False
    Call GetQueryDate()
	Call LayerShowHide(1)

    With frm1

		if .ZeroFg1.checked = True Then
			strZeroFg = "Y"
		Else
			strZeroFg = "N"
		End IF

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			
			strSp_Id	= ""

			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt'.txtFromGlDt.Year	& Right("0" & frm1.txtFromGlDt.Month,2)		& Right("0" & frm1.txtFromGlDt.Day,2)			
			strVal = strVal & "&txtToGlDt="      & lgToGlDt'.txtToGlDt.Year		& Right("0" & frm1.txtToGlDt.Month,2)		& Right("0" & frm1.txtToGlDt.Day,2)				
			strVal = strval & "&txtClassType="   & .txtClassType.value 
			
			strVal = strVal & "&txtBizArea1Cd="	 & .txtBizArea1Cd.value
			strVal = strVal & "&txtBizArea2Cd="	 & .txtBizArea2Cd.value
			strVal = strVal & "&txtBizArea3Cd="	 & .txtBizArea3Cd.value
			strVal = strVal & "&txtBizArea4Cd="	 & .txtBizArea4Cd.value
			strVal = strVal & "&txtBizArea5Cd="	 & .txtBizArea5Cd.value												
			
			strVal = strVal & "&strZeroFg="		 & strZeroFg
        	strVal = strVal & "&strUserId="		 & Parent.gUsrID
        Else
			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt
        End If

    '--------- Developer Coding Part (End) ------------------------------------------------------------

        strVal = strVal & "&lgSp_Id="       & strSp_Id
        strVal = strVal & "&lgPageNo="       & lgPageNo
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectLIstDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd		' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd			' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd		' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID			' 개인 
'msgbox strVal
        Call RunMyBizASP(MyBizASP, strVal)	
    End With
'msgbox "DbQuery-1"
    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()

	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE
    lgSaveRow        = 1
    Call SetToolBar("1100000000011111")	
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function

'############################################################################################################
'-------------------------------------  OpenBizAreaCd1()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea1Cd.Value)	' Code Condition
	arrParam(3) = ""
								' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea1Cd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,11)
	End If

End Function




'-------------------------------------  OpenBizAreaCd2()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea2Cd.Value)	' Code Condition
	arrParam(3) = ""
								' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea2Cd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,12)
	End If

End Function



'-------------------------------------  OpenBizAreaCd3()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd3()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea3Cd.Value)	' Code Condition
	arrParam(3) = ""
								' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea3Cd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,13)
	End If

End Function



'-------------------------------------  OpenBizAreaCd4()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd4()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea4Cd.Value)	' Code Condition
	arrParam(3) = ""
								' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea4Cd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,14)
	End If

End Function


'-------------------------------------  OpenBizAreaCd5()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd5()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea5Cd.Value)	' Code Condition
	arrParam(3) = ""
								' Name Cindition
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea5Cd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,15)
	End If

End Function


'###################################################################################################################




'-------------------------------------  OpenClassTypeCd()  -----------------------------------------------
'	Name : OpenClassTypeCd()
'	Description : Acct Class Type PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenClassTypeCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "재무제표유형 팝업"			' 팝업 명칭 
	arrParam(1) = "A_ACCT_CLASS_TYPE"			' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtClassType.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("BS%", "''", "S") & " "		' Where Condition
	arrParam(5) = "재무제표코드"

    arrField(0) = "CLASS_TYPE"					' Field명(0)
    arrField(1) = "CLASS_TYPE_NM"				' Field명(1)

    arrHeader(0) = "재무제표코드"			' Header명(0)
	arrHeader(1) = "재무제표명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtClassType.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If

End Function


'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
			case 11
				.txtBizArea1Cd.focus
				.txtBizArea1Cd.Value	= arrRet(0)
				.txtBizAreaNm1.Value	= arrRet(1)
			case 12
				.txtBizArea2Cd.focus
				.txtBizArea2Cd.Value	= arrRet(0)
				.txtBizAreaNm2.Value	= arrRet(1)
			case 13
				.txtBizArea3Cd.focus
				.txtBizArea3Cd.Value	= arrRet(0)
				.txtBizAreaNm3.Value	= arrRet(1)
			case 14
				.txtBizArea4Cd.focus
				.txtBizArea4Cd.Value	= arrRet(0)
				.txtBizAreaNm4.Value	= arrRet(1)
			case 15
				.txtBizArea5Cd.focus
				.txtBizArea5Cd.Value	= arrRet(0)
				.txtBizAreaNm5.Value	= arrRet(1)																
			case 2
				.txtClassType.focus
				.txtClassType.Value	= arrRet(0)
				.txtClassTypeNm.Value	= arrRet(1)
		End select
	End With

End Function


''========================================================================================================
''	Name : SetPrintCond()
''	Description : Group Condition PopUp
''========================================================================================================
'Function SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
'
'	StrEbrFile = "A5108MA1"
'
'	' 권한관리 추가 
'	Dim IntRetCD
'	
'	varBizArea = UCASE(Trim(frm1.txtBizArea1.value))
'
'	If varBizArea = "" Then
'		If lgAuthBizAreaCd <> "" Then			
'			varBizArea  = lgAuthBizAreaCd
'		Else
'			varBizArea = "*"
'		End If			
'	Else
'		If lgAuthBizAreaCd <> "" Then			
'			If UCASE(lgAuthBizAreaCd) <> varBizArea Then
'				IntRetCD = DisplayMsgBox("124200","x","x","x")
'				SetPrintCond =  False
'				Exit Function
'			End If			
'		End If			
'	End If
'
'	varFromDt		 = lgFromGlDt
'	varToDt			 = lgToGlDt
'
'	SetPrintCond =  True
'
'End Function    

''========================================================================================
'Function BtnPreview()
'	Dim StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea, VarClassType
'	Dim StrUrl
'	Dim lngPos
'	Dim intCnt,IntRetCD
'	Dim arrParam, arrField, arrHeader
'
'	if lgIntFlgMode <> Parent.OPMD_UMODE then
'		IntRetCD = DisplayMsgBox("900002","x","x","x")   ' 조회를 먼저 하십시오.	
'		Exit Function
'	end if		
'
'    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
'       Exit Function
'    End If
'
'	if frm1.vspddata.MaxRows < 1 then
'		IntRetCD = DisplayMsgBox("900014","x","x","x")
'		Exit Function
'	end if
'
'    IntRetCD = SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
'	If IntRetCD = False Then
'	    Exit Function
' 	End If
'
'    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
'
'    lngPos = 0
'
'	For intCnt = 1 To 3
'	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
'	Next
'
'	StrUrl = StrUrl & "varFromDt|"	& varFromDt
'	StrUrl = StrUrl & "|varToDt|"			& varToDt
'	StrUrl = StrUrl & "|varPreFromDt|"      & varPreFromDt
'	StrUrl = StrUrl & "|varPreToDt|"		& varPreToDt
'	StrUrl = StrUrl & "|BizAreaCd|"			& varBizArea
'	StrUrl = StrUrl & "|strSp_Id|"			& strSp_Id
'
'	Call FncEBRPreview(ObjName, StrUrl)
'
'End Function
'
'
'Function BtnPrint()
'	Dim IntRetCD,intCnt	
'	Dim StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea, VarClassType
'	Dim StrUrl
'	Dim lngPos
'
'	if lgIntFlgMode = Parent.OPMD_UMODE then
'		IntRetCD = DisplayMsgBox("900019", Parent.VB_YES_NO,"x","x")
'		If IntRetCD = vbNo Then	Exit Function
'	else
'		IntRetCD = DisplayMsgBox("900002","x","x","x")
'		 Exit Function
'	end if
'
'    If Not chkField(Document, "1") Then	
'       Exit Function
'    End If
'
'	if frm1.vspddata.MaxRows < 1 then
'		IntRetCD = DisplayMsgBox("900014","x","x","x")
'		Exit Function
'	end if
'
'    IntRetCD = SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
'	If IntRetCD = False Then
'	    Exit Function
' 	End If
'
'    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
'
'    lngPos = 0
'
'	For intCnt = 1 To 3
'	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
'	Next
'
'	StrUrl = StrUrl & "varFromDt|" & varFromDt
'	StrUrl = StrUrl & "|varToDt|"			& varToDt
'	StrUrl = StrUrl & "|varPreFromDt|"      & varPreFromDt
'	StrUrl = StrUrl & "|varPreToDt|"		& varPreToDt
'	StrUrl = StrUrl & "|BizAreaCd|"			& varBizArea
'	StrUrl = StrUrl & "|strSp_Id|"			& strSp_Id
'
'	Call FncEBRPrint(EBAction,ObjName,StrUrl)
'
'End Function

'========================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If
End Function


'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
	
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
'	Dim ii

	Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If
    End If

	lgCookValue = ""
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

'========================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

	
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If

End Sub



'=======================================================================================================
'   Event Name : txtFromGlDt_KeyDown(KeyCode, Shift)
'   Event Desc :
'=======================================================================================================
Sub txtFromGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub
'=======================================================================================================
'   Event Name : txtToGlDt_KeyDown(KeyCode, Shift)
'   Event Desc :
'=======================================================================================================
Sub txtToGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


'=======================================================================================================
'   Event Name : txtFromGlDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromGlDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToGlDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToGlDt.Focus
    End If
End Sub





'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetFloatByCellOfCur C_ItemAmt,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec	
	End With
End Sub

'========================================================================================================
'   Event Name : GetQueryDate()
'   Event Desc : 
'========================================================================================================
Sub GetQueryDate()

	Dim strFromYYYY, strFromMM, strFromDD
	Dim strToYYYY, strToMM, strToDD

	Call ExtractDateFrom(frm1.txtFromGlDt.text,		Parent.gDateFormat,	Parent.gComDateType,	strFromYYYY,	strFromMM,		strFromDD)
	Call ExtractDateFrom(frm1.txtToGlDt.text,		Parent.gDateFormat,	Parent.gComDateType,	strToYYYY,		strToMM,		strToDD)


	lgFiscStart		= GetFiscDate(frm1.txtFromGlDt.Text)
	lgFromGlDt		= strFromYYYY		& strFromMM			& strFromDD
	lgToGlDt		= strToYYYY			& strToMM			& strToDD

End Sub

'========================================================================================================
'   Event Name : GetFiscDate()
'   Event Desc : 
'========================================================================================================
Function GetFiscDate( ByVal strFromDate)

	Dim strFiscYYYY, strFiscMM, strFiscDD
	Dim strFromYYYY, strFromMM, strFromDD

	GetFiscDate				="19000101"	

	Call ExtractDateFrom(Parent.gFiscStart,	Parent.gServerDateFormat,	Parent.gServerDateType,	strFiscYYYY,	strFiscMM,	strFiscDD)
	Call ExtractDateFrom(strFromDate,	Parent.gDateFormat,		Parent.gComDateType,		strFromYYYY,	strFromMM,	strFromDD)

	strFiscYYYY =  strFromYYYY

	If isnumeric(strFromYYYY) And isnumeric(strFromMM) And isnumeric(strFiscMM) Then

		If Cint(strFiscMM) > Cint(strFromMM)  then
		   GetFiscDate	= Cstr(Cint(strFromYYYY) - 1) & strFiscMM & strFiscDD
		Else
		   GetFiscDate	= strFromYYYY & strFiscMM & strFiscDD
		End If

	End If

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일(당기)</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일(당기)" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
												           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일(당기)" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									</TD>
									<TD CLASS="TD5">재무제표코드</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtClassType"   NAME="txtClassType"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="재무제표코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenClassTypeCd()"> <INPUT TYPE=TEXT ID="txtClassTypeNm" NAME="txtClassTypeNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizArea1"   NAME="txtBizArea1Cd"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd1()"> <INPUT TYPE=TEXT ID="txtBizAreaNm1" NAME="txtBizAreaNm1" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
									<TD CLASS="TD5">사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizArea2"   NAME="txtBizArea2Cd"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd2()"> <INPUT TYPE=TEXT ID="txtBizAreaNm2" NAME="txtBizAreaNm2" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizArea3"   NAME="txtBizArea3Cd"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd3()"> <INPUT TYPE=TEXT ID="txtBizAreaNm3" NAME="txtBizAreaNm3" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
									<TD CLASS="TD5">사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizArea4"   NAME="txtBizArea4Cd"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd4()"> <INPUT TYPE=TEXT ID="txtBizAreaNm4" NAME="txtBizAreaNm4" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizArea5"   NAME="txtBizArea5Cd"   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizAreaCd5()"> <INPUT TYPE=TEXT ID="txtBizAreaNm5" NAME="txtBizAreaNm5" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
									<TD CLASS="TD5" NOWRAP>조회구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" ID="ZeroFg1" VALUE="Y" tag="15"><LABEL FOR="ZeroFg1">전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" CHECKED ID="ZeroFg2" VALUE="N" tag="15"><LABEL FOR="ZeroFg2">발생금액</LABEL></SPAN></TD>
								</TR>																
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
<!--	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>-->
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>		
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"   TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"  TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date"     TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

