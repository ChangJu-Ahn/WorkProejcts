<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5112MA1
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
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">				  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/JpQuery.vbs">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================

Const BIZ_PGM_ID 		= "a5112MB1.asp"
Const BIZ_PGM_ID_SP 	= "a5112MB2.asp"


'========================================================================================
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 

Dim C_ListSeq
Dim C_LBal
Dim C_LSum
Dim C_LThis
Dim C_Result 
Dim C_RThis
Dim C_RSum
Dim C_RBal

'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================
Dim lgIsOpenPop


Dim lgMaxFieldCount


Dim lgCookValue

Dim lgFiscStart
Dim lgStartDt
Dim lgEndDt

Dim lgSaveRow 

Dim strSp_Id

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

	frm1.hSum.value = "합계"
	frm1.hUnBalance.value = "대차착오"
End Sub

'========================================================================================
Sub SetDefaultVal()
	frm1.txtStartDT.Text	= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtEndDT.Text		= UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>
End Sub


'========================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 0 Then
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
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Function

'========================================================================================
Sub InitComboBox()

End Sub

'========================================================================================
Sub initSpreadPosVariables()  
	 C_ListSeq= 1
	 C_LBal   = 2
	 C_LSum   = 3
	 C_LThis  = 4
	 C_Result = 5
	 C_RThis  = 6
	 C_RSum   = 7
	 C_RBal   = 8
End Sub

'========================================================================================
Sub InitSpreadSheet()


	Call SetZAdoSpreadSheet("A5112MA1_GRD01","S","A","V20060620",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	
	Call initSpreadPosVariables()
    With frm1.vspdData2
        .ReDraw = False

         .MaxRows = 1
         .MaxCols = 9
         .Col = .MaxCols             '☜: 공통콘트롤 사용 Hidden Column
         .ColHidden = True

         .RowHeaderDisplay = 0
         .Row = 0
        .RowHidden = True

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021220",,parent.gAllowDragDropSpread

        ggoSpread.ClearSpreadData

		Call GetSpreadColumnPos("B")

         ggoSpread.SSSetEdit C_ListSeq, "", 15
         ggoSpread.SSSetFloat C_LBal, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_LSum, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_LThis, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetEdit C_Result, "", 25, 2, , 40
         ggoSpread.SSSetFloat C_RThis, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_RSum, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_RBal, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

         .ScrollBars = 0     'ScrollBarsNone

		Call ggoSpread.SSSetColHidden(C_ListSeq ,C_ListSeq	,True)
			
         .ReDraw = True
    End With

	Call SetSpreadLock

End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = True
    End With
    
    With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = True
    End With

    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ListSeq			= iCurColumnPos(1)
			C_LBal   		    = iCurColumnPos(2)
			C_LSum   		    = iCurColumnPos(3)
			C_LThis  		    = iCurColumnPos(4)
			C_Result 		    = iCurColumnPos(5)
			C_RThis  			= iCurColumnPos(6)
			C_RSum   		    = iCurColumnPos(7)
			C_RBal   			= iCurColumnPos(8)
    End Select
End Sub


'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
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
    Call InitComboBox()
    Call CookiePage(0)
    
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
    
    frm1.txtStartDT.focus
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

	If CompareDateByFormat(frm1.txtStartDt.text,frm1.txtEndDt.text,frm1.txtStartDt.Alt,frm1.txtEndDt.Alt,"970025", _
		frm1.txtStartDt.UserDefinedFormat,Parent.gComDateType,True) = False Then
		Exit Function
	End If

   IF frm1.PrintOpt1.value = "1" Then
	If frm1.txtClassType.value <> "" Then
		IntRetCD = CommonQueryRs(" CLASS_TYPE_NM, CLASS_TYPE"," A_ACCT_CLASS_TYPE ","  CLASS_TYPE = " & FilterVar(frm1.txtClassType.Value, "''", "S") & " and CLASS_TYPE Like " & FilterVar("TB%", "''", "S") & "  " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		If IntRetCD = False  Then
		    Call DisplayMsgBox("110500","X","X","X")
			frm1.txtClassType.Focus
			Exit Function
		End If
	End If
   End If

    '-----------------------
    'Query function call area
    '-----------------------
	Call ClearVspdData2()
    If DbQuery = False Then Exit Function

    FncQuery = True

End Function

'========================================================================================
Function FncNew()
    Dim IntRetCD
    FncNew = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncNew = True
End Function
	
'========================================================================================
Function FncDelete()
    Dim intRetCD
    FncDelete = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncDelete = True
End Function


'========================================================================================
Function FncSave()
    Dim IntRetCD
    FncSave = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncSave = True
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
	Dim strValSp, strZeroFg

    Err.Clear
    DbQuery = False

    Call GetQueryDate()
	Call LayerShowHide(1)

	With frm1

		if .PrintOpt1.checked = True Then
			.txtPrintOpt.value = "1"
		ElseIf .PrintOpt2.checked =  True Then
			.txtPrintOpt.value = "2"
		Elseif .PrintOpt3.checked = True Then
			.txtPrintOpt.value = "3"
		End IF

		if .ZeroFg1.checked = True Then
			strZeroFg = "Y"
		Else
			strZeroFg = "N"
		End IF

		strValSp	= BIZ_PGM_ID_SP

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			
			strSp_Id	= ""
			'sp를 호출한다.        				
			strValSp = strValSp & "?lgFiscStart="	& Trim(lgFiscStart)
			strValSp = strValSp & "&lgStartDt="     & Trim(lgStartDt)
			strValSp = strValSp & "&lgEndDt="       & Trim(lgEndDt)
        	strValSp = strValSp & "&txtClassType=" & Trim(.txtClassType.value)
        	strValSp = strValSp & "&txtBizArea="	& Trim(.txtBizArea.value)
        	strValSp = strValSp & "&strHqBrchFg="   & "N"
        	strValSp = strValSp & "&strZeroFg="		& strZeroFg
        	strValSp = strValSp & "&txtPrintOpt="   & Trim(.txtPrintOpt.value)
        	strValSp = strValSp & "&strUserId="		& Parent.gUsrID

			' 권한관리 추가 
			strValSp = strValSp & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
			strValSp = strValSp & "&lgInternalCd="		& lgInternalCd				' 내부부서 
			strValSp = strValSp & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
			strValSp = strValSp & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

        	Call RunMyBizASP(MyBizASP, strValSp)

        End If

    '--------- Developer Coding Part (End) ------------------------------------------------------------

    End With

    DbQuery = True

End Function



'========================================================================================
Function DbQuery2()
	Dim strVal

    Err.Clear
    DbQuery2 = False

	Call LayerShowHide(1)

	With frm1
        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtClassType="   & .txtClassType.value
           	strVal = strVal & "&txtBizArea="	 & .txtBizArea.value
       Else
			strVal = strVal & "?txtClassType="   & .htxtClassType.value
           	strVal = strVal & "&txtBizArea="	 & .htxtBizArea.value
        End If   

        strVal = strVal & "&strSp_Id="	 & strSp_Id

		strVal = strVal & "&lgPageNo="       & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

       Call RunMyBizASP(MyBizASP, strVal)
    '--------- Developer Coding Part (End) ------------------------------------------------------------        

    End With

    DbQuery2 = True

End Function


'========================================================================================
Function DbQueryOk()
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
	Call DbQuery2()
End Function

'========================================================================================
Function DbQuery2Ok()
	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    Call SetToolBar("1100000000011111")	
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
dim strgChangeOrgId

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"					' Header명(1)

		Case 1
			arrParam(0) = "재무제표코드팝업"			' 팝업 명칭 
			arrParam(1) = "A_ACCT_CLASS_TYPE" 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("TB%", "''", "S") & " "		' Where Condition
			arrParam(5) = "재무제표코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "CLASS_TYPE"					' Field명(0)
			arrField(1) = "CLASS_TYPE_NM"				' Field명(1)

			arrHeader(0) = "재무제표코드"				' Header명(0)
			arrHeader(1) = "재무제표명"				' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtBizArea.focus
			Case 1
				frm1.txtClassType.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If

End Function

'========================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		'BIZ_AREA
				.txtBizArea.focus
				.txtBizArea.value		= arrRet(0)
				.txtBizAreaNm.value		= arrRet(1)
			Case 1	
				.txtClassType.focus
				.txtClassType.value		= arrRet(0)
				.txtClassTypeNm.value	= arrRet(1)
		End Select

		'lgBlnFlgChgValue = True
	End With
End Function


'========================================================================================
Function SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)

	StrEbrFile = "a5112ma1"

	' 권한관리 추가 
	Dim IntRetCD
	
	varBizArea = UCASE(Trim(frm1.txtBizArea.value))

	If varBizArea = "" Then
		If lgAuthBizAreaCd <> "" Then			
			varBizArea  = lgAuthBizAreaCd
		Else
			varBizArea = "*"
		End If			
	Else
		If lgAuthBizAreaCd <> "" Then			
			If UCASE(lgAuthBizAreaCd) <> varBizArea Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				SetPrintCond =  False
				Exit Function
			End If			
		End If			
	End If

	ClassType	= frm1.txtClassType.value
	varString2	= frm1.hSum.value
	varString3	= frm1.hUnBalance.value

'	당기시작일은 DB(AP)Server Format의 날짜이다.
	varFiscStartDt	= lgFiscStart
	varFromDt		= lgStartDt	
	varToDt			= lgEndDt

	SetPrintCond =  True

End Function

'========================================================================================
Function BtnPreview()
	Dim StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType
	Dim StrUrl,IntRetCD

	if lgIntFlgMode <> Parent.OPMD_UMODE then
		IntRetCD = DisplayMsgBox("900002","x","x","x")
		Exit Function
	end if

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    IntRetCD =  SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)
	If IntRetCD = False Then
	    Exit Function
 	End If
 	
    ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varFiscStartDt|" & varFiscStartDt
	StrUrl = StrUrl & "|ClassType|"		& ClassType
	StrUrl = StrUrl & "|VarBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varString2|"	& varString2
	StrUrl = StrUrl & "|varString3|"	& varString3
	'@@

	With frm1.vspdData2
		.Row = 1

		.Col  = 2
		StrUrl = StrUrl & "|bal_lamt|"		& .Text

		.Col  = 3
		StrUrl = StrUrl & "|tot_lamt|"		& .Text

		.Col  = 4
		StrUrl = StrUrl & "|this_lamt|"		& .Text

		.Col  = 6
		StrUrl = StrUrl & "|this_ramt|"		& .Text

		.Col  = 7
		StrUrl = StrUrl & "|tot_ramt|"		& .Text

		.Col  = 8
		StrUrl = StrUrl & "|bal_ramt|"		& .Text

	End With

	StrUrl = StrUrl & "|strSp_Id|"			& strSp_Id

	Call FncEBRPreview(ObjName,StrUrl)

End Function

'========================================================================================
Function BtnPrint()
	Dim StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType
	Dim StrUrl,IntRetCD

	if lgIntFlgMode <> Parent.OPMD_UMODE then
		IntRetCD = DisplayMsgBox("900002","x","x","x")
		Exit Function
	end if

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    IntRetCD =  SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)
	If IntRetCD = False Then
	    Exit Function
 	End If

    ObjName = AskEBDocumentName(StrEBrFile, "ebr")
    
	StrUrl = StrUrl & "varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varFiscStartDt|" & varFiscStartDt
	StrUrl = StrUrl & "|ClassType|"		& ClassType
	StrUrl = StrUrl & "|VarBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varString2|"	& varString2
	StrUrl = StrUrl & "|varString3|"	& varString3
	'@@
	With frm1.vspdData2
		.Row = 1

		.Col  = 2
		StrUrl = StrUrl & "|bal_lamt|"		& .Text

		.Col  = 3
		StrUrl = StrUrl & "|tot_lamt|"		& .Text

		.Col  = 4
		StrUrl = StrUrl & "|this_lamt|"		& .Text

		.Col  = 6
		StrUrl = StrUrl & "|this_ramt|"		& .Text

		.Col  = 7
		StrUrl = StrUrl & "|tot_ramt|"		& .Text

		.Col  = 8
		StrUrl = StrUrl & "|bal_ramt|"		& .Text

	End With

	StrUrl = StrUrl & "|strSp_Id|"			& strSp_Id

	Call FncEBRPrint(EBAction,ObjName,StrUrl)

End Function	



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
Function PrintOpt1_OnClick() 
	if frm1.PrintOpt1.checked = True then
		bs_pl_fg.innerHTML = "재무제표코드"
		Call ElementVisible(frm1.txtClassType, 1)
		Call ElementVisible(frm1.txtClassTypeNm, 1)
		Call ElementVisible(frm1.btnClassType, 1)

		frm1.txtClassType.value		= ""
		frm1.txtClassTypeNm.value	= ""
	end if
End Function

'========================================================================================
Function PrintOpt2_OnClick() 
	if frm1.PrintOpt2.checked = True then
		bs_pl_fg.innerHTML = ""
		Call ElementVisible(frm1.txtClassType, 0)
		Call ElementVisible(frm1.txtClassTypeNm, 0)
		Call ElementVisible(frm1.btnClassType, 0)

		frm1.txtClassType.value		= "*"
	end if
End Function

'========================================================================================
Function PrintOpt3_OnClick() 
	if frm1.PrintOpt3.checked = True then
		bs_pl_fg.innerHTML = ""
		Call ElementVisible(frm1.txtClassType, 0)
		Call ElementVisible(frm1.txtClassTypeNm, 0)
		Call ElementVisible(frm1.btnClassType, 0)

		frm1.txtClassType.value		= "*"
	end if
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
	ggoSpread.Source = frm1.vspdData
	Call JumpPgm()
	
	
End Function

Function JumpPgm()
	
	Dim pvSelmvid, pvFB_fg,pvKeyVal,StrNVar,StrNPgm,pvSingle
	
	if lgIntFlgMode     <> Parent.OPMD_UMODE then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	ggoSpread.Source = frm1.vspdData
     
    
    frm1.vspddata.row = Frm1.vspdData.ActiveRow
    frm1.vspddata.col = 1
	
	
	if 	frm1.vspddata.value <> "" then
		if frm1.PrintOpt3.checked then 
			
			pvKeyVal =  frm1.vspddata.value
			pvSingle  =	frm1.vspddata.value  & chr(11) & _
						frm1.txtBizArea.value & chr(11) & _
						frm1.txtBizArea.value & chr(11) & _ 
						frm1.fpDateTime1.text & chr(11) & _ 
						frm1.fpDateTime2.text & chr(11)
			
			pvFB_fg   = "F"
			pvSelmvid = "ACCT_CD"
	
				Call Jump_Pgm (	pvSelmvid, _
								pvFB_fg, _
								pvSingle,  _
								pvKeyVal)				
		End if				
	End if 
	
	
End Function
	
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
'	Dim ii

	Call SetPopupMenuItemInf("00000000001") 
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
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
           
           If DbQuery2 = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If

End Sub


'========================================================================================
Sub txtStartDT_DblClick(Button)
	If Button = 1 Then
       frm1.txtStartDT.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtStartDT.Focus
	End If
End Sub

'========================================================================================
Sub txtStartDT_Change()
	
End Sub
'========================================================================================
Sub txtEndDT_DblClick(Button)
	If Button = 1 Then
       frm1.txtEndDT.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtEndDT.Focus
	End If
End Sub



'========================================================================================
Sub txtStartDT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtEndDT.focus
	   Call MainQuery()
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub txtEndDT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtStartDT.focus
	   Call MainQuery()
	End If
End Sub

'========================================================================================
Sub ClearVspdData2()
	With frm1.vspdData2
		.Row = 1

		.Col  = C_LBal
		.Text = ""

		.Col  = C_LSum
		.Text = ""

		.Col  = C_LThis
		.Text = ""

		.Col  = C_Result
		.text = ""

		.Col  = C_RThis
		.Text =	""

		.Col  = C_RSum
		.Text = ""

		.Col  = C_RBal
		.Text = ""

	End With 

End Sub


'========================================================================================================
'   Event Name : GetQueryDate()
'   Event Desc : 
'========================================================================================================
Sub GetQueryDate()

	Dim strFromYYYY, strFromMM, strFromDD
	Dim strToYYYY, strToMM, strToDD

	Call ExtractDateFrom(frm1.txtStartDT.text,	Parent.gDateFormat,	Parent.gComDateType,	strFromYYYY,	strFromMM,	strFromDD)
	Call ExtractDateFrom(frm1.txtEndDT.text,	Parent.gDateFormat,	Parent.gComDateType,	strToYYYY,		strToMM,	strToDD)

	lgFiscStart		= GetFiscDate(frm1.txtStartDT.Text)
	lgStartDt		= strFromYYYY	& strFromMM		& strFromDD
	lgEndDt			= strToYYYY		& strToMM		& strToDD

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

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
									<TD CLASS="TD5" NOWRAP>회계일자</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> Name=txtStartDT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일자" tag="12" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtEndDT" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일자" tag="12" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>조회유형</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="15"><LABEL FOR="PrintOpt1">재무제표구분</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="15"><LABEL FOR="PrintOpt2">계정그룹</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt3" VALUE="N" tag="15"><LABEL FOR="PrintOpt3">계정코드</LABEL></SPAN></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" ID="bs_pl_fg" NOWRAP>재무제표코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtClassType" NAME="txtClassType"   SIZE=10 MAXLENGTH=4 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="재무제표코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtClassType.Value, 1)">&nbsp;<INPUT TYPE=TEXT ID="txtClassTypeNm" NAME="txtClassTypeNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
									<TD CLASS="TD5" NOWRAP>조회구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" ID="ZeroFg1" VALUE="Y" tag="15"><LABEL FOR="ZeroFg1">전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" CHECKED ID="ZeroFg2" VALUE="N" tag="15"><LABEL FOR="ZeroFg2">발생금액</LABEL></SPAN></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizArea" NAME="txtBizArea" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNm" NAME="txtBizAreaNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X" ></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD HEIGHT="94%"><!--94%-->
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="6%"><!--6%-->
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
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
					<TD><BUTTON NAME="bttnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()"  Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()"    Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		<!--<TD WIDTH=100% HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>		-->
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizArea"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassType"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFiscStart"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStartDT"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hEndDT"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hSum"				tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hUnBalance"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrintOpt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtClassType"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea"		tag="24" TABINDEX="-1">

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

