<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : sale
'*  2. Function Name        : SDNEPES001
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         :  Ado query Sample with DBAgent(Sort)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : SONG TAE HO
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance


'========================================================================================

Const BIZ_PGM_ID 		= "a5109MB1_nepes.asp"

'========================================================================================
Const C_MaxKey          = 3	

Const C_ThisLeftAmt		= 3
Const C_ThisRightAmt	= 4
Const C_PreLeftAmt		= 5
Const C_PreRightAmt		= 6
'========================================================================================
'=                       4.3 Common variables 
'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================
'=                       4.4 User-defind Variables
'========================================================================================
Dim lgIsOpenPop
Dim lgMaxFieldCount

Dim lgCookValue 

Dim lgFiscStart
Dim lgFromGlDt
Dim lgToGlDt
Dim lgPreFromGlDt
Dim lgPreToGlDt


Dim lgSaveRow 

Dim strSp_Id

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

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

	Dim ServerDate
	Dim PreStartDate
	Dim PreEndDate
	DIm strYear, strMonth, strDay

 	ServerDate	= "<%=GetSvrDate%>"
    PreStartDate = UNIDateAdd("m", -12, Parent.gFiscStart,Parent.gServerDateFormat)
	PreEndDate   = UNIDateAdd("m", -12, Parent.gFiscEnd,  Parent.gServerDateFormat)	

	Call ggoOper.FormatDate(frm1.txtFromGlDt,    Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToGlDt,      Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtPreFromGlDt, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtPreToGlDt,   Parent.gDateFormat, 2)

	Call ExtractDateFrom(Parent.gFiscStart, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtFromGlDt.Year = strYear
	frm1.txtFromGlDt.Month = strMonth

	'Call ExtractDateFrom(Parent.gFiscEnd, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	'frm1.txtToGlDt.Year = strYear
	'frm1.txtToGlDt.Month = strMonth
    'frm1.txtToGlDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ExtractDateFrom(ServerDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtToGlDt.Year = strYear
	frm1.txtToGlDt.Month = strMonth
    
	Call ExtractDateFrom(PreStartDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtPreFromGlDt.Year = strYear
	frm1.txtPreFromGlDt.Month = strMonth

	Call ExtractDateFrom(PreEnddate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtPreToGlDt.Year = strYear
	frm1.txtPreToGlDt.Month = strMonth

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
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Function

'========================================================================================
Sub InitComboBox()

End Sub

'========================================================================================
Sub InitSpreadSheet()
    
	Call SetZAdoSpreadSheet("A5109MA1_GRD01","S","A","V20021220",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
    
    call SetUrl()

    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")
    Call InitComboBox()
    Call CookiePage(0)

    ' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing

    frm1.txtFromGlDt.focus
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

'    If CompareDateByFormat(frm1.txtFromGlDt.Text,frm1.txtToGlDt.Text,frm1.txtFromGlDt.Alt,frm1.txtToGlDt.Alt, _
'	 "970024", frm1.txtFromGlDt.UserDefinedFormat,Parent.gComDateType, true)=False then
'		frm1.txtToGlDt.Focus
'		Exit Function
'	End If

'	If CompareDateByFormat(frm1.txtPreFromGlDt.Text,frm1.txtPreToGlDt.Text,frm1.txtPreFromGlDt.Alt,frm1.txtPreToGlDt.Alt, _
'	 "970024", frm1.txtPreFromGlDt.UserDefinedFormat,Parent.gComDateType, true)=False then
'		frm1.txtPreToGlDt.Focus
'		Exit Function
'	End If

'    If CompareDateByFormat(frm1.txtPreToGlDt.Text,frm1.txtFromGlDt.Text,frm1.txtPreToGlDt.Alt,frm1.txtFromGlDt.Alt, _
'	 "970024", frm1.txtPreToGlDt.UserDefinedFormat,Parent.gComDateType, true)=False then
'		frm1.txtFromGlDt.Focus
'		Exit Function
'	End If

    If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If

    If frm1.txtClassType.value <> "" Then
		IntRetCD = CommonQueryRs(" CLASS_TYPE_NM, CLASS_TYPE"," A_ACCT_CLASS_TYPE ","  CLASS_TYPE = " & FilterVar(frm1.txtClassType.Value, "''", "S") & " and CLASS_TYPE Like " & FilterVar("PL%", "''", "S") & "  " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
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

    Err.Clear
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

			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt
			strVal = strVal & "&txtToGlDt="      & lgToGlDt
			strVal = strVal & "&txtPreFromGlDt=" & lgPreFromGlDt
			strVal = strVal & "&txtPreToGlDt="   & lgPreToGlDt
			strVal = strval & "&txtClassType="   & .txtClassType.value 
			strVal = strVal & "&txtBizAreaCd="	 & .txtBizAreaCd.value
			strVal = strVal & "&strHqBrchFg="	 & "N"
			strVal = strVal & "&strZeroFg="		& strZeroFg
        	strVal = strVal & "&strUserId="		 & Parent.gUsrID
        Else
			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt
        End If   

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgSp_Id="       & strSp_Id
        strVal = strVal & "&lgPageNo="       & lgPageNo
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()

	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgSaveRow        = 1
    Call SetToolBar("1100000000011111")	
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement

End Function


'========================================================================================
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "����� �ڵ�"

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,1)
	End If

End Function


'-------------------------------------  OpenClassTypeCd()  -----------------------------------------------
'	Name : OpenClassTypeCd()
'	Description : Acct Class Type PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenClassTypeCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�繫��ǥ���� �˾�"			' �˾� ��Ī 
	arrParam(1) = "A_ACCT_CLASS_TYPE"			' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtClassType.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("PL%", "''", "S") & " "		' Where Condition
	arrParam(5) = "�繫��ǥ�ڵ�"

    arrField(0) = "CLASS_TYPE"					' Field��(0)
    arrField(1) = "CLASS_TYPE_NM"				' Field��(1)

    arrHeader(0) = "�繫��ǥ�ڵ�"				' Header��(0)
	arrHeader(1) = "�繫��ǥ��"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.Value	= arrRet(0)
				.txtBizAreaNm.Value	= arrRet(1)
			case 2
				.txtClassType.focus
				.txtClassType.Value	= arrRet(0)
				.txtClassNm.Value	= arrRet(1)
		End select
	End With

End Function


'========================================================================================
Function SetPrintCond(StrEbrFile, var1, var2, var3, var4, var5)

	StrEbrFile = "a5109ma1"
	
	' ���Ѱ��� �߰� 
	Dim IntRetCD
	
	var1 = UCASE(Trim(frm1.txtBizAreaCD.value))

	If var1 = "" Then
		If lgAuthBizAreaCd <> "" Then			
			var1  = lgAuthBizAreaCd
		Else
			var1 = "*"
			frm1.txtBizAreaNM.value = ""
		End If			
	Else
		If lgAuthBizAreaCd <> "" Then			
			If UCASE(lgAuthBizAreaCd) <> var1 Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				SetPrintCond =  False
				Exit Function
			End If
		End If			
	End If

'����� �����⶧���� UniConvDateToYYYYMMDD �� ����Ҽ� ����. EBR ������ YYYMMDD ������ �ʿ�� �Ѵ� 
	var2	= lgFromGlDt & "01"
	var3	= lgToGlDt & "01"
	var4	= lgPreFromGlDt & "01"
	var5	= lgPreToGlDt & "01"

	SetPrintCond =  True

End Function

'========================================================================================
Function FncBtnPreview()
	On Error Resume Next

	Dim var1, var2, var3, var4, var5
	Dim strUrl

	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim IntRetCD
	Dim lngPos

	' �繫��ǥ�� �ݵ�� ��ȸ�� ���� ������Ѽ� Data�� ������� ����ϵ��� �Ѵ�.	
	If frm1.vspdData.MaxRows = 0 Then
		IntRetCD = DisplayMsgBox("900002", "X", "X", "X")			'��: "Will you destory previous data"
			' ��ȸ�� ���� �Ͻʽÿ�.
		Exit Function
	End if

	lngPos = 0	

	If Not chkField(Document, "1") Then									'��: This function check indispensable field
	    Exit Function
	End If

    IntRetCD = SetPrintCond(StrEbrFile, var1, var2, var3, var4, var5)
	If IntRetCD = False Then
	    Exit Function
 	End If

    ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "BizAreaCd|"	& var1
	StrUrl = StrUrl & "|FromThisGlDt|"	& var2
	StrUrl = StrUrl & "|ToThisGlDt|"	& var3
	StrUrl = StrUrl & "|FromPreGlDt|"	& var4
	StrUrl = StrUrl & "|ToPreGlDt|"		& var5
	StrUrl = StrUrl & "|strSp_Id|"			& strSp_Id

	Call FncEBRPreview(ObjName,StrUrl)

End Function

'-------------------------------------  FncBtnPrint()  --------------------------------------------------
'	Name : FncBtnPrint()
'	Description : 
'---------------------------------------------------------------------------------------------------------

Function FncBtnPrint()
	On Error Resume Next

	Dim Var1, Var2, Var3, Var4, Var5 

	Dim strUrl

	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim IntRetCD
	Dim lngPos

	' �繫��ǥ�� �ݵ�� ��ȸ�� ���� ������Ѽ� Data�� ������� ����ϵ��� �Ѵ�.	
    If frm1.vspdData.MaxRows = 0 Then
		IntRetCD = DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
    End if

    lngPos = 0	

    If Not chkField(Document, "1") Then
        Exit Function
    End If

    Call SetPrintCond(StrEbrFile, var1, var2, var3, var4, var5)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "BizAreaCd|"	& var1
	StrUrl = StrUrl & "|FromThisGlDt|"	& var2
	StrUrl = StrUrl & "|ToThisGlDt|"	& var3
	StrUrl = StrUrl & "|FromPreGlDt|"	& var4
	StrUrl = StrUrl & "|ToPreGlDt|"		& var5
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
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
'   Event Name : txtPreFromGlDt_KeyDown(KeyCode, Shift)
'   Event Desc :
'=======================================================================================================
Sub txtPreFromGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub
'=======================================================================================================
'   Event Name : txtPreToGlDt_KeyDown(KeyCode, Shift)
'   Event Desc :
'=======================================================================================================
Sub txtPreToGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


'=======================================================================================================
'   Event Name : txtFromGlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToGlDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtPreFromGlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPreFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPreFromGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPreFromGlDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtPreToGlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPreToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPreToGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPreToGlDt.Focus
    End If
End Sub


'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
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
	Dim strPreFromYYYY, strPreFromMM, strPreFromDD
	Dim strPreToYYYY, strPreToMM, strPreToDD

	lgFromGlDt		= frm1.txtFromGlDt.Year		& Right("0" & frm1.txtFromGlDt.Month,2)
	lgToGlDt		= frm1.txtToGlDt.Year		& Right("0" & frm1.txtToGlDt.Month,2)
	lgPreFromGlDt	= frm1.txtPreFromGlDt.Year	& Right("0" & frm1.txtPreFromGlDt.Month,2)
	lgPreToGlDt		= frm1.txtPreToGlDt.Year	& Right("0" & frm1.txtPreToGlDt.Month,2)
End Sub

Sub SetUrl()
    Dim b
    b = mid( gADODBConnString, InStr(gADODBConnString,"Catalog") + 15 ,InStr(gADODBConnString,";Data Source")  - (InStr(gADODBConnString,"Catalog") + 15))     
    
    if b = "nepes"  Then
       execScript("GetInfo()")
    elseif b = "nepes_display" Then
       execScript("GetInfo_display()")
    elseif b = "nepes_rigma" Then
       execScript("GetInfo_rigma()")
    elseif b = "nepes_led" Then
       execScript("GetInfo_led()")  
    End if
End Sub


</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<iframe src="http://192.168.31.231:369/ERPAddition/SM/sm_s1001/web_sm_s10001.aspx" width="100%" height="100%">
</BODY>
</HTML>