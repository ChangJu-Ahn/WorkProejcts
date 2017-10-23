<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2112QA1
'*  4. Program Name         : 그룹별 품목판매계획현황조회 
'*  5. Program Desc         : 그룹별 품목판매계획현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID        = "s2112qb1.asp"
Const BIZ_PGM_JUMP_ID   = "s2111ma1"				  	       '☆: 비지니스 로직 ASP명 

Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value

Dim lsCreditGrp                                            '☆: Jump시 Cookie로 보낼 Grid value

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

	frm1.txtConSalesGrp.value = Parent.gSalesGrp
	frm1.txtConCurr.value = Parent.gCurrency
	frm1.txtConSpYear.value = Year(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))
	frm1.cboSpMonth.value = Month(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))
	frm1.txtConSalesGrp.focus
	
End Sub

'===========================================================================================================
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "QA") %>
End Sub


'===========================================================================================================
Sub InitComboBox()
    Call SetCombo(frm1.cboSpMonth, "01", "01")
    Call SetCombo(frm1.cboSpMonth, "02", "02")
    Call SetCombo(frm1.cboSpMonth, "03", "03")
    Call SetCombo(frm1.cboSpMonth, "04", "04")
    Call SetCombo(frm1.cboSpMonth, "05", "05")
    Call SetCombo(frm1.cboSpMonth, "06", "06")
    Call SetCombo(frm1.cboSpMonth, "07", "07")
    Call SetCombo(frm1.cboSpMonth, "08", "08")
    Call SetCombo(frm1.cboSpMonth, "09", "09")
    Call SetCombo(frm1.cboSpMonth, "10", "10")
    Call SetCombo(frm1.cboSpMonth, "11", "11")
    Call SetCombo(frm1.cboSpMonth, "12", "12")
End Sub

'===========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S2112QA1","S","A","V20030828", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock
End Sub

'===========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'===========================================================================================================
Function OpenPlanNumber()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "계획차수"		<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"				<%' TABLE 명칭 %>

	arrParam(2) = Trim(frm1.txtConPlanNum.Value)  <%' Code Condition%>
	arrParam(3) = ""					<%' Name Cindition%>

	arrParam(4) = "MAJOR_CD=" & FilterVar("S2001", "''", "S") & ""    <%' Where Condition%>
	arrParam(5) = "계획차수"		<%' TextBox 명칭 %>
		 
	arrField(0) = "MINOR_CD"			<%' Field명(0)%>
	arrField(1) = "MINOR_NM"			<%' Field명(1)%>
		    
	arrHeader(0) = "계획차수"       <%' Header명(0)%>
	arrHeader(1) = "계획차수명"     <%' Header명(1)%>

	frm1.txtConPlanNum.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
	
		Exit Function
	
	Else
	
		Call SetPlanNumber(arrRet)
	
	End If 
 
End Function

'===========================================================================================================
Function SetPlanNumber(Byval arrRet)

	With frm1	
		.txtConPlanNum.value = arrRet(0) 
		.txtConPlanNumNm.value = arrRet(1)
	
		'lgBlnFlgChgValue = True
	End With

End Function

'===========================================================================================================
Function OpenSaleGrp()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "영업그룹"						<%' 팝업 명칭 %>
	arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConSalesGrp.Value)		<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtConSalesGrpNm.Value)     <%' Name Cindition%>
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "	                    <%' Where Condition%>
	arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>
	
	arrField(0) = "SALES_GRP"							<%' Field명(0)%>
	arrField(1) = "SALES_GRP_NM"						<%' Field명(1)%>
    
	arrHeader(0) = "영업그룹"						<%' Header명(0)%>
	arrHeader(1) = "영업그룹명"						<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	frm1.txtConSalesGrp.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSaleGrp(arrRet)
	End If	
	
End Function

'===========================================================================================================
'pis 품목팝업 수정 
Function OpenSaleItem()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(1) = "b_item"									<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConSalesItem.Value)			<%' Code Condition%>
	arrParam(3) = ""                             			<%' Name Cindition%>
	arrParam(4) = ""										<%' Where Condition%>
	arrParam(5) = "품목"								<%' TextBox 명칭 %>
	
	arrField(0) = "Item_cd"									<%' Field명(0)%>
	arrField(1) = "Item_nm"									<%' Field명(1)%>
	arrField(2) = "Spec"	
    
	arrHeader(0) = "품목"								<%' Header명(0)%>
	arrHeader(1) = "품목명"								<%' Header명(1)%>
	arrHeader(2) = "규격"	
	    
	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	frm1.txtConSalesItem.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSaleItem(arrRet)
	End If	
	
End Function

'===========================================================================================================
Function OpenPlanType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "계획구분"							<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConPlanTypeCd.Value)			<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S4089", "''", "S") & ""						<%' Where Condition%>
	arrParam(5) = "계획구분"							<%' TextBox 명칭 %>
	
	arrField(0) = "MINOR_CD"								<%' Field명(0)%>
	arrField(1) = "MINOR_NM"								<%' Field명(1)%>
    
	arrHeader(0) = "계획구분"							<%' Header명(0)%>
	arrHeader(1) = "계획구분명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	frm1.txtConPlanTypeCd.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlanType(arrRet)
	End If	
	
End Function


'===========================================================================================================
Function OpenDealType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "거래구분"							<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConDealTypeCd.Value)			<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S4225", "''", "S") & ""						<%' Where Condition%>
	arrParam(5) = "거래구분"							<%' TextBox 명칭 %>
	
	arrField(0) = "MINOR_CD"								<%' Field명(0)%>
	arrField(1) = "MINOR_NM"								<%' Field명(1)%>
    
	arrHeader(0) = "거래구분"							<%' Header명(0)%>
	arrHeader(1) = "거래구분명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	frm1.txtConDealTypeCd.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDealType(arrRet)
	End If	
	
End Function

'===========================================================================================================
Function OpenCurr()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "화폐"								<%' 팝업 명칭 %>
	arrParam(1) = "B_CURRENCY"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConCurr.Value)				<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = ""										<%' Where Condition%>
	arrParam(5) = "화폐"								<%' TextBox 명칭 %>
	
	arrField(0) = "CURRENCY"								<%' Field명(0)%>
	arrField(1) = "CURRENCY_DESC"							<%' Field명(1)%>
    
	arrHeader(0) = "화폐"								<%' Header명(0)%>
	arrHeader(1) = "화폐명"								<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtConCurr.value = arrRet(0)
	End If	
	
End Function


'===========================================================================================================
Function SetSaleGrp(Byval arrRet)

	With frm1
		.txtConSalesGrp.value = arrRet(0) 
		.txtConSalesGrpNm.value = arrRet(1)
	End With

End Function


'===========================================================================================================
Function SetSaleItem(Byval arrRet)

	With frm1
		.txtConSalesItem.value = arrRet(0) 
		.txtConSalesItemNm.value = arrRet(1)
	End With

End Function


'===========================================================================================================
Function SetPlanType(Byval arrRet)

	With frm1
		.txtConPlanTypeCd.value = arrRet(0) 
		.txtConPlanTypeNm.value = arrRet(1)   
	End With

End Function


'===========================================================================================================
Function SetDealType(Byval arrRet)

	With frm1
		.txtConDealTypeCd.value = arrRet(0) 
		.txtConDealTypeNm.value = arrRet(1)
	End With

End Function

'===========================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877									<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then											<% 'Jump로 화면을 이동할 경우 %>

		If frm1.vspdData.MaxRows = 0 Then Exit Function

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		strTemp = ""
		For i = 0 To frm1.vspdData.MaxCols - 3
			Select Case i
			Case 0
				strTemp = lgKeyPosVal(i)
			Case Else
				strTemp = strTemp & Parent.gRowSep & lgKeyPosVal(i)
			End Select
		Next
    

		WriteCookie CookieSplit , strTemp						<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then										<% 'Jump로 화면이 이동해 왔을경우 %>

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

<%'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------%>
		<% '자동조회되는 조건값과 검색조건부 Name의 Match %>
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("여신관리그룹")
				frm1.txtCreditGrp.value =  arrVal(iniSep + 1)
			Case UCase("여신관리그룹명")
				frm1.txtCreditGrpNm.value =  arrVal(iniSep + 1)
			End Select
		Next
<%'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------%>

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'========================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub


'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitComboBox
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>
   
	Call CookiePage(0)
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
End Sub

'===========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'===========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If    
		Exit Sub     
    End If
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>

<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
    
End Sub

'===========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'===========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
				
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End if
	End if	    


End Sub

'===========================================================================================================
Function NumericCheck()

	Dim objEl, KeyCode
	
	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode

	Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
	Case Else
		window.event.keycode = 0
	End Select

End Function

'===========================================================================================================
Sub txtConSpYear_onKeyPress()
	Call NumericCheck()
End Sub

'===========================================================================================================
Sub cboSpMonth_onKeyPress()
	Call NumericCheck()
End Sub

'===========================================================================================================
Sub txtConPlanNum_onKeyPress()
	Call NumericCheck()
End Sub


'===========================================================================================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   
    lgIntFlgMode     = Parent.OPMD_CMODE 
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'===========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'===========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'===========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'===========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'===========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'===========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If
    
    With frm1
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtConSalesGrp=" & Trim(.txtHConSalesGrp.value)
			strVal = strVal & "&txtConSalesItem=" & Trim(.txtHConSalesItem.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtHConSpYear.value)
			strVal = strVal & "&cboSpMonth=" & Trim(.cboHSpMonth.value)		
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtHConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtHConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtHConCurr.value)
			strVal = strVal & "&txtConPlanNum=" & Trim(.txtHConPlanNum.value)
		Else
			strVal = BIZ_PGM_ID & "?txtConSalesGrp=" & Trim(.txtConSalesGrp.value)
			strVal = strVal & "&txtConSalesItem=" & Trim(.txtConSalesItem.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtConSpYear.value)
			strVal = strVal & "&cboSpMonth=" & Trim(.cboSpMonth.value)		
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtConCurr.value)
			strVal = strVal & "&txtConPlanNum=" & Trim(.txtConPlanNum.value)		
		End If
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")        
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function

'===========================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
	Call SetToolbar("11000000000111")
    
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		End If
		lgIntFlgMode = Parent.OPMD_UMODE	
    Else
       frm1.txtConSalesGrp.focus
    End If  

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목판매계획현황</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleGrp()">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								    <TD CLASS="TD5" NOWRAP>계획년도</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSpYear" ALT="계획년도" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSalesItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSaleItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleItem()">&nbsp;<INPUT NAME="txtConSalesItemNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>계획월</TD>
									<TD CLASS="TD6"><SELECT NAME="cboSpMonth" ALT="계획월" STYLE="Width: 97px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>		
								<TR>
									<TD CLASS="TD5" NOWRAP>계획구분</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConPlanTypeCd" ALT="계획구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanType()">&nbsp;<INPUT NAME="txtConPlanTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>거래구분</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConDealTypeCd" ALT="거래구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDealType()">&nbsp;<INPUT NAME="txtConDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>							
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계획차수</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConPlanNum" ALT="계획차수" TYPE="Text" MAXLENGTH=2 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanNumber()">&nbsp;<INPUT NAME="txtConPlanNumNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>화폐</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConCurr" ALT="화폐" TYPE="Text" MAXLENGTH=3 SiZE=10 tag="14XXXU"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s2112qa1_vaSpread1_vspdData.js'></script>
									</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
	<INPUT TYPE=HIDDEN NAME="txtHConSalesGrp" tag="24" TABINDEX="-1"> 
	<INPUT TYPE=HIDDEN NAME="txtHConSalesItem" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConSpYear" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="cboHSpMonth" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConPlanTypeCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConDealTypeCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConCurr" tag="14" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConPlanNum" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
