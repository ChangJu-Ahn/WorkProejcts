<%@ LANGUAGE="VBSCRIPT" %>
<%'********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S5212QA1
'*  4. Program Name         : B/L상세조회 
'*  5. Program Desc         : B/L상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim IscookieSplit

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gDateFormat)

Const BIZ_PGM_ID        = "s5212qb1_ko441.asp"
Const BIZ_PGM_JUMP_ID	= "s5212ma1"
Const C_MaxKey          = 4                                    '☆☆☆☆: Max key value

Const C_PopSalesGrp		= 1			' 영업그룹 
Const C_PopSoldToParty	= 2			' 주문처 
Const C_PopItemCd		= 3			' 품목코드 

'=========================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = Parent.OPMD_CMODE	
End Sub
'==========================================  2.2 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	
<%'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtBLFrDt.text = StartDate
	frm1.txtBLToDt.text = EndDate

<%'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>
	
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGroup, "Q") 
        	frm1.txtSalesGroup.value = lgSGCd
	End If
	
	frm1.txtSalesGroup.focus


End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub
'========================================= 2.6 InitSpreadSheet() =========================================
Sub InitSpreadSheet()
 
	Call SetZAdoSpreadSheet("s5212QA1","S","A","V20030714", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    
    Call SetSpreadLock 

End Sub
'========================================= 2.7 SetSpreadLock() ===========================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenEXBLNoPop()
	Dim iCalledAspName
	Dim strRet
		
	If lgIsOpenPop = True Or UCase(frm1.txtBLNo.className) = "PROTECTED" Then Exit Function
		
	iCalledAspName = AskPRAspName("s5211pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5211pa1", "x")
		lgIsOpenPop = False
		exit Function
	end if

	lgIsOpenPop = True

	frm1.txtBLNo.focus 
			
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
		
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetExBLNo(strRet)
	End If	
End Function
'===========================================================================
Function OpenConPopup(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	OpenConPopup = False
	
	lgIsOpenPop = True

	Select Case iWhere
	Case C_PopSoldToParty
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				<%' Where Condition%>
		arrParam(5) = "수입자"							<%' TextBox 명칭 %>
	
		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>
    
		arrHeader(0) = "수입자"							<%' Header명(0)%>
		arrHeader(1) = "수입자명"						<%' Header명(1)%>
		
		frm1.txtconBp_cd.focus

	Case C_PopItemCd
		OpenConPopup = OpenConItemPopup(C_PopItemCd, frm1.txtItem_cd.value)
		frm1.txtItem_cd.focus
		Exit Function

	Case C_PopSalesGrp

                If frm1.txtSalesGroup.className = "protected" Then
                	lgIsOpenPop = False
                        Exit Function
                End If 	

		arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>
	
		arrField(0) = "SALES_GRP"							<%' Field명(0)%>
		arrField(1) = "SALES_GRP_NM"							<%' Field명(1)%>
    
		arrHeader(0) = "영업그룹"						<%' Header명(0)%>
		arrHeader(1) = "영업그룹명"							<%' Header명(1)%>
		
		frm1.txtSalesGroup.focus
		
	End Select

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		OpenConPopup = SetConPopup(arrRet, iWhere)
	End If	
	
End Function

' Item Popup
'========================================
Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName

	OpenConItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub
'========================================================================================================
Function OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	
	ON ERROR RESUME NEXT
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetExBLNo(strRet)
	frm1.txtBLNo.value = strRet(0)
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetConPopup(Byval arrRet, Byval iWhere)

	SetConPopup = False
	
	With frm1
		Select Case iWhere
		
		Case C_PopSalesGrp
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
			
		Case C_PopSoldToParty
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
			
		Case C_PopItemCd
			.txtItem_cd.value = arrRet(0) 
			.txtItem_Nm.value = arrRet(1)   
			
		End Select
	End With

	SetConPopup = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then								<% 'Jump로 화면을 이동할 경우 %>

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		
		WriteCookie CookieSplit , IsCookieSplit					<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							<% 'Jump로 화면이 이동해 왔을경우 %>

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		
		Dim iniSep

<%'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------%>
		<% '자동조회되는 조건값과 검색조건부 Name의 Match %>
		frm1.txtconBp_cd.value =  arrVal(0)
		frm1.txtconBp_Nm.value =  arrVal(1)
		frm1.txtBillType.value =  arrVal(2)
		frm1.txtBillTypeNm.value = arrVal(3) 
		frm1.txtSalesOrg.value =  arrVal(4)
		frm1.txtSalesOrgNm.value = arrVal(5) 
		frm1.txtSalesGroup.value =  arrVal(6)
		frm1.txtSalesGroupNm.value = arrVal(7) 
		frm1.txtItem_cd.value =  arrVal(8)
		frm1.txtItem_Nm.value = arrVal(9)

<%'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------%>

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function
'========================================================================================
Function FncSplitColumn()
   
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 4
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	    
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub Form_Load()
    Call LoadInfTB19029	

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                '⊙: Lock  Suitable  Field
    
  	Call InitVariables
        Call GetValue_ko441() 														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("00000000001")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    Set gActiveSpdSheet = frm1.vspdData

	If Row <= 0 Then
	     
	     If lgSortKey = 1 Then
	         ggoSpread.SSSort Col				'Sort in Ascending
	         lgSortKey = 2
	     Else
	         ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	         lgSortKey = 1
	     End If
		 Exit Sub     
	 End If
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>
	If Row <> 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		IscookieSplit = frm1.vspdData.text
	Else
		IscookieSplit = ""
	End if
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>    
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess Then	Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DBQuery
		End if
	End if	    


End Sub
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub txtBLFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBLFrDt.Action = 7
		Call SetFocusToDocument("M")   
		frm1.txtBLFrDt.Focus				
	End If
End Sub

Sub txtBLToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBLToDt.Action = 7
		Call SetFocusToDocument("M")   
		frm1.txtBLToDt.Focus				
	End If
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub txtBLFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtBLToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function FncQuery() 
	FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
   	lgIntFlgMode = Parent.OPMD_CMODE	

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function FncPrint() 
    Call parent.FncPrint()
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function
'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
		strVal = strVal & "&txtBLNo=" & Trim(.txtBLNo.value)
		strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)
		strVal = strVal & "&txtBLFrDt=" & Trim(.txtBLFrDt.text)
		strVal = strVal & "&txtBLToDt=" & Trim(.txtBLToDt.text)
		
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
                strVal = strVal & "&gBizArea=" & lgBACd 
                strVal = strVal & "&gPlant=" & lgPLCd 
                strVal = strVal & "&gSalesGrp=" & lgSGCd 
                strVal = strVal & "&gSalesOrg=" & lgSOCd       
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
  Call SetToolBar("11000000000111")
    If frm1.vspdData.MaxRows > 0 Then
    	frm1.vspdData.Focus
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		End If
		lgIntFlgMode = Parent.OPMD_UMODE		
    Else
       frm1.txtSalesGroup.focus
    End If  
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L상세</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>수입자</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="수입자" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSoldToParty">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>B/L관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=20 MAXLENGTH=18 TAG="11XXXU" ALT="B/L관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEXBLNoPop">
									<TD CLASS=TD5 NOWRAP>B/L발행일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s5212qa1_fpDateTime1_txtBLFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s5212qa1_fpDateTime2_txtBLToDt.js'></script>
									</TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6><INPUT NAME="txtItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopItemCd">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
								<script language =javascript src='./js/s5212qa1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<td <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%"><TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">B/L내역등록</a></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HconItem_cd" tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="HValid_from_dt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconCurrency" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconDeal_type" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconPay_terms" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconSales_unit" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14" TABINDEX="-1">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>

