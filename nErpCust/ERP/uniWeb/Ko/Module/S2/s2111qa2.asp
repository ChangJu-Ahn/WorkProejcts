<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111QA2
'*  4. Program Name         : 조직별 품목판매계획실적조회 
'*  5. Program Desc         : 조직별 품목판매계획실적조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho Song Hyun
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

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

<%'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   %>
Dim lgStrPrevKey_A                                          <%'☜: Next Key tag                          %>
Dim lgSortKey_A                                             <%'☜: Sort상태 저장변수                     %> 

<%'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   %>
Dim lgStrPrevKey_B                                          <%'☜: Next Key tag                          %>
Dim lgSortKey_B                                             <%'☜: Sort상태 저장변수                     %> 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s2111qb21.asp"
Const BIZ_PGM_ID1       = "s2111qb22.asp"							'☆: Biz logic spread sheet for #2

Const C_MaxKey            = 7										'☆☆☆☆: Max key value
Const C_MaxKey1           = 5										'☆☆☆☆: Max key value

Dim lsCreditGrp														'☆: Jump시 Cookie로 보낼 Grid value
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey_A   = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgStrPrevKey_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	Dim ii,kk	
	Dim iCast

	frm1.txtConSalesOrg.value = Parent.gSalesOrg
	frm1.txtConCurr.value = Parent.gCurrency
	frm1.txtConSpYear.value = Year(UniConvDateToYYYYMMDD(EndDate,Parent.gDateFormat,Parent.gServerDateType))
	frm1.txtConSalesOrg.focus

End Sub

'========================================================================================================= 
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S2111QA21","S","A","V20030711", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetZAdoSpreadSheet("S2111QA22","S","B","V20030711", Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey1, "X", "X" )

	Call SetSpreadLock("A")
	Call SetSpreadLock("B")
End Sub

'========================================================================================================= 
Sub SetSpreadLock(Byval iOpt)
    If iOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
End Sub

'========================================================================================================= 
Function OpenSaleOrg()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "영업조직"						<%' 팝업 명칭 %>
	arrParam(1) = "B_SALES_ORG"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConSalesOrg.Value)		<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
	arrParam(5) = "영업조직"						<%' TextBox 명칭 %>
	
	arrField(0) = "SALES_ORG"							<%' Field명(0)%>
	arrField(1) = "SALES_ORG_NM"						<%' Field명(1)%>
    
	arrHeader(0) = "영업조직"						<%' Header명(0)%>
	arrHeader(1) = "영업조직명"						<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	frm1.txtConSalesOrg.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSaleOrg(arrRet)
	End If	
	
End Function

'========================================================================================================= 
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


'========================================================================================================= 
Function OpenDealType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "거래구분"							<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtConDealTypeCd.Value)		<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
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

'========================================================================================================= 
Function SetSaleOrg(Byval arrRet)

	With frm1
		.txtConSalesOrg.value = arrRet(0) 
		.txtConSalesOrgNm.value = arrRet(1)
	End With

End Function

'========================================================================================================= 
Function SetPlanType(Byval arrRet)

	With frm1
		.txtConPlanTypeCd.value = arrRet(0) 
		.txtConPlanTypeNm.value = arrRet(1)   
	End With

End Function

'========================================================================================================= 
Function SetDealType(Byval arrRet)

	With frm1
		.txtConDealTypeCd.value = arrRet(0) 
		.txtConDealTypeNm.value = arrRet(1)
	End With

End Function

'========================================================================================================= 
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	If gActiveSpdSheet.id = "vspdData1" Then
		Call OpenOrderByPopup("A")
	Else
		Call OpenOrderByPopup("B")
	End If 
		
End Sub

'========================================================================================================= 
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

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
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	

	Call InitSpreadSheet()
	
    Call SetToolBar("11000000000011")							'⊙: 버튼 툴바 제어 
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>
   
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
    
	If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort Col, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort Col, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If
    
   	Dim ii

	If Row < 1 Then Exit Sub

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		   	

	If Col < 1 Then Exit Sub

	frm1.vspdData.Row = Row

	<% '품목 %>
	frm1.vspdData.Col = GetKeyPos("A",1)
	frm1.txtItemCd.value = UCase(Trim(frm1.vspdData.text))
	<% '품목명 %>
	frm1.vspdData.Col = GetKeyPos("A",2)
	frm1.txtItemNm.value = frm1.vspdData.text
	<% '총실적금액 %>
	frm1.vspdData.Col = GetKeyPos("A",3)
	frm1.txtUseAmt.text = frm1.vspdData.text
	<% '총계획금액 %>
	frm1.vspdData.Col = GetKeyPos("A",4)
	frm1.txtPlanAmt.text = frm1.vspdData.text
	<% '총달성율 %>
	frm1.vspdData.Col = GetKeyPos("A",5)
	frm1.txtRate.text = frm1.vspdData.text
	<% '화폐 %>
	frm1.txtCurr.value = frm1.txtConCurr.value

    Call DbQuery("B")

    frm1.vspdData2.MaxRows = 0
    lgStrPrevKey_B   = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
    
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP2C"

    Set gActiveSpdSheet = frm1.vspdData2
    
	If frm1.vspdData2.MaxRows = 0 Then 
		Exit Sub
	End If


    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort Col, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort Col, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If

End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If

End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)			
			If DBQuery("A") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
   End if
    
End Sub

'==========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If				
			Call DisableToolBar(Parent.TBC_QUERY)			
			If DBQuery("B") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
   End if
    
End Sub

'==========================================================================================================
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

'==========================================================================================================
Sub txtConSpYear_onKeyPress()
	Call NumericCheck()
End Sub

'==========================================================================================================
Function FncQuery() 

	Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

   	lgIntFlgMode = Parent.OPMD_CMODE
   		
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If   

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

    Call DbQuery("A")															'☜: Query db data

    FncQuery = True		
End Function

'==========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'==========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'==========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'==========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==========================================================================================================
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

'==========================================================================================================
Function DbQuery(ByVal iOpt) 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

    With frm1

        If iOpt = "A" Then
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		  If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtConSalesOrg=" & Trim(.txtHConSalesOrg.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtHConSpYear.value)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtHConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtHConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtHConCurr.value)
		  Else
			strVal = BIZ_PGM_ID & "?txtConSalesOrg=" & Trim(.txtConSalesOrg.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtConSpYear.value)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtConCurr.value)		  
		  End if
        Else   
		  If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID1 & "?txtConSalesOrg=" & Trim(.txtHConSalesOrg.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtHConSpYear.value)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtHConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtHConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtHConCurr.value)		  				
		  Else        
			strVal = BIZ_PGM_ID1 & "?txtConSalesOrg=" & Trim(.txtConSalesOrg.value)
			strVal = strVal & "&txtConSpYear=" & Trim(.txtConSpYear.value)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(.txtConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(.txtConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(.txtConCurr.value)		
		  End if
        	strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
        End If   

<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        If iOpt = "A" Then
           strVal = strVal & "&lgStrPrevKey_A=" & lgStrPrevKey_A                      '☜: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Else
           strVal = strVal & "&lgStrPrevKey_B=" & lgStrPrevKey_B                      '☜: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
        End If   

        Call RunMyBizASP(MyBizASP, strVal)	

    End With
    
    DbQuery = True
End Function

'==========================================================================================================
Function DbQueryOk(ByVal iOpt)														'☆: 조회 성공후 실행로직 

    If iOpt = "A" Then
        Call SetToolBar("11000000000111")
		If frm1.vspdData.MaxRows > 0 Then
			frm1.vspdData.Focus
			frm1.vspdData.SelModeSelected = True
			If lgIntFlgMode <> Parent.OPMD_UMODE Then
				frm1.vspdData.Row = 1
				Call vspdData_Click(1, 1)
			End If
			lgIntFlgMode = Parent.OPMD_UMODE		
		Else
			frm1.txtConSalesOrg.focus 
		End If
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
					<TD CLASS="CLSLTAB">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>품목판매계획대실적</font></td>
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
					<TD HEIGHT=20 WIDTH=100% COLSPAN=2>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업조직</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSalesOrg" ALT="영업조직" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleOrg()">&nbsp;<INPUT NAME="txtConSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>계획년도</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSpYear" ALT="계획년도" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12X1XU"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계획구분</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConPlanTypeCd" ALT="계획구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanType()">&nbsp;<INPUT NAME="txtConPlanTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>거래구분</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConDealTypeCd" ALT="거래구분" TYPE="Text" MAXLENGTH=1 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDealType()">&nbsp;<INPUT NAME="txtConDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>							</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>화폐</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConCurr" ALT="화폐" TYPE="Text" MAXLENGTH=3 SiZE=10 tag="14XXXU"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=45% HEIGHT=100% ROWSPAN="2" valign=top>
						<TABLE WIDTH=100% HEIGHT=100%>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/s2111qa2_vspdData1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=55% HEIGHT=20% valign=top>
						<FIELDSET CLASS="CLSFLD">
						<TABLE WIDTH=100% HEIGHT=100% CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemCd" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="24XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목명</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>총계획금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0 STYLE="PADDING-BOTTOM:2px;PADDING-TOP:2px">
										<TR>
											<TD>
												<script language =javascript src='./js/s2111qa2_fpDoubleSingle1_txtPlanAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtCurr" TYPE="Text" MAXLENGTH=3 SiZE=5 tag="24XXXU">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>총실적금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0 STYLE="PADDING-BOTTOM:2px">
										<TR>
											<TD>
												<script language =javascript src='./js/s2111qa2_fpDoubleSingle2_txtUseAmt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>총달성율</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0 STYLE="PADDING-BOTTOM:2px">
										<TR>
											<TD>
												<script language =javascript src='./js/s2111qa2_fpDoubleSingle3_txtRate.js'></script>
											</TD>
											<TD>
												&nbsp;
											</TD>
											<TD>
												%
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=55% HEIGHT=80% valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%">
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/s2111qa2_vspdData2_vspdData2.js'></script>
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
	<INPUT TYPE=HIDDEN NAME="txtHConSalesOrg" tag="24" TABINDEX="-1"> 
	<INPUT TYPE=HIDDEN NAME="txtHConSpYear" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConPlanTypeCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConDealTypeCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHConCurr" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
