<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1111QA1
'*  4. Program Name         : 품목단가조회 
'*  5. Program Desc         : 품목단가조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/07/04
'*  8. Modified date(Last)  : 2005/05/03
'*  9. Modifier (First)     : SonBumYeol		
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : -2002/12/06 : UI성능향상(include) 반영 강준구 
'*                            -2002/12/12 : UI성능향상(include) 다시 반영 강준구 
'**********************************************************************************************
%>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim gblnWinEvent

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim prDBSYSDate
Dim EndDate ,StartDate
prDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)

Const BIZ_PGM_ID        = "s1111qb1.asp"

Const C_MaxKey          = 13                                  '☆☆☆☆: Max key value
                                            '☆: Jump시 Cookie로 보낼 Grid value

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = parent.OPMD_CMODE						   'Indicates that current mode is Create mode

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtconValid_from_dt.text = EndDate
	frm1.txtconValid_from_dt.focus
End Sub

'===========================================================================================================
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub


'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S1111QA1","S","A","V20050503", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub


'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub


'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case iWhere
	Case 0
		arrParam(1) = "b_item"									<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconItem_cd.Value)			<%' Code Condition%>
		arrParam(3) = ""                             			<%' Name Cindition%>
		arrParam(4) = ""										<%' Where Condition%>
		arrParam(5) = "품목"								<%' TextBox 명칭 %>
	
		arrField(0) = "Item_cd"									<%' Field명(0)%>
		arrField(1) = "Item_nm"									<%' Field명(1)%>
		arrField(2) = "Spec"									<%' Field명(1)%>
    
		arrHeader(0) = "품목"								<%' Header명(0)%>
		arrHeader(1) = "품목명"								<%' Header명(1)%>
		arrHeader(2) = "규격"								<%' Header명(1)%>

		frm1.txtconItem_cd.focus 
	Case 1
		arrParam(1) = "B_UNIT_OF_MEASURE"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconSales_unit.Value)			<%' Code Condition%>
		arrParam(3) = ""										<%' Name Cindition%>
		arrParam(4) = ""										<%' Where Condition%>
		arrParam(5) = "단위"								<%' TextBox 명칭 %>
	
		arrField(0) = "UNIT"									<%' Field명(0)%>
		arrField(1) = "UNIT_NM"									<%' Field명(1)%>
		
		arrHeader(0) = "단위"								<%' Header명(0)%>
		arrHeader(1) = "단위명"								<%' Header명(1)%>

		frm1.txtconSales_unit.focus
	Case 2
		arrParam(1) = "B_CURRENCY"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconCurrency.Value)			<%' Code Condition%>
		arrParam(3) = ""										<%' Name Cindition%>
		arrParam(4) = ""										<%' Where Condition%>
		arrParam(5) = "화폐"								<%' TextBox 명칭 %>
	
		arrField(0) = "CURRENCY"								<%' Field명(0)%>
		arrField(1) = "CURRENCY_DESC"							<%' Field명(1)%>
    
		arrHeader(0) = "화폐"								<%' Header명(0)%>
		arrHeader(1) = "화폐명"								<%' Header명(1)%>

		frm1.txtconCurrency.focus 
	Case 3
		arrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconDeal_type.Value)			<%' Code Condition%>
		arrParam(3) = ""                                		<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & "" 						<%' Where Condition%>
		arrParam(5) = "판매유형"							<%' TextBox 명칭 %>
	
		arrField(0) = "MINOR_CD"								<%' Field명(0)%>
		arrField(1) = "MINOR_NM"								<%' Field명(1)%>
    
		arrHeader(0) = "판매유형"							<%' Header명(0)%>
		arrHeader(1) = "판매유형명"							<%' Header명(1)%>

		frm1.txtconDeal_type.focus 
	Case 4
		arrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconPay_terms.Value)			<%' Code Condition%>
		arrParam(3) = ""                                 		<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""						<%' Where Condition%>
		arrParam(5) = "결제방법"							<%' TextBox 명칭 %>
	
		arrField(0) = "MINOR_CD"								<%' Field명(0)%>
		arrField(1) = "MINOR_NM"								<%' Field명(1)%>
    
		arrHeader(0) = "결제방법"							<%' Header명(0)%>
		arrHeader(1) = "결제방법명"							<%' Header명(1)%>

		frm1.txtconPay_terms.focus 
	End Select

	arrParam(0) = arrParam(5)									<%' 팝업 명칭 %>


	Select Case iWhere
	Case 0
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select


	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================= 
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
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


'========================================================================================================= 
Function SetConSItemDC(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		Case 0
			.txtconItem_cd.value = arrRet(0) 
			.txtconItem_nm.value = arrRet(1)   
		Case 1
			.txtconSales_unit.value = arrRet(0) 
		Case 2
			.txtconCurrency.value = arrRet(0) 
		Case 3
			.txtconDeal_type.value = arrRet(0) 
			.txtconDeal_type_nm.value = arrRet(1)   
		Case 4
			.txtconPay_terms.value = arrRet(0) 
			.txtconPay_terms_nm.value = arrRet(1)   
		End Select

	End With
	
End Function


'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolBar("11000000000011")							'⊙: 버튼 툴바 제어 
End Sub

'=======================================================================================================
 Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
		End If
	End if    
    
End Sub

<%
'==========================================================================================
'   Event Desc : OCX_DbClick() 시 Calendar Popup
'==========================================================================================
%>
Sub txtconValid_from_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtconValid_from_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtconValid_from_dt.Focus
	End If
End Sub

<%
'==========================================================================================
'   Event Desc : 조회조건부의 OCX_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
%>
Sub txtconValid_from_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
Function FncQuery() 
    Dim IntRetCD
    FncQuery = False                                                        '⊙: Processing is NG    
    Err.Clear                                                               '☜: Protect system from crashing

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
       
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
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

'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
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
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function


<%
'========================================================================================
' Function Desc : This function is related to Print Button
'========================================================================================
%>
Function BtnPrint()
	Dim IntRetCD
	Dim var1,var2,var3,var4,var5,var6
	Dim StrUrl
	Dim arrParam, arrField, arrHeader
	
	if lgIntFlgMode <> parent.OPMD_UMODE then	
		IntRetCD = DisplayMsgBox("900002","x","x","x")   ' 조회를 먼저 하십시오.	
		Exit Function
	end if		
	
    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If	
    
	if frm1.vspddata.MaxRows < 1 then
		IntRetCD = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
		'데이타가 없습니다.
		Exit Function
	end if
	
				
	var1 = UniConvDateToYYYYMMDD(frm1.txtconValid_from_dt.Text,parent.gDateFormat,parent.gServerDateType)
	
	If UCase(frm1.txtconItem_cd.value) = "" Then
		var2 = "%"	' "%"
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtconItem_cd.value)), "" ,  "SNM")
	End If
    
    If UCase(frm1.txtconPay_terms.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtconPay_terms.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconDeal_type.value) = "" Then
		var4 = "%"
	Else
		var4 = FilterVar(Trim(UCase(frm1.txtconDeal_type.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconSales_unit.value) = "" Then
		var5 = "%"
	Else
		var5 = FilterVar(Trim(UCase(frm1.txtconSales_unit.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconCurrency.value) = "" Then
		var6 = "%"
	Else
		var6 = FilterVar(Trim(UCase(frm1.txtconCurrency.value)), "" ,  "SNM")
	End If
    
			
	strUrl = strUrl & "VALID_FROM_DT|" & var1     
	strUrl = strUrl & "|ITEM_CD|" & var2
	strUrl = strUrl & "|PAY_METH|" & var3
	strUrl = strUrl & "|DEAL_TYPE|" & var4
	strUrl = strUrl & "|UNIT|" & var5
	strUrl = strUrl & "|CUR|" & var6 

	OBjName = AskEBDocumentName("s1111oa2","ebr")    
	Call FncEBRprint(EBAction, OBjName, strUrl)

End Function

<%
'========================================================================================
' Function Desc : This function is related to Preview Button
'========================================================================================
%>
Function BtnPreview()
	Dim IntRetCD
	Dim var1,var2,var3,var4,var5,var6
	Dim StrUrl
	Dim arrParam, arrField, arrHeader
	
	if lgIntFlgMode <> parent.OPMD_UMODE then	
		IntRetCD = DisplayMsgBox("900002","x","x","x")   ' 조회를 먼저 하십시오.	
		Exit Function
	end if		
	
    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If	
    
	if frm1.vspddata.MaxRows < 1 then
		IntRetCD = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
		'데이타가 없습니다.
		Exit Function
	end if
	
			
	var1 = UniConvDateToYYYYMMDD(frm1.txtconValid_from_dt.Text,parent.gDateFormat,parent.gServerDateType)
	
	If UCase(frm1.txtconItem_cd.value) = "" Then
		var2 = "%"	' "%"
	Else
		var2 = FilterVar(Trim(UCase(frm1.txtconItem_cd.value)), "" ,  "SNM")
	End If
    
    If UCase(frm1.txtconPay_terms.value) = "" Then
		var3 = "%"
	Else
		var3 = FilterVar(Trim(UCase(frm1.txtconPay_terms.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconDeal_type.value) = "" Then
		var4 = "%"
	Else
		var4 = FilterVar(Trim(UCase(frm1.txtconDeal_type.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconSales_unit.value) = "" Then
		var5 = "%"
	Else
		var5 = FilterVar(Trim(UCase(frm1.txtconSales_unit.value)), "" ,  "SNM")
	End If

    If UCase(frm1.txtconCurrency.value) = "" Then
		var6 = "%"
	Else
		var6 = FilterVar(Trim(UCase(frm1.txtconCurrency.value)), "" ,  "SNM")
	End If
    
			
	strUrl = strUrl & "VALID_FROM_DT|" & var1     
	strUrl = strUrl & "|ITEM_CD|" & var2
	strUrl = strUrl & "|PAY_METH|" & var3
	strUrl = strUrl & "|DEAL_TYPE|" & var4
	strUrl = strUrl & "|UNIT|" & var5
	strUrl = strUrl & "|CUR|" & var6 

    
	OBjName = AskEBDocumentName("s1111oa2","ebr")    
	Call FncEBRPreview(OBjName, strUrl)		

End Function


'========================================================================================
Function DbQuery() 
	
	Dim strVal

    DbQuery = False
    
   	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtHconItem_cd.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtconDeal_type=" & Trim(frm1.txtHconDeal_type.value)
		strVal = strVal & "&txtconPay_terms=" & Trim(frm1.txtHconPay_terms.value)
		strVal = strVal & "&txtconValid_from_dt=" & Trim(frm1.txtHconValid_from_dt.value)
		strVal = strVal & "&txtconSales_unit=" & Trim(frm1.txtHconSales_unit.value)
		strVal = strVal & "&txtconCurrency=" & Trim(frm1.txtHconCurrency.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtconItem_cd.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtconDeal_type=" & Trim(frm1.txtconDeal_type.value)
		strVal = strVal & "&txtconPay_terms=" & Trim(frm1.txtconPay_terms.value)
		strVal = strVal & "&txtconValid_from_dt=" & Trim(frm1.txtconValid_from_dt.text)
		strVal = strVal & "&txtconSales_unit=" & Trim(frm1.txtconSales_unit.value)
		strVal = strVal & "&txtconCurrency=" & Trim(frm1.txtconCurrency.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			
	End If
		
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True    

End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		End If
		lgIntFlgMode = parent.OPMD_UMODE
	Else
		Call SetFocusToDocument("M")
		frm1.txtconValid_from_dt.focus
	End If
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolBar("11000000000111")
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목단가조회</font></td>
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
									<TD CLASS="TD5" NOWRAP>유효일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/s1111qa1_fpDateTime1_txtconValid_from_dt.js'></script></TD>	
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value, 0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>결제방법</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconPay_terms" ALT="결제방법" TYPE="Text" MAXLENGTH=5 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconPay_terms.value,4">&nbsp;<INPUT NAME="txtconPay_terms_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD> 
									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconDeal_type" ALT="판매유형" TYPE="Text" MAXLENGTH=5 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconDeal_type.value,3">&nbsp;<INPUT NAME="txtconDeal_type_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>단위</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconSales_unit" ALT="단위" TYPE="Text" MAXLENGTH=3 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconSales_unit.value,1"></TD>
									<TD CLASS="TD5" NOWRAP>화폐</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconCurrency" ALT="화폐" TYPE="Text" MAXLENGTH=3 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconCurrency.value,2"></TD>
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
										<script language =javascript src='./js/s1111qa1_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
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
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconDeal_type" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconPay_terms" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconValid_from_dt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconSales_unit" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHconCurrency" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX="-1">
    <input type="hidden" name="dbname" TABINDEX="-1">
    <input type="hidden" name="filename" TABINDEX="-1">
    <input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>
