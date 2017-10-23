<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : s3134ma1
'*  4. Program Name         : 출고현황조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/08/08
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/12/14 Include 성능향상 강준구 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop
Dim lgLngStartRow

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

Const BIZ_PGM_ID        = "s3134mb1.asp"
Const C_MaxKey          = 25                                    '☆☆☆☆: Max key value


'=========================================
Sub InitVariables()
    lgPageNo     = ""                                  
    lgSortKey        = 1
    lgIntFlgMode     = parent.OPMD_CMODE						   

    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 

End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtSoDtFrom.text = StartDate
	frm1.txtSoDtTo.text = EndDate

	frm1.txtSoDtFrom.focus
End Sub

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub


'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3134MA1","S","A","V20030711", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================
Function OpenConDnPopup(ByVal iWhere)

	Dim arrRet, i
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 1					' 영업그룹 
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGrp.Value)			
		arrParam(3) = ""									
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"
		
		frm1.txtSalesGrp.focus						

	Case 2					' 거래처 
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtShipToParty.Value)		
		arrParam(3) = ""									
		arrParam(4) = "BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"						
		arrParam(5) = "납품처"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "납품처"							
		arrHeader(1) = "납품처명"						

		frm1.txtShipToParty.focus
		
	Case 3					' 품목 
		arrParam(1) = "B_ITEM"								
		arrParam(2) = Trim(frm1.txtItemCode.Value)			
		arrParam(4) = "PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "									
		arrParam(5) = "품목"							
	
		arrField(0) = "ITEM_CD"								
		arrField(1) = "ITEM_NM"								
		arrField(2) = "SPEC"								
    
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"		
		arrHeader(2) = "규격"		
		
		frm1.txtItemCode.focus							

	Case 4					'공장 
		arrParam(1) = "B_PLANT"								
		arrParam(2) = Trim(frm1.txtPlantCode.value)				
		arrParam(4) = ""									
		arrParam(5) = "공장"							
	
		arrField(0) = "PLANT_CD"							
		arrField(1) = "PLANT_NM"							
    
		arrHeader(0) = "공장"							
		arrHeader(1) = "공장명"							
		
		frm1.txtPlantCode.focus

	Case 5					'출하형태 
		arrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"				
		arrParam(2) = Trim(frm1.txtDnType.value)
		arrParam(3) = ""
		arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
		arrParam(5) = "출하형태"

		arrField(0) = "A.MINOR_CD"
		arrField(1) = "A.MINOR_NM"

		arrHeader(0) = "출하형태"
		arrHeader(1) = "출하형태명"

		frm1.txtDNType.focus
	
	Case 6	'tracking no
	
	'	Dim strRet
		
'		Dim arrTNParam(5), i
		
		Dim iCalledAspName, IntRetCD

		For i = 0 to UBound(arrParam)
			arrParam(i) = ""
		Next	

		'20021227 kangjungu dynamic popup
		iCalledAspName = AskPRAspName("s3135pa1")	
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
			lgIsOpenPop = False
			exit Function
		end if

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If arrRet = "" Then
			Exit Function
		Else
			frm1.txtTrackingNo.value = arrRet 
		End If		
		
		frm1.txtTrackingNo.focus
		Exit Function
	
	End Select

	arrParam(0) = arrParam(5)								

	Select Case iWhere
	Case 3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetConDnPopup(arrRet,iWhere)
	End If	
	
End Function

'=========================================
Function SetConDnPopup(Byval arrRet,ByVal iWhere)

	With frm1
		Select Case iWhere
		Case 1
			.txtSalesGrp.value = arrRet(0) 
			.txtSalesGrpNm.value = arrRet(1)   
		Case 2
			.txtShipToParty.value = arrRet(0) 
			.txtShipToPartyNm.value = arrRet(1)
		Case 3
			.txtItemCode.value = arrRet(0) 
			.txtItemCodeNm.value = arrRet(1)   
		Case 4
			.txtPlantCode.value = arrRet(0) 
			.txtPlantName.value = arrRet(1)   
		Case 5
			.txtDNType.value = arrRet(0) 
			.txtDNTypeNm.value = arrRet(1)   
		End Select
	End With
	
End Function

'=========================================
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

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
End Sub

'=======================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================
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

End Sub

'=======================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'=======================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub

	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
		If CheckRunningBizProcess = True Then Exit Sub
    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If
End Sub

'=======================================================
Sub rdoQueryFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg1.value
End Sub

'=======================================================
Sub rdoQueryFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg2.value
End Sub

'=======================================================
Sub rdoQueryFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg3.value
End Sub

'=======================================================
Sub txtSoDtFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtSoDtFrom.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtSoDtFrom.Focus
	End If
End Sub

'=======================================================
Sub txtSoDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtSoDtTo.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtSoDtTo.Focus
	End If
End Sub

'=======================================================
Sub txtSoDtFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=======================================================
Sub txtSoDtTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=======================================================
Function GridHeadColName(strRadio)

	Select Case strRadio
	
	Case UCase(frm1.rdoGiYes.value)
		frm1.vspdData.Col = GetKeyPos("A",8)
		frm1.vspdData.Row = 0
		frm1.vspdData.Text = "출고량"

		frm1.vspdData.Col = GetKeyPos("A",9)
		frm1.vspdData.Row = 0
		frm1.vspdData.Text = "덤출고량"

	Case UCase(frm1.rdoGiNo.value)

		frm1.vspdData.Col = GetKeyPos("A",8)
		frm1.vspdData.Row = 0
		frm1.vspdData.Text = "PICKING수량"

		frm1.vspdData.Col = GetKeyPos("A",9)
		frm1.vspdData.Row = 0
		frm1.vspdData.Text = "덤PICKING수량"

	End Select

End Function

'=======================================================
Function FncQuery() 
    Dim IntRetCD
    FncQuery = False                                                            
    Err.Clear                                                               

    If Not chkField(Document, "1") Then	Exit Function

	If ValidDateCheck(frm1.txtSoDtFrom, frm1.txtSoDtTo) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")	         						
    Call InitVariables 														
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=======================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=======================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=======================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'=======================================================
Function FncExit()
    FncExit = True
End Function

'=======================================================
Function DbQuery() 
	On Error Resume Next

	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    With frm1
		If .rdoGiYes.checked = True Then
			.txtGiFlag.value = .rdoGiYes.value
			Call GridHeadColName(UCase(.rdoGiYes.value))
		ElseIf .rdoGiNo.checked = True Then
			.txtGiFlag.value = .rdoGiNo.value
			Call GridHeadColName(UCase(.rdoGiNo.value))
		End If

		If lgIntFlgMode = parent.OPMD_UMODE Then    
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			strVal = strVal & "&txtSalesGrp=" & Trim(.HSalesGrp.value)					
			strVal = strVal & "&txtShipToParty=" & Trim(.HShipToParty.value)
			strVal = strVal & "&txtItemCode=" & Trim(.HItemCode.value)
			strVal = strVal & "&txtPlantCode=" & Trim(.HPlantCode.value)
			strVal = strVal & "&txtDNType=" & Trim(.HDNType.value)
			strVal = strVal & "&txtGiFlag=" & Trim(.txtGiFlag.value)
			strVal = strVal & "&txtSoDtFrom=" & Trim(.HSoDtFrom.value)
			strVal = strVal & "&txtSoDtTo=" & Trim(.HSoDtTo.value)
			strVal = strVal & "&txtTrackingNO=" & Trim(.HtxtTrackingNo.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)				
			strVal = strVal & "&txtShipToParty=" & Trim(.txtShipToParty.value)
			strVal = strVal & "&txtItemCode=" & Trim(.txtItemCode.value)
			strVal = strVal & "&txtPlantCode=" & Trim(.txtPlantCode.value)
			strVal = strVal & "&txtGiFlag=" & Trim(.txtGiFlag.value)
			strVal = strVal & "&txtDNType=" & Trim(.txtDNType.value)
			strVal = strVal & "&txtSoDtFrom=" & Trim(.txtSoDtFrom.Text)
			strVal = strVal & "&txtSoDtTo=" & Trim(.txtSoDtTo.Text)
			strVal = strVal & "&txtTrackingNO=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo

		End If	

		lgLngStartRow = .vspdData.MaxRows + 1
				    
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True    

End Function

'=====================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	lgIntFlgMode = parent.OPMD_UMODE
	
	Call SetToolbar("11000000000111")
	
    If frm1.vspdData.MaxRows > 0 Then
		Call FormatSpreadCellByCurrency()
       frm1.vspdData.Focus
    End if  	

End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency()
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,GetKeyPos("A",10),GetKeyPos("A",11),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,GetKeyPos("A",10),GetKeyPos("A",12),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",13),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",14),"A", "Q" ,"X","X")		
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출고현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>납기일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSoDtFrom" CLASS=FPDTYYYYMMDD tag="12X1" Alt="납기시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;~&nbsp;</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtSoDtTo" CLASS=FPDTYYYYMMDD tag="12X1" Alt="납기종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSalesGrp" TYPE="Text" ALT="영업그룹" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemCode" TYPE="Text" ALT="품목" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 3">&nbsp;<INPUT NAME="txtItemCodeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6"><INPUT NAME="txtPlantCode" TYPE="Text" ALT="공장" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 4">&nbsp;<INPUT NAME="txtPlantName" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShipToParty" TYPE="Text" ALT="납품처" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 2">&nbsp;<INPUT NAME="txtShipToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>출고여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoGiFlag" id="rdoGiYes" VALUE="Y" tag = "11" CHECKED>
											<LABEL FOR="rdoGiYes">출고</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoGiFlag" id="rdoGiNo" VALUE="N" tag = "11">
											<LABEL FOR="rdoGiNo">미출고</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5>출하형태</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtDNType" SIZE=10 MAXLENGTH=3 TAG="11XXXU" ALT="출하형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSORef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 5">&nbsp;<INPUT TYPE=TEXT NAME="txtDNTypeNm" SIZE=25 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDnPopup 6"></TD>	
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="HSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HShipToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="HItemCode" tag="24">
<INPUT TYPE=HIDDEN NAME="HPlantCode" tag="24">
<INPUT TYPE=HIDDEN NAME="HDNType" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtTrackingNo" tag="24">

<INPUT TYPE=HIDDEN NAME="txtGiFlag" tag="24">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
