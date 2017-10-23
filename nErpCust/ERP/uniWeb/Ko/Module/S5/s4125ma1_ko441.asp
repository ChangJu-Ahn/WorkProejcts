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

Const BIZ_PGM_ID        = "s4125mb1_ko441.asp"
Const C_MaxKey          = 100                                    '☆☆☆☆: Max key value


'=========================================
Sub InitVariables()
    lgPageNo     = ""                                  
    lgSortKey        = 1
    lgIntFlgMode     = parent.OPMD_CMODE						   

    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 

End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtSoDtFrom.text = EndDate

   Call ggoOper.FormatDate(frm1.txtSoDtFrom, Parent.gDateFormat, 2)

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
	Dim ii, Tempval

	Call SetZAdoSpreadSheet("S4125MA1_ko441","S","A","V20030711", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	With 	frm1.vspdData
	    .ColHeaderRows = 2

 		For ii = 6 to .MaxCols
			.Col = ii
			.Row = 0
			TempVal =  .text
			.row = 1
			.Text = Tempval
			.Row = 0
			.Text =  ""

		Next 

		Call .AddCellSpan(0,-1000, 1, 2)					'컬럼제목을 -1000 은 합치고 -1000+1 은 놔둔다
		Call .AddCellSpan(1,-1000, 1, 2) 
		Call .AddCellSpan(2,-1000, 1, 2) 
		Call .AddCellSpan(3,-1000, 1, 2) 
		Call .AddCellSpan(4,-1000, 1, 2) 
		Call .AddCellSpan(5,-1000, 1, 2) 

		Call .AddCellSpan(6,-1000, 2, 1)
		.Row = -1000 : .Col = 6 : .Text = "BOH"
		Call .AddCellSpan(8,-1000, 2, 1)
		.Row = -1000 : .Col = 8 : .Text = "IN"
		Call .AddCellSpan(10,-1000, 2, 1)
		.Row = -1000 : .Col = 10 : .Text = "OUT"
		Call .AddCellSpan(12,-1000, 2, 1)
		.Row = -1000 : .Col = 12 : .Text = "RETURN"
		Call .AddCellSpan(14,-1000, 2, 1)
		.Row = -1000 : .Col = 14 : .Text = "LOSS"
		Call .AddCellSpan(16,-1000, 2, 1)
		.Row = -1000 : .Col = 16 : .Text = "BONUS"
		Call .AddCellSpan(18,-1000, 2, 1)
		.Row = -1000 : .Col = 18 : .Text = "EOH"

		.RowHeight(-1000+1) = 14
	End with				

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
		Case 2
			.txtShipToParty.value = arrRet(0) 
			.txtShipToPartyNm.value = arrRet(1)
		Case 4
			.txtPlantCode.value = arrRet(0) 
			.txtPlantName.value = arrRet(1)   
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


Function txtPlantCode_OnChange()
    If  frm1.txtPlantCode.value <> "" Then
        if   CommonQueryRs(" plant_nm "," B_PLANT "," plant_cd =  " & FilterVar(frm1.txtPlantCode.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtPlantName.value = ""
            Call  DisplayMsgBox("970000", "x","공장코드","x")
	        frm1.txtPlantCode.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtPlantName.value = Replace(lgF0, Chr(11), "")
	    End If
	else 
		 frm1.txtPlantName.value=""
    End If

End Function


Function txtShipToParty_OnChange()
    If  frm1.txtShipToParty.value <> "" Then
        if   CommonQueryRs(" bp_nm "," B_BIZ_PARTNER "," bp_cd =  " & FilterVar(frm1.txtShipToParty.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtShipToPartyNm.value = ""
            Call  DisplayMsgBox("970000", "x","업체코드","x")
	        frm1.txtShipToParty.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtShipToPartyNm.value = Replace(lgF0, Chr(11), "")
	    End If
	else 
		 frm1.txtShipToPartyNm.value=""
    End If

End Function

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
Function FncQuery() 
    Dim IntRetCD
    FncQuery = False                                                            
    Err.Clear                                                               

    If Not chkField(Document, "1") Then	Exit Function


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
		If lgIntFlgMode = parent.OPMD_UMODE Then    
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			strVal = strVal & "&txtShipToParty=" & Trim(.HShipToParty.value)
			strVal = strVal & "&txtGiFlag=" & Trim(.txtGiFlag.value)
			strVal = strVal & "&txtSoDtFrom=" & Trim(.HSoDtFrom.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			strVal = strVal & "&txtShipToParty=" & Trim(.txtShipToParty.value)
			strVal = strVal & "&txtGiFlag=" & Trim(.txtGiFlag.value)
			strVal = strVal & "&txtSoDtFrom=" & replace(Trim(.txtSoDtFrom.Text),"-","")
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
	'Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,GetKeyPos("A",6),GetKeyPos("A",10),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",6),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",7),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",8),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",9),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",10),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",11),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",12),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",13),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",14),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",15),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",16),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",17),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",18),"A", "Q" ,"X","X")		
	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,lgLngStartRow,frm1.vspdData.MaxRows,parent.gCurrency,GetKeyPos("A",19),"A", "Q" ,"X","X")		
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>마감-수불양식</font></td>
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
									<TD CLASS="TD5" NOWRAP>수불년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSoDtFrom" CLASS=FPDTYYYYMMDD tag="12X1" Alt="출하시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>업체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShipToParty" TYPE="Text" ALT="업체" MAXLENGTH="10" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 2">&nbsp;<INPUT NAME="txtShipToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
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
