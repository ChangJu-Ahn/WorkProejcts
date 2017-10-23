<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s4212qa1
'*  4. Program Name         : 통관상세조회 
'*  5. Program Desc         : 통관상세조회 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 
'**********************************************************************************************************
Dim lgIsOpenPop                                             '☜: Popup status                          
Dim lgSortTitleNm                                           '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD1                                          '☜: Orderby popup용 데이타(필드코드)      
Dim IscookieSplit 
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)
Const BIZ_PGM_ID        = "s4212qb1_KO441.asp"
Const BIZ_PGM_JUMP_ID	= "s4212ma1"
Const C_MaxKey          = 4                                    
'========================================================================================================= 
Sub InitVariables()	
	lgBlnFlgChgValue = False 
    lgStrPrevKey     = ""
    lgSortKey        = 1
    lgIntFlgMode = parent.OPMD_CMODE   
End Sub
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtOpenFrDt.Text = StartDate
	frm1.txtOpenToDt.Text = EndDate
End Sub
'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub
'==========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S4212QA1","S","A","V20030321",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock 
        
End Sub
'=========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		arrParam(3) = ""									
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "수입자"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "수입자"							
		arrHeader(1) = "수입자명"						
	Case 1
		arrParam(1) = "B_ITEM"								
		arrParam(2) = Trim(frm1.txtItem_cd.Value)			
		arrParam(3) = ""									
		arrParam(4) = ""									
		arrParam(5) = "품목"							
	
		arrField(0) = "ITEM_CD"								
		arrField(1) = "ITEM_NM"								
    
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"							

	Case 2
		arrParam(1) = "B_SALES_ORG"							
		arrParam(2) = Trim(frm1.txtSalesOrg.Value)			
		arrParam(3) = Trim(frm1.txtSalesOrgNm.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업조직"						
	
		arrField(0) = "SALES_ORG"							
		arrField(1) = "SALES_ORG_NM"						
    
		arrHeader(0) = "영업조직"						
		arrHeader(1) = "영업조직명"						

	Case 3
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		
		arrParam(3) = Trim(frm1.txtSalesGroupNm.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"							
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"							

	End Select

	arrParam(0) = arrParam(5)								
	arrParam(3) = ""
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	With frm1
		Select Case iWhere
		    Case 0
		    	.txtconBp_cd.focus
		    Case 1
		    	.txtItem_cd.focus
		    Case 2
		    	.txtSalesOrg.focus
		    Case 3
		    	.txtSalesGroup.focus
		End Select
	End With	

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenIvNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "송장번호"						
	arrParam(1) = "S_CC_HDR"							
	arrParam(2) = Trim(frm1.txtIvNo.value)					
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									
	arrParam(5) = "송장번호"						

	arrField(0) = "IV_NO"								
	arrField(1) = "CC_NO"								

	arrHeader(0) = "송장번호"						
	arrHeader(1) = "통관관리번호"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
		
	frm1.txtIvNo.focus

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetIvNo(arrRet)
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
'--------------------------------------------------------------------------------------------------------- 
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 1
			.txtItem_cd.value = arrRet(0) 
			.txtItem_Nm.value = arrRet(1)   
		Case 2
			.txtSalesOrg.value = arrRet(0)
			.txtSalesOrgNm.value = arrRet(1)  
		Case 3
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
		End Select
	End With
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetIvNo(arrRet)
	frm1.txtIvNo.Value = arrRet(0)
End Function
'==================================================================================================== 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						 

	If Kubun = 1 Then								 

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		
		WriteCookie CookieSplit , IsCookieSplit		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							

		strTemp = ReadCookie(CookieSplit)
		

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		
		Dim iniSep

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

'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================================= 
Sub Form_Load()    
    
    Call LoadInfTB19029															    
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                       

	Call InitVariables
        Call GetValue_ko441()															
	Call SetDefaultVal		
	Call InitSpreadSheet()	
    Call SetToolbar("11000000000011")							
    
    frm1.txtSalesGroup.focus
    
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

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
    
    If Row <> 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		IscookieSplit = frm1.vspdData.text
	Else
		IscookieSplit = ""
	End if
	
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

End Sub

'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
					
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End if
	End if	    
End Sub
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==========================================================================================
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg1.value
End Sub

Sub rdoTexIssueFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg2.value
End Sub

Sub rdoTexIssueFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg3.value
End Sub
'========================================================================================================
Sub txtOpenFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenFrDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtOpenFrDt.Focus
	End If
End Sub

Sub txtOpenToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenToDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtOpenToDt.Focus
	End If
End Sub
'==========================================================================================
Sub txtOpenFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtOpenToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.ClearSpreadData
	
	Call InitVariables 														
		
	If ValidDateCheck(frm1.txtOpenFrDt, frm1.txtOpenToDt) = False Then Exit Function
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
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "x", "x")
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
    
    Err.Clear                                                               
			
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
        
    With frm1

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtHconBp_cd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtHSalesGroup.value)
			strVal = strVal & "&txtItem_cd=" & Trim(.txtHItem_cd.value)
			strVal = strVal & "&txtIvNo=" & Trim(.txtHIvNo.value)
			strVal = strVal & "&txtOpenFrDt=" & Trim(.txtHOpenFrDt.value)
			strVal = strVal & "&txtOpenToDt=" & Trim(.txtHOpenToDt.value)		
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)
			strVal = strVal & "&txtIvNo=" & Trim(.txtIvNo.value)
			strVal = strVal & "&txtOpenFrDt=" & Trim(.txtOpenFrDt.Text)
			strVal = strVal & "&txtOpenToDt=" & Trim(.txtOpenToDt.Text)
		End If
		
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                              
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
                strVal = strVal & "&gBizArea=" & lgBACd 
                strVal = strVal & "&gPlant=" & lgPLCd 
                strVal = strVal & "&gSalesGrp=" & lgSGCd 
                strVal = strVal & "&gSalesOrg=" & lgSOCd 
       
        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    
    DbQuery = True


End Function
'========================================================================================
Function DbQueryOk()														
	lgIntFlgMode = parent.OPMD_UMODE
    Call SetToolbar("11000000000111")							
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
    End if  
        
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관상세</font></td>
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
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>수입자</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="수입자" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6><INPUT NAME="txtItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>송장번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=35 MAXLENGTH=35 TAG="11XXXU" ALT="송장번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvNo" align=top TYPE="BUTTON" OnClick="vbscript:OpenIvNo"></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>통관일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s4212qa1_fpDateTime1_txtOpenFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s4212qa1_fpDateTime2_txtOpenToDt.js'></script>
									</TD>
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
								<script language =javascript src='./js/s4212qa1_vaSpread1_vspdData.js'></script>
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
				<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">통관내역등록</a></TD>
				<TD WIDTH=50>&nbsp;</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>		
		
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHconBp_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHOpenFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHOpenToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>

</BODY>
</HTML>
