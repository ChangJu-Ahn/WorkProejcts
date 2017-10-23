<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S3133MA1
'*  4. Program Name         : 미출하생성현황조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : S31133QuerySoNotDnSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/12/16 Include 성능향상 강준구 
'*                            -2002/12/20 : Get방식을 Post방식으로 변경 
'**********************************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                               
Const BIZ_PGM_ID 		= "s3133mb1.asp"                              
Const C_MaxKey          = 4                                           

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                    
Dim IsOpenPop     

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'=========================================
Sub InitVariables()

    lgStrPrevKey	 = ""
    lgPageNo         = ""
    lgIntFlgMode	 = parent.OPMD_CMODE               

    lgSortKey        = 1                        '정렬상태 저장 변수(내림차순,오름차순)
    
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    
End Sub

'===========================================
Sub SetDefaultVal()
	frm1.txtSoDtFrom.text = StartDate
	frm1.txtSoDtTo.text = EndDate
	
	frm1.txtDlvyDtFrom.text = StartDate
	frm1.txtDlvyDtTo.text = EndDate
	
	frm1.txtSoDtFrom.focus
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'==========================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S3133QA1","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'===========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Function OpenConDnPopup(ByVal iWhere)

	Dim arrRet, i
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 1				
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

	Case 2				
		arrParam(1) = "B_BIZ_PARTNER"					
		arrParam(2) = Trim(frm1.txtSoldToParty.Value)	
		arrParam(3) = ""								
		arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""				
		arrParam(5) = "주문처"					
	
		arrField(0) = "BP_CD"						
		arrField(1) = "BP_NM"						
    
		arrHeader(0) = "주문처"					
		arrHeader(1) = "주문처명"				

		frm1.txtSoldToParty.focus
		
	Case 3		
		arrParam(0) = "품목"
		arrParam(1) = "b_item item,b_item_by_plant item_plant"
		arrParam(2) = Trim(frm1.txtItemCode.Value)		
		arrParam(3) = ""								
		arrParam(4) = "item.item_cd=item_plant.item_cd"	
		arrParam(5) = "품목"						

		arrField(0) = "item.item_cd"					
		arrField(1) = "item.item_nm"					
		arrField(2) = "item.spec"					
    
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"							
		arrHeader(2) = "규격"							
		
		frm1.txtItemCode.focus
		
	Case 4				
		arrParam(0) = "수주형태"					
		arrParam(1) = "S_SO_TYPE_CONFIG"				
		arrParam(2) = Trim(frm1.txtSoType.value)		
		arrParam(3) = Trim(frm1.txtSoTypeNm.value)		
		arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "				
		arrParam(5) = "수주형태"					
		
	    arrField(0) = "SO_TYPE"			
	    arrField(1) = "SO_TYPE_NM"		
	    
	    arrHeader(0) = "수주형태"					
	    arrHeader(1) = "수주형태명"					

		frm1.txtSoType.focus
	Case 5	'tracking no
	
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
			IsOpenPop = False
			exit Function
		end if

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

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

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConDnPopup(arrRet,iWhere)
	End If	
	
End Function

'========================================
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

'========================================
Function SetConDnPopup(Byval arrRet,ByVal iWhere)

	With frm1
		Select Case iWhere
		Case 1
			.txtSalesGrp.value = arrRet(0) 
			.txtSalesGrpNm.value = arrRet(1)   
		Case 2
			.txtSoldToParty.value = arrRet(0) 
			.txtSoldToPartyNm.value = arrRet(1)
		Case 3
			.txtItemCode.value = arrRet(0) 
			.txtItemCodeNm.value = arrRet(1)   
		Case 4
			.txtSoType.value = arrRet(0)
			.txtSoTypeNm.value = arrRet(1)
		End Select
	End With
	
End Function

'==========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														
	Call SetDefaultVal
	Call InitSpreadSheet
End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
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

'=========================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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

'=========================================
Sub txtSoDtFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtSoDtFrom.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtSoDtFrom.Focus
	End If
End Sub

'=========================================
Sub txtSoDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtSoDtTo.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtSoDtTo.Focus
	End If
End Sub

'=========================================
Sub txtDlvyDtFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyDtFrom.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtDlvyDtFrom.Focus
	End If
End Sub

'=========================================
Sub txtDlvyDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyDtTo.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtDlvyDtTo.Focus
	End If
End Sub

'=========================================
Sub txtSoDtFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=========================================
Sub txtSoDtTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=========================================
Sub txtDlvyDtFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=========================================
Sub txtDlvyDtTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=========================================
Function FncQuery() 
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkField(Document, "1") Then	Exit Function

	If ValidDateCheck(frm1.txtSoDtFrom, frm1.txtSoDtTo) = False Then Exit Function

	If ValidDateCheck(frm1.txtDlvyDtFrom, frm1.txtDlvyDtTo) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    If DbQuery = False Then Exit Function

    FncQuery = True															

End Function

'=========================================
Function FncPrint()
    FncPrint = False                                                             
    Err.Clear                                                                    
	Call Parent.FncPrint()                                                       
    FncPrint = True                                                              
End Function

'========================================
Function FncExcel() 
    FncExcel = False                                                             
    Err.Clear                                                                    

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              
End Function

'========================================
Function FncFind() 
    FncFind = False                                                              
    Err.Clear                                                                    

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               
End Function

'========================================
Function FncExit()
    FncExit = True                                                               
End Function

'========================================
Function DbQuery() 
	
	Dim strVal
	
	DbQuery = False                                                         
    
    Err.Clear                                                               
    Call LayerShowHide(1) 
    
	frm1.txtFlgMode.Value = lgIntFlgMode
	frm1.OPMD_UMODE.Value = parent.OPMD_UMODE
	frm1.txtMode.Value = Parent.UID_M0001				
	frm1.txt_lgPageNo.Value = lgPageNo                      '☜: Next key tag
	frm1.txt_lgStrPrevKey.Value = lgStrPrevKey                      '☜: Next key tag
	frm1.txt_lgSelectListDT.Value = GetSQLSelectListDataType("A")			 
	frm1.txt_lgTailList.Value = MakeSQLGroupOrderByList("A")
	frm1.txt_lgSelectList.Value = EnCoding(GetSQLSelectList("A"))

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          

    DbQuery = True																	

End Function

'=========================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
	lgIntFlgMode	 = parent.OPMD_UMODE												<%'⊙: Indicates that current mode is Update mode%>
	
	Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미출하생성현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*">&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>수주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSoDtFrom" CLASS=FPDTYYYYMMDD tag="12X1" Alt="수주시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;~&nbsp;</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtSoDtTo" CLASS=FPDTYYYYMMDD tag="12X1" Alt="수주종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS="TD5" NOWRAP>납기일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtDlvyDtFrom" CLASS=FPDTYYYYMMDD tag="12X1" Alt="납기시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;~&nbsp;</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyDtTo" CLASS=FPDTYYYYMMDD tag="12X1" Alt="납기종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemCode" TYPE="Text" Alt="품목" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 3">&nbsp;<INPUT NAME="txtItemCodeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" TYPE="Text" Alt="주문처" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 2">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSalesGrp" TYPE="Text" Alt="영업그룹" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>수주형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoType" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 4">&nbsp;<INPUT NAME="txtSoTypeNm" TYPE="Text" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDnPopup 5"></TD>	
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="OPMD_UMODE" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HDlvyDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HDlvyDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoldToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="HItemCode" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoType" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtTrackingNo" tag="24">

<INPUT TYPE=HIDDEN NAME="txt_lgPageNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgStrPrevKey" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgMaxCount" tag="24" TABINDEX="-1">  
<INPUT TYPE=HIDDEN NAME="txt_lgSelectListDT" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgTailList" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgSelectList" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
