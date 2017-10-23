<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01oa1.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim IsOpenPop
Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
End Sub

'========================== 2.2.6 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1002", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemClass1, lgF0, lgF1, Chr(11))
	Call SetCombo2(frm1.cboItemClass2, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================
' Function Name : parent.TB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "OA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("10000000000011")	
    Call InitComboBox		
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    frm1.cboAccount.focus
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
Function FncQuery()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncSave()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncNew()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncDelete()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncInsertRow()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncDeleteRow()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncCopy()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncCancel()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

Function FncFind()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint() 
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	
	Dim var7
	Dim var8
	Dim var9
	Dim var11
	Dim var12
	Dim strUrl, strEbrFile
	

	Call BtnDisabled(1)
	
	If frm1.txtItemCd1.value = "" Then 
		var1 = "0"	
	Else
		var1 = UCase(Trim(frm1.txtItemCd1.value))
	End If
	
	If frm1.txtItemCd2.value = "" Then 
		var2 = "zzzzzzzzzzzzzzzzzz"	
	else
		var2 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var3 = "B_ITEM.ITEM_CD"										 
	Else 
		var3 = "B_ITEM.ITEM_NM"	
	End If

	If frm1.rdoValidFlg2.checked  = True Then	  
		var4 = "Y"									
		var5 = "Y"											 
	ElseIf frm1.rdoValidFlg3.checked  = True Then
		var4 = "N"	
		var5 = "N"								
	Else
		var4 = "Y"	
		var5 = "N"											 
	End If
	
	If frm1.cboAccount.value = "" Then 
		var7 = "0"	
	Else
		var7 = UCase(Trim(frm1.cboAccount.value))
	End If
	
	If frm1.cboAccount.value = "" Then 
		var8 = "zz"	
	Else
		var8 = UCase(Trim(frm1.cboAccount.value))
	End If
	
	If Trim(frm1.txtItemGroupCd.value) = "" Then 
		var9 = ""	
	Else
		var9 = "and b_item.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp( " & FilterVar(UCase(frm1.txtItemGroupCd.value), "''", "S") & "))"
	End If
	
	If frm1.cboItemClass1.value = "" Then 
		var11 = "0"	
	Else
		var11 = UCase(Trim(frm1.cboItemClass1.value))
	End If
	
	If frm1.cboItemClass2.value = "" Then 
		var12 = "zzzzzzzzzzzz"	
	Else
		var12 = UCase(Trim(frm1.cboItemClass2.value))
	End If
	
	strEbrFile = AskEBDocumentName("B1B01OA1", "EBR")

	strUrl = "item_cd1|" & var1 
	strUrl = strUrl & "|item_cd2|" & var2 
	strUrl = strUrl & "|sort_by|" & var3
	strUrl = strUrl & "|valid_flg|" & var4 
	strUrl = strUrl & "|valid_flg2|" & var5
	strUrl = strUrl & "|item_acct1|" & var7 
	strUrl = strUrl & "|item_acct2|" & var8
	strUrl = strUrl & "|cond|" & var9
	strUrl = strUrl & "|item_class1|" & var11
	strUrl = strUrl & "|item_class2|" & var12
	
'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
	call FncEBRprint(EBAction, strEbrFile, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
    
    frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement
    
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 
	Dim strUrl, strEbrFile

	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5

	Dim var7
	Dim var8
	Dim var9
	Dim var11
	Dim var12
	
	Call BtnDisabled(1)
	
	if frm1.txtItemCd1.value = "" then 
		var1 = "0"	
	else
		var1 = UCase(Trim(frm1.txtItemCd1.value))
	End If
	
	if frm1.txtItemCd2.value = "" then 
		var2 = "zzzzzzzzzzzzzzzzzz"	
	else
		var2 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	If frm1.rdoSortBy1.checked  = True Then	  
		var3 = "B_ITEM.ITEM_CD"										 
	Else 
		var3 = "B_ITEM.ITEM_NM"	
	End If

	If frm1.rdoValidFlg2.checked  = True Then	  
		var4 = "Y"									
		var5 = "Y"											 
	ElseIf frm1.rdoValidFlg3.checked  = True Then
		var4 = "N"	
		var5 = "N"								
	Else
		var4 = "Y"	
		var5 = "N"											 
	End If
		
	If frm1.cboAccount.value = "" Then 
		var7 = "0"	
	Else
		var7 = UCase(Trim(frm1.cboAccount.value))
	End If
	
	If frm1.cboAccount.value = "" Then 
		var8 = "zz"	
	Else
		var8 = UCase(Trim(frm1.cboAccount.value))
	End If
	
	If Trim(frm1.txtItemGroupCd.value) = "" Then 
		var9 = ""	
	Else
		var9 = "and b_item.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp( " & FilterVar(UCase(frm1.txtItemGroupCd.value), "''", "S") & "))"
	End If
	
	If frm1.cboItemClass1.value = "" Then 
		var11 = "0"	
	Else
		var11 = UCase(Trim(frm1.cboItemClass1.value))
	End If
	
	If frm1.cboItemClass2.value = "" Then 
		var12 = "zzzzzzzzzzzz"	
	Else
		var12 = UCase(Trim(frm1.cboItemClass2.value))
	End If
	
	strEbrFile = AskEBDocumentName("B1B01OA1", "EBR")

	strUrl = "item_cd1|" & var1 
	strUrl = strUrl & "|item_cd2|" & var2 
	strUrl = strUrl & "|sort_by|" & var3
	strUrl = strUrl & "|valid_flg|" & var4 
	strUrl = strUrl & "|valid_flg2|" & var5
	strUrl = strUrl & "|item_acct1|" & var7 
	strUrl = strUrl & "|item_acct2|" & var8
	strUrl = strUrl & "|cond|" & var9
	strUrl = strUrl & "|item_class1|" & var11
	strUrl = strUrl & "|item_class2|" & var12
	
	call FncEBRPrevIew(strEbrFile, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

Function SetParmameter()

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'----------  Coding part  -------------------------------------------------------------
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenItemCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Or UCase(frm1.txtitemcd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd1.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtItemCd1.value = arrRet(0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd1.focus
		
End Function

Function OpenItemCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Or UCase(frm1.txtitemcd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd2.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtItemCd2.value = arrRet(0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd2.focus
		
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  " 			
	arrParam(5) = "품목그룹"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemGroupCd.Value	= arrRet(0)		
		lgBlnFlgChgValue		= True
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목정보출력</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=2>								
								<TR>
									<TD CLASS=TD5 NOWRAP>품목계정</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="품목계정" STYLE="Width: 168px;" tag="X1XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="X1XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>집계용품목클래스</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass1" ALT="집계용품목클래스" STYLE="Width: 168px;" tag="X1XXXU"><OPTION VALUE=""></OPTION></SELECT>&nbsp;~&nbsp;
														<SELECT NAME="cboItemClass2" ALT="집계용품목클래스" STYLE="Width: 168px;" tag="X1XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=30 MAXLENGTH=18 tag="X1XXXU"  ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd1()">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
  									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=30 MAXLENGTH=18 tag="X1XXXU"  ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd2()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
					<TD HEIGHT=10 WIDTH=100%>
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>	
								<TR>
									<TD CLASS="TD5" NOWRAP>유효구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" CHECKED ID="rdoValidFlg1" tag="21"><LABEL FOR="rdoValidFlg1">전체</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" tag="21"><LABEL FOR="rdoValidFlg2">예</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg3" tag="21"><LABEL FOR="rdoValidFlg3">아니오</LABEL><BR>																	
									</TD>
								</TR>							
								<TR>
									<TD CLASS="TD5" NOWRAP>출력순서</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSortBy" CHECKED ID="rdoSortBy1" tag="21"><LABEL FOR="rdoSortBy1">품목코드순</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSortBy" ID="rdoSortBy2" tag="21"><LABEL FOR="rdoSortBy2">품목명순</LABEL><BR>																	
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON></TD>		
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
