<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가관리 
'*  3. Program ID           : c2216oa1
'*  4. Program Name         : 표준원가 출력 
'*  5. Program Desc         : 표준원가 출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/15
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Hyo Seok, Seo
'* 10. Modifier (Last)      : Cho Ig Sung
'* 11. Comment              :
'=======================================================================================================
 -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================	 -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim lgBlnFlgChgValue
Dim lgIntFlgMode
Dim lgIntGrpCount


Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
         
    IsOpenPop = False     
End Sub

Sub SetDefaultVal()
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "PA") %>
End Sub

Function OpenMinor(ByVal iMinor)
	Dim arrRet,itemacct
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iMinor
	Case 0
		arrParam(0) = "품목계정팝업"
		arrParam(1) = "B_MINOR a,b_item_acct_inf b"
		arrParam(2) = Trim(frm1.txtItemAccntCd.value)
		'arrParam(3) = Trim(frm1.txtItemAccntNm.value)
		arrParam(4) = "MAJOR_CD=" & FilterVar("P1001", "''", "S") & " and a.minor_cd = b.item_acct and b.item_acct_group <> " & FilterVar("6MRO","''","S")	 			' Where Condition
		arrParam(5) = "품목계정"
		
	    arrField(0) = "MINOR_CD"
	    arrField(1) = "MINOR_NM"
	    
	    arrHeader(0) = "품목계정코드"
	    arrHeader(1) = "품목계정명"
	Case 1
		arrParam(0) = "품목팝업"
		arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"
		arrParam(2) = Trim(frm1.txtItemCdFrom.value)
		'arrParam(3) = Trim(frm1.txtItemNmFrom.value)

		itemacct = Trim(frm1.txtItemAccntCd.value)
		IF itemacct = "" Then
				 itemacct = "%"
		END If
	
		arrParam(4) = "a.item_cd = b.item_cd and b.plant_cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "" _
			& " and a.valid_flg = " & FilterVar("y", "''", "S") & "  and a.valid_from_dt <= getdate() and a.valid_to_dt >= getdate() " _
			& " and b.item_acct LIKE  " & FilterVar(itemacct, "''", "S") & " order by a.item_cd"
		
		arrParam(5) = "품목"
		
	    arrField(0) = "a.item_cd"
	    arrField(1) = "a.item_nm"
	    
	    arrHeader(0) = "품목코드"
	    arrHeader(1) = "품목명"
	Case 2
		arrParam(0) = "품목팝업"
		arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"
		arrParam(2) = Trim(frm1.txtItemCdTo.value)
		'arrParam(3) = Trim(frm1.txtItemNmTo.value)

		itemacct = Trim(frm1.txtItemAccntCd.value)
		IF itemacct = "" Then
				 itemacct = "%"
		END If
	
		arrParam(4) = "a.item_cd = b.item_cd and b.plant_cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "" _
			& " and a.valid_flg = " & FilterVar("y", "''", "S") & "  and a.valid_from_dt <= getdate() and a.valid_to_dt >= getdate() " _
			& " and b.item_acct LIKE  " & FilterVar(itemacct, "''", "S") & " order by a.item_cd"
		
		arrParam(5) = "품목"
		
	    arrField(0) = "a.item_cd"
	    arrField(1) = "a.item_nm"
	    
	    arrHeader(0) = "품목코드"
	    arrHeader(1) = "품목명"
	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If iMinor = 0 Then
			frm1.txtItemAccntCD.focus
		End If
		Exit Function
	Else
		Call SetMinor(arrRet,iMinor)
	End If	
End Function

Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCD.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

Function SetMinor(Byval arrRet,ByVal iMinor)

If arrRet(0) <> "" Then 
	Select Case iMinor
	Case 0
		frm1.txtItemAccntCD.focus
		frm1.txtItemAccntCd.value = arrRet(0)
		frm1.txtItemAccntNm.value = arrRet(1)
	Case 1
		frm1.txtItemCdFrom.value = arrRet(0)
		frm1.txtItemNmFrom.value = arrRet(1)
	Case 2
		frm1.txtItemCdTo.value = arrRet(0)
		frm1.txtItemNmTo.value = arrRet(1)
	end select
End If

End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
			
End Function

Function OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("125000","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True
	
	select case iWhere

	case 0
		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
		arrParam(1) = Trim(frm1.txtItemCdFrom.value)	' Item Code
		arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
	

		arrField(0) = 1 								' Field명(0) :"ITEM_CD"
		arrField(1) = 2									' Field명(1) :"ITEM_NM"

		arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
				
		IsOpenPop = False
	
		If arrRet(0) = "" Then
			frm1.txtItemCDFrom.focus
			Exit Function
		Else
			Call SetPopUp(arrRet,iWhere)
		End If
	case 1		
		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
		arrParam(1) = Trim(frm1.txtItemCdTo.value)	' Item Code
		arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
	

		arrField(0) = 1 								' Field명(0) :"ITEM_CD"
		arrField(1) = 2									' Field명(1) :"ITEM_NM"

		arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
				
		IsOpenPop = False
	
		If arrRet(0) = "" Then
			frm1.txtItemCdTo.focus
			Exit Function
		Else
			Call SetPopUp(arrRet,iWhere)
		End If
	end select

End Function

 '==========================================  2.4.3 SetPopup()  =============================================
'	Name : SetPopup()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet,Byval iWhere)
	With frm1
		select case iWhere
		case 0
			frm1.txtItemCDFrom.focus
			.TxtItemCdFrom.Value = arrRet(0)
			.TxtItemNmFrom.Value = arrRet(1)
		case 1
			frm1.TxtItemCdTo.focus
			.TxtItemCdTo.Value = arrRet(0)
			.TxtItemNmTo.Value = arrRet(1)
		end select 
		lgBlnFlgChgValue = True
		
	End With
	
End Function


Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call SetDefaultVal
    Call SetToolbar("10000000000011")
    
	frm1.OptSumFlag1.checked = True
    frm1.txtPlantCd.focus
   	Set gActiveElement = document.activeElement		    
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


Function FncQuery() 
    BtnPreview() 
End Function

Function FncBtnPrint() 
	dim strUrl, StrEbrFile
	dim var1,var2,var3,var4
	
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

	Call BtnDisabled(1)

	if frm1.OptSumFlag1.checked = True then
		StrEbrFile = "c2210oa1"
	else
		StrEbrFile = "c2210oa2"
	end if
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
	var1 = Trim(frm1.txtPlantCd.value)
	var2 = Trim(frm1.txtItemAccntCD.value)
	var3 = Trim(frm1.txtItemCDFrom.value)
	var4 = Trim(frm1.txtItemCDTo.value)
	
	if var2 = "" then
		var2 = "%"
	End if	

	if var3 = "" then
		var3 = "000000"
	End if	

	if var4 = "" then
		var4 = "zzzzzz"
	End if	
	
	

	strUrl = strUrl & "plantcd|" & var1
	strUrl = strUrl & "|itemacct|" & var2
	strUrl = strUrl & "|itemcdFrom|" & var3
	strUrl = strUrl & "|itemcdTo|" & var4
	
	call FncEBRprint(EBAction, ObjName, strUrl)

	Call BtnDisabled(0)	
		
End Function

Function BtnPreview() 
    
	dim strUrl, StrEbrFile
	dim var1,var2,var3,var4
	
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

	Call BtnDisabled(1)

	if frm1.OptSumFlag1.checked = True then
		StrEbrFile = "c2210oa1"
	else
		StrEbrFile = "c2210oa2"
	end if
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
	var1 = Trim(frm1.txtPlantCd.value)
	var2 = Trim(frm1.txtItemAccntCD.value)
	var3 = Trim(frm1.txtItemCDFrom.value)
	var4 = Trim(frm1.txtItemCDTo.value)
	
	if var2 = "" then
		var2 = "%"
	End if	

	if var3 = "" then
		var3 = "0"
	End if	

	if var4 = "" then
		var4 = "ZZZZZZZZZZZZZ"
	End if	
	
	strUrl = strUrl & "plantcd|" & var1
	strUrl = strUrl & "|itemacct|" & var2
	strUrl = strUrl & "|itemcdFrom|" & var3
	strUrl = strUrl & "|itemcdTo|" & var4
		
	call FncEBRPreview(ObjName, strUrl)

	Call BtnDisabled(0)	
	
End Function

Function FncExit()
	FncExit = True
End Function

Function FncPrint()
    Call parent.FncPrint()
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>표준원가출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="OptSumFlag" CHECKED ID="OptSumFlag1" VALUE="Y" tag="25"><LABEL FOR="OptSumFlag1">집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="OptSumFlag" ID="OptSumFlag2" VALUE="N" tag="25"><LABEL FOR="OptSumFlag2">상세</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">공장</TD>
								<TD CLASS="TD6"><INPUT  ClASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
												<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=30  ALT ="공장명" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">품목계정</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemAccntCD" MAXLENGTH="2" SIZE=10  ALT ="품목계정" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAccntCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMinor(0)">
												<INPUT NAME="txtItemAccntNM" MAXLENGTH="30" SIZE=30  ALT ="품목계정명" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">품목</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemCDFrom" MAXLENGTH="18" SIZE=10  ALT ="품목" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(0)">
												<INPUT NAME="txtItemNMFrom" MAXLENGTH="30" SIZE=30  ALT ="품목명" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">~</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemCDTo" MAXLENGTH="18" SIZE=10  ALT ="품목" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
												<INPUT NAME="txtItemNMTo" MAXLENGTH="30" SIZE=30  ALT ="품목명" tag="14X"></TD>
							</TR>

						</TABLE>
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
					<TD>
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

