<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Comon Popup																*
'*  3. Program ID           : TermDeptPopup.asp															*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 기간별 부서팝업															*
'*  7. Modified date(First) : 2000/08/30																*
'*  8. Modified date(Last)  : 2000/08/30																*
'*  9. Modifier (First)     : Hwang Jeong Won															*
'* 10. Modifier (Last)      : Hwang Jeong Won															*
'* 11. Comment              :																			*
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/eventpopup.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js">  </SCRIPT>

<Script Language="VBScript">
Option Explicit   

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

	Const BIZ_PGM_ID = "TermDeptPopupBiz.asp"						 '☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
	Const C_SHEETMAXROWS = 30							              '--- 한화면에 보일수 있는 최대 Row 수 
	Const CODE_CON = 0												  '--- Index of Code Condition value
	Const INTERNAL = 1
	Const C_OrgId = 1
	Const C_ChangeDt = 2
	Const C_DeptCd = 3
	Const C_DeptNm = 4
	Const C_InternalCd = 5

<% EndDate= GetSvrDate %>		
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
	Dim lgOrgId, lgCode, lgChangeDt, lgName
	
	Dim arrParent
	Dim arrParam
	Dim arrReturn
	Dim gintDataCnt
	Dim lgStrPrevKey
	Dim lgInternal
	Dim lgIntFlgMode
		
	arrParent = window.dialogArguments
	arrParam = arrParent(0)
	
	top.document.title = "기간별 부서 Popup"
			
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()
	lgStrPrevKey     = ""
    vspdData.MaxRows = 0
    lgIntFlgMode = OPMD_CMODE
End Function

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
<%	
	On Error Resume Next
	
	Dim pB29013
	Dim dtDate

	Set	pB29013 = Server.CreateObject("B29013.B29013LookupAcctDeptCurrDt")

	pB29013.Execute
		
	dtDate = pB29013.ExportBAcctDeptOrgChangeDt
	   
	Set pB29013 = Nothing
	
    IF IsEmpty(dtDate) Then
		dtDate = EndDate 'default value
	End If
    
	On Error Goto 0
%>		
	txtFromDate.text = "<%=UNIDateClientFormat(dtDate)%>"
	txtToDate.text   = "<%=UNIDateClientFormat(EndDate)%>"
			
	lblDate.innerHTML = "기준일자"
	lblTitle.innerHTML = "부서Popup"
	txtCd.value = arrparam(CODE_CON)
	lgInternal = arrparam(INTERNAL)	
	Self.Returnvalue = Array("")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="ComLoadInfTB19029.asp" -->
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	    
    vspdData.ReDraw = False
		    
    ggoSpread.Source = vspdData
    vspdData.OperationMode = 3
		
    vspdData.MaxCols = 5
    vspdData.MaxRows = 0
	    
	ggoSpread.Spreadinit		
			
	ggoSpread.SSSetEdit C_OrgId     , "변경ID"      , 8 , ,,5
    ggoSpread.SSSetDate C_ChangeDt  , "변경일"      , 10,2,gDateFormat	    
    ggoSpread.SSSetEdit C_DeptCd    , "부서코드"    , 14, ,,10
	ggoSpread.SSSetEdit C_DeptNm    , "부서명"      , 30, ,,200
	ggoSpread.SSSetEdit C_InternalCd, "내부부서코드", 10, ,,10

	vspdData.ReDraw = True
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	
	Call MM_preloadImages("../image/Query.gif","../image/OK.gif","../image/Cancel.gif")
	Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
	Call InitVariables
		
	Call SetDefaultVal()
	Call InitSpreadSheet()
	lgCode = Trim(txtCd.value)
	lgName = Trim(txtNm.value)
	Call DbQuery()
	
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

	
'========================================================================================================
' Name :
' Desc : 
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
    If vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then
		If lgStrPrevKey <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
     			DbQuery
   		End If
    End if
End Sub
	
'========================================================================================================
' Name :
' Desc : 
'========================================================================================================
Function DbQuery()
    Dim strVal
	
	If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
    DbQuery = False			                                                   '⊙: Processing is NG

	If lgIntFlgMode = OPMD_CMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
		strVal = strVal & "&txtFromDate=" & txtFromDate.text
		strVal = strVal & "&txtToDate="   & txtToDate.text
		strVal = strVal & "&txtOrgId="    & lgOrgId
		strVal = strVal & "&txtCode="     & lgCode
		strVal = strVal & "&txtChangeDt=" & txtFromDate.text
		strVal = strVal & "&txtName="     & lgName
		strval = strval & "&txtUser="     & arrparam(1)
		strval = strval & "&txtInternal=" & lgInternal
    Else
		If txtNm.value = "" Then lgName = ""
		
		strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
		strVal = strVal & "&txtFromDate=" & txtFromDate.text
		strVal = strVal & "&txtToDate="   & txtToDate.text
		strVal = strVal & "&txtOrgId="    & lgOrgId
		strVal = strVal & "&txtCode="     & lgCode
		strVal = strVal & "&txtChangeDt=" & lgChangeDt
		strVal = strVal & "&txtName="     & lgName
		strval = strval & "&txtUser="     & arrparam(1)
		strval = strval & "&txtInternal=" & lgInternal
    End If
    
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    DbQuery = True                                                          '⊙: Processing is NG
	    
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call FncQuery()
	End If
End sub

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Function vspdData_DblClick( Col,  Row)
	If vspdData.MaxRows > 0 Then
       If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
		  Call OKClick()
       End If
	End If
End Function
'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function
'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Function Document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function
'========================================================================================================
' Function Name : MousePointer
' Function Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
           case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function OKClick()
	Dim intColCnt
	
	If vspdData.MaxRows < 1 Then
		self.close()
		Exit Function
	End If
		
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
				
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
		
	Self.Close()
End Function

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0

	Call InitVariables
	
	lgCode = Trim(txtCd.value)
	lgName = Trim(txtNm.value)
	lgChangeDt = txtFromDate.text

	Call DbQuery()

End Function

'========================================================================================================
' Name : txtFromDate_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtFromDate_DblClick(Button)
	If Button = 1 Then
		txtFromDate.Action = 7
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDate_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtFromDate_Keypress(Key)
    On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================================
' Name : txtToDate_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtToDate_DblClick(Button)
	If Button = 1 Then
		txtToDate.Action = 7
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDate_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtToDate_Keypress(Key)
    On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>				
				<TD CLASS="TD5" STYLE="WIDTH:35%"><SPAN CLASS="normal" ID="lblDate">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" STYLE="WIDTH:65%">
				<script language =javascript src='./js/termdeptpopup_Date_txtFromDate.js'></script>
				 ~		
				<script language =javascript src='./js/termdeptpopup_Date_txtToDate.js'></script>
			</TR>
			<TR>
				<TD CLASS="TD5" STYLE="WIDTH:35%"><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" STYLE="WIDTH:65%"><INPUT TYPE="Text" Name="txtCd" MAXLENGTH=10 SIZE=20 tag="12XXXU" onkeypress="ConditionKeypress"></TD>
			</TR>		
			<TR>
				<TD CLASS="TD5" STYLE="WIDTH:35%">&nbsp;</TD>
				<TD CLASS="TD6" STYLE="WIDTH:65%"><INPUT TYPE="Text" NAME="txtNm" MAXLENGTH=200 SIZE=30 tag="12" onkeypress="ConditionKeypress"></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/termdeptpopup_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="TermDeptBiz.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

