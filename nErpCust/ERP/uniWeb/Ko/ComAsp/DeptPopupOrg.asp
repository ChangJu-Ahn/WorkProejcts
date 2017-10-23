
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common Module
*  2. Function Name        : Common Function
*  3. Program ID           : DEPTPOPUPORG
*  4. Program Name         : 부서공통팝업 
*  5. Program Desc         : txtdate field is protected after main screen query 
*  6. Component List       :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2002/09/02
*  9. Modifier (First)     : Hwang Jeong Won
* 10. Modifier (Last)      : Hwang Jeong Won,namyo lee
* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :  2002/11/25 : ASP Standard for Include improvement
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../inc/incSvrCcm.inc"  -->
<!-- #Include file="../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<!--
========================================================================================================
=                          3.3 Client Side Script
========================================================================================================-->


<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAMain.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAEvent.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAOperation.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliVariables.vbs">			</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../inc/incImage.js">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit            

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
	Const BIZ_PGM_ID = "DeptPopupOrgBiz.asp"							'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
	Const C_SHEETMAXROWS	= 30
	Const DtForm		= 0
	Const DtTo			= 1
	Const INTERNAL		= 2
	Const DeptCd		= 3
	Const DtClassName	= 4
	
	Const C_DeptCd		= 1
	Const C_DeptNm		= 2
	Const C_Internal	= 4
	Const C_OrgChangeID	= 3

<% StartDate	= GetSvrDate %>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

	Dim arrParent
	Dim arrParam
	Dim arrReturn
	Dim gintDataCnt
	Dim lgPageNo
	Dim lgInternal
	Dim lgIntFlgMode

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
		
	'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
	arrParent		= window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)
		
	top.document.title = "부서 Popup"
	
			
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
Sub InitVariables()
	lgPageNo = ""
    ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData
    lgIntFlgMode = PopupParent.OPMD_CMODE

	' 권한관리 추가 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd	= arrParam(5)
		lgInternalCd	= arrParam(6)
		lgSubInternalCd	= arrParam(7)
		lgAuthUsrID		= arrParam(8)    
	End If

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()


	txtDeptCd.value		= arrparam(DeptCd)
	txtFromDate.text	= arrparam(DtForm)
	txtToDate.text		= arrparam(DtTo)
	lgInternal			= arrparam(INTERNAL)

' 권한관리 추가 
'	lgInternalCd		= arrparam(INTERNAL+1)
'	lgSubInternalCd		= arrparam(INTERNAL+2)
	
	if arrparam(DtClassName) = "T" Then
		Call ggoOper.SetReqAttr(txtFromDate, "Q")
		Call ggoOper.SetReqAttr(txtToDate, "Q")
	Else
		Call ggoOper.SetReqAttr(txtFromDate, "N")
		Call ggoOper.SetReqAttr(txtToDate, "N")
	End If

	Self.Returnvalue = Array("")
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	    
    vspdData.ReDraw = False
		    
    ggoSpread.Source = vspdData
    vspdData.OperationMode = 3
		
    vspdData.MaxCols = C_Internal
    vspdData.MaxRows = 0
	    
	ggoSpread.Spreadinit
		    
    ggoSpread.SSSetEdit C_DeptCd, "부서코드", 10,,,10
	ggoSpread.SSSetEdit C_DeptNm, "부서명"  , 20,,,40
	ggoSpread.SSSetEdit C_OrgChangeID, "조직개편ID", 14,,,10   
	ggoSpread.SSSetEdit C_Internal, "내부부서코드", 16,,,10
	ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
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

	Call LoadInfTB19029
	
		<% ' 이미지 효과 자바스크립트 함수 호출  %>
	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)

	Call InitVariables
				
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call DbQuery()
		
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="./LoadInfTB19029.asp"  -->
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
' Name :vspdData_TopLeftChange
' Desc : 
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	    
   
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    

End Sub
	
'========================================================================================================
' Name :DbQuery
' Desc : 
'========================================================================================================
Function DbQuery()
    Dim strVal
	
	If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
	
    
    DbQuery = False                                                         '⊙: Processing is NG
	
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal     & "&txtFromDate=" & txtFromDate.text
		strVal = strVal     & "&txtToDate=" & txtToDate.text
		strVal = strVal     & "&txtCode=" & Trim(txtDeptCd.value)
		strVal = strVal     & "&txtName=" & Trim(txtDeptNm.value)
		strVal = strVal     & "&txtInternal=" & lgInternal
	Else
				
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal     & "&txtFromDate=" & txtFromDate.text
		strVal = strVal     & "&txtToDate=" & txtToDate.text
		strVal = strVal     & "&txtCode=" & Trim(txtDeptCd.value)
		strVal = strVal     & "&txtName=" & Trim(txtDeptNm.value)
		strVal = strVal     & "&txtInternal=" & lgInternal
	End If
		strVal = strVal & "&lgPageNo="       & lgPageNo   

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장			
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

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
		Redim arrReturn(vspdData.MaxCols+1)
		
		vspdData.Row = vspdData.ActiveRow
				
		For intColCnt = 0 To vspdData.MaxCols -1

			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = Trim(vspdData.Text)
		Next

		arrReturn(intColCnt) = txtFromDate.text
		arrReturn(intColCnt+1) = txtToDate.text
		
		Self.Returnvalue = arrReturn
	End If
		
	Self.Close()
End Function

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()

	Call InitVariables
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    If CompareDateByFormat(txtFromDate.text,txtToDate.text,txtFromDate.Alt,txtToDate.Alt, _
                        "970025",txtFromDate.UserDefinedFormat,PopupParent.gComDateType,True) = False Then			
		Exit Function
    End If
	Call DbQuery()

End Function

'========================================================================================================
' Name : txtFromDate_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtFromDate_DblClick(Button)
	if arrparam(DtClassName) <> "T" Then
		If Button = 1 Then
			txtFromDate.Action = 7
			Call SetFocusToDocument("P")
			txtFromDate.Focus
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDate_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtFromDate_Keypress(KeyAscii)
    On Error Resume Next      
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
		txtToDate.focus
       vspdData.MaxRows = 0
       Call FncQuery()
    End if    
End Sub'========================================================================================================
' Name : txtFromDate_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtToDate_DblClick(Button)
	if arrparam(DtClassName) <> "T" Then
		If Button = 1 Then
			txtToDate.Action = 7
			Call SetFocusToDocument("P")
			txtToDate.Focus
		End If
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDate_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtToDate_Keypress(KeyAscii)
    On Error Resume Next      
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
	txtFromDate.focus
      vspdData.MaxRows = 0
       Call FncQuery()
    End if    
End Sub


</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>				
				<TD CLASS="TD5" >일자</TD>
				<TD CLASS="TD6" ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDATETIME tag="13X1" ALT="시작일자" id=OBJECT1></OBJECT>');</SCRIPT>
				&nbsp;~ &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDATETIME tag="13X1" ALT="끝일자" id=OBJECT1></OBJECT>');</SCRIPT>
				</TD>
			</TR>
			<TR>
				<TD CLASS="TD5" >부서코드</TD>
				<TD CLASS="TD6" ><INPUT TYPE=TEXT" Name="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" onkeypress="ConditionKeypress"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" >&nbsp;</TD>
				<TD CLASS="TD6" ><INPUT TYPE=TEXT NAME="txtDeptNm"   SIZE=30 MAXLENGTH=40 tag="11" ALT="부서명" onkeypress="ConditionKeypress"></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


