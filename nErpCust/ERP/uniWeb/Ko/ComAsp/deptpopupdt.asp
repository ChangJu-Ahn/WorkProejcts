<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common Module
*  2. Function Name        : Common Function
*  3. Program ID           : DeptPopup
*  4. Program Name         : 부서공통팝업 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     : Hwang Jeong Won
* 10. Modifier (Last)      : Lee Seok Min
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../inc/IncSvrCcm.inc" -->
<!-- #Include file="../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliVariables.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../inc/incImage.js"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                       

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
	Const BIZ_PGM_ID = "DeptPopupDtBiz.asp"							'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
	Const C_SHEETMAXROWS = 30
	Const CODE_CON = 0
	Const DATE_CON = 1
	Const INTERNAL = 2
	Const NAME_CON = 3
	
	Dim C_DeptCd
	Dim C_DeptNm
	Dim C_Internal

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

<!-- #Include file="../inc/lgvariables.inc" -->		
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
	Dim lgCode					  '--- Next code
	Dim lgName					  '--- Next name
	Dim arrParent
	Dim arrParam
	Dim arrReturn
	Dim gintDataCnt
	Dim lgInternal
		
	arrParent = window.dialogArguments
	arrParam = arrParent(1)
	Set PopupParent = arrParent(0)
	top.document.title = "부서 Popup"
	
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgStrPrevKey = ""
    vspdData.MaxRows = 0
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

	Dim strYear, strMonth, strDay
	lblDate.innerHTML  = "기준일자"
	lblTitle.innerHTML = "부서"	

	txtDeptCd.value    = arrparam(CODE_CON)
	'by lsm : 세번째인자 추가 - 이름으로 조회할수 있게 
	if UBound(arrParam) >= 3 then
		txtDeptNm.value    = arrParam(NAME_CON)
	end if
	If arrparam(DATE_CON) = "" Then
'		Call ExtractDateFrom("<%=GetSvrDate%>", popupparent.gServerDateFormat, popupparent.gServerDateType, strYear,strMonth,strDay)	
'		txtDate.text	=  UniConvYYYYMMDDToDate(popupparent.gDateFormat, strYear, strMonth, strDay)
		txtDate.Text    = UNIConvDateAToB("<%=GetSvrDate%>" ,popupParent.gServerDateFormat,gDateFormat)
		'txtDate.text = PopupParent.UNIDateClientFormat("<%=GetSvrDate%>")
	Else
		txtDate.text = arrparam(DATE_CON)
	End If	
		
	lgInternal		   = arrparam(INTERNAL)
	Self.Returnvalue = Array("")
End Sub

'========================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'========================================================================================================
sub InitSpreadPosVariables()
	C_DeptCd    = 1
	C_DeptNm    = 2
	C_Internal  = 3

end sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_DeptCd    = iCurColumnPos(1)
			C_DeptNm    = iCurColumnPos(2)
			C_Internal  = iCurColumnPos(3)

    End Select    
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	call InitSpreadPosVariables()
    vspdData.ReDraw = False
		    
    ggoSpread.Source = vspdData
    ggoSpread.Spreadinit "V20021212", , Popupparent.gAllowDragDropSpread
		
    vspdData.MaxCols = C_Internal + 1
    vspdData.MaxRows = 0
	vspdData.Col = vspdData.MaxCols
    vspdData.ColHidden = True
    vspdData.lock = false    
	Call GetSpreadColumnPos("A")    
	    	    
    ggoSpread.SSSetEdit C_DeptCd, "부서코드", 14,,,10
	ggoSpread.SSSetEdit C_DeptNm, "부서명"  , 44,,,40
	ggoSpread.SSSetEdit C_Internal, "내부부서코드", 16,,,10
	ggoSpread.SSSetProtected	-1,-1,-1 
	vspdData.ReDraw = True

    Call SetSpreadLock 	
    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029
		
	Call MM_preloadImages("../../Cshared/image/Query.gif","../../Cshared/image/OK.gif","../../Cshared/image/Cancel.gif")

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)

	Call InitVariables
				
	Call SetDefaultVal()
	Call InitSpreadSheet()
	lgCode = Trim(txtDeptCd.value)
	lgName = Trim(txtDeptNm.value)
	
	Call DbQuery()
		
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "A", "NOCOOKIE", "PA")%>
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")	
    gMouseClickStatus = "SPC" 
	
    Set gActiveSpdSheet = vspdData
	if vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	
	end if
	'vspdData.Row = Row
	
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if vspdData.MaxRows = 0 then
		exit sub
	end if
	If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
		  Call OKClick()
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = vspdData
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub


'========================================================================================================
' Name :vspdData_TopLeftChange
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

Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

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
		strVal = strVal     & "&txtDate=" & txtDate.text
		strVal = strVal     & "&txtCode=" & lgCode
		strVal = strVal     & "&txtName=" & lgName
		strVal = strVal     & "&txtInternal=" & lgInternal
	Else
		If txtDeptNm.value = "" Then lgName = ""
		
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal     & "&txtDate=" & txtDate.text
		strVal = strVal     & "&txtCode=" & lgCode
		strVal = strVal     & "&txtName=" & lgName
		strVal = strVal     & "&txtInternal=" & lgInternal
	End If
    strVal = strVal & "&txtMaxRows= " & vspdData.MaxRows

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
	Dim iCurColumnPos
	If vspdData.MaxRows < 1 Then
		self.close()
		Exit Function
	End If
		
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = iCurColumnPos(intColCnt + 1)
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
	set PopupParent = nothing
	Self.Close()
End Function

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0
	
	Call InitVariables
	
	lgCode = Trim(txtDeptCd.value)
	lgName = Trim(txtDeptNm.value)
		
	Call DbQuery()

End Function

'========================================================================================================
' Name : txtDate_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtDate_DblClick(Button)
	If Button = 1 Then
		txtDate.Action = 7
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDate_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDate_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		txtDeptCd.focus
		Call FncQuery()
	End If
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

</SCRIPT>
<!-- #Include file="../inc/uni2kcmcom.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>				
				<TD CLASS="TD5" ><SPAN CLASS="normal" ID="lblDate">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDATETIME tag="13X1" ALT="기준일자"></OBJECT>');</SCRIPT></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" ><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
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
				<IMG SRC="../../Cshared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../Cshared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../Cshared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


