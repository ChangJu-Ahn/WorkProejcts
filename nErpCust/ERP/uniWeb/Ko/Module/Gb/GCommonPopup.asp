<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GCommonPopup
*  4. Program Name         : 경영손익 작업실행 에러 조회화면
*  5. Program Desc         : 경영손익 작업실행 에러 조회화면
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/04
*  8. Modified date(Last)  : 2002/01/04
*  9. Modifier (First)     : Kwon Ki Soo
* 10. Modifier (Last)      : Kwon Ki Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/UNI2KCM.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/eventpopup.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID = "GCommonPopupBiz.asp"                                      '☆: 비지니스 로직 ASP명

Const POPUP_TITLE = 0                                                           '--- Index of POP-UP Title
Const TABLE_NAME  = 1                                                           '--- Index of DB table name to query
Const CODE_CON    = 2                                                           '--- Index of Code Condition value
Const NAME_CON    = 3                                                           '--- Index of Name Condition value
Const WHERE_CON   = 4                                                           '--- Index of Where Clause
Const TEXT_NAME   = 5                                                           '--- Index of Textbox Name

Const C_SHEETMAXROWS = 30                                                       '--- 한화면에 보일수 있는 최대 Row 수

Dim lgStrCodeKey
Dim lgStrNameKey
Dim lgStrPrevKeyIndex

Dim arrParent
Dim arrParam					 '--- First Parameter Group
Dim arrTblField				 '--- Second Parameter Group(DB Table Field Name)
Dim arrGridHdr				 '--- Third Parameter Group(Column Captions of the SpreadSheet)
Dim arrReturn				 '--- Return Parameter Group
Dim gintDataCnt				 '--- Data Counts to Query

		'------ Set Parameters from Parent ASP ------
		arrParent = window.dialogArguments
		arrParam = arrParent(0)
		arrTblField = arrParent(1)
		arrGridHdr = arrParent(2)

		top.document.title = arrParam(POPUP_TITLE)

'========================================================================================================
Function InitVariables()
		Dim intLoopCnt

		lgStrCodeKey = ""
		lgStrNameKey = ""

		gintDataCnt = 0

		For intLoopCnt = 0 To Ubound(arrTblField)
			If arrTblField(intLoopCnt) <> "" Then
				gintDataCnt = gintDataCnt + 1
			Else
				Exit For
			End If
		Next
End Function

'========================================================================================================
Sub SetDefaultVal()

	lblTitle.innerHTML = arrParam(TEXT_NAME)
	txtCd.value = arrParam(CODE_CON)
	txtNm.value = arrParam(NAME_CON)

	Self.Returnvalue = Array("")
End Sub

'========================================================================================================
Sub InitSpreadSheet()
    Dim i
    Dim iArr
    Dim iLen

    vspdData.ReDraw = False

	    ggoSpread.Source = vspdData
	    vspdData.OperationMode = 3

	    vspdData.MaxCols = gintDataCnt
	    vspdData.MaxRows = 0

		ggoSpread.Spreadinit

		ggoSpread.SSSetEdit 1, arrGridHdr(0), 8	' 코드

		ggoSpread.SSSetEdit 2, arrGridHdr(1), 100,,,200

		vspdData.ReDraw = True
	End Sub

'========================================================================================================
Sub Form_Load()

    Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

	Call InitVariables
	Call SetDefaultVal()
	Call InitSpreadSheet()

	Call Search_OnClick()
End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/ComLoadInfTB19029.asp" -->
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
Function OKClick()
	Dim intColCnt

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
Function CancelClick()
	Self.Close()
End Function

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
Sub Search_OnClick()

    vspdData.MaxRows = 0
    lgStrCodeKey = ""
    lgStrNameKey = ""

	Call DbQuery()

End Sub

'========================================================================================================
Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call Search_OnClick()
	End If
End sub

'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop)  Then
       If lgStrCodeKey <> ""  Then
 		  Call DbQuery
       End If
    End if
End Sub

'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strPreWhere, strWhere
    Dim iLoop
    Dim arrStrVal

    DbQuery = False
    
    strPreWhere = ""
    strWhere = ""

	Call LayerShowHide(1)	
	
    If UCase(Trim(arrParam(WHERE_CON))) <> "" Then
		strPreWhere = UCase(Trim(arrParam(WHERE_CON))) & " AND "
    End If


   '----- Code가 있을 경우는 Name에 상관없이 Code로만 조회하고, Code가 없는 경는 Name으로 조회한다.
	If lgStrCodeKey <> "" Or lgStrNameKey <> "" Then
		If Trim(txtNm.value) <> "" Then
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(1)) & ">=" & Trim(UCase(lgStrNameKey)) & " Order by " & Trim(arrTblField(1))
		Else
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(0)) & ">=" & Trim(UCase(lgStrCodeKey)) & " Order by " & Trim(arrTblField(0))
		End If
	Else
		if Trim(txtNm.value) <> "" Then
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(1)) & ">=0" & " Order by " & Trim(arrTblField(1))
		Else
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(0)) & ">=0" & " Order by " & Trim(arrTblField(0))
			End If
		End If

	    DbQuery = False                                                         '⊙: Processing is NG

	    arrStrVal = ""

	    For iLoop = 0 To gintDataCnt - 1
	        arrStrVal = arrStrVal & Trim(arrTblField(iLoop)) & gColSep
	    Next
	    
	    strVal=""
	    strVal = BIZ_PGM_ID & "?txtTable=" & Trim(arrParam(TABLE_NAME))
	    strVal = strVal & "&txtWhere="    & strWhere
	    strVal = strVal & "&gintDataCnt=" & gintDataCnt
		strVal = strVal & "&arrField="    & arrStrVal

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동

	    DbQuery = True                                                          '⊙: Processing is NG
	End Function

'========================================================================================================
Function DbQueryOk()
   Dim IntRetCD

   If vspdData.MaxRows = 0 Then
      IntRetCD = DisplayMsgBox("900014","X","X","X")
      If Trim(txtCd.value) > "" Then
         txtCd.Select
         txtCd.Focus
      Else
         txtNm.Select
         txtNm.Focus
     End If
   End If

End Function

</SCRIPT>
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" STYLE="WIDTH:35%"><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
				<TD CLASS="TD6" STYLE="WIDTH:65%"><INPUT TYPE="Text" Name="txtCd" SIZE=20 tag="12XXXU" onkeypress="ConditionKeypress"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" STYLE="WIDTH:35%">&nbsp;</TD>
				<TD CLASS="TD6" STYLE="WIDTH:65%"><INPUT TYPE="Text" NAME="txtNm" SIZE=30 tag="12" onkeypress="ConditionKeypress"></TD>
			</TR>
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/gcommonpopup_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=YES noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

