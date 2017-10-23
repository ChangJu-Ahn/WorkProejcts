<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common
*  2. Function Name        : Single Sample
*  3. Program ID           : CommonPopup
*  4. Program Name         : CommonPopup
*  5. Program Desc         : CommonPopup
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     : .
* 10. Modifier (Last)      : .
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>

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
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "GCommonPopupBiz.asp"                                      '☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================

Const POPUP_TITLE = 0                                                           '--- Index of POP-UP Title
Const TABLE_NAME  = 1                                                           '--- Index of DB table name to query
Const CODE_CON    = 2                                                           '--- Index of Code Condition value
Const NAME_CON    = 3                                                           '--- Index of Name Condition value
Const WHERE_CON   = 4                                                           '--- Index of Where Clause
Const TEXT_NAME   = 5                                                           '--- Index of Textbox Name

Const C_SHEETMAXROWS = 30                                                       '--- 한화면에 보일수 있는 최대 Row 수 

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgStrCodeKey
Dim lgStrNameKey

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
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	lblTitle.innerHTML = arrParam(TEXT_NAME)
	txtCd.value = arrParam(CODE_CON)
	txtNm.value = arrParam(NAME_CON)

	Self.Returnvalue = Array("")
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
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

		ggoSpread.SSSetEdit 1, arrGridHdr(0), 18	' 코드 

	    For i = 1 To gintDataCnt - 1
			If InStr(1, UCase(arrTblField(i)), "CONVERT") > 0 And InStr(1, UCase(arrTblField(i)), "CHAR") > 0 Then
				ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 25, 1
			Else
				ggoSpread.SSSetEdit i + 1, arrGridHdr(i), 50
			End If
	    Next

	    For i = 0 To gintDataCnt - 1
			If InStr(1, UCase(arrTblField(i)), gColSep) > 0 Then
               iArr = Split(UCase(arrTblField(i)),gColSep)

               iLen = 0

               If Len(Trim(iArr(0))) > 2 Then
                  iLen = Cint(Mid(iArr(0),3,2))
               End If

               Select Case Mid(iArr(0),1,2)
                    Case "ED"   '일반문자 
                           If iLen > 0 Then
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), iLen
                           Else
                              ggoSpread.SSSetEdit   i + 1,arrGridHdr(i), 50
                           End If
                           arrTblField(i) = iArr(1)
                    Case "DD"   '날짜 
                           If iLen > 0 Then
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),iLen,2,gDateFormat
                           Else
                              ggoSpread.SSSetDate   i + 1,arrGridHdr(i),  12,2,gDateFormat
                           End If
                           arrTblField(i) = iArr(1)
                    Case "F2","F3","F4","F5"
                           If iLen > 0 Then
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),iLen,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           Else
                              ggoSpread.SSSetFloat  i + 1,arrGridHdr(i),  17,Mid(iArr(0),2,1),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
                           End If
                           arrTblField(i) = iArr(1)
                    Case "TT"   ' Time
                           If iLen > 0 Then
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),iLen,,1,1
                           Else
                              ggoSpread.SSSetTime   i + 1,arrGridHdr(i),  12,,1,1
                           End If
                           arrTblField(i) = iArr(1)
               End Select
           End If
	    Next

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

	Call InitVariables
	Call SetDefaultVal()
	Call InitSpreadSheet()

	Call Search_OnClick()
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="ComLoadInfTB19029.asp" -->
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Function Name :
' Function Desc :
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
' Function Name :
' Function Desc :
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'========================================================================================================
' Function Name :
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
Sub Search_OnClick()

    vspdData.MaxRows = 0
    lgStrCodeKey = ""
    lgStrNameKey = ""

	Call DbQuery()

End Sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call Search_OnClick()
	End If
End sub
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

    If vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then
       If lgStrCodeKey <> "" Or lgStrNameKey <> "" Then
 		  DbQuery
       End If
    End if
End Sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strPreWhere, strWhere
    Dim iLoop
    Dim arrStrVal

    strPreWhere = ""
    strWhere = ""

    If UCase(Trim(arrParam(WHERE_CON))) <> "" Then
		strPreWhere = UCase(Trim(arrParam(WHERE_CON))) & " AND "
    End If


   '----- Code가 있을 경우는 Name에 상관없이 Code로만 조회하고, Code가 없는 경는 Name으로 조회한다.
	If lgStrCodeKey <> "" Or lgStrNameKey <> "" Then
		If Trim(txtNm.value) <> "" Then
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(1)) & ">='" & Trim(UCase(lgStrNameKey)) & "'" & " Order by " &   Trim(arrTblField(1))
		Else
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(0)) & ">='" & Trim(UCase(lgStrCodeKey)) & "'" & " Order by " &   Trim(arrTblField(0))
		End If
	Else
		if Trim(txtNm.value) <> "" Then
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(1)) & ">='" & Trim(UCase(txtNm.value)) & "'" & " Order by " &   Trim(arrTblField(1))
		Else
			strWhere = "WHERE " & strPreWhere & Trim(arrTblField(0)) & ">='" & Trim(UCase(txtCd.value)) & "'" & " Order by " &   Trim(arrTblField(0))
			End If
		End If

	    DbQuery = False                                                         '⊙: Processing is NG

	    arrStrVal = ""

	    For iLoop = 0 To gintDataCnt - 1
	        arrStrVal = arrStrVal & Trim(arrTblField(iLoop)) & gColSep
	    Next

	    strVal = BIZ_PGM_ID & "?txtTable=" & Trim(arrParam(TABLE_NAME))
	    strVal = strVal & "&txtWhere="    & strWhere
	    strVal = strVal & "&gintDataCnt=" & gintDataCnt
		strVal = strVal & "&arrField="    & arrStrVal

		Call LayerShowHide(1)
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	    DbQuery = True                                                          '⊙: Processing is NG
	End Function
'========================================================================================================
' Function Name :
' Function Desc :
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
<!-- #Include file="../inc/UNI2KCMCom.inc" -->
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
				<IMG SRC="../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

