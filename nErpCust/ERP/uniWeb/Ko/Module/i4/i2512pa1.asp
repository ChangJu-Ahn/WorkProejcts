<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'********************************************************************************************************
'*  1. Module Name          : Inventory																*
'*  2. Function Name        : Popup Item By Plant														*	
'*  3. Program ID           : i2512pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Item by Plant Popup														*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2003/05/26																*
'*  9. Modifier (First)     : Im Hyun Soo																*
'* 10. Modifier (Last)      : Lee Seung Wook																*
'* 11. Comment              :																			*
=======================================================================================================-->
<HTML>
<HEAD>
<!--
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js">  </SCRIPT>

<Script Language="VBScript">
Option Explicit   

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "i2512pb1.asp"						

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim ITEM_CD		
Dim ITEM_NM		
Dim SPECIFICATION	
Dim BASIC_UNIT	
Dim ITEM_ACCT		
Dim ITEM_CLASS	
Dim LOT_FLG		
Dim TRACKING_FLG	
Dim MAJOR_SL_CD	
Dim ISSUED_SL_CD	
Dim VALID_FLG		
Dim VALID_FROM_DT	
Dim VALID_TO_DT	

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim strReturn			

Dim lgCurDate

Dim gblnWinEvent			
										
Dim arrReturn
Dim arrParam					
Dim PlantCd
			
Dim lgItemAcct
Dim lgDefItemAcct

<!-- #Include file="../../inc/lgvariables.inc" -->
'------ Set Parameters from Parent ASP ------ 
arrParam = window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName	
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()

	lgStrPrevKeyIndex = ""
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	'------ Coding part ------
	gblnWinEvent = False
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
	
	
'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()

	Call CommonQueryRS(" MINOR_CD, MINOR_NM "," B_MINOR ", _
					   " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	Call SetCombo2(cboItemAccount, lgF0, LgF1, Chr(11))
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
		
	PlantCd         = arrParam(1)
	txtItemCd.value = arrParam(2)		

	'-------------------------------
	' 품목계정 Setting
	'-------------------------------
		ReDim lgItemAcct(0) 
		lgItemAcct(0) = "N"
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Sub


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

    vspdData.ReDraw = False
		    
    vspdData.MaxCols = VALID_TO_DT + 1
    vspdData.MaxRows = 0
	Call GetSpreadColumnPos("A")    
    
	ggoSpread.SSSetEdit ITEM_CD,      "품목",          15												
	ggoSpread.SSSetEdit ITEM_NM,      "품목명",        25												
	ggoSpread.SSSetEdit ITEM_ACCT,    "계정",           5											
	ggoSpread.SSSetEdit ITEM_CLASS,   "품목그룹",       8											
	ggoSpread.SSSetEdit BASIC_UNIT,   "단위",           5											
	ggoSpread.SSSetEdit LOT_FLG,      "LOT관리",        8					
	ggoSpread.SSSetEdit TRACKING_FLG, "Tracking 구분", 10
	ggoSpread.SSSetEdit SPECIFICATION,"규격",          10											
	ggoSpread.SSSetEdit MAJOR_SL_CD,  "입고창고",       8										
	ggoSpread.SSSetEdit ISSUED_SL_CD, "출고창고",       8										
	ggoSpread.SSSetEdit VALID_FLG,    "유효구분",       8
	ggoSpread.SSSetDate VALID_FROM_DT,"시작일",        10, 2, PopupParent.gDateFormat										
	ggoSpread.SSSetDate VALID_TO_DT,  "종료일",        10, 2, PopupParent.gDateFormat											
																						
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	ggoSpread.SSSetSplit2(2)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	ITEM_CD		  = 1
	ITEM_NM		  = 2
	SPECIFICATION = 3
	BASIC_UNIT	  = 4
	ITEM_ACCT	  = 5
	ITEM_CLASS	  = 6
	LOT_FLG		  = 7
	TRACKING_FLG  = 8
	MAJOR_SL_CD	  = 9
	ISSUED_SL_CD  = 10
	VALID_FLG	  = 11
	VALID_FROM_DT = 12
	VALID_TO_DT	  = 13
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		ITEM_CD		  = iCurColumnPos(1)
		ITEM_NM		  = iCurColumnPos(2)
		SPECIFICATION = iCurColumnPos(3)
		BASIC_UNIT	  = iCurColumnPos(4)
		ITEM_ACCT	  = iCurColumnPos(5)
		ITEM_CLASS	  = iCurColumnPos(6)
		LOT_FLG		  = iCurColumnPos(7)
		TRACKING_FLG  = iCurColumnPos(8)
		MAJOR_SL_CD	  = iCurColumnPos(9)
		ISSUED_SL_CD  = iCurColumnPos(10)
		VALID_FLG	  = iCurColumnPos(11)
		VALID_FROM_DT = iCurColumnPos(12)
		VALID_TO_DT	  = iCurColumnPos(13)
	End Select

End Sub


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = vspdData
	vspdData.ReDraw = False
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
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call SetDefaultVal()
	Call InitVariables
	Call InitComboBox()
	Call InitSpreadSheet()

	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Sub
	End If
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
	
    DbQuery = False                                       
	
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	Call LayerShowHide(1)												
	    
    Dim strVal
          
  	strVal = BIZ_PGM_ID &	"?txtMode="        & PopupParent.UID_M0001	& _			
							"&PlantCd="        & PlantCd				& _							
							"&txtItemCd="      & Trim(txtItemCd.value)	& _			
							"&txtitemNm="      & txtItemNm.value		& _
							"&cboItemAccount=" & Trim(cboItemAccount.value)
	
	If lgItemAcct(0) = "Y" Then
		strVal = strVal & "&FromItemAcct=" & lgItemAcct(1)
		strVal = strVal & "&ToItemAcct="   & lgItemAcct(2)
	Else
		strVal = strVal & "&FromItemAcct=" & ""
		strVal = strVal & "&ToItemAcct="   & "zz"
	End If
	
	strVal = strVal & "&txtItemSpec=" & txtItemSpec.value 
	
		
	If rdoLotItem1.checked = True Then
		strVal = strVal & "&rdoLotItem= %" 
	Elseif rdoLotItem2.checked = True then
		strVal = strVal & "&rdoLotItem= Y" 
		Else
		strVal = strVal & "&rdoLotItem= N" 
	End If
		
	strVal = strVal & "&txtMaxRows="         & vspdData.MaxRows
	strVal = strVal & "&lgStrPrevKey="       & lgStrPrevKey
        
	Call RunMyBizASP(MyBizASP, strVal)										
		
    DbQuery = True                                                   
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	lgIntFlgMode = PopupParent.OPMD_UMODE								
    vspdData.focus
	Call ggoOper.LockField(Document, "Q")								
End Function

'========================================================================================================
'	Name : FncQuery
'	Desc : 
'========================================================================================================
Function FncQuery()

	FncQuery = False

	Call InitVariables()
			
	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
	lgStrPrevKey = ""
		
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
	
	FncQuery = True
	
End Function


'========================================================================================================
'	Name : OKClick
'	Desc : 
'========================================================================================================
Function OKClick()
	Dim i
	
	If vspdData.MaxRows > 0 Then
		
		vspdData.row = vspdData.ActiveRow 
		Redim arrReturn(1)		

		vspddata.Col = ITEM_CD
		arrReturn(0) = vspdData.Text
		vspddata.Col = ITEM_NM
		arrReturn(1) = vspdData.Text

		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
					
End Function

'========================================================================================================
'	Name : CancelClick
'	Desc : 
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : MousePointer
'	Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = vspdData
   
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		
			lgSortKey = 1
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	'------ Developer Coding part (Start)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	If KeyAscii=27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 and vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then						
				If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
				End If
			End If
		End If
	End With
End Sub
	

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then						
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
-->
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목</TD>
				<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=45 MAXLENGTH=40 tag="11" ALT="품목명"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목계정</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAccount" ALT="품목계정" STYLE="Width: 98px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
				<TD CLASS=TD5 NOWRAP>LOT관리여부</TD>
				<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotItem" tag="11" ID="rdoLotItem1" VALUE="%"><LABEL FOR="rdoLotItem1">전체</LABEL>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotItem" tag="11" CHECKED ID="rdoLotItem2" VALUE="Y"><LABEL FOR="rdoLotItem2">예</LABEL>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotItem" tag="11" ID="rdoLotItem3" VALUE="N"><LABEL FOR="rdoLotItem3">아니오</LABEL></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>규격</TD>
				<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=40 tag="11" ALT="규격">&nbsp;</TD>
				<TD CLASS=TD5 NOWRAP></TD>
				<TD CLASS=TD6 NOWRAP></TD>
			</TR>
			
		</TABLE></FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/i2512pa1_OBJECT1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"	SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hItemNm" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hItemAccount" tag="24" TABINDEX="-1">

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
