<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PA4
'*  4. Program Name         : 
'*  5. Program Desc         : 공장별 품목 팝업	
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "q1211pb4.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim ITEM_CD		
Dim ITEM_NM		
Dim SPEC	

Dim strReturn
Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgCurDate

Dim gblnWinEvent
											   
Dim arrReturn

Dim ArrParent
Dim arrParam

Dim PlantCd
Dim lgConFlg 

Dim lgInspClassCd			
Dim search_char

ArrParent = window.dialogArguments

ReDim arrParam(UBOUND(ArrParent) - 2)		'품목검색조건 추가 arrParam(5)

Set PopupParent = ArrParent(0)
Dim i
For i = 0 To UBOUND(ArrParent) - 2 
	arrParam(i) = ArrParent(i + 1)
Next
top.document.title = PopupParent.gActivePRAspName
'top.document.title = "공장별 품목 팝업"

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	ITEM_CD		= 1
    ITEM_NM		= 2
    SPEC	= 3
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""
	lgConFlg = ""
	lgSortKey    = 1                            '⊙: initializes sort direction
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	hItemCd.value = ""
	hItemNm.value = ""

	
	gblnWinEvent = False
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
		
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspClassCd , lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	
	txtPlantCd.value	= arrParam(0)
	txtPlantNm.value	= arrParam(1)
	txtItemCd.value		= arrParam(2)
	txtItemNm.value		= arrParam(3)
	cboInspClassCd.value = arrParam(4)
	lgInspClassCd = arrParam(4)
	
	If UBOUND(ArrParent) > 6 then
		search_char = arrParam(5)
	End If
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021216",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False

    vspdData.MaxCols = SPEC + 1
    vspdData.MaxRows = 0
	    
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit ITEM_CD,"품목",20												
	ggoSpread.SSSetEdit ITEM_NM,"품목명",40									
	ggoSpread.SSSetEdit SPEC,"규격",40								
			
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	vspdData.ReDraw = True
		
	Call SetSpreadLock()
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

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

			ITEM_CD		= iCurColumnPos(1)
			ITEM_NM		= iCurColumnPos(2)
			SPEC		= iCurColumnPos(3)
			
    End Select    

End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call InitComboBox()
	Call SetDefaultVal()
	Call InitVariables
	Call InitSpreadSheet()
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call DbQuery()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    'Err.Clear                                                               '☜: Protect system from crashing
	
    DbQuery = False                                                         '⊙: Processing is NG
	
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then	Exit Function								
    
	Call LayerShowHide(1)												<%'⊙: 작업진행중 표시 %>	
	    
    Dim strVal
  	
	If lgStrPrevKey1 <> "" Then
		strVal = BIZ_PGM_ID & "?PlantCd=" & Trim(txtPlantCd.value) & _
							  "&txtItemCd=" & lgStrPrevKey1 & _
							  "&txtItemNm=" & lgStrPrevKey2
	ELSE
		strVal = BIZ_PGM_ID & "?PlantCd=" & Trim(txtPlantCd.value) & _
							  "&txtItemCd=" & Trim(txtItemCd.value) & _
							  "&txtItemNm=" & Trim(txtItemNm.value)					'☆: 조회 조건 데이타 
	End If     
	strVal = strVal & "&lgInspClassCd=" & Trim(lgInspClassCd) _
					& "&cboInspClassCd=" & Trim(cboInspClassCd.value) _
					& "&lgConFlg=" & lgConFlg _
					& "&search_char=" & search_char _
					& "&txtMaxRows=" & vspdData.MaxRows		
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    vspdData.focus
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
End Function

'========================================================================================================
'	Name : FncQuery
'	Desc : 
'========================================================================================================
Function FncQuery()

	FncQuery = False

	Call InitVariables()
			
	vspdData.MaxRows = 0						'Grid 초기화 
		
	Call DbQuery()
	
	FncQuery = True
	
End Function

'========================================================================================================
'	Name : OKClick
'	Desc : 
'========================================================================================================
Function OKClick()

	Dim intColCnt, iCurColumnPos
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 2)
	
		ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
				
		For intColCnt = 0 To vspdData.MaxCols - 2
			vspddata.Col = iCurColumnPos(CInt(intColCnt + 1))
			arrReturn(intColCnt) = vspdData.Text
		Next
			
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000011111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then Exit Sub                                                   'If there is no data.
   	
   	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
   	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function             ' 타이틀 cell을 dblclick했거나....
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then Exit Sub

		If NewRow = .MaxRows Then
			If lgStrPrevKey1 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				DbQuery
			End If
		End If
	End With
End Sub	

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub
	
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			If DBQuery = False Then Exit Sub
		End If
	End If
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	vspdData.Redraw = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">

	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
			<TR>
				<td CLASS="TD5" NOWPAP>공장</td>
				<td CLASS="TD6" NOWPAP>
					<input TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="14XXXU" >&nbsp;<input TYPE=TEXT NAME="txtPlantNm" SIZE="20" tag="14" >
				</td>
				<td CLASS="TD5" NOWPAP>&nbsp;</td>
				<td CLASS="TD6" NOWPAP>&nbsp;</td>										
			</TR>
			<TR>
				<TD CLASS=TD5 NOWPAP>품목</TD>
				<TD CLASS=TD6 NOWPAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="11" ALT="품목명"></TD>
				<td CLASS="TD5" NOWPAP>검사분류</td>
				<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="11"><OPTION VALUE="" selected></OPTION></SELECT></TD>
			</TR>
					
		</TABLE></FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/q1211pa4_I128669810_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hItemNm" tag="24" TABINDEX="-1">

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
