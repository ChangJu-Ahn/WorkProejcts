<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p2313pa1.asp																*
'*  4. Program Name         : MPS History No. Popup														*			*
'*  5. Program Desc         : MPS�̷¹�ȣ �˾�															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/19																*
'*  8. Modified date(Last)  : 2002/12/10																*
'*  9. Modifier (First)     : Kim GyoungDon																*
'* 10. Modifier (Last)      : Jung Yu Kyung																*			*
'* 11. Comment              :																			*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>MPS RUN POPUP</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	

'========================================================================================================
Const BIZ_PGM_QRY_ID	= "p2213pb1.asp"

Dim C_RunNo
Dim C_Date 

'========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim arrReturn
Dim lgPlantCD
Dim lgRunNo
Dim ArgArray

ArgArray = window.dialogArguments
Set PopupParent = ArgArray(0)
lgPlantCD = ArgArray(1)
lgRunNo= ArgArray(2)
top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_RunNo	=  1
	C_Date  =  2
	
End Sub

'========================================================================================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE
	
	lgIntGrpCount = 0
	lgStrPrevKey = ""
	hRunNo.value = ""
	lgSortKey    = 1

	Self.Returnvalue = Array("")
	vspdData.MaxRows = 0
End Function
	
Function InitSetting()
	txtRunNo.value = lgRunNo
End Function
	
'========================================================================================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021216",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False

	vspdData.MaxCols = C_Date + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit 	C_RunNo, "RUN NO.", 18
	ggoSpread.SSSetDate 	C_Date, "����", 20, 2, PopupParent.gDateFormat
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	Call SetSpreadLock()
	vspdData.ReDraw = True
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	vspdData.ReDraw = False
	ggoSpread.SpreadLock -1,	-1
	vspdData.ReDraw = True
End Sub

'========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_RunNo		= iCurColumnPos(1)
			C_Date		= iCurColumnPos(2)
			
    End Select    

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000011111")

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
        Exit Sub
    End If
   	
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'========================================================================================================
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			
			If DBQuery = False Then 
               Exit Sub
            End If 
		End If
    End if
    
End Sub


'========================================================================================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()

	Dim arrTmp
	Dim intRowCnt
	Dim intColCnt
	Dim intSelCnt

	intSelCnt = 0
	Redim arrReturn(0)
	
	vspdData.Row = vspdData.ActiveRow

	If vspdData.SelModeSelected = True Then
		arrReturn(0) = GetSpreadText(vspdData,C_RunNo,vspdData.ActiveRow,"X","X")
	End If

	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
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

Sub vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii=27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub
	
'========================================================================================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call InitSpreadSheet()
	Call InitSetting()
	Call FncQuery()
End Sub


'========================================================================================================
Function FncQuery()

    FncQuery = False
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	Call Initvariables
	
	If DBQuery = False Then 
       Exit Function
    End If 
	
	FncQuery = True
End Function

'========================================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	vspdData.Redraw = True
End Sub

'========================================================================================================
'*	Object���� �߻� �ϴ� Event ó��																		*
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' Ÿ��Ʋ cell�� dblclick�߰ų�....
	   Exit Function
	End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function


'========================================================================================================
' Function Name : DbQuery																				=
' Function Desc : This function is data query and display												=
'========================================================================================================
Function DbQuery()
    Err.Clear
    
    DbQuery = False

    Call LayerShowHide(1)
    
    Dim strVal
    
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtRunNo=" & Trim(hRunNo.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows="   & vspdData.MaxRows	    
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & lgPlantCD
		strVal = strVal & "&txtRunNo=" & Trim(txtRunNo.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
		strVal = strVal & "&txtMaxRows="   & vspdData.MaxRows
	End If

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode = PopupParent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
	vspdData.focus

End Function

Function FncExit()
	FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
			<TR>
				<TD CLASS=TD5 NOWRAP>RUN NO.</TD>
				<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtRunNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="RUN NO."></TD>
			</TR>
		</TABLE></FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/p2213pa1_vspdData_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hRunNo" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
