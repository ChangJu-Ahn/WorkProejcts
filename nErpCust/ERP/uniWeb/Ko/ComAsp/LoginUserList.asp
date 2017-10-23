<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common Module
*  2. Function Name        : Common Function
*  3. Program ID           : LoginUserList
*  4. Program Name         : LoginUserList
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/07
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Sim Hae Young
* 10. Modifier (Last)      : 
* 11. Comment              :
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

<!-- #Include file="../inc/IncSvrCcm.inc" -->
<!-- #Include file="../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================-->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">

Option Explicit            

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
	Const BIZ_PGM_ID = "LoginUserListBiz.asp"							'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
    Dim C_IP_ADDRESS
    Dim C_DB_SERVER
    Dim C_DB_NAME
    Dim C_LANGUAGE
    Dim C_USERID

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
	Dim arrParent
	Dim arrParam
	Dim arrReturn
	Dim lgStrPrevKey	
	Dim lgIntFlgMode
	
    arrParent = window.dialogArguments
    arrParam = arrParent(1)
    Set PopupParent = arrParent(0)

	top.document.title = "현재 접속중인 사용자 리스트 Popup"
			
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_DB_SERVER     = 1
	C_DB_NAME       = 2
	C_IP_ADDRESS	= 3
	C_LANGUAGE      = 4
	C_USERID        = 5
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	vspdData.MaxRows = 0
    ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData
    lgIntFlgMode = PopupParent.OPMD_CMODE
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="LoadInfTB19029.asp" -->
<%Call loadInfTB19029A("Q", "B", "NOCOOKIE", "PA")%>
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()  
    
    With vspdData
	    ggoSpread.Source = vspdData	
    'patch version
     ggoSpread.Spreadinit "V20030107",,Popupparent.gAllowDragDropSpread    
     
	 .ReDraw = false

     .MaxCols = C_USERID + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	 .Col = .MaxCols														'☆: 사용자 별 Hidden Column
	 .ColHidden = True    
	           
     .MaxRows = 0
     ggoSpread.ClearSpreadData
	
     Call GetSpreadColumnPos("A")  
     
     ggoSpread.SSSetEdit C_DB_SERVER, ""  , 20,,,20	
	 ggoSpread.SSSetEdit C_DB_NAME, "", 15,,,15
	 
     ggoSpread.SSSetEdit C_IP_ADDRESS, "IP ADDRESS", 15,,,15	
	 ggoSpread.SSSetEdit C_LANGUAGE, "LANGUAGE", 15,,,15
	 ggoSpread.SSSetEdit C_USERID, "User ID", 15,,,15

     Call ggoSpread.SSSetColHidden(C_DB_SERVER,C_DB_SERVER,True)
     Call ggoSpread.SSSetColHidden(C_DB_NAME,C_DB_NAME,True)

	.ReDraw = True
	    
	Call SetSpreadLock 
    	    
	End With
End Sub

<%'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================%>
Sub SetSpreadLock()
    vspdData.ReDraw = False
    ggoSpread.SpreadLock C_DB_SERVER, -1, C_DB_SERVER
    ggoSpread.SpreadLock C_DB_NAME, -1, C_DB_NAME    
    ggoSpread.SpreadLock C_IP_ADDRESS, -1, C_IP_ADDRESS    
    ggoSpread.SpreadLock C_LANGUAGE, -1, C_LANGUAGE    
    ggoSpread.SpreadLock C_USERID, -1, C_USERID    
	ggoSpread.SSSetProtected vspdData.MaxCols, -1, -1
    vspdData.ReDraw = True
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
			C_DB_SERVER     = iCurColumnPos(1)
			C_DB_NAME       = iCurColumnPos(2)
			C_IP_ADDRESS	= iCurColumnPos(3)
			C_LANGUAGE      = iCurColumnPos(4)
			C_USERID        = iCurColumnPos(5)
    End Select    
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
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field

	Call InitVariables
				
	Call SetDefaultVal()
	Call InitSpreadSheet()
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
'                        5.3 Common Method
'========================================================================================================
'========================================================================================================


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000011111")	
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
	
	vspdData.Row = Row
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If vspdData.MaxRows = 0 Then
        Exit Sub
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
' Name :vspdData_TopLeftChange
' Desc : 
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	 
    	If lgStrPrevKey <> "" Then 
      		DbQuery
    	End If

    End if

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
   
	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001

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
'    On Error Resume Next
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
		
	Self.Close()
End Function

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0
	
	Call InitVariables
	
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
Sub txtDate_Keypress(Key)
'    On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub


</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=20>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" ><SPAN CLASS="normal">&nbsp;</SPAN></TD>
				<TD CLASS="TD5" ><SPAN CLASS="normal">&nbsp;</SPAN></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" ><SPAN CLASS="normal">현재 접속중인 사용자 리스트&nbsp;</SPAN></TD>
				<TD CLASS="TD5" ><SPAN CLASS="normal">&nbsp;</SPAN></TD>
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/loginuserlist_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="ZipPopupBiz.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


