<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Comon Popup																*
'*  3. Program ID           : i1411pa1																	*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수불유형팝업																*
'*  7. Modified date(First) : 2000/02/29																*
'*  8. Modified date(Last)  : 2000/02/29																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/02/29 : Coding Start													*
'********************************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "i1411pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgQueryFlag				
Dim lgMovType				
Dim lgMovTypeNm				
Dim lgTrnsType

Dim arrParam				
Dim arrReturn				

Dim C_MovType
Dim C_MovTypeNm
Dim C_DebitCreditFlag
Dim C_PriceCtrlFlag
Dim C_PostCtrlFlag
Dim C_MatlCostDistIndctr

arrParam = window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName

'=======================================================================================================
' Function Name : LoadInfTB19029
'=======================================================================================================
Function LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	txtMovType.value   =   arrParam(1)
	txtMovTypeNm.value =   arrParam(2)
	txtTrnsType.Value  =   arrParam(3)	
	Self.Returnvalue   = Array("")
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'========================================================================================================
Sub InitSpreadSheet()

 	Call InitSpreadPosVariables()
    
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread
    
    vspdData.ReDraw = False
    
    vspdData.MaxCols = C_MatlCostDistIndctr + 1
    vspdData.MaxRows = 0
    	    
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit C_MovType,				"수불유형",			14,2	
	ggoSpread.SSSetEdit C_MovTypeNm,			"수불유형명",		30
	ggoSpread.SSSetEdit C_DebitCreditFlag,		"재고증감구분",		18,2
	ggoSpread.SSSetEdit C_PriceCtrlFlag,		"재고단가반영구분", 18,2	
	ggoSpread.SSSetEdit C_PostCtrlFlag,			"회계Posting구분",	18,2
	ggoSpread.SSSetEdit C_MatlCostDistIndctr,	"비용발생여부",		18,2		

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SSSetSplit2(1)
	vspdData.ReDraw = True
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_MovType				= 1
	C_MovTypeNm				= 2
	C_DebitCreditFlag		= 3
	C_PriceCtrlFlag			= 4
	C_PostCtrlFlag			= 5
	C_MatlCostDistIndctr	= 6
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_MovType			= iCurColumnPos(1)
		C_MovTypeNm			= iCurColumnPos(2)	
		C_DebitCreditFlag	= iCurColumnPos(3)
		C_PriceCtrlFlag		= iCurColumnPos(4)
		C_PostCtrlFlag		= iCurColumnPos(5)
		C_MatlCostDistIndctr= iCurColumnPos(6)
		
 	End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
	
		vspdData.Row = vspdData.ActiveRow
				
		vspdData.Col = C_MovType
		arrReturn(0) = vspdData.Text
		vspdData.Col = C_MovTypeNm
		arrReturn(1) = vspdData.Text
		vspdData.Col = C_DebitCreditFlag
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_PriceCtrlFlag
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_PostCtrlFlag
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_MatlCostDistIndctr
		arrReturn(5) = vspdData.Text
									
		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call ggoOper.LockField(Document, "N")                     
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = vspdData
    
 	If vspdData.MaxRows = 0 Then Exit Sub
 	
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

'========================================================================================
' Function Name : vspdData_KeyPress
'========================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================
' Function Name : vspdData_TopLeftChange
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) And lgQueryFlag <> "1" Then
		DbQuery
    End if
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    
 	If Row <= 0 Then Exit Sub
  	If vspdData.MaxRows = 0 Then Exit Sub

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If

End Sub
 	
'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
Function FncQuery()
 
    ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData

	lgQueryFlag		= "1"
	lgMovType		= Trim(txtMovType.value)
	lgMovTypeNm		= Trim(txtMovTypeNm.value)
	lgTrnsType		= Trim(txtTrnsType.value)
	
	If Not chkField(Document, "1") Then Exit Function
	
	If DbQuery() = False Then Exit Function

End Function

'********************************************  5.2 DbQuery()  *******************************************
' Function Name : DbQuery																				*
'********************************************************************************************************
Function DbQuery()

   	Call LayerShowHide(1) 

    Dim strVal

    DbQuery = False                                              
    
    strVal = BIZ_PGM_ID &	"?txtMovType="		& lgMovType			& _    
							"&txtTrnsType="		& lgTrnsType		& _
							"&Flag="			& lgQueryFlag		& _
							"&lgStrPrevKey="	& lgStrPrevKey		& _
							"&txtMaxRows="		& vspdData.MaxRows
    
     Call RunMyBizASP(MyBizASP, strVal)								
	
    DbQuery = True 
                                                    
End Function

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQueryOk																				*
'********************************************************************************************************
Function DbQueryOk()								
	vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" WIDTH=30%>수불유형</TD>
				<TD CLASS="TD6" WIDTH=70%><INPUT TYPE="Text" Name="txtMovType" SIZE=10 MAXLENGTH=3  tag="11XXXU" >&nbsp;<INPUT TYPE="Text" NAME="txtMovTypeNm" MAXLENGTH=40 tag="14" ></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" WIDTH=30%>수불구분</TD>
				<TD CLASS="TD6" WIDTH=70%><INPUT TYPE="Text" Name="txtTrnsType" SIZE=10 MAXLENGTH=2  tag="14XXXU" ></TD>
			</TR>
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/i1411pa1_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

