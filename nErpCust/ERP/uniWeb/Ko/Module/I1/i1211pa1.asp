<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I1211pa1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 수불품목 팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/04/03
'*  8. Modified date(Last)  : 2005/02/17
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit           

Const BIZ_PGM_ID = "i1211pb1.asp"

Dim C_ItemCode      
Dim C_ItemName   
Dim C_ItemSpec   
Dim C_InvUnit    
Dim C_TrackingNo 
Dim C_LotNo      
Dim C_LotSubNo   
Dim C_GoodQty    
Dim C_BadQty     
Dim C_InspQty    
Dim C_TrnsQty    
Dim C_PickingQty 
Dim C_EntryUnit  

Dim arrReturn
Dim arrParam
Dim arrField
Dim arrParent

Dim arrUserFlag
  
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)
arrField		= arrParent(2)

top.document.title = PopupParent.gActivePRAspName

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim gblnWinEvent
Dim strReturn

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
 
	lgIntFlgMode      = PopupParent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
	lgStrPrevKeyIndex = ""
	lgLngCurRows      = 0
 
	gblnWinEvent = False
	Self.Returnvalue = Array("")    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
 
 txtPlantCd.value = arrParam(0)
 txtPlantNm.value = arrParam(1) 
 txtSLCd.value    = arrParam(2)
 txtSLNm.value    = arrParam(3)
 txtItemCd1.value = arrParam(4)
 txtItemNm1.value = arrParam(5)
 
 Self.Returnvalue = Array("")
End Sub 

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20030423", , PopupParent.gAllowDragDropSpread
	
	With  vspdData
		.ReDraw = false
		.MaxCols = C_PickingQty+1        
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
		     
		ggoSpread.SSSetEdit C_ItemCode,		"품목",			18
		ggoSpread.SSSetEdit C_ItemName,		"품목명",		25
		ggoSpread.SSSetEdit C_ItemSpec,		"규격",			20
		ggoSpread.SSSetEdit C_InvUnit,		"단위",			10
		ggoSpread.SSSetEdit C_TrackingNo,	"Tracking No",	20    
		ggoSpread.SSSetEdit C_LotNo,		"Lot No.",		12
		ggoSpread.SSSetFloat C_LotSubNo,	"순번",			6, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_GoodQty,		"양품재고량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_BadQty,		"불량재고량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_InspQty,		"검사중수량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_TrnsQty,		"이동재고량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PickingQty,	"PICKING수량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)	  
		ggoSpread.SSSetSplit2(3)

		ggoSpread.SpreadLockWithOddEvenRowColor

		.ReDraw = true

	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode   = 1 
	C_ItemName   = 2
	C_ItemSpec   = 3
	C_InvUnit    = 4
	C_TrackingNo = 5
	C_LotNo      = 6
	C_LotSubNo   = 7
	C_GoodQty    = 8
	C_BadQty     = 9
	C_InspQty    = 10
	C_TrnsQty    = 11
	C_PickingQty = 12
	C_EntryUnit  = 13
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode   = iCurColumnPos(1)
		C_ItemName   = iCurColumnPos(2)
		C_ItemSpec   = iCurColumnPos(3)
		C_InvUnit    = iCurColumnPos(4)
		C_TrackingNo = iCurColumnPos(5)
		C_LotNo      = iCurColumnPos(6)
		C_LotSubNo   = iCurColumnPos(7)
		C_GoodQty    = iCurColumnPos(8)
		C_BadQty     = iCurColumnPos(9)
		C_InspQty    = iCurColumnPos(10)
		C_TrnsQty    = iCurColumnPos(11)
		C_PickingQty = iCurColumnPos(12)
		C_EntryUnit  = iCurColumnPos(13)		
	End Select

End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim i, iCurColumnPos
  
	If vspdData.ActiveRow > 0 Then 
		Redim arrReturn(UBound(arrField))

        ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
			
		For i = 0 To UBound(arrField)
			If arrField(i) <> "" Then
				vspddata.Col = iCurColumnPos(CInt(arrField(i)))
				arrReturn(i) = vspdData.Text
			End If
		Next

		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
	
End Function


'=========================================  2.3.2 CancelClick()  ========================================
 Function CancelClick()  
  Self.Close()
 End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029          
    Call ggoOper.LockField(Document, "N")    
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                               
    Call SetDefaultVal()
    If DbQuery = False Then Exit Sub
	txtItemCd1.focus
 
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
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   
	If Row <= 0 Then Exit Sub
	If vspdData.MaxRows = 0 Then Exit Sub

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
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
 
'========================================================================================================
'   Event Name : vspdData_LeaveCell
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then Exit Sub
		
		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then
				If DbQuery = False Then	Exit Sub
			End if
		End if
	
	End With
End Sub 

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) and lgStrPrevKeyIndex <> "" Then 
		If DbQuery = False Then Exit Sub
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)

	On error Resume Next
	
	If KeyAscii = 27 Then Call CancelClick()

	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then 
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                             
    
    Err.Clear                                                    

    '-----------------------
    'Erase contents area
    '-----------------------
 	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
	
	Call InitVariables            
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then Exit Function
       
    FncQuery = True            
    
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call PopupParent.FncPrint()
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim strVal
    Dim strFlag
    
	If RadioOutputType.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
    
    Call LayerShowHide(1)
    DbQuery = False
    
    Err.Clear                                                             
    
    strVal = BIZ_PGM_ID &	"?txtMode="           & PopupParent.UID_M0001		& _
							"&txtPlantCd="        & Trim(txtPlantCd.value)		& _
							"&txtPlantNm="        & Trim(txtPlantNm.value)		& _
							"&txtSLCd="           & Trim(txtSLCd.value)			& _
							"&txtSlNm="           & Trim(txtSLNm.value)			& _
							"&txtItemCd1="        & Trim(txtItemCd1.value)		& _
							"&txtItemNm1="        & Trim(txtItemNm1.value)		& _
							"&txtFlag="			  & strFlag						& _
							"&txtMaxRows="        & vspdData.MaxRows			& _
							"&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex

    Call RunMyBizASP(MyBizASP, strVal)
       
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()            
    Call ggoOper.LockField(Document, "Q")    
	vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
	<TABLE CELLSPACING=0 CLASS="basicTB">
		<TR>
			<TD HEIGHT=40>
				<FIELDSET>
					<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
						<TR>
							<TD CLASS="TD5" NOWRAP>공장</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtPlantCd" SIZE=10 MAXLENGTH=7  tag="14XXXU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
							<TD CLASS="TD5" NOWRAP>창고</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtSLCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="창고" >&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="14"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE="Text" NAME="txtItemCd1" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 MAXLENGTH=40 tag="11XXXU" ALT="품목명"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수량유무여부</TD>
							<TD CLASS="TD6">
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" checked><LABEL FOR="rdoCase1">수량있음</LABEL>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X"><LABEL FOR="rdoCase2">전품목</LABEL>
							</TD>
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
						</TR>
					</TABLE>
				</FIELDSET>
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=100%>
				<script language =javascript src='./js/i1211pa1_OBJECT1_vspdData.js'></script>
			</TD>
		</TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
												  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>  
			</TD>
		</TR>
	</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


