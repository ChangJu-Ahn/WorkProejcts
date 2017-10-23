<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Physical Inventory header Popup											*
'*  3. Program ID           : i2121pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 실사품목팝업																*
'*  7. Modified date(First) : 2001/02/23																*
'*  8. Modified date(Last)  : 2003/06/02																*
'*  9. Modifier (First)     : Lee Seung Wook															*
'* 10. Modifier (Last)      : Lee Seung Wook															*
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "i2121pb1.asp"					

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemSpec
Dim C_BasicUnit
Dim C_TrackingNo
Dim C_LotNO
Dim C_LotSubNo
Dim C_ABCFlag

Dim arrParam 
Dim arrReturn				
Dim arrSL_Cd
Dim arrSL_Nm
Dim arrItem_Cd
Dim arrItem_Nm
Dim arrTracking_No
Dim arrPlant_Cd
Dim arrPlant_Nm
Dim arrLotNo
Dim arrLotSubNo

Dim lgUserFlag			

Dim lgStrKey1			
Dim lgStrKey2
Dim lgStrKey3
Dim lgStrKey4

Dim lsOpenPop

arrParam = window.dialogArguments
Set PopupParent = arrParam(0)

arrPlant_Cd   =   arrParam(1)
arrItem_Cd    =   arrParam(2)
arrSl_Cd      =   arrParam(3)
arrPlant_Nm   =   arrParam(4)
arrSl_Nm      =   arrParam(5)

top.document.title = PopupParent.gActivePRAspName
'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
	lgIntGrpCount = 0
	lgStrKey1 = ""
    lgStrKey2 = ""
    lgStrKey3 = ""
    lgStrKey4 = ""
    
    lgLngCurRows = 0                          
    Self.Returnvalue = Array("")
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
'=======================================================================================================
Function LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	txtPlantCd.Value = arrPlant_Cd
	txtItemCd1.Value = arrItem_Cd
	txtSLCd.Value    = arrSL_Cd
	txtPlantNm.Value = arrPlant_Nm
	txtSLNm.Value    = arrSL_Nm
	Self.Returnvalue = Array("")
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	
 	Call InitSpreadPosVariables()

    With vspdData	

	ggoSpread.Source = vspdData

	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

	.ReDraw = False
	.MaxCols = C_ABCFlag+1
	.MaxRows = 0
	 
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_ItemCd, 		"품목", 15
	ggoSpread.SSSetEdit C_ItemNm,		"품목명", 20	
	ggoSpread.SSSetEdit C_ItemSpec,		"규격", 15
	ggoSpread.SSSetEdit C_BasicUnit,	"단위", 8,2
	ggoSpread.SSSetEdit C_TrackingNo,	"Tracking No", 20
	ggoSpread.SSSetEdit C_LotNo, 		"LOT NO", 12
	Call AppendNumberPlace("6", "3", "0")
	ggoSpread.SSSetFloat	C_LotSubNo, "순번", 5, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	ggoSpread.SSSetEdit C_ABCFlag,   "ABC", 8,2

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	ggoSpread.SpreadLock -1, -1

	ggoSpread.SSSetSplit2(2)

	.ReDraw = True

   End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ItemCd = 1
	C_ItemNm = 2
	C_ItemSpec = 3
	C_BasicUnit = 4
	C_TrackingNo = 5
	C_LotNO = 6
	C_LotSubNo = 7
	C_ABCFlag = 8
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_ItemCd		= iCurColumnPos(1)
		C_ItemNm		= iCurColumnPos(2)	
		C_ItemSpec		= iCurColumnPos(3)
		C_BasicUnit		= iCurColumnPos(4)
		C_TrackingNo	= iCurColumnPos(5)
		C_LotNO			= iCurColumnPos(6)
		C_LotSubNo		= iCurColumnPos(7)
		C_ABCFlag		= iCurColumnPos(8)
		
	End Select
End Sub

'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If hlgPlantCd = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		txtPlantCd.focus  
	Exit Function
	End if

	If lsOpenPop = True Then Exit Function

	lsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(txtSLCd.Value)
	arrParam(3) = ""	
	arrParam(4) = "PLANT_CD = " & FilterVar(hlgPlantCd, "''", "S")	
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lsOpenPop = False
	
	If arrRet(0) = "" Then
		txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
Function SetSL(byval arrRet)
	txtSLCd.Value    = arrRet(0)		
	txtSLNm.Value    = arrRet(1)
	txtSLCd.focus		
End Function

'-----------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lsOpenPop = True Then Exit Function

	lsOpenPop = True

	arrParam(0) = "공장팝업"							
	arrParam(1) = "B_PLANT"									 
	arrParam(2) = Trim(txtPlantCd.Value)		
	arrParam(3) = ""							
	arrParam(4) = ""							
	arrParam(5) = "공장"
	
    arrField(0) = "B_PLANT.PLANT_CD"			
    arrField(1) = "B_PLANT.PLANT_NM"			
    
    arrHeader(0) = "공장"					
    arrHeader(1) = "공장명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lsOpenPop = False
	
	If arrRet(0) = "" Then
		txtPlantCd.focus
		Exit Function
	Else
		txtPlantCd.Value = arrRet(0)
		txtPlantNm.Value = arrRet(1)
		txtPlantCd.focus
	End If	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If Trim(txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901", "X", "X")
		txtPlantCd.focus 	
		Exit Function
	End If

	If lsOpenPop = True Then Exit Function	

	iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
    End If

	lsOpenPop = True
	
	arrParam(0) = Trim(txtPlantCd.value)	
	arrParam(1) = Trim(txtItemCd.Value)	
	arrParam(2) = ""				
	arrParam(3) = ""				
	
	arrField(0) = 1 
	arrField(1) = 2 
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lsOpenPop = False
	
	If arrRet(0) = "" Then
		txtItemCd.focus
		Exit Function
	Else
		Call SetItem(arrRet)
	End If	

End Function

'------------------------------------------  SetItem()  --------------------------------------------------
Function SetItem(arrParam)
	txtItemCd.Value = arrParam(0)
	txtItemNm.Value = arrParam(1)
	txtItemCd.focus
End Function

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
	
		vspdData.Row = vspdData.ActiveRow
				
		vspdData.Col = C_ItemCd
		arrReturn(0) = vspdData.Text
		vspdData.Col = C_ItemNm
		arrReturn(1) = vspdData.Text
		vspdData.Col = C_ItemSpec
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_BasicUnit
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_TrackingNo
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_LotNO
		arrReturn(5) = vspdData.Text
		vspdData.Col = C_LotSubNo
		arrReturn(6) = vspdData.Text
		vspdData.Col = C_ABCFlag
		arrReturn(7) = vspdData.Text

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
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")  
	Call LoadInfTB19029
	
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call InitVariables
	
	Call SetDefaultVal()
	Call InitSpreadSheet()
	
	Call FncQuery()
End Sub

'*********************************************  3.2 Tag 처리  *******************************************
Function FncQuery() 

    FncQuery = False                                                       
    Err.Clear                                                              

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
	
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop)  Then
		If (lgStrKey1<> "" and lgStrKey2 <> "" and lgStrKey3 <> "" and lgStrKey4 <> "") Then
			DbQuery
		End if
	End if
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If vspdData.MaxRows = 0 Then Exit Sub
 	
	If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
		Call OKClick()
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

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()

    Call LayerShowHide(1)       
    
    DbQuery = False
    
    Dim strVal
    Dim txtMaxRows
    
		strVal = BIZ_PGM_ID &	"?txtSLCd="			& Trim(txtSLCd.value)		& _
								"&txtItemCd1="		& Trim(txtItemCd1.value)	& _
								"&txtPlantCd="		& Trim(arrPlant_Cd)			& _
								"&txtLotNo1="		& Trim(arrLotNo)			& _
								"&txtLotSubNo1="	& Trim(arrLotSubNo)			& _
								"&txtTrackingNo1="	& Trim(arrTracking_No)		& _
								"&txtItemCd2="		& Trim(lgStrKey1)			& _
								"&txtTrackingNo2="	& Trim(lgStrKey2)			& _
								"&txtLotNo2="		& Trim(lgStrKey3)			& _
								"&txtLotSubNo2="	& Trim(lgStrKey4)			& _
								"&lgStrUserFlag="	& lgUserFlag				& _					
								"&txtMaxRows="		& vspdData.MaxRows

    	Call RunMyBizASP(MyBizASP, strVal)						
    
    DbQuery = True
End Function

'********************************************  5.1 DbQueryOk()  *******************************************
Function DbQueryOk()
	Call ggoOper.LockField(Document, "Q")						
	vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtPlantCd" SIZE=10 MAXLENGTH=7  tag="14XXXU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtSLCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="창고" >&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYP="Text" NAME="txtItemCd1" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% HEIGHT=100%>
									<script language =javascript src='./js/i2121pa1_I637129789_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_01%>></TD>
				</TR>
				<TR HEIGHT=20>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD WIDTH=10>&nbsp;</TD>
								<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
								<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                              <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
								<TD WIDTH=10>&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
					</TD>
				</TR>		
</TABLE	>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
