<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : LOT Related Order REF
'*  3. Program ID           : I2511ra1.asp
'*  4. Program Name         : LOT Order REF 화면 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/10/09
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/10/09 : 4th Iteration
'**********************************************************************************************-->
<HTML>
<HEAD>
<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                

'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "i2511rb1.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_MvmtFlag
Dim C_Qty
Dim C_SoNo
Dim C_SoSeqNo
Dim C_PoNo
Dim C_PurOrdNo
Dim C_PurOrdSeqNo
Dim C_TrasactionDt
Dim C_Squence

Dim arrReturn
Dim arrParam

arrParam   = window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName
<!-- #Include file="../../inc/lgvariables.inc" -->

'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntGrpCount		= 0                          
    lgStrPrevKey		= ""
    lgLngCurRows		= 0                           
    Self.Returnvalue	= Array("")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	txtPlantCd.value  = UCase(arrParam(1))
	txtPlantNm.value  = arrParam(5)
	txtItemCd.value   = UCase(arrParam(2))
	txtItemNm.value   = arrParam(6)     
	txtLotNo.value    = UCase(arrParam(3))
	txtLotSubNo.value = arrParam(4) 
	
	Self.Returnvalue = Array("")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

	With  vspdData
		.ReDraw = false
		.MaxCols = C_Squence+1          
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit C_MvmtFlag,		"수불유형",				10, 40     
		ggoSpread.SSSetFloat C_Qty,			"수량",					15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetEdit C_SoNo,			"수주번호",				15, 18 
		ggoSpread.SSSetEdit C_SoSeqNo,		"수주순번",				10, 3
		ggoSpread.SSSetEdit C_PoNo,			"모품목제조오더번호",	15, 18 
		ggoSpread.SSSetEdit C_PurOrdNo,		"발주번호",				15, 18
		ggoSpread.SSSetEdit C_PurOrdSeqNo,	"발주순번",				10, 4
		ggoSpread.SSSetDate	C_TrasactionDt,	"수불일자",				12,	2,Parent.gDateFormat	
		ggoSpread.SSSetFloat C_Squence,		"순번",					13, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		
		'ggoSpread.MakePairsColumn()
		Call ggoSpread.SSSetColHidden(C_Squence, C_Squence, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(1)
		.ReDraw = true
		  
		Call SetSpreadLock 
	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_MvmtFlag		= 1
	C_Qty			= 2
	C_SoNo			= 3
	C_SoSeqNo		= 4
	C_PoNo			= 5
	C_PurOrdNo		= 6
	C_PurOrdSeqNo	= 7
	C_TrasactionDt	= 8
	C_Squence		= 9
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

		C_MvmtFlag		= iCurColumnPos(1)
		C_Qty			= iCurColumnPos(2)
		C_SoNo			= iCurColumnPos(3)
		C_SoSeqNo		= iCurColumnPos(4)
		C_PoNo			= iCurColumnPos(5)
		C_PurOrdNo		= iCurColumnPos(6)
		C_PurOrdSeqNo	= iCurColumnPos(7)
		C_TrasactionDt	= iCurColumnPos(8)
		C_Squence		= iCurColumnPos(9)		
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	vspdData.ReDraw = False    
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.ReDraw = True
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'= Name : OkClick()                     =
'= Description : Return Array to Opener Window when OK button click         =
'========================================================================================================
Function OKClick()
	Dim intColCnt
	  
	If vspdData.ActiveRow > 0 Then 
		Redim arrReturn(vspdData.MaxCols - 1)
		  
		vspdData.Row = vspdData.ActiveRow
		     
		vspdData.Col = C_MvmtFlag
		arrReturn(0) = vspdData.Text			
		vspdData.Col = C_Qty
		arrReturn(1) = vspdData.Text			
		vspdData.Col = C_SoNo
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_SoSeqNo
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_PoNo
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_PurOrdNo
		arrReturn(5) = vspdData.Text
		vspdData.Col = C_PurOrdSeqNo
		arrReturn(6) = vspdData.Text
		vspdData.Col = C_TrasactionDt
		arrReturn(7) = vspdData.Text
		vspdData.Col = C_Squence
		arrReturn(8) = vspdData.Text

		Self.Returnvalue = arrReturn
	End If
	  
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'= Name : CancelClick()                    =
'= Description : Return Array to Opener Window for Cancel button click         =
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029            
    Call ggoOper.LockField(Document, "N")                               
    Call InitSpreadSheet
    Call InitVariables                                                     
    Call SetDefaultVal()
    Call FncQuery()
End Sub

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
		Exit Sub
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
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
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

Function vspdData_KeyPress(KeyAscii)
On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	End if
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) And lgStrPrevKey <> "" Then
		DbQuery
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
On Error Resume Next                                                 

	lgStrPrevKey = ""
	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
	
	Call DbQuery()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtMaxRows

	Call LayerShowHide(1)  

	DbQuery = False 
	txtMaxRows = vspdData.MaxRows
	 
	strVal = BIZ_PGM_ID &	"?txtPlantCd="    & Trim(txtPlantCd.Value)	& _
							"&txtItemCd="     & Trim(txtItemCd.Value)	& _
							"&txtLotNo="      & Trim(txtLotNo.Value)	& _
							"&txtLotSubNo="   & Trim(txtLotSubNo.Value)	& _
							"&txtSeq="        & Trim(txthSeq.Value)		& _
							"&lgStrPrevKey=" & lgStrPrevKey			& _
							"&txtMaxRows="    & txtMaxRows
	  
	Call RunMyBizASP(MyBizASP, strVal)
	    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()            
	vspdData.focus 
End Function

'----------  Coding part  -------------------------------------------------------------
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD">
			<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
				<TR>
					<TD CLASS="TD5" NOWRAP>공장</TD>
					<TD CLASS="TD6" NOWRAP>
					 <input TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="14">&nbsp;<input TYPE=TEXT NAME="txtPlantNm" SIZE="20" tag="14" >
					</TD>
					<TD CLASS="TD5" NOWRAP></TD>
					<TD CLASS="TD6" NOWRAP></TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP>품목</TD>
					<TD CLASS="TD6" NOWRAP>
					 <input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="품목" tag="14">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="20" tag="14" >
					</TD>     
					<TD CLASS="TD5" NOWRAP>LOT NO</TD>
					<TD CLASS="TD6" NOWRAP><input NAME="txtLotNo" TYPE="Text" SIZE="12" MAXLENGTH="12" STYLE="Text-Transform: uppercase" tag="14" ALT = "Lot No.">&nbsp;<input NAME="txtLotSubNo" TYPE="Text" SIZE="5" MAXLENGTH="3" STYLE="Text-Transform: uppercase" tag="14" ALT = "지번"></TD>
				</TR>   
			</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=*>
		<script language =javascript src='./js/i2511ra1_I422671407_vspdData.js'></script>
	</TD></TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>     </TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>  
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txthSeq" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


