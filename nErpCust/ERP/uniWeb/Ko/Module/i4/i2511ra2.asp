<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List onhand stock detail
'*  3. Program ID           : I2511ra2.asp
'*  4. Program Name         : 현 재고 상세 조회 
'*  5. Program Desc         : 현재 창고에 있는 품목의 상세정보를 조회한다.
'*  6. Comproxy List        : 
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/01
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/01 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<!--'#########################################################################################################
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
Const BIZ_PGM_ID = "i2511rb2.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'=========================================================================================================
Dim C_SLCd               
Dim C_SLNm     
Dim C_TrackingNo 
Dim C_GoodQty    
Dim C_BadQty     
Dim C_InspQty    
Dim C_TrnsQty    
Dim C_PrevGoodQty
Dim C_PrevBadQty 
Dim C_PrevInspQty
Dim C_PrevTrnsQty

Dim arrReturn
Dim arrParam

Dim arrPlant_Cd
Dim arrPlant_Nm
Dim arrItem_Nm
Dim arrItem_Cd
Dim arrLot_No
Dim arrLotSub_No
  
'------ Set Parameters from Parent ASP ------
arrParam   = window.dialogArguments
Set PopupParent = arrParam(0)

arrPlant_Cd  = arrParam(1)
arrItem_Cd   = arrParam(2)
arrLot_No    = arrParam(3)
arrLotSub_No = arrParam(4)
arrPlant_Nm  = arrParam(5)
arrItem_Nm   = arrParam(6)

top.document.title = PopupParent.gActivePRAspName
'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevSubKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop          
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntGrpCount	= 0                          
    lgLngCurRows	= 0                           
    Self.Returnvalue = Array("")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	txtItem_Cd.value	= arrItem_Cd 
	txtPlant_Cd.value	= arrPlant_Cd
	txtPlant_Nm.value	= arrPlant_Nm
	txtItem_Nm.value	= arrItem_Nm
	txtLot_No.value		= arrLot_No
	txtLotSub_No.value  = arrLotSub_No
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
		.MaxCols = C_PrevTrnsQty+1           
		.MaxRows = 0
		Call GetSpreadColumnPos("A")	  
		     
		ggoSpread.SSSetEdit  C_SLCd,        "창고",            7
		ggoSpread.SSSetEdit  C_SLNm,        "창고명",         20 
		ggoSpread.SSSetEdit  C_TrackingNo,  "Tracking No",    20
		ggoSpread.SSSetFloat C_GoodQty,     "양품재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_BadQty,      "불량재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_InspQty,     "검사중수량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_TrnsQty,     "이동재고량",     15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevGoodQty, "전월양품재고량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevBadQty,  "전월불량재고량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevInspQty, "전월검사중수량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_PrevTrnsQty, "전월이동중수량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		  
		'ggoSpread.MakePairsColumn()
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
		  
		Call SetSpreadLock 
	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_SLCd        = 1   
	C_SLNm        = 2
	C_TrackingNo  = 3
	C_GoodQty     = 4
	C_BadQty      = 5
	C_InspQty     = 6
	C_TrnsQty     = 7
	C_PrevGoodQty = 8
	C_PrevBadQty  = 9
	C_PrevInspQty = 10
	C_PrevTrnsQty = 11
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

		C_SLCd        = iCurColumnPos(1)  
		C_SLNm        = iCurColumnPos(2)
		C_TrackingNo  = iCurColumnPos(3)
		C_GoodQty     = iCurColumnPos(4)
		C_BadQty      = iCurColumnPos(5)
		C_InspQty     = iCurColumnPos(6)
		C_TrnsQty     = iCurColumnPos(7)
		C_PrevGoodQty = iCurColumnPos(8)
		C_PrevBadQty  = iCurColumnPos(9)
		C_PrevInspQty = iCurColumnPos(10)
		C_PrevTrnsQty = iCurColumnPos(11)
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
		     
		vspdData.Col = C_SLCd
		arrReturn(0) = vspdData.Text
		vspdData.Col = C_SLNm     
		arrReturn(1) = vspdData.Text
		vspdData.Col = C_TrackingNo 
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_GoodQty    
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_BadQty     
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_InspQty    
		arrReturn(5) = vspdData.Text
		vspdData.Col = C_TrnsQty    
		arrReturn(6) = vspdData.Text
		vspdData.Col = C_PrevGoodQty
		arrReturn(7) = vspdData.Text
		vspdData.Col = C_PrevBadQty 
		arrReturn(8) = vspdData.Text
		vspdData.Col = C_PrevInspQty
		arrReturn(9) = vspdData.Text
		vspdData.Col = C_PrevTrnsQty
		arrReturn(10) = vspdData.Text
		    
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

    Call LoadInfTB19029             
    Call ggoOper.LockField(Document, "N")                                    
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                                    
    Call SetDefaultVal()
    Call DbQuery()
 
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

Function vspdData_KeyPress(KeyAscii)
On Error Resume Next
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	if  vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then 
		If lgStrPrevKey <> "" and lgStrPrevSubKey <> "" Then    
			DbQuery
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
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
    Call DbQuery              
       
    FncQuery = True           
    
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      
    Call LayerShowHide(1)
    DbQuery = False
    
    Err.Clear                                                             

    Dim strVal    
    strVal = BIZ_PGM_ID &	"?txtPlant_Cd="		& Trim(txtPlant_Cd.value)	& _
							"&txtItem_Cd="      & Trim(txtItem_Cd.value)	& _
							"&txtLot_No="       & Trim(arrLot_No)			& _
							"&txtLotSub_No="    & Trim(arrLotSub_No)		& _
							"&lgStrPrevKey="    & lgStrPrevKey				& _
							"&lgStrPrevSubKey=" & lgStrPrevSubKey			& _ 
							"&txtMaxRows="      & vspdData.MaxRows
    
    Call RunMyBizASP(MyBizASP, strVal)        
        
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================
Function DbQueryOk()             
    Call ggoOper.LockField(Document, "Q")
    vspdData.focus        
End Function


'----------  Coding part  -------------------------------------------------------------
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
 <TR>
 <TD HEIGHT=40>
  <FIELDSET CLASS="CLSFLD">
   <TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
    <TR>
        <TD CLASS="TD5" NOWRAP>공장</TD>
        <TD CLASS="TD6" NOWRAP><input NAME="txtPlant_Cd" TYPE="Text" size="10" MAXLENGTH="4" tag="14XXXU" ALT = "공장">&nbsp;<input NAME="txtPlant_Nm" TYPE="Text" SIZE="20" MAXLENGTH="40" tag="14N"></td>
        <TD CLASS="TD5" NOWRAP>LOT번호</td>
        <TD CLASS="TD6" NOWRAP><input NAME="txtLot_No" TYPE="Text" size="12" MAXLENGTH="12" tag="14XXXU" ALT = "LOT번호">&nbsp;<input NAME="txtLotSub_No" TYPE="Text" MAXLENGTH="3" tag="14" ALT = "LOTSUB번호" size=5></TD>    
    </TR>
    <TR>
        <TD CLASS="TD5" NOWRAP>품목</td>
        <TD CLASS="TD6" NOWRAP><input NAME="txtItem_Cd" TYPE="Text" size="15" MAXLENGTH="18" tag="14XXXU" ALT = "품목">&nbsp;<input NAME="txtItem_Nm" TYPE="Text" SIZE="20" MAXLENGTH="40" tag="14N"></TD>        
        <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
        <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
    </TR>
   </TABLE>
  </FIELDSET>
 </TD>
 </TR>
 <TR>
 <TD HEIGHT=*>
  <script language =javascript src='./js/i2511ra2_OBJECT1_vspdData.js'></script>
 </TD>
 </TR>
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
 <DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
 
 
 

