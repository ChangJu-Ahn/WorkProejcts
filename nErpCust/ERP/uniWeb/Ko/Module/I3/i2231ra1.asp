<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I2231ra1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 마감후 발생수불 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/01
'*  8. Modified date(Last)  : 2001/11/10
'*  9. Modifier (First)     : Han
'* 10. Modifier (Last)      : Han
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>

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

Const BIZ_PGM_ID = "i2231rb1.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_ItemDocumentNo          
Dim C_DocumentDt
Dim C_TrnsType
Dim C_SlCd
Dim C_ItemCd
Dim C_LotNo
Dim C_LotSubNo
Dim C_SeqNo
Dim C_SubSeqNo
Dim C_DnNo
Dim C_PoNo
Dim C_Qty
Dim C_Price


Dim arrReturn
Dim arrParam

Dim arrPlantCd
Dim arrPlantNm
Dim arrUserFlag
Dim arrInvClsDt

'------ Set Parameters from Parent ASP ------ 
arrParam   = window.dialogArguments
Set PopupParent = arrParam(0)

arrPlantCd   = arrParam(1)
arrPlantNm   = arrParam(2)
arrInvClsDt  = arrParam(3)

top.document.title = PopupParent.gActivePRAspName
'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

Dim lgUserFlag      

'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop          
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgStrPrevKey1 = ""                          
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0                          
    '---- Coding part--------------------------------------------------------------------    
    lgLngCurRows = 0                           
    Self.Returnvalue = Array("")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

    txtPlantCd.value	= arrPlantCd 
    txtPlantNm.value	= arrPlantNm 
	txtInvClsDt.value	= arrInvClsDt

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
		.MaxCols = C_Price+1          
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")

		ggoSpread.SSSetEdit C_ItemDocumentNo, "수불번호", 16
		ggoSpread.SSSetEdit C_TrnsType, "수불유형", 8
		ggoSpread.SSSetDate C_DocumentDt, "수불발생일", 12 , 2 , PopupParent.gDateFormat
		ggoSpread.SSSetEdit C_SlCd, "창고", 7 
		ggoSpread.SSSetEdit C_DnNo, "출하번호", 18
		ggoSpread.SSSetEdit C_PoNo, "구매번호", 18
		ggoSpread.SSSetEdit C_ItemCd, "품목", 18
		ggoSpread.SSSetEdit C_LotNo, "Lot No.", 12
		ggoSpread.SSSetFloat C_SeqNo,    "수불일련번호", 12, "6",                      ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_SubSeqNo, "수불상세번호", 12, "6",                      ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_LotSubNo, "순번",          6, "6",                      ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_Price,    "단가",         15, PopupParent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec 
		ggoSpread.SSSetFloat C_Qty,      "수량",         15, PopupParent.ggQtyNo,      ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		  
		'ggoSpread.MakePairsColumn()
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
	C_ItemDocumentNo	= 1
	C_DocumentDt		= 2
	C_TrnsType			= 3
	C_SlCd				= 4
	C_ItemCd			= 5
	C_LotNo				= 6
	C_LotSubNo			= 7
	C_SeqNo				= 8
	C_SubSeqNo			= 9
	C_DnNo				= 10
	C_PoNo				= 11
	C_Qty				= 12
	C_Price				= 13
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

		C_ItemDocumentNo = iCurColumnPos(1)
		C_DocumentDt     = iCurColumnPos(2)
		C_TrnsType       = iCurColumnPos(3)
		C_SlCd           = iCurColumnPos(4)
		C_ItemCd         = iCurColumnPos(5)
		C_LotNo          = iCurColumnPos(6)
		C_LotSubNo       = iCurColumnPos(7)
		C_SeqNo          = iCurColumnPos(8)
		C_SubSeqNo       = iCurColumnPos(9)
		C_DnNo           = iCurColumnPos(10)
		C_PoNo           = iCurColumnPos(11)
		C_Qty            = iCurColumnPos(12)
		C_Price          = iCurColumnPos(13)		
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	vspdData.ReDraw = False    

	ggoSpread.Source = vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()      
	vspdData.ReDraw = True
End Sub


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

    Call InitSpreadSheet()
    Call InitVariables()                                                    
    Call SetDefaultVal()
   
    Call DbQuery()
   
End Sub



'******************************  3.2.1 Object Tag 처리  *********************************************
' Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************
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
 
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    if  vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then 
		If lgStrPrevKey1 <> ""  Or lgStrPrevKey2 <> "" Or lgStrPrevKey3 <> "" Or lgStrPrevKey4 <> "" Then     
			DbQuery
		End If
	End if 
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
On error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	End if
End Function

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
    Call InitVariables              
    ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then        
       Exit Function
    End If
    
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
    
    strVal = BIZ_PGM_ID & "?txtPlantCd="		& Trim(txtPlantCd.value)	& _
						"&txtMaxRows="			& vspdData.MaxRows			& _			
						"&lgStrPrevKeya1="		& lgStrPrevKey1				& _
						"&lgStrPrevKeya2="		& lgStrPrevKey2				& _
						"&lgStrPrevKeya3="		& lgStrPrevKey3				& _
						"&lgStrPrevKeya4="		& lgStrPrevKey4
   
    Call RunMyBizASP(MyBizASP, strVal)          
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()           
 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")         

 vspdData.Focus
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR><TD HEIGHT=40>
  <TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
		<TD <%=HEIGHT_TYPE_02%> >
		</TD>
   </TR>
   <TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%> > 
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><input NAME="txtPlantCd" TYPE="Text" MAXLENGTH="18" tag="14XXXU" ALT = "공장" size=15>&nbsp;<input NAME="txtPlantNm" TYPE="Text" MAXLENGTH="40" tag="14N"></TD>
						<TD CLASS="TD5" NOWRAP>최종마감년월</TD>
						<TD CLASS="TD6" NOWRAP><input NAME="txtInvClsDt" TYPE="Text" MAXLENGTH="7" tag="14XXXU" ALT = "최종마감년월" size=10></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> 
        </TD>
    </TR>
  </TABLE>
 </TD>
 </TR>
 <TR>
	<TD HEIGHT=* VALIGN=TOP>
		<TABLE <%=LR_SPACE_TYPE_20%>>
			<TR>
				<TD>
					<script language =javascript src='./js/i2231ra1_I477616244_vspdData.js'></script>
				</TD>
			</TR>  
		</TABLE>
	</TD>
 </TR>
 <TR>
     <TD <%=HEIGHT_TYPE_01%> >
     </TD>
 </TR>
 <TR HEIGHT=20 >
     <TD>
         <TABLE <%=LR_SPACE_TYPE_30%> >
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD WIDTH=* ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>
        </TD>
        <TD WIDTH=10>&nbsp;</TD>
    </TR>
         </TABLE>
     </TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemDocumentNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDocumentYear" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hSeqNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hSubSeqNo" tag="24" TABINDEX="-1">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

