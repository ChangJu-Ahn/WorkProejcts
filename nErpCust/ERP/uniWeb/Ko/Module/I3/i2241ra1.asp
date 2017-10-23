<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List onhand stock detail
'*  3. Program ID           : I2241ra1.asp
'*  4. Program Name         : �� ��� �� ��ȸ 
'*  5. Program Desc         : ���� â�� �ִ� ǰ���� �������� ��ȸ�Ѵ�.
'*  6. Comproxy List        : 
'                             
'                             
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/01
'*  8. Modified date(Last)  : 2000/04/01
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Nam hoon kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/01 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                

'******************************************  1.2 Global ����/��� ����  ***********************************
' 1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

Const BIZ_PGM_ID = "i2241rb1.asp"


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_MovType      
Dim C_MovTypeNm 
Dim C_Qty 
Dim C_Amount 

Dim arrParam

Dim strPlantCd
Dim strPlantNm
Dim strYyyyMm
Dim strItemCd
Dim strItemNm
Dim strItemSpec
Dim strUnit
Dim strDateFormat

arrParam   = window.dialogArguments
Set PopupParent = arrParam(0)

strPlantCd    = arrParam(1)
strPlantNm    = arrParam(2)
strYyyyMm     = arrParam(3)
strItemCd     = arrParam(4)
strItemNm     = arrParam(5)
strItemSpec   = arrParam(6)
strUnit       = arrParam(7)
strDateFormat = arrParam(8)

top.document.title = PopupParent.gActivePRAspName
'==========================================  1.2.2 Global ���� ����  =====================================
' 1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
' 2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntGrpCount = 0                      
    lgStrPrevKey = ""
    lgLngCurRows = 0                            
    Self.Returnvalue = Array("")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	txtPlantCd.Value	= strPlantCd
	txtPlantNm.Value	= strPlantNm
	txtYyyyMm.Value		= strYyyyMm
	txtItemCd.Value		= strItemCd
	txtItemNm.Value		= strItemNm 
	txtUnit.Value		= strUnit
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
		.MaxCols = C_Amount+1           
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
				     
		ggoSpread.SSSetEdit C_MovType,   "�̵�����",   10, 2
		ggoSpread.SSSetEdit C_MovTypeNm, "�̵�������", 31
		ggoSpread.SSSetFloat C_Qty,      "����",       22, PopupParent.ggQtyNo,        ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_Amount,   "�ݾ�",       22, PopupParent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec

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
	C_MovType   = 1
	C_MovTypeNm = 2
	C_Qty       = 3
	C_Amount    = 4
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

		C_MovType   = iCurColumnPos(1)
		C_MovTypeNm = iCurColumnPos(2)
		C_Qty       = iCurColumnPos(3)
		C_Amount    = iCurColumnPos(4)

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


'=========================================  2.3.2 CancelClick()  ========================================
'= Name : CancelClick()                    =
'= Description : Return Array to Opener Window for Cancel button click         =
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
 
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
	Select case UCase(pstr1)
	case "PON"
		window.document.search.style.cursor = "wait"
	case "POFF"
		window.document.search.style.cursor = ""
	End Select
End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029          
	Call ggoOper.LockField(Document, "N")          
	Call InitSpreadSheet
	Call InitVariables                                                 
	Call SetDefaultVal()
	Call DbQuery()
End Sub


'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
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
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If vspdData.MaxRows = 0 Then
		Exit Sub
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
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
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
	'----------  Coding part  -------------------------------------------------------------
	if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then '��: ������ üũ 
		If lgStrPrevKey <> "" Then       
			DbQuery
		End If
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress()
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(keyAscii)
On error Resume Next
	if KeyAscii = 27 Then
		Call CancelClick()
	End IF
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
    'Query function call area
    '-----------------------
    'Call DbQuery               
       
    FncQuery = True               
    
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim strYear
	Dim strMonth
	Dim strDay

	 
	Call ExtractDateFrom(txtYyyyMm.value, strDateFormat, PopupParent.gComDateType, strYear, strMonth, strDay)

	DbQuery = False
	 
	Call LayerShowHide(1)
	 
Err.Clear                                                             
	 
	strVal = BIZ_PGM_ID &	"?txtPlantCd="	& strPlantCd		& _
							"&txtYyyy="     & strYear			& _
							"&txtMm="       & strMonth			& _
							"&txtItemCd="   & strItemCd			& _
							"&lgStrPrevKey=" & lgStrPrevKey		& _
							"&txtMaxRows=" & vspdData.MaxRows
	 
	Call RunMyBizASP(MyBizASP, strVal)         
	        
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
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
							<TD CLASS="TD5" NOWRAP>����</TD>      
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH="4" tag="14XXXU" ALT = "����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>    
							<TD CLASS="TD5" NOWRAP>����</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtYyyyMm" SIZE=7 MAXLENGTH="6" CLASS=FPDTYYYMM tag="14" ALT="����"></TD>       
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>ǰ��</td>
							<TD CLASS="TD6" NOWRAP><input NAME="txtItemCd" TYPE="Text" SIZE=15 MAXLENGTH="18" tag="14XXXU" ALT = "ǰ��">&nbsp;<input NAME="txtItemNm" TYPE="Text" SIZE=20 MAXLENGTH="30" tag="14XXXU"></TD>
							<TD CLASS="TD5" NOWRAP>������</TD>
							<TD CLASS="TD6" NOWRAP><input NAME="txtUnit" TYPE="Text" SIZE=10 MAXLENGTH="10" tag="14XXXU"></td>
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
						<script language =javascript src='./js/i2241ra1_OBJECT1_vspdData.js'></script>
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
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
