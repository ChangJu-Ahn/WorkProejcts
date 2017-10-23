<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PA1
'*  4. Program Name         : 
'*  5. Program Desc         : ǰ�� �˻��׸� �˾� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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

Const BIZ_PGM_ID = "q1211pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim C_InspOrder
Dim C_InspItemCd 
Dim C_InspItemNm 
Dim C_InspMthdCd 
Dim C_InspMthdNm 

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim lgInspClassCd
Dim lgInspMthdCd

Dim hInspItemCd
Dim ArrParent
Dim arrParam					 '--- First Parameter Group 
ReDim arrParam(11)
Dim arrReturn				 '--- Return Parameter Group 

Dim IsOpenPop 

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)
arrParam(1) = ArrParent(2)
arrParam(2) = ArrParent(3)
arrParam(3) = ArrParent(4)
arrParam(4) = ArrParent(5)
arrParam(5) = ArrParent(6)
arrParam(6) = ArrParent(7)
arrParam(7) = ArrParent(8)
arrParam(8) = ArrParent(9)
arrParam(9) = ArrParent(10)
arrParam(10) = ArrParent(11)
arrParam(11) = ArrParent(12)

top.document.title = PopupParent.gActivePRAspName
'top.document.title = "ǰ�� �˻��׸� �˾�"

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_InspOrder = 1
	C_InspItemCd = 2
    C_InspItemNm = 3
    C_InspMthdCd = 4
    C_InspMthdNm = 5
	
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgStrPrevKey = ""										'initializes Previous Key		
	lgSortKey    = 1                            '��: initializes sort direction
	 '------ Coding part ------ 
	Self.Returnvalue = Array("")
	vspdData.MaxRows = 0
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()

	txtPlantCd.value = UCase(arrParam(0))
	txtPlantNm.Value = arrParam(1)
	txtItemCd.Value = UCase(arrParam(2))
	txtItemNm.Value = arrParam(3)
	lgInspClassCd = UCase(arrParam(4))
	txtInspClassNm.Value = arrParam(5)
	If lgInspClassCd = "P" then
		txtRoutNo.value = UCase(arrParam(6))
		txtRoutNoDesc.value = arrParam(7)
		txtOprNo.value  = UCase(arrParam(8))
	End If
	txtInspItemCd.Value = UCase(arrParam(9))
	txtInspItemNm.Value = arrParam(10)
	lgInspMthdCd = arrParam(11)
		
	Self.Returnvalue = Array("")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20040518",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False
	vspdData.MaxCols = C_InspMthdNm + 1
	vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	    
	ggoSpread.SSSetEdit C_InspOrder, "�˻����", 10, 1
	ggoSpread.SSSetEdit C_InspItemCd, "�˻��׸��ڵ�", 15
	ggoSpread.SSSetEdit C_InspItemNm, "�˻��׸��", 30	
	ggoSpread.SSSetEdit C_InspMthdCd, "�˻����ڵ�", 15
	ggoSpread.SSSetEdit C_InspMthdNm, "�˻��ĸ�", 30	
	
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
			C_InspOrder			= iCurColumnPos(1)
			C_InspItemCd		= iCurColumnPos(2)
			C_InspItemNm		= iCurColumnPos(3)
			C_InspMthdCd		= iCurColumnPos(4)
			C_InspMthdNm		= iCurColumnPos(5)
    End Select
End Sub

'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : InspItemPlant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�˻��׸� �˾�"		' �˾� ��Ī 
	arrParam(1) = "Q_INSPECTION_ITEM"		' TABLE ��Ī 
	arrParam(2) = Trim(txtInspItemCd.Value)		' Code Condition
	arrParam(3) = ""				' Name Cindition
	arrParam(4) = ""				' Where Condition
	arrParam(5) = "�˻��׸�"			
	
	arrField(0) = "INSP_ITEM_CD"			' Field��(0)
	arrField(1) = "INSP_ITEM_NM"		' Field��(1)
	
	arrHeader(0) = "�˻��׸��ڵ�"			' Header��(0)
	arrHeader(1) = "�˻��׸��"		' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtInspItemCd.Focus	
	If arrRet(0) <> "" Then
		txtInspItemCd.Value    = arrRet(0)
		txtInspItemNm.Value    = arrRet(1)
	End If	
	
	txtInspItemCd.focus
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
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

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'============================================= EnableField()  ======================================
'=	Event Name : EnableField
'=	Event Desc :
'========================================================================================================
Sub EnableField(Byval strInspClass)
	If	strInspClass = "P" Then
		Process.style.display = ""
	Else	
		Process.style.display = "none"
	End if
End Sub


'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field
	Call InitVariables()
	Call SetDefaultVal()
	Call EnableField(lgInspClassCd)
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'==========================================  3.2.1 FncQuery =======================================
'
'
'========================================================================================================
Function FncQuery()
	FncQuery = False
    vspdData.MaxRows = 0

	lgQueryFlag = "1"
	lgStrPrevKey = ""
	
    lgIntFlgMode = PopupParent.OPMD_CMODE								'Indicates that current mode is Create mode	
	Self.Returnvalue = Array("")
		
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If DbQuery = false Then
		Exit Function		
	End If

	FncQuery = True
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' Ÿ��Ʋ cell�� dblclick�߰ų�....
	   Exit Function
	End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DbQuery = False Then
				Exit Sub
			End If
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

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()

	Dim strVal
	Dim txtMaxRows
	
   	Call LayerShowHide(1)  

	DbQuery = False 
	txtMaxRows = vspdData.MaxRows
	
	if lgStrPrevKey <> "" Then
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(txtPlantCd.Value)
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.Value)
		strVal = strVal & "&txtInspClassCd=" & lgInspClassCd
		strVal = strVal & "&txtRoutNo=" & Trim(txtRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(txtOprNo.value)
		strVal = strVal & "&txtInspMthdCd=" & lgInspMthdCd
		strVal = strVal & "&txtInspItemCd=" & hInspItemCd
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows
	else
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(txtPlantCd.Value)
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.Value)
		strVal = strVal & "&txtInspClassCd=" & lgInspClassCd
		strVal = strVal & "&txtRoutNo=" & Trim(txtRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(txtOprNo.value)
		strVal = strVal & "&txtInspMthdCd=" & lgInspMthdCd
		strVal = strVal & "&txtInspItemCd=" & Trim(txtInspItemCd.Value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows	
	end if     
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True                                                          '��: Processing is NG
End Function

'********************************************  5.1 DbQueryOk()  *******************************************
' Function Name : DbQueryOk																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQueryOk()								'��: ��ȸ ������ ������� 

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
					<TR>
						<td CLASS="TD5" NOWPAP>����</td>
						<td CLASS="TD6" NOWPAP>
							<input TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" tag="14">&nbsp;<input TYPE=TEXT NAME="txtPlantNm" SIZE="20" tag="14" >
						</td>
						<td CLASS="TD5" NOWPAP>�˻�з�</td>
						<td CLASS="TD6" NOWPAP>
							<input TYPE=TEXT NAME="txtInspClassNm" SIZE="20" ALT="�˻�з�" tag="14">
						</td>					
					</TR>
					<TR>
						<td CLASS="TD5" NOWPAP>ǰ��</td>
						<td CLASS="TD6" NOWPAP>
							<input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="ǰ��" tag="14">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="20" tag="14" >
						</td>
						<TD CLASS="TD5" NOWPAP></TD>
						<TD CLASS="TD6" NOWPAP></TD>
					</TR>
					<TR ID="Process">
						<TD CLASS="TD5" NOWRAP>�����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="14" ALT="�����">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="14" ALT="����"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>�˻��׸�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="�˻��׸�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
							<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" >
						</TD>
						<TD CLASS="TD5" NOWPAP></TD>
						<TD CLASS="TD6" NOWPAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=*>
			<script language =javascript src='./js/q1211pa1_I373058909_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif" Style="CURSOR: hand" ALT="Search" NAME="search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

