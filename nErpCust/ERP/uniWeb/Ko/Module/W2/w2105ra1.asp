
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<!--
'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���Աݾ����� 
'*  3. Program ID           : W1111RA1
'*  4. Program Name         : W1111RA1.asp
'*  5. Program Desc         : ���Թ��� ��ȸ�˾� 
'*  6. Modified date(First) : 2004/12/31
'*  7. Modified date(Last)  : 2004/12/31
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
 -->
<HTML>
<HEAD>
<TITLE>���� �ݾ� ��ȸ</TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID 		= "w2105rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 100                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 18												'��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              

Dim  IsOpenPop                                                  '��: ��ũ                                  

Dim arrReturn
Dim arrParent
Dim arrParam

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)

	 '------ Set Parameters from Parent ASP ------ 
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6

	top.document.title = "���� �ݾ� ��ȸ"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    Redim arrReturn(0, 0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn
 
End Sub

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1				= 1
    C_W2				= 2
    C_W3				= 3
    C_W4				= 4
    C_W5				= 5
    C_W6				= 6
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

End Sub


'======================================================================================================
'   Function Name : EscPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.focus
		End Select
	End With
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
				.txtDealBpCd.focus
		End Select
	End With
End Function
'========================================  2.3 LoadInfTB19029()  =========================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		'patch version
		'ggoSpread.Spreadinit "V20041222",,PopupParent.gAllowDragDropSpread    
			 
		.ReDraw = false

		.MaxCols = C_W5 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols														'��: ����� �� Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData
			 
		ggoSpread.SSSetEdit		C_W1,	"�����ڵ�",	10,,,100,1
		ggoSpread.SSSetEdit		C_W2,	"��������",	10,,,100,1
		ggoSpread.SSSetFloat	C_W3,	"�ݾ�",	10,		PopupParent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		PopupParent.gComNum1000,		PopupParent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetEdit		C_W4,	"��ǥ�����ڵ�",	15,,,100,1
		ggoSpread.SSSetEdit		C_W5,	"��ǥ����",	10,,,100,1
		ggoSpread.SSSetEdit		C_W6,	"����/��������",	10,,,100,1
			 
		'Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		'Call ggoSpread.SSSetColHidden(C_W4,C_W4,True)
		.Col = C_W4													'��: ����� �� Hidden Column
		.ColHidden = True 
		.Col = C_W6													'��: ����� �� Hidden Column
		.ColHidden = True 
							
		'Call InitSpreadComboBox()
		.OperationMode = 5	
		.ReDraw = true
			
		Call SetSpreadLock 
    
    End With
    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_W1, -1, C_W5

    .vspdData.ReDraw = True

    End With
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow, iLngSelectedRows, sW4, dblAmt, Idx, arrTemp

	With frm1.vspdData
		iLngSelectedRows = .SelModeSelCount

		If iLngSelectedRows > 0 Then 
			intInsRow = 0

			Redim arrTemp(iLngSelectedRows, 5)

			For intRowCnt = 1 To .MaxRows

				.Row = intRowCnt
				.Col = C_W4
				If .SelModeSelected Then
					.Col = C_W4	
					Idx = CheckSavedW4(arrTemp, .Text)
					If Idx > -1 Then

						.Col = C_W3		: arrTemp(Idx, 2) = arrTemp(Idx, 2) + UNICDbl(.Value)
			
					Else
						.Col = C_W1		: arrTemp(intInsRow, 0) = .Value
						.Col = C_W2		: arrTemp(intInsRow, 1) = .Value
						.Col = C_W3		: arrTemp(intInsRow, 2) = arrTemp(intInsRow, 2) + UNICDbl(.Value)
						.Col = C_W4		: arrTemp(intInsRow, 3) = .Value
						.Col = C_W5		: arrTemp(intInsRow, 4) = .Value
						.Col = C_W6		: arrTemp(intInsRow, 5) = .Value
					
						intInsRow = intInsRow + 1
					End If
				End IF
			Next
			
			intRowCnt = GetGPCount(arrTemp)
			ReDim arrReturn(intRowCnt, 5)
			
			For intRowCnt = 0 To  intRowCnt -1
				arrReturn(intRowCnt, 0) = arrTemp(intRowCnt, 0)
				arrReturn(intRowCnt, 1) = arrTemp(intRowCnt, 1)
				arrReturn(intRowCnt, 2) = arrTemp(intRowCnt, 2)
				arrReturn(intRowCnt, 3) = arrTemp(intRowCnt, 3)
				arrReturn(intRowCnt, 4) = arrTemp(intRowCnt, 4)
				arrReturn(intRowCnt, 5) = arrTemp(intRowCnt, 5)
			Next
		End if			
	End With

	Self.Returnvalue = arrReturn
	Self.Close()
End Function

Function CheckSavedW4(Byref arrRet2, Byval pW4)	' -- ����� W4�� �˻��Ѵ�..
	Dim iRow, iMaxRows, sW4, iGPCnt

	iMaxRows = UBound(arrRet2)
		
	For iRow = 0 To iMaxRows-1
			
		If arrRet2(iRow, 3) = pW4 Then
			CheckSavedW4 = iRow
			Exit Function
		End If
	Next

	CheckSavedW4 = -1
End Function

Function GetGPCount(Byref arrRet2)	' -- �迭�� ���� üũ 
	Dim iRow, iMaxRows, sW4, iGPCnt

	iMaxRows = UBound(arrRet2)
		
	For iRow = 0 To iMaxRows-1
			
		If arrRet2(iRow, 3) = "" Then
			GetGPCount = iRow
			Exit Function
		End If
	Next

	GetGPCount = iRow
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "1", PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")
    
	Call InitVariables()														
	Call SetDefaultVal()	
	Call InitSpreadSheet()
    Call InitComboBox()
    Call FncQuery
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

	Dim IntRetCD
    FncQuery = False                                            
    
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables() 											
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    '-----------------------
    'Check condition area
    '-----------------------

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '��: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & PopupParent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & arrParam(1)      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & arrParam(2)      '��: Query Key   
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  						
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If
		
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell�� dblclick�߰ų�....
		Exit Function
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Function
	End If
	Call OKClick
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	    
End Sub
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

End Sub

Function vspdData_KeyPress(KeyAscii)

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG ��																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD>
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Call FncQuery()"></IMG>&nbsp;
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" ></IMG>
					</TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=  <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

