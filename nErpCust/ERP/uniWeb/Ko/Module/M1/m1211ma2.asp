<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211MA2
'*  4. Program Name         : ����ó����к��� 
'*  5. Program Desc         : ����ó����к��� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/01/09
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID	= "m1211mb2.asp"
CONST BIZ_PGM_ID2	= "m1211mb201.asp"												'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'spdData
Dim C_PlantCd 	             '���� 
Dim C_PlantNm 	             '����� 
Dim C_ItemCd 	             'ǰ�� 
Dim C_ItemNm 	             'ǰ��� 
Dim C_SpplSpec 	             'ǰ��԰� 

'spdData2
Dim C_SpplCd                 '����ó 
Dim C_SpplNm 	             '����ó�� 
Dim C_Quota_Rate             '��к��� 
Dim C_Purpriority            '���ֹ�������ġ 
Dim C_Defflg                 '�ְ��޾�ü���� 
Dim C_SpplDlvylt             '����L/T
Dim C_GrpCd 	             '���ű׷� 
Dim C_GrpNm 	             '���ű׷�� 
Dim C_ParentPlantCd
Dim C_ParentItemCd
Dim C_ParentRowNo
Dim C_RecordCnt

Dim lgIntFlgModeM           'Variable is for Operation Status
Dim lgStrPrevKeyM()			'Multi���� �������� ���� ���� 
Dim lglngHiddenRows()		'Multi���� �������� ���� ����	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgSortKey1
Dim lgSortKey2

Dim lgPageNo1
Dim lgCurrRow
Dim lgSpdHdrClicked

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE   
    lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode3            
    lgBlnFlgChgValue = False                
    lgIntGrpCount = 0                       
    lgStrPrevKey1 = ""						'initializes Previous Key
    lgStrPrevKey2 = ""						'initializes Previous Key
    
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    lgPageNo = 0
    lgPageNo1 = 0
    
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.Value = Parent.gPlant
	frm1.txtPlantNm.Value = Parent.gPlantNm
	Call SetToolbar("1110000000001111")
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
	Set gActiveSpdSheet = frm1.vspdData
End Sub
'====================================================================================================
Sub ReadCookiePage()
	
	if Trim(ReadCookie("m1211qa1_plantcd")) = "" then Exit Sub
	
	frm1.txtPlantCd.Value	 = ReadCookie("m1211qa1_plantcd")
	frm1.txtItemCd.Value	 = ReadCookie("m1211qa1_itemcd")
	
	Call MainQuery()
	
	Call WriteCookie("m1211qa1_plantcd","")
	Call WriteCookie("m1211qa1_itemcd","")
	Call WriteCookie("m1211qa1_suppliercd","")
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub
'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Then	
		C_PlantCd 	  = 1             '���� 
		C_PlantNm 	  = 2             '����� 
		C_ItemCd 	  = 3             'ǰ�� 
		C_ItemNm 	  = 4             'ǰ��� 
		C_SpplSpec 	  = 5             'ǰ��԰� 
		
	Else
		C_SpplCd      = 1             '����ó 
		C_SpplNm 	  = 2             '����ó�� 
		C_Quota_Rate  = 3             '��к��� 
		C_Purpriority = 4             '���ֹ�������ġ 
		C_Defflg      = 5            '�ְ��޾�ü���� 
		C_SpplDlvylt  = 6            '����L/T
		C_GrpCd 	  = 7            '���ű׷� 
		C_GrpNm 	  = 8            '���ű׷�� 
		C_ParentPlantCd = 9
		C_ParentItemCd	= 10
		C_ParentRowNo	= 11
		C_RecordCnt		= 12
	End If
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
    Call InitSpreadPosVariables(pvSpdNo)
	
	If pvSpdNo = "A" Then
		With frm1.vspdData	
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20030219",,Parent.gAllowDragDropSpread
	
			.ReDraw = false
			.MaxCols = C_SpplSpec + 1							
			.Col = .MaxCols:	.ColHidden = True
			.MaxRows = 0
    
			Call GetSpreadColumnPos("A")
 
			ggoSpread.SSSetEdit 	C_PlantCd, "����", 15
			ggoSpread.SSSetEdit 	C_PlantNm,"�����",20
			ggoSpread.SSSetEdit 	C_ItemCd,"ǰ��",20
			ggoSpread.SSSetEdit 	C_ItemNm, "ǰ���", 25
			ggoSpread.SSSetEdit 	C_SpplSpec, "ǰ��԰�", 25
				
			Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
			Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)

			Call SetSpreadLock("A") 

			.ReDraw = true
		End With
	
	Elseif  pvSpdNo = "B" Then
		With frm1.vspdData2	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20030219",,Parent.gAllowDragDropSpread

			.ReDraw = false
			.MaxCols = C_RecordCnt + 1							
			.MaxRows = 0
    
			Call GetSpreadColumnPos("B")
 
			ggoSpread.SSSetEdit		C_SpplCd		,"����ó"			, 10,,,10,2
			ggoSpread.SSSetEdit 	C_SpplNm		,"����ó��"			, 18
			SetSpreadFloatLocal		C_Quota_Rate	,"��к���(%)"		,15,1,5
			ggoSpread.SSSetEdit		C_Purpriority	,"���ֹ�������ġ"	,15
			ggoSpread.SSSetEdit 	C_Defflg		,"�ְ��޾�ü����"	,15, 2
			ggoSpread.SSSetEdit		C_SpplDlvylt	,"����L/T"			,15
			ggoSpread.SSSetEdit 	C_GrpCd			,"���ű׷�"			,15,,,4,2
			ggoSpread.SSSetEdit 	C_GrpNm			,"���ű׷��"		,20
			ggoSpread.SSSetEdit 	C_ParentPlantCd	, ""		, 10
			ggoSpread.SSSetEdit 	C_ParentItemCd	, ""		, 10
			ggoSpread.SSSetEdit     C_ParentRowNo	, ""		, 25,2,,,2
			ggoSpread.SSSetEdit     C_RecordCnt		, ""		, 25,2,,,2
	
			Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
			Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)
			
			Call ggoSpread.SSSetColHidden(C_ParentPlantCd,	C_RecordCnt,	True)		
			Call ggoSpread.SSSetColHidden(.MaxCols,			.MaxCols,	True)		
			
			Call SetSpreadLock("B") 
    
			.ReDraw = true
		End with
	End If
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
			ggoSpread.SpreadLock		-1 , -1
			Call SetSpreadColor(-1,-1)
		Else
			.vspdData.ReDraw = False
			ggoSpread.SpreadLock		-1 , -1
			ggoSpread.SpreadLock		C_PlantCd , -1
			ggoSpread.SpreadUnLock		C_Quota_Rate , -1, -1
			ggoSpread.SSSetRequired		C_Quota_Rate, -1, -1                  '��к� 
			.vspdData.ReDraw = True
		End IF
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected		C_PlantCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_PlantNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ItemCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SpplSpec, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm 		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_SpplSpec		= iCurColumnPos(5)
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpplCd 		= iCurColumnPos(1)
			C_SpplNm		= iCurColumnPos(2)
			C_Quota_Rate	= iCurColumnPos(3)
			C_Purpriority	= iCurColumnPos(4)
			C_Defflg        = iCurColumnPos(5)
			C_SpplDlvylt    = iCurColumnPos(6)
			C_GrpCd         = iCurColumnPos(7)
			C_GrpNm         = iCurColumnPos(8)
			C_ParentPlantCd = iCurColumnPos(9)
			C_ParentItemCd  = iCurColumnPos(10)
			C_ParentRowNo   = iCurColumnPos(11)
			C_RecordCnt     = iCurColumnPos(12)
	End Select
End Sub	

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp  ���� 
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value   = arrRet(0)		
		frm1.txtPlantNm.value	= arrret(1)
		frm1.txtPlantCd.focus
	End If	
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Plant PopUp ǰ�� 
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If
End Function

'------------------------------------------  OpenBP()  ---------------------------------------------
'	Name : OpenBP()
'	Description : SpplCd PopUp ����ó 
'---------------------------------------------------------------------------------------------------------
Function OpenBP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_SpplCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "����ó"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "����ó"		
    arrHeader(1) = "����ó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_SpplCd
		frm1.vspdData.Row = frm1.vspdData.ActiveRow

		frm1.vspdData.Text = arrRet(0)		
		frm1.vspdData.Col  = C_SpplNm
		frm1.vspdData.Text = arrret(1)
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
		Call SpplChange()	
	End If		
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )

   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"   
        Case 7                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "7" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","99"  
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                
    Call ggoOper.LockField(Document, "N")              
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)                               
    Call SetDefaultVal
    Call InitVariables  
    Call ReadCookiePage()
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    Else
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
    End IF

	Set gActiveSpdSheet = frm1.vspdData
	    
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey1 = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey1 = 2
		Else
			ggoSpread.SSSort Col, lgSortKey1	'Sort in Descending
			lgSortkey1 = 1
		End If
	Else
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If    			
End Sub

'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
 	Dim strShowDataFirstRow
 	Dim strShowDataLastRow
 	Dim lngStartRow
 	Dim i,k
 	Dim strFlag,strFlag1
 	Dim iActiveRow
 	
 	gMouseClickStatus = "SP2C"   

 	Set gActiveSpdSheet = frm1.vspdData2
 	
 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    Else
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
    End IF
 	
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 AND Col <> 0 Then	'2003-03-01 Release �߰� 
 		ggoSpread.Source = frm1.vspdData2

 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = frm1.vspdData.MaxCols

 		iActiveRow = CInt(frm1.vspdData.Text)
 		
 		frm1.vspdData2.Redraw = False
		lngStartRow = CInt(ShowFromData(iActiveRow, CInt(lglngHiddenRows(iActiveRow - 1))))
		frm1.vspdData2.Redraw = True
		
		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Descending
 			lgSortKey2 = 1
		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
 	
 	With frm1.vspdData2
 		For i = 1 to .MaxRows
 			.Row = i
 			.col = 0	
 			If .Rowhidden = False Then
 				k = K + 1
 				if .text <> ggoSpread.UpdateFlag  then
 					.text = k
 				end if
 			End If
 		Next
 	End With 	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release �߰� 
	    lgSpdHdrClicked = 1 
	End If
	
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
   
   If Button = 2 And gMouseClickStatus = "SP2C" Then
      gMouseClickStatus = "SP2CR"
   End If
End Sub    

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

''========================================================================================
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
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
	Dim lngRangeFrom
	Dim lngRangeTo	
	Dim lRow

    ggoSpread.Source = gActiveSpdSheet
    
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("A")
		Call ggoSpread.ReOrderingSpreadData
    ElseIf gActiveSpdSheet.Name = "vspdData2" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("B")
		frm1.vspdData2.Redraw = False
		
		Call ggoSpread.ReOrderingSpreadData("F")

		Call DbQuery2(frm1.vspdData.ActiveRow,False)
		
		lngRangeFrom = Clng(ShowDataFirstRow2)
		lngRangeTo = Clng(ShowDataLastRow2)
		
		lRow = frm1.vspdData.ActiveRow	'###�׸��� ������ ���Ǻκ�###
		frm1.vspdData2.Redraw = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo lRow
    End If
    
 	'------ Developer Coding part (Start)	
 	'------ Developer Coding part (End) 	
End Sub

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim strMark
	Dim iparentrow

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
	With frm1.vspdData2
		.Row = Row
		.Col = C_ParentRowNo
		iparentrow = .text
		.Col = 0
		strMark = .Text
		.Col = C_RecordCnt 
		.Text = strMark
	
		Call QuotaRateChange(Row)   
	End With
	
	With frm1.vspdData
		If strMark = ggoSpread.UpdateFlag Then
			.Row = iparentrow
			.Col = 0
			.Text = ggoSpread.UpdateFlag
		End if
	End With
End Sub	
'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If
	
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)	
End Sub

'=======================================================================================================
'   Event Name : Sub_vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	Dim lRow
	if Row = 0 then exit sub
	If Row <> NewRow And NewRow > 0 Then
		With frm1        
			If CheckRunningBizProcess = True Then
				Call SetActiveCell(frm1.vspdData,1,Row,"M","X","X")
				Exit Sub
			End If
			lgCurrRow = NewRow	
		End With
		
		With frm1.vspdData2
			.ReDraw = False
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.RowHidden = True
			.BlockMode = False
			.ReDraw = True
		End With
		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '/* �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* �ػ󵵿� ������� �������ǵ��� ���� - END */
		if Trim(lgPageNo) = "" then exit sub
		If lgPageNo > 0   Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
				
			Call DisableToolBar(Parent.TBC_QUERY)
			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End If
End Sub

'=======================================================================================================
' Function Name : DefaultCheck
' Function Desc : 
'=======================================================================================================
Function DefaultCheck()
	DefaultCheck = False
	Dim i
	Dim j
	Dim RequiredColor 

	ggoSpread.Source = frm1.vspdData2
	RequiredColor = ggoSpread.RequiredColor
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				.Col = 0
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
					For j = 1 To .MaxCols
						.Col = j
						If .BackColor = RequiredColor Then
							If Len(Trim(.Text)) < 1 Then
								.Row = 0
								Call DisplayMsgBox("970021","X",.Text,"")
								Call SetActiveCell(frm1.vspdData2,j,i,"M","X","X")
								Exit Function
							End If
						End If			
					Next
				End If
			End If
		Next
	End With
	DefaultCheck = True
End Function

'==========================================   QuotaRateChange()  ======================================
'	Name : QuotaRateChange()
'	Description : 
'=================================================================================================
Sub QuotaRateChange(ByVal Row)
    Dim iparentrow
    Dim iReqQty,iApportionQty,iquotarate 
    Dim totalquotarate,totalApportionQty
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim index 

	with frm1.vspdData2
		.Row		= Row    
		.Col		= C_ParentRowNo
		iparentrow  = Trim(.text)
		
		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)
		
		totalquotarate = 0
		
		.Row		= Row    
		.Col		= 0
		
		for index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0 
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then
				.Col = C_Quota_Rate
				totalquotarate = totalquotarate + Unicdbl(.text)
		    end if
		next 
	End with
End Sub

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow()
	Dim i
	ShowDataFirstRow = 0
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow2
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow2()
	ShowDataFirstRow2 = 0
	Dim i
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow2 = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow()
	Dim i
	ShowDataLastRow = 0
	
	With frm1.vspdData
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow2
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow2()
	ShowDataLastRow2 = 0
	Dim i
	
	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow2 = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DataFirstRow
' Function Desc : 
'=======================================================================================================
Function DataFirstRow(ByVal Row)
	Dim i
	DataFirstRow = 0
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DataLastRow
' Function Desc : 
'=======================================================================================================
Function DataLastRow(ByVal Row)
	Dim i
	DataLastRow = 0
	
	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
'   Function Name : ShowFromData
'   Function Desc : 
'=======================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)	'###�׸��� ������ ���Ǻκ�###
'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.
	ShowFromData = 0
	
	Dim lngRow
	Dim lngStartRow
	
	With frm1.vspdData2
		
		Call SortSheet()
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		lngStartRow = 0
		If .MaxRows < 1 Then Exit Function
		
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = CInt(.Text) Then
				lngStartRow = lngRow
				ShowFromData = lngRow
				Exit For
			End If    
		Next
		'------------------------------------
		' Show Data
		'------------------------------------ 
		
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.Col = C_RecordCnt
			.Col2 = C_RecordCnt
			.DestCol = 0
			.DestRow = 1
			.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False
			
			.BlockMode = False
			
			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� ù��° ���� 2��° ������ Row�� �����.
			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 7��° ���� ������ ������ Row�� �����.
			If lngStartRow < .MaxRows Then
				If lngStartRow + lngShowingRows <= .MaxRows Then
					.BlockMode = True
					.Row = lngStartRow + lngShowingRows
					.Row2 = .MaxRows
					.RowHidden = True
					.BlockMode = False
				End If
			End If
			
			.BlockMode = False
			.Row = lngStartRow	'2003-03-01 Release �߰� 
			.Col = 0			'2003-03-01 Release �߰� 
			.Action = 0			'2003-03-01 Release �߰� 
		End If
	End With	
End Function

'======================================================================================================
' Function Name : SortSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortSheet()
	SortSheet = false
    With frm1.vspdData2
        .BlockMode = True
        .Col = 0
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .SortBy = 0 'SS_SORT_BY_ROW

        .SortKey(1) = C_ParentRowNo
        .SortKey(2) = C_RecordCnt
        
        .SortKeyOrder(1) = 0 'SS_SORT_ORDER_ASCENDING
        .SortKeyOrder(2) = 0 'SS_SORT_ORDER_ASCENDING

        .Col = 1	'C_SupplierCd	'###�׸��� ������ ���Ǻκ�###
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .Action = 25 'SS_ACTION_SORT
        
        .BlockMode = False
    End With     

    SortSheet = true
End Function

'=======================================================================================================
' Function Name : ChangeCheck
' Function Desc : 
'=======================================================================================================
Function ChangeCheck()
	Dim i
	ChangeCheck = False
	
	ggoSpread.Source = frm1.vspdData2
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then
				ChangeCheck = True
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc : 
'=======================================================================================================
Function CheckDataExist()
	Dim i
	CheckDataExist = False
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End IF
		Next
	End With
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear                                                          
    
    FncQuery = False                                                   
    
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables
    
    ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    												
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
	Set gActiveElement = document.activeElement
    FncQuery = True									
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    Err.Clear                                                           
    
    FncNew = False                                                      
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                              
    Call ggoOper.LockField(Document, "N")                               
    
    ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call SetDefaultVal
    Call InitVariables                                                  
	Set gActiveElement = document.activeElement
    FncNew = True                                                       
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Err.Clear         

    FncSave = False                                                         
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                      
    
    ggoSpread.Source = frm1.vspdData
    
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    If DefaultCheck = False Then
    	Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then	
		Exit Function
	End If			
	  
	Set gActiveElement = document.activeElement
    FncSave = True                                                       
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
	FncCancel = false
	Dim lRow
	Dim i,k,iCnt
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow
	Dim strFlag
	
	iActiveRow = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = iActiveRow
	frm1.vspdData.Col = frm1.vspdData.MaxCols
	iConvActiveRow = frm1.vspdData.Text
	
	If frm1.vspdData.MaxRows < 1 then
	    FncCancel = True
		Exit function
	End If
	
	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
	    FncCancel = True
    	Exit function
    End If
	
	If gActiveSpdSheet.ID = "B" Then

		ggoSpread.Source = frm1.vspdData2	
		With frm1.vspdData2
			
			'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START	    
		    lngRangeFrom = .SelBlockRow
		    .Row = lngRangeFrom

			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			
			.Redraw = False
			ggoSpread.EditUndo                                                 '��: Protect system from crashing
			.Redraw = True

			iCnt=0
			For k=lngRangeFrom To lngRangeTo
				.Row=k
				.col=0
				if .text = ggoSpread.UpdateFlag then
					iCnt = iCnt + 1
				End if	
			Next
			
			If iCnt = 0 Then
				ggoSpread.Source = frm1.vspdData
				ggoSpread.EditUndo iActiveRow                                                '��: Protect system from crashing
			End If	
		End With
	Else
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo                                                  '��: Protect system from crashing

		ggoSpread.Source = frm1.vspdData2	
		With frm1.vspdData2
			'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START	    
		    lngRangeFrom = .SelBlockRow
		    .Row = lngRangeFrom
			.Redraw = False
			
			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			
			iCnt=1
			For k=lngRangeFrom to lngRangeTo
				.Row=k
				ggoSpread.EditUndo k                                                 '��: Protect system from crashing
			Next
			.Redraw = True
		End WIth	
	End If
	
	lRow = frm1.vspdData.ActiveRow
	If lngRangeTo = 0 Then
		lglngHiddenRows(lRow - 1) = 0
	Else
		lglngHiddenRows(lRow - 1) = lngRangeTo - lngRangeFrom  + 1
	End If
	'**********///// END
	'********** START
	If lglngHiddenRows(lRow - 1) = 0 Then
		frm1.cmdInsertSampleRows.Disabled = False
	End If
	
	k = 0 
	For i = lngRangeFrom To lngRangeTo
	    frm1.vspdData2.Row = i 
	    frm1.vspdData2.Col = 0
	    strFlag = Trim(frm1.vspdData2.Text)
	    If strFlag = ggoSpread.UpdateFlag Then 
	        k = 1
	        Exit For
	    End If
	next 
	
	Call vspdData2_Click(frm1.vspdData2.ActiveCol,frm1.vspdData2.ActiveRow)

	Set gActiveElement = document.activeElement
	FncCancel = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()                        
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)		
	Set gActiveElement = document.activeElement
 	FncExcel = True
 End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement
    FncFind = True
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* %>
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    Err.Clear 

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value            '���� 
	    strVal = strVal & "&txtItemCd=" & .hdnItem.value              'ǰ�� 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
    Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
    End If

	Call RunMyBizASP(MyBizASP, strVal)	
    
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)							
	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows
	Dim ii
	
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolbar("1110100100011111")				'��ư ���� ���� 

	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then 
				ReDim lgStrPrevKeyM(intARow)	
				ReDim lglngHiddenRows(intARow)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lgStrPrevKeyM(intTRow+intARow)	
				ReDim lglngHiddenRows(intTRow+intARow)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
				For i = 0 To intTRow
					lgStrPrevKeyM(i) = TmpArrPrevKey(i)
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next 
			End If

			For i = intTRow To intTRow+intARow
				lglngHiddenRows(i) = 0
			Next 

			if lgIntFlgModeM = Parent.OPMD_CMODE then
			    If DbQuery2(1, false) = False Then	Exit Function
		    end if
	
		    lgIntFlgModeM = Parent.OPMD_UMODE
		End If
	End With
	
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
	DbQueryOk = true
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk2(Byval DataCount)
	DbQueryOk2 = false
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim Index
	
	With frm1.vspdData2
		
		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_RecordCnt
		
		.Col2 = C_RecordCnt
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With
	
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement
	
	DbQueryOk2 = true
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim txtPlantCd, txtItemCd
	
	DbQuery2 = False

	'/* 9�� ������ġ: ���� ���������� �ణ �̵� �� �̹� ��ȸ�� �ڷᳪ �Էµ� �ڷḦ �о� ���� ������ '' â ���� - START */
	Call LayerShowHide(1)
	
	With frm1
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		Row = CInt(.vspdData.Text)	
		If lglngHiddenRows(Row - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			 lngRet = ShowFromData(Row, lglngHiddenRows(Row - 1))	'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.
			
			Call SetToolbar("1110100100011111")				'��ư ���� ���� 
			Call LayerShowHide(0)
			
			lngRangeFrom = ShowDataFirstRow
			lngRangeTo = ShowDataLastRow		
						
			.vspdData2.ReDraw = True
			DbQuery2 = True
			Exit Function
		End If

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		.vspdData.Row = CInt(Row)
		.vspdData.Col = C_PlantCd		    
		strVal = strVal & "&txtPlantCd=" & Trim(.vspdData.text)
		.vspdData.Col = C_ItemCd		    
		strVal = strVal & "&txtItemCd=" & Trim(.vspdData.text)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag 
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(Row)
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery2 = True
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt
	Dim strVal 
	Dim lngRangeFrom
    Dim lngRangeTo
    Dim parentRow
    Dim totalRate
	Dim Zsep
	Dim iColSep
	Dim iRowSep

	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size
	Dim ii
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
	
	DbSave = False                                                          '��: Processing is NG
    
	Call LayerShowHide(1)

	frm1.txtMode.value = Parent.UID_M0002

	lGrpCnt = 1
	strVal = ""
    Zsep = "@"
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0

	'-----------------------
	'Data manipulate area
	'-----------------------
	With frm1
	    For parentRow = 1 To .vspdData.MaxRows
			If Trim(GetSpreadText(.vspdData,0,parentRow,"X","X")) = ggoSpread.UpdateFlag Then
				
			    lngRangeFrom = DataFirstRow(parentRow)
			    lngRangeTo   = DataLastRow(parentRow)
			   
			    totalRate = 0
			    for lRow = lngRangeFrom To lngRangeTo
					totalRate = totalRate + UNICDbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))
				Next
				
			    If UniCdbl(totalRate) <> 100  then       '���� ���� ���� ǰ�� ��к����� 100���� 
				    Call DisplayMsgBox("171325", "X", Trim(GetSpreadText(.vspdData,C_ItemCd,parentRow,"X","X")) & "(" & parentRow & "Row)" , "X")
				    Call LayerShowHide(0)
				    Call RemovedivTextArea
				    Exit Function
				End if 
					
			    for lRow = lngRangeFrom To lngRangeTo
			        If Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then

						strVal = strVal & "U" & iColSep		
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentPlantCd,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentItemCd,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep
			
						If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
							strVal = strVal & "0" & iColSep
						Else
							strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
						End If
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_ParentRowNo,lRow,"X","X")) & iColSep
						strVal = strVal & Trim(GetSpreadText(.vspdData2,C_RecordCnt,lRow,"X","X")) & iColSep & iRowSep
							
						lGrpCnt = lGrpCnt + 1
				
					End If			
   			    Next
				
				strVal = strVal & Zsep
				Select Case Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
				    Case ggoSpread.UpdateFlag
				         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
					                            
				            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
				            objTEXTAREA.name = "txtCUSpread"
				            objTEXTAREA.value = Join(iTmpCUBuffer,"")
				            divTextArea.appendChild(objTEXTAREA)     
					 
				            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
				            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
				            iTmpCUBufferCount = -1
				            strCUTotalvalLen  = 0
				         End If
					       
				         iTmpCUBufferCount = iTmpCUBufferCount + 1
					      
				         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
				            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
				            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				         End If   
				         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
				         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				End Select   
			End If
			strVal  = ""
		Next     
	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'��: �����Ͻ� ASP �� ���� 

	DbSave = True	
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()								
	Call InitVariables
	
    lgIntFlgMode	 = Parent.OPMD_UMODE		
	lgBlnFlgChgValue = False
	Call MainQuery		    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--########################################################################################################
'       					6. Tag�� 
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ó����к�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU" ALT="�� ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="�� ��"></TD>
							    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
												   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			
			<TR>
			 <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
    
			<TR HEIGHT= 40%>
			 <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			  <TABLE <%=LR_SPACE_TYPE_60%>>
			   <TR>
			    <TD HEIGHT=100% WIDTH=100% COLSPAN=4>
			     <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
			    </TD>
			   </TR>
			  </TABLE>
			 </TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
      <td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m1211mb2.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtColsep" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRowsep" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
