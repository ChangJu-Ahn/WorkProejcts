<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1302ma1_ko119.asp
'*  4. Program Name         : �ð��뺰 �۾�����Ȯ��(S) 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2006/04/19
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Const BIZ_PGM_QRY_ID  = "p1302mb1_ko119.asp"												'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_SAVE_ID = "p1302mb1_ko119.asp"
Const BIZ_PGM_ID = "p4119mb1_ko119.asp"	                                 'Biz Logic ASP
Const BIZ_PGM_ID2 = "p4119mb2_ko119.asp"	                                 'Biz Logic ASP 					
'Const BIZ_PGM_JUMPSHIFTEXECEPTION_ID  = "p1504ma1"						'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_CheckBox					'Ȯ������
Dim C_ItemCd					'��ǰ�ڵ�
Dim C_ItemNm					'��ǰ��
Dim C_Spec						'Spec
Dim C_JobPlanDt                 '�۾���ȹ�ð�
Dim C_JobLineCd					
Dim C_JobLine                   'Job Line
Dim C_JobPlanTime				'�۾���ȹ�ð�
Dim C_JobQty					'�۾�����
Dim C_JobSeq					'�켱����
Dim C_JobOrderNo				'�۾����ù�ȣ
Dim C_SecItemCd                 '�Ｚ�ڵ�				
Dim C_ProdtOrderNo			    '����������ȣ
Dim C_RoutNo				    '�����

Dim lgShiftCnt
Dim lgRadio
Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
'Dim EndDate

    iDBSYSDate = "<%=GetSvrDate%>"			'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'	StartDate = UNIDateAdd("D",7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'	EndDate = UNIDateAdd("D", 7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop          
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables() 
		C_CheckBox			= 1		'Ȯ������
		C_ItemCd			= 2		'��ǰ�ڵ�
		C_ItemNm			= 3		'��ǰ��
		C_Spec				= 4		'Spec
		C_JobPlanDt         = 5     '�۾���ȹ�ð�
		C_JobLineCd			= 6		
		C_JobLine           = 7     'Job Line
		C_JobPlanTime		= 8		'�۾���ȹ�ð�
		C_JobQty			= 9		'�۾�����
		C_JobSeq			= 10	'�켱����
		C_JobOrderNo		= 11	'�۾����ù�ȣ
		C_SecItemCd         = 12    '�Ｚ�ڵ�				
		C_ProdtOrderNo		= 13	'����������ȣ
		C_RoutNo			= 14	'�����
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
'    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey = ""
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
	if frm1.NoConfirm.checked = True then
		lgRadio = "N"
	elseif frm1.Confirm.checked = True then
		lgRadio = "Y"
	elseif frm1.All.checked = True then
		lgRadio = "A"
	else
	end if 
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	   frm1.txtProdFromDt.text = StartDate
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()     
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub


'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
		
			.Col = C_JobLineCd
			intIndex = .Value
			.Col = C_JobLine
			.Value = intIndex
			
		Next	
	End With
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	Call AppendNumberPlace("6", "18", "0") 
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    
		
		.MaxCols = C_RoutNo+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		.ReDraw = false
		
		Call ggoSpread.ClearSpreadData()    '��: Clear spreadsheet data 
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck    C_CheckBox ,"Ȯ������",8 , ,"",true 
		ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 15
		ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 25
		ggoSpread.SSSetEdit		C_Spec, "�԰�", 25				'5
		ggoSpread.SSSetDate 	C_JobPlanDt, "�۾���ȹ����", 11, 2, parent.gDateFormat	'15
		ggoSpread.SSSetCombo	C_JOBLINECD, "LineCd", 10
		ggoSpread.SSSetCombo	C_JOBLINE, "Line", 8
		ggoSpread.SSSetEdit		C_JobPlanTime, "�ð���" , 6,2,,5
		ggoSpread.SSSetFloat	C_JobQty, "�۾�����" ,8,"6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit	    C_JobSeq, "�켱����" ,8, 2, -1, 5
		ggoSpread.SSSetEdit		C_JobOrderNo, "�۾����ù�ȣ" , 13,,,18	
		ggoSpread.SSSetEdit		C_SecItemCd, "�Ｚ�ڵ�", 10			'20
		ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 10,,,18
		ggoSpread.SSSetEdit		C_RoutNo, "�����", 10 

'		Call ggoSpread.MakePairsColumn(C_JOBLINECD,C_JOBLINE)

		Call ggoSpread.SSSetColHidden(C_JOBLINECD ,C_JOBLINECD	,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

		ggoSpread.SSSetSplit2(3)										'frozen ����߰� 
		Call SetSpreadLock
		Call InitData
'		Call initComboBox_two 
		.ReDraw = true
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
	
		.vspdData.ReDraw = False
'		ggoSpread.SpreadLock	 -1, -1 
		ggoSpread.SSSetProtected C_ItemCd	, -1, C_ItemCd
		ggoSpread.SSSetProtected C_ItemNm		, -1, C_ItemNm
		ggoSpread.SSSetProtected C_Spec	, -1, C_Spec
		ggoSpread.SSSetProtected C_JobPlanDt	, -1, C_JobPlanDt
		ggoSpread.SSSetProtected C_JOBLINECD	, -1, C_JOBLINECD
		ggoSpread.SSSetProtected C_JOBLINE	, -1, C_JOBLINE
		ggoSpread.SSSetProtected C_JobPlanTime	, -1, C_JobPlanTime
		ggoSpread.SSSetProtected C_JobQty	, -1, C_JobQty
		ggoSpread.SSSetProtected C_JobSeq	, -1, C_JobSeq
		ggoSpread.SSSetProtected C_JobOrderNo	, -1, C_JobOrderNo
		ggoSpread.SSSetProtected C_SecItemCd	, -1, C_SecItemCd
		ggoSpread.SSSetProtected C_ProdtOrderNo	, -1, C_ProdtOrderNo
		ggoSpread.SSSetProtected C_RoutNo	, -1, C_RoutNo
		
		.vspdData.ReDraw = True
	End With	
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
'		ggoSpread.SSSetRequired 	C_ShiftCd,			pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected	C_ShiftNm,			pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_Line_Group		,pvStartRow	,pvEndRow
        ggoSpread.SSSetRequired		C_Work_Line			,pvStartRow	,pvEndRow
'		ggoSpread.SSSetRequired		C_Work_Line_Desc	,pvStartRow	,pvEndRow
		
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
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
			C_CheckBox			= iCurColumnPos(1)
			C_ItemCd			= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_JobPlanDt         = iCurColumnPos(5)
			C_JobLineCd			= iCurColumnPos(6)
			C_JobLine           = iCurColumnPos(7)
			C_JobPlanTime		= iCurColumnPos(8)
			C_JobQty			= iCurColumnPos(9)
			C_JobSeq			= iCurColumnPos(10)
			C_JobOrderNo		= iCurColumnPos(11)
			C_SecItemCd         = iCurColumnPos(12)
			C_ProdtOrderNo		= iCurColumnPos(13)
			C_RoutNo			= iCurColumnPos(14)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strWhere
	
	if frm1.txtPlantCd.value = "" then
	strWhere    = ""
	strWhere	= strWhere &  " plant_cd >= '" & frm1.txtPlantCd.value & "'"
	else
	strWhere    = ""
	strWhere	= strWhere &  " plant_cd = '" & frm1.txtPlantCd.value & "'"
	end if
	strWhere	= strWhere & " order by line_group, work_line "

	Call CommonQueryRs(" WORK_LINE, WORK_LINE_DESC ", " p_work_line_ko119 ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboLine, lgF0, lgF1, Chr(11)) 

End Sub

Sub InitComboBox3()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strWhere

	strWhere    = ""
	strWhere	= strWhere & " a.major_cd = b.major_cd and a.major_cd = 'M2110' and b.seq_no = 99 and a.minor_cd = b.minor_cd "
	strWhere	= strWhere & " order by b.reference "

	Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM ", " b_minor a (NOLOCK), b_configuration b (nolock)  ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboTime, lgF0, lgF1, Chr(11)) 

End Sub

Sub InitComboBox2()
	Dim i
	Dim strWhere
	Dim strSelect
	Dim strFrom
	Dim arrVal1, arrVal2
	Dim ii, jj
	Dim strPlantCd
	
	strSelect	=			 " plant_cd "
	strFrom		=			 " b_plant (NOLOCK) "
	strWhere	=			 " plant_cd > '' "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1) 
       For ii = 0 To jj - 1 
			arrVal2			= Split(arrVal1(ii), chr(11))
			strPlantCd		= Ucase(Trim(arrVal2(1)))
'			strMinorNm		= Trim(arrVal2(2))

	
		strWhere    = ""
		strWhere	=	strWhere &  " plant_cd = '" & strPlantCd & "'"
   	   	If  CommonQueryRs(" count(WORK_LINE) "," p_work_line_ko119 (NOLOCK) ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
	      	If UNICdbl(Replace(lgF0,Chr(11),"")) = 0 Then
	      	   if lgShiftCnt > 0 then
	      	   else
					lgShiftCnt = 0
			   end if	
			Else
			   if lgShiftCnt = 0 then
			     if int(Replace(lgF0,Chr(11),"")) > int(lgShiftCnt) then
			     	lgShiftCnt = Replace(lgF0,Chr(11),"")
			     end if	
			   end if  
			End if	
		End if
	  Next    
   end if	       

'	If Trim(frm1.txtPlantCd.value) = "" Then
'		frm1.txtPlantNm.value = ""
'	Else
		For i = lgShiftCnt + 1 To 1 Step -1
			frm1.cboLine.remove(i) 
		Next
'	End If
End Sub

'========================== 2.2.6 InitSpreadComboBox()  ========================================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strWhere
	
	if frm1.txtPlantCd.value = "" then
	strWhere	=  ""
	strWhere	= strWhere & " plant_cd >= '" & frm1.txtPlantCd.value & "'"
	else
	strWhere	= ""
	strWhere	= strWhere &  " plant_cd = '" & frm1.txtPlantCd.value & "'"
	end if
	strWhere	= strWhere & " order by line_group, work_line "

	'****************************
	'List Minor code
	'****************************
	Call CommonQueryRs(" WORK_LINE, WORK_LINE_DESC ", " p_work_line_ko119 ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobLineCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobLine
End Sub



'Sub InitComboBox_two()
'	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("PX901", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DiFlag			'COLM_DATA_TYPE
'    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DiFlagNm
'End Sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
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
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
		Call initCombobox2()
		Call initCombobox()
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)

    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'-------------------------------------  JumpShiftException()  -----------------------------------------
'	Name : JumpShiftException()
'	Description : Shift ���ܵ������ Jump�Ѵ�.
'--------------------------------------------------------------------------------------------------------- 

Function JumpShiftException()
    Dim IntRetCd, strVal

    '-----------------------
    'Precheck area
    '-----------------------
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900017",parent.VB_YES_NO,"X","X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
		
    If frm1.vspdData.ActiveRow <= 0 Then 
		Call DisplayMsgBox("181216",parent.VB_YES_NO,"X","X")
		Exit Function
	End If
	
	'-----------------------------
	' Write Cookie
	'-----------------------------	
	WriteCookie "txtPlantCd", FilterVar(UCase(Trim(frm1.txtPlantCd.value)),,"SNM")
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
'	WriteCookie "txtResourceCd", FilterVar(UCase(Trim(frm1.txtResourceCd.value)),,"SNM")
'	WriteCookie "txtResourceNm", frm1.txtResourceNm.value 
	
'	frm1.vspdData.Row = frm1.vspdData.ActiveRow
'	frm1.vspdData.Col = C_ShiftCd
	
'	WriteCookie "txtShiftCd", UCase(Trim(frm1.vspdData.Text))
	
'	frm1.vspdData.Row = frm1.vspdData.ActiveRow
'	frm1.vspdData.Col = C_ShiftNm
	
'	WriteCookie "txtShiftNm", UCase(Trim(frm1.vspdData.Text))
	
'	PgmJump(BIZ_PGM_JUMPSHIFTEXECEPTION_ID)	
	
End Function

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetCookieVal()
   	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value	  = ReadCookie("txtPlantCd")
	End If
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value	  = ReadCookie("txtPlantNm")
	End If
	If ReadCookie("txtProdOrderNo") <> "" Then
		frm1.txtProdOrderNo.Value	  = ReadCookie("txtProdOrderNo")
	End If
	If ReadCookie("txtItemcd") <> "" Then
		frm1.txtItemCd.Value	  = ReadCookie("txtItemcd")
	End If
	If ReadCookie("txtProdFromDt") <> "" Then
		frm1.txtProdFromDt.Text	  = ReadCookie("txtProdFromDt")
	End If

	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtProdOrderNo", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtProdFromDt", ""

End Sub

'-------------------------------------  CntMaxRows()  -----------------------------------------
'	Name : CntMaxRows()
'	Description : �ڿ��� ��ϵ� �� �ִ� Shift������ �����Ѵ�. iPos : �߰��� Row�� 
'------------------------------------------------------------------------------------------------- 

Function CntMaxRows(iPos)
	Dim TotRowCnt
	Dim iRows
	
	On Error Resume Next
	Err.Clear
	
	CntMaxRows = False
		
	TotRowCnt = frm1.vspdData.MaxRows
	
	ggoSpread.Source = frm1.vspdData
	
	'--------------------------------------------------------------------------
	' ������ ���� ������ �� Row�� ������ ����Ѵ�.
	'--------------------------------------------------------------------------
	For iRows = 1 To TotRowCnt
		frm1.vspdData.Col = 0
		frm1.vspdData.Row = iRows
		
		If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			TotRowCnt = TotRowCnt - 1
		End If
	Next
	
	'--------------------------------------------------------------------------
	' FncInsertRow�� FncCopy�̸� ���� �߰��ϱ� ���� ���� �ϳ� ���ؼ� ����Ѵ�.
	'--------------------------------------------------------------------------
	TotRowCnt = TotRowCnt + iPos	
	
	If TotRowCnt > 4 Then
		Call DisplayMsgBox("181814","X","X","X")
		Exit Function
	End If

	CntMaxRows = True
	
End Function
'########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal

    '----------  Coding part  -------------------------------------------------------------
   
    Call SetToolbar("11000000000011")								'��: ��ư ���� ���� 
        
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus  
		Set gActiveElement = document.activeElement 
	End IF

    Call SetCookieVal()

    Call InitComboBox    
    Call InitComboBox3
    Call InitSpreadComboBox
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub



'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'------------------------------------------  txtPlantCd_OnChange -----------------------------------------
'	Name : txtPlantCd_OnChange()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlantCd_OnChange()
	Call initCombobox2()
	Call InitComboBox
End Sub

'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtProdFromDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

Sub cboLine_onclick()

	if frm1.txtPlantCd.value = "" then
	 Call DisplayMsgBox("169901","X", "����","X")
	 exit sub
	end if 
    
End Sub

Function Radio1_onChange()
	
	IF lgRadio = "N" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData	
	
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "N"
	
	lgBlnFlgChgValue = True
End Function

Function Radio2_onChange()

	IF lgRadio = "Y" Then
		Exit Function
	ENd IF

	ggoSpread.Source = frm1.vspdData
	
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, False)			
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, False)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()


	lgRadio = "Y"
	
	lgBlnFlgChgValue = True
End Function

Function Radio3_onChange()
	
	IF lgRadio = "A" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData	
	
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItem, C_MCSDTLItem, True)			
'	Call ggoSpread.SSSetColHidden( C_MCSDTLItemNm, C_MCSDTLItemNm, True)			
	
	ggoSpread.ClearSpreadData		
	call initVariables()
	
	lgRadio = "A"
	
	lgBlnFlgChgValue = True
End Function


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

   	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0

	select case Col
	Case   C_CheckBox

			If  frm1.vspdData.Text = ggoSpread.UpdateFlag Then
		
			    If frm1.vspdData.value = "0" Then

				    frm1.vspdData.Row = Row
				    frm1.vspdData.Col = 0	
				    frm1.vspdData.text = "" 
			    else
				    frm1.vspdData.Row = Row
				    frm1.vspdData.Col = 0	
				    frm1.vspdData.text = "" 
			    End If 
			else 
					frm1.vspdData.Row = Row
				    frm1.vspdData.Col = 0	
				    frm1.vspdData.text = ggoSpread.UpdateFlag     
			end if    
'			elseif frm1.vspdData2.Text = ggoSpread.DeleteFlag Then
'				frm1.vspdData2.Col = C_CheckBox
'			        			
'			    If frm1.vspdData2.value = "1" Then
'	
'				    frm1.vspdData2.Row = Row
'				    frm1.vspdData2.Col = 0	
'			        frm1.vspdData2.text = frm1.vspdData2.Row
'			    End If
'			elseif frm1.vspdData2.Text = ggoSpread.InsertFlag Then
'				frm1.vspdData2.Col = C_CheckBox
'			        			
'			    If frm1.vspdData2.value = "0" Then
'				    frm1.vspdData2.Row = Row
'				    frm1.vspdData2.Col = 0	
'				    frm1.vspdData2.text = frm1.vspdData2.Row
'			    End If  

'	    	End If  
	
	end select


'	CopyToHSheet Row
	
End Sub	


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspddata_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("1101111111")	
	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
	   Exit Sub
	End If
	   	    
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspddata_DblClick(ByVal Col , ByVal Row )       
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'==========================================================================================
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
		
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii)
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------   

    End With

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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
		    
			Case  C_LINE_GROUP
				.Col = Col
				intIndex = .Value
				.Col = C_LINE_GROUPCD
				.Value = intIndex
					
				
			Case  C_LINE_GROUPCD
				.Col = Col
				intIndex = .Value
				.Col = C_LINE_GROUP
				.Value = intIndex
			
		End Select
	End With
End Sub


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False															'��: Processing is NG

    Err.Clear																    '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData 
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     														'��: Query db data

    FncQuery = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData 
    
    If ggoSpread.SSCheckChange = False Then 
       IntRetCD = DisplayMsgBox("900001","X","X","X")                            '��: No data changed!!
       Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData 
	If Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then
			Exit Function
		End If
    End IF
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     													'��: Save db data
    
    FncSave = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	'Row �ϳ��� �߰��Ҷ� 4���� �ʰ��ϴ��� üũ 
'	If Not CntMaxRows(1) Then Exit Function
	
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    frm1.vspdData.EditMode = True
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	
	  On Error Resume Next
    
    Dim iDx
	
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
 
    
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo 
    call InitData 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
   Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    With frm1.vspdData 
    
    .focus
    Set gActiveElement = document.activeElement 
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData 
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False

    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal

    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdFromDt=" & UNIConvDate(Trim(.hProdFromDt.Value))
		strVal = strVal & "&cboLine=" & Trim(.hcboLine.Value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.htxtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.htxtItemCd.Value)
		strVal = strVal & "&cboTime=" & Trim(.hcboTime.Value)
		strVal = strVal & "&txtRadio="			& lgRadio
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdFromDt=" & UNIConvDate(Trim(.txtProdFromDt.Text))
		strVal = strVal & "&cboLine=" & Trim(.cboLine.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
		strVal = strVal & "&cboTime=" & Trim(.cboTime.Value)
		strVal = strVal & "&txtRadio="			& lgRadio
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    Call InitSpreadComboBox
    Call InitData()
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
'	Call SetToolbar("11001111000111")
	Call SetToolbar("11001000000111")
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    
    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal
	Dim strDel
	
    DbSave = False                                                          '��: Processing is NG
    
'    If Not CntMaxRows(0) Then Exit Function
    
    LayerShowHide(1) 
		
    'On Error Resume Next                                                   '��: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    '-----------------------
    'Data manipulate area
    '-----------------------
     
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'��: �ű� 
				
'				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep					'��: C=Create
				
				
'               .vspdData.Col = C_Line_GroupCd	'2
'                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
'               .vspdData.Col = C_Work_Line	'3
'               strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
'               .vspdData.Col = C_Work_Line_Desc  '4	
'               strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
'                .vspdData.Col = C_Remark	'5
'                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                                
'                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag

				strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep					'��: U=Update
				
                .vspdData.Col = C_CheckBox	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ItemCD	         '3
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_JobPlanDt	'4
                strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                
                .vspdData.Col = C_JobLineCd
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_JobPlanTime
                strVal = strVal & Replace(Trim(.vspdData.Text),":","") & parent.gColSep
                
                .vspdData.Col = C_JobOrderNo
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_JobSeq
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_JobQty
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
                
                .vspdData.Col = C_RoutNo
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
				.vspdData.Col = C_ProdtOrderNo
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                                               
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag												'��: ���� 

'				strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep
				
'                .vspdData.Col = C_Line_GroupCd	'2
'                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                
'                .vspdData.Col = C_Work_Line	'3
'                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                                
'                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True																	'��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()																	'��: ���� ������ ���� ���� 
   
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk2()																	'��: ���� ������ ���� ���� 
   
    frm1.confirm.checked = true
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function

Function DbDeleteOk()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
Function ExeReflect()
	Dim IntRetCD,iRow, RetFlag  
    Dim strVal, strYear, strPYear
    Dim lGrpCnt
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim Flag
	Dim StartDate2
	Dim strWhere
	
    Err.Clear
    ExeReflect = False	
    
    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
     
  ' �ѹ� Ȯ���Ѱ��� ��Ȯ���Ǽ��� �ȵ�(��Ȯ�� ������翩�� üũ) 
    strWhere = " plant_cd = " & FilterVar(Trim(frm1.txtPlantCd.value),"''","S")
	strWhere = strWhere & " and CONFIRM_FLG = 'N'"
	strWhere = strWhere & " and JOB_PLAN_DT = " & FilterVar(UniConvDate(frm1.txtProdFromDt.text), "''", "S")
	
	if frm1.txtProdOrderNo.value <> "" then
		strWhere = strWhere & " AND PRODT_ORDER_NO  = " & FilterVar(Trim(frm1.txtProdOrderNo.value),"''","S")
    end if
    
    if frm1.txtItemCd.value <> "" then
		strWhere = strWhere & " AND ITEM_CD  = " & FilterVar(Trim(frm1.txtItemCd.value),"''","S")
    end if	

	IF frm1.cboLine.value <> "" then
		strWhere = strWhere & " AND JOB_LINE  = " & FilterVar(Trim(frm1.cboLine.value),"''","S")
	end if
	
	IF frm1.cboTime.value <> "" then
		strWhere = strWhere & " AND JOB_PLAN_TIME  = " & FilterVar(Trim(frm1.cboTime.value),"''","S")
	end if

	If  CommonQueryRs("top 1 ITEM_CD"," P_PROD_TIME_ORDER_KO119 (NOLOCK) ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
		
	else
		CAll DisplayMsgBox("XX1051", "X","X","X")
		exit function
	end if 
    
    
    IntRetCD = DisplayMsgBox("XX1013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If	
    	
       
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	strVal = ""
	lGrpCnt = 0
	
	With Frm1
		strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0002							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdFromDt=" & UNIConvDate(Trim(.txtProdFromDt.Text))
		strVal = strVal & "&cboLine=" & Trim(.cboLine.Value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
		strVal = strVal & "&cboTime=" & Trim(.cboTime.Value)

'	Call ExecMyBizASP(strVal, BIZ_PGM_ID2)										'��: �����Ͻ� ASP �� ���� 
	Call RunMyBizASP(MyBizASP, strVal)
		
	End With


    ExeReflect = True 
    LayerShowHide(0)        

End Function





</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ð��뺰�۾�����Ȯ��(S)</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4119ma1_ko119_I847991329_txtProdFromDt.js'></script>
									</TD>																						
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLine" ALT="����" STYLE="Width: 80px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>�ð���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboTime" ALT="�ð���" STYLE="Width: 100px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=NoConfirm Checked tag = 2 value="N" onclick=radio1_onchange()><LABEL FOR=NoConfirm>��Ȯ��</LABEL>&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Confirm tag = 2 value="Y" onclick=radio2_onchange()><LABEL FOR=Confirm>Ȯ��</LABEL>&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=All tag = 2 value="A" onclick=radio3_onchange()><LABEL FOR=All>��ü</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
								<TD HEIGHT=100%>
									<script language =javascript src='./js/p4119ma1_ko119_I510278844_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>Ȯ��</BUTTON>
 		</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%>  FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboLine" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboTime" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

