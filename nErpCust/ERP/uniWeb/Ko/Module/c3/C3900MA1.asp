
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3900ma1
'*  4. Program Name         : �򰡱ݾ׹ݿ� 
'*  5. Program Desc         : ����ǰ������ ��ȸ, ���ݾ��򰡹ݿ� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit	
																'��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================

'@PGM_ID
Const BIZ_PGM_ID = "c3900mb1.asp"												'�����Ͻ� ���� ASP�� 
Const BIZ_PGM_EXE_ID = "c3900mb2.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_CNC_ID = "c3900mb3.asp"												'��: �����Ͻ� ���� ASP�� 
 
     
Dim C_PlantCd		 															'Spread Sheet�� Column�� ��� 
Dim C_TrnsTypeCd
Dim C_TrnsTypeNm
Dim C_MovTypeCd
Dim C_MovTypeNm
Dim C_CostCd
Dim C_ItemAcctNm
Dim C_ItemCd		 
Dim C_ItemNm		 
Dim C_DocNo		 
Dim C_SeqNo		 
Dim C_DiffAmt		 
Dim C_TrnsPlantCd
Dim C_TrnsSlCd
Dim C_TrnsItemCd
Dim C_TrnsItemNm
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6



'@Global_Var
Dim lgBlnFlgChgValue           'Variable is for Dirty flag
Dim lgIntGrpCount              'Group View Size�� ������ ���� 
Dim lgIntFlgMode               'Variable is for Operation Status
Dim lgIsOpenPop          

Dim lgStrPlantPrevKey
Dim lgStrTrnsPrevKey
Dim lgStrMovPrevKey
Dim lgStrCostPrevKey
Dim lgStrItemPrevKey
Dim lgStrTrnsPlantPrevKey
Dim lgStrTrnsSlPrevKey
Dim lgStrTrnsItemPrevKey

Dim lgLngCurRows
Dim lgSortKey

'======================================================================================================
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_PlantCd		=1
	C_TrnsTypeCd	=2		
	C_TrnsTypeNm	=3
	C_MovTypeCd		=4
	C_MovTypeNm		=5
	C_CostCd		=6
	C_ItemAcctNm	=7
	C_ItemCd		=8 
	C_ItemNm		=9 
	C_DocNo			=10
	C_SeqNo			=11
	C_DiffAmt		=12 
	C_TrnsPlantCd	=13
	C_TrnsSlCd		=14
	C_TrnsItemCd	=15
	C_TrnsItemNm	=16
End Sub


'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
	lgStrPlantPrevKey	= ""
	lgStrTrnsPrevKey	= ""
	lgStrMovPrevKey		= ""
	lgStrCostPrevKey	= ""
	lgStrItemPrevKey	= ""
	lgStrTrnsPlantPrevKey = ""
	lgStrTrnsSlPrevKey	= ""
	lgStrTrnsItemPrevKey=""    
    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
	
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "BA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
    .MaxCols = C_TrnsItemNm+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
    .Col = .MaxCols															'������Ʈ�� ��� Hidden Column
    .ColHidden = True
    
   
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread 
	
	ggoSpread.ClearSpreadData   

	.ReDraw = false

	Call GetSpreadColumnPos("A")

	Call AppendNumberPlace("6","3","0")
	
    ggoSpread.SSSetEdit C_PlantCd, "�����ڵ�", 10
    ggoSpread.SSSetEdit C_TrnsTypeCd, "���ұ���", 10
    ggoSpread.SSSetEdit C_TrnsTypeNm, "���ұ��и�", 15
    ggoSpread.SSSetEdit C_MovTypeCd, "��������", 15
    ggoSpread.SSSetEdit C_MovTypeNm, "����������", 15        		
    ggoSpread.SSSetEdit C_CostCd, "C/C", 15        		
	ggoSpread.SSSetEdit C_ItemAcctNm, "ǰ�����", 10
	ggoSpread.SSSetEdit C_ItemCd, "ǰ���ڵ�", 18
    ggoSpread.SSSetEdit C_ItemNm, "ǰ���", 27
	ggoSpread.SSSetEdit C_DOCNO, "���ҹ�ȣ", 15

	ggoSpread.SSSetFloat C_SeqNo, "����SEQ", 8, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat C_DiffAmt, "�����ݾ�", 20, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetEdit C_TrnsPlantCd, "�̵�����", 15
	ggoSpread.SSSetEdit C_TrnsSlCd, "�̵�â��", 15
	ggoSpread.SSSetEdit C_TrnsItemCd, "�̵�ǰ��", 18
	ggoSpread.SSSetEdit C_TrnsItemNm, "�̵�ǰ���", 27
		
	'Call ggoSpread.SSSetColHidden(C_TrnsPlantCd ,C_TrnsPlantCd	,True)
	'Call ggoSpread.SSSetColHidden(C_TrnsSlCd	,C_TrnsSlCd		,True)
	'Call ggoSpread.SSSetColHidden(C_TrnsItemCd	,C_TrnsItemCd	,True)
	'Call ggoSpread.SSSetColHidden(C_TrnsItemNm	,C_TrnsItemNm	,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1

    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_PlantCd, -1, C_PlantCd
    ggoSpread.SpreadLock C_TrnsTypeCd, -1, C_TrnsTypeCd
    ggoSpread.SpreadLock C_TrnsTypeNm, -1, C_TrnsTypeNm
    ggoSpread.SpreadLock C_MovTypeCd, -1, C_MovTypeCd
    ggoSpread.SpreadLock C_MovTypeNm, -1, C_MovTypeNm
    ggoSpread.SpreadLock C_CostCd, -1, C_CostCd
    ggoSpread.SpreadLock C_ItemAcctNm, -1, C_ItemAcctNm
    ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
    ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
    ggoSpread.SpreadLock C_DocNo, -1, C_DocNo
    ggoSpread.SpreadLock C_SeqNo, -1, C_SeqNo
    ggoSpread.SpreadLock C_DiffAmt, -1, C_DiffAmt
    ggoSpread.SpreadLock C_TrnsPlantCd, -1, C_TrnsPlantCd
    ggoSpread.SpreadLock C_TrnsSlCd, -1, C_TrnsSlCd
    ggoSpread.SpreadLock C_TrnsItemCd, -1, C_TrnsItemCd
    ggoSpread.SpreadLock C_TrnsItemNm, -1, C_TrnsItemNm
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
	select case iWhere
		case 1
			arrParam(0) = "���� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_PLANT"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""		' Where Condition
			arrParam(5) = "����"			
	
			arrField(0) = "PLANT_CD"					' Field��(0)
			arrField(1) = "PLANT_NM"					' Field��(1)
    
			arrHeader(0) = "����"				' Header��(0)
			arrHeader(1) = "�����"				' Header��(1)
		case 2
			arrParam(0) = "�ڽ�Ʈ���� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_COST_CENTER"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtCostCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""		' Where Condition
			arrParam(5) = "�ڽ�Ʈ��Ÿ"			
	
			arrField(0) = "COST_CD"					' Field��(0)
			arrField(1) = "COST_NM"					' Field��(1)
    
			arrHeader(0) = "�ڽ�Ʈ����"				' Header��(0)
			arrHeader(1) = "�ڽ�Ʈ���͸�"				' Header��(1)
		case 3
			arrParam(0) = "���ұ��� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_MINOR a"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtTrnsTypeCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "a.major_cd = " & FilterVar("I0002", "''", "S") & " "		' Where Condition
			arrParam(5) = "���ұ���"			
	
			arrField(0) = "a.MINOR_CD"					' Field��(0)
			arrField(1) = "a.MINOR_NM"					' Field��(1)
    
			arrHeader(0) = "���ұ���"				' Header��(0)
			arrHeader(1) = "���ұ��и�"				' Header��(1)

		case 4
			arrParam(0) = "�������� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_MINOR a,i_movetype_configuration b"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtMovTypeCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			IF Trim(frm1.txtTrnsTypeCd.value) <> "" Then
				arrParam(4) = "a.major_cd = " & FilterVar("I0001", "''", "S") & "  and a.minor_cd = b.mov_type and b.trns_type = " & FilterVar(frm1.txtTrnsTypeCd.value, "''", "S")		' Where Condition
			Else
				arrParam(4) = "a.major_cd = " & FilterVar("I0001", "''", "S") & "  and a.minor_cd = b.mov_type "		' Where Condition
			END IF
			arrParam(5) = "��������"			
	
			arrField(0) = "a.MINOR_CD"					' Field��(0)
			arrField(1) = "a.MINOR_NM"					' Field��(1)
    
			arrHeader(0) = "��������"				' Header��(0)
			arrHeader(1) = "����������"				' Header��(1)
		case 5
			arrParam(0) = "ǰ����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_MINOR a,b_item_acct_inf b" 					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			IF frm1.cboProcurType.value = "P" Then
				arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('3RAW','4SUB','5GOODS') "
			ELSE
				arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('1FINAL','2SEMI') "
			END IF			

			arrParam(5) = "ǰ�����"			
	
			arrField(0) = "MINOR_CD"					' Field��(0)
			arrField(1) = "MINOR_NM"					' Field��(1)
    
			arrHeader(0) = "ǰ�����"				' Header��(0)
			arrHeader(1) = "ǰ�������"				' Header��(1)
		case 6
			arrParam(0) = "ǰ�� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtItemCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			IF frm1.cboProcurType.value = "P" Then
				arrParam(4) = "a.item_cd = b.item_cd and b.procur_type = " & FilterVar("P", "''", "S") & " "		' Where Condition
			ELSE
				arrParam(4) = "a.item_cd = b.item_cd and b.procur_type <> " & FilterVar("P", "''", "S") & " "		' Where Condition
			END IF			
			arrParam(5) = "ǰ��"			
	
			arrField(0) = "a.ITEM_CD"					' Field��(0)
			arrField(1) = "a.ITEM_NM"					' Field��(1)
    
			arrHeader(0) = "ǰ��"				' Header��(0)
			arrHeader(1) = "ǰ���"				' Header��(1)
	end select
		
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
      Select case iWhere
		case 1
			frm1.txtPlantCD.focus		
		case 2
			frm1.txtCostCd.focus		
		case 3
			frm1.txtTrnsTypeCd.focus		
		case 4
			frm1.txtMovTypeCd.focus		
		case 5
			frm1.txtItemAcctCd.focus		
		case 6
			frm1.txtItemCd.focus
	  End Select		
		Exit Function
	Else
		Call SetReturnVal(iWhere,arrRet)
	End If	

End Function'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function SetReturnVal(byval iwhere,byval arrRet)
	With frm1
		select case iWhere	
			case 1
				.txtPlantCD.focus	
				.txtplantCd.Value	= arrRet(0)
				.txtPlantNm.Value	= arrRet(1)
			case 2
				.txtCostCd.focus
				.txtCostCd.Value	= arrRet(0)
				.txtCostNm.Value	= arrRet(1)
			case 3
				.txtTrnsTypeCd.focus
				.txtTrnsTypeCd.Value	= arrRet(0)
				.txtTrnsTypeNm.Value	= arrRet(1)
			case 4
				.txtMovTypeCd.focus
				.txtMovTypeCd.Value	= arrRet(0)
				.txtMovTypeNm.Value	= arrRet(1)
			case 5
				.txtItemAcctCd.focus
				.txtItemAcctCd.Value	= arrRet(0)
				.txtItemAcctNm.Value	= arrRet(1)
			case 6
				.txtItemCd.focus
				.txtItemCd.Value	= arrRet(0)
				.txtItemNm.Value	= arrRet(1)
		end select 		
	End With

End Function

'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboProcurType, "P", "����ǰ")								
	Call SetCombo(frm1.cboProcurType, "M", "����ǰ")
End Sub

'=======================================================================================================
'	Name : ExeReflect()
'	Description : �򰡱ݾ� �ݿ��۾� 
'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD

    ExeReflect = False															'��: Processing is NG

'	if frm1.cboProcurType.value = "" then
'		IntRetCD = DisplayMsgBox("232520","X","X","X") '���ð��� �����ϴ� 
'	end if
    
    If Not chkField(Document, "1") Then                             '��: Check contents area
       Exit Function
    End If


	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	    
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
	Dim strYear
	Dim strMonth
	Dim strDay
    
    With frm1
    
		Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	    
		strVal = BIZ_PGM_EXE_ID 
		strVal = strVal & "?txtMode=" & Parent.UID_M0002
		strVal = strVal & "&txtYyyymm=" & strYear & strMonth
		strVal = strVal & "&cboProcurType=" & Trim(.cboProcurType.value)


		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
'	Call ExecMyBizASP(frm1, BIZ_PGM_EXE_ID)

    End With	
    
    ExeReflect = True         
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc :
'=======================================================================================================
Function ExeReflectOk()
Dim IntRetCD 

	window.status = "�ݿ� �۾� �Ϸ�"

	IntRetCD =DisplayMsgBox("990000","X","X","X")

	MainQuery
			
End Function

'=======================================================================================================
'	Name : ExeCancel()
'	Description : �򰡱ݾ� �ݿ��۾� 
'=======================================================================================================
Function ExeCancel()
	Dim IntRetCD

    ExeCancel = False															'��: Processing is NG

'	if frm1.cboProcurType.value = "" then
'		IntRetCD = DisplayMsgBox("232520","X","X","X") '���ð��� �����ϴ� 
'	end if
    
    If Not chkField(Document, "1") Then                             '��: Check contents area
       Exit Function
    End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	    
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
	Dim strYear
	Dim strMonth
	Dim strDay
    
    With frm1
    
		Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	    
		strVal = BIZ_PGM_CNC_ID 
		strVal = strVal & "?txtMode=" & Parent.UID_M0003
		strVal = strVal & "&txtYyyymm=" & strYear & strMonth
		strVal = strVal & "&cboProcurType=" & Trim(.cboProcurType.value)
		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
'	Call ExecMyBizASP(frm1, BIZ_PGM_EXE_ID)
		
    End With	
    
    ExeCancel = True         
End Function

'======================================================================================================
' Function Name : ExeCancelOk
' Function Desc :
'=======================================================================================================
Function ExeCancelOk()
Dim IntRetCD

	window.status = "��� �۾� �Ϸ�"

	IntRetCD =DisplayMsgBox("990000","X","X","X")

	MainQuery

End Function


Function OpenPopupGL()
	Dim iCalledAspName
	Dim IntRetCD

 
	Dim arrRet
	Dim arrParam(1)	
    
	If lgIsOpenPop = True Then Exit Function

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_DocNo
	
	if Trim(frm1.vspdData.Value) = "" then
       Call DisplayMsgBox("169804","X", "X", "X")    '���ҹ�ȣ�� �ʿ��մϴ� 
       Exit Function
    End If
	
		   	
	arrParam(0) = ""			'ȸ����ǥ��ȣ 
	arrParam(1) = Trim(frm1.vspdData.Value) & "-" & Trim(frm1.txtYyyymm.Year)   '���ҹ�ȣ'
    
	lgIsOpenPop = True

	
	iCalledAspName = AskPRAspName("A5120RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5120RA1","x")
		IsOpenPop = False
		Exit Function
	End If
   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

 	
End Function


Function OpenMoveDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim strYear,strMonth,strDay
	Dim arrRet
	Dim Param1  '������ڵ� 
	Dim Param2  '������ 
	Dim Param3  '���ҹ�ȣ 
	Dim Param4	'���ҹ߻��� 
	Dim Param5	'ȸ����ǥ�߻��� 
	Dim Param6  '���ұ��� 
	Dim Param7  '���ұ��� 
	
	If lgIsOpenPop = True Then Exit Function


	IF lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("800167","X", "X", "X")    '��ȸ�� ���� �ϼ��� 
		Exit Function	    
	ENd IF	
	
	With frm1.vspdData	
	


		.Row = .ActiveRow 
		.Col = C_MovTypeNm
		
	    Param7 = .Value


		
		
		Param4 = UniDateAdd("D",-1,UniDateAdd("m", 1,UniConvYYYYMMDDToDate(parent.gDateFormat ,frm1.txtYyyymm.Year, frm1.txtYyyymm.Month, "01"),parent.gDateFormat),parent.gDateFormat) 
		Param5 = UniDateAdd("D",-1,UniDateAdd("m", 1,UniConvYYYYMMDDToDate(parent.gDateFormat ,frm1.txtYyyymm.Year, frm1.txtYyyymm.Month, "01"),parent.gDateFormat),parent.gDateFormat) 

		.Row = .ActiveRow 
		.Col = C_TrnsTypeNm
		Param6 = Trim(.Value)

		
  
	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169804","X", "X", "X")    '���ҹ�ȣ�� �ʿ��մϴ� 
			Exit Function
		else
		   .Col = C_DocNo			: .Row = .ActiveRow : Param3 = Trim(.Text )
			IF Param3 = "" Then
				Call DisplayMsgBox("169804","X", "X", "X")    '���ҹ�ȣ�� �ʿ��մϴ� 
				Exit Function
			END IF
		End If	
		
		ggoSpread.Source = frm1.vspdData    
		.Row = .ActiveRow 
		.Col = C_DocNo
		
		IntRetCD = CommonQueryRs("a.biz_area_cd,b.biz_area_nm","i_goods_movement_header a,b_biz_area b","a.biz_area_cd = b.biz_area_cd and a.item_document_no = " & FilterVar(.Value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If IntRetCD = False Then
			Call DisplayMsgBox("169803","X", "X", "X")    '������ڷᰡ �ʿ��մϴ� 
			Exit Function	    
		Else
		    Param1 = Trim(Replace(lgF0,Chr(11),""))
		    Param2 = Trim(Replace(lgF1,Chr(11),""))
		End If
		
    End With
    	
    if Param3 = "" then
       Call DisplayMsgBox("169804","X", "X", "X")    '���ҹ�ȣ�� �ʿ��մϴ� 
    	Exit Function
    End If
	
	lgIsOpenPop = True

	iCalledAspName = AskPRAspName("I1711RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1711RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
    		'Call SetPartRef(arrRet)
	End If	
	
End Function



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
            
			C_PlantCd				= iCurColumnPos(1)
			C_TrnsTypeCd			= iCurColumnPos(2)
			C_TrnsTypeNm			= iCurColumnPos(3)
			C_MovTypeCd			    = iCurColumnPos(4)    
			C_MovTypeNm			    = iCurColumnPos(5)
			C_CostCd			    = iCurColumnPos(6)
			C_ItemAcctNm			= iCurColumnPos(7)
			C_ItemCd			    = iCurColumnPos(8)
			C_ItemNm				= iCurColumnPos(9)
			C_DocNo					= iCurColumnPos(10)
			C_SeqNo					= iCurColumnPos(11)
			C_DiffAmt				= iCurColumnPos(12)
			C_TrnsPlantCd			= iCurColumnPos(13)
			C_TrnsSlCd				= iCurColumnPos(14)
			C_TrnsItemCd			= iCurColumnPos(15)
			C_TrnsItemNm			= iCurColumnPos(16)
			
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
                                                                            'Format Numeric Contents Field                                                                            
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    
    Call SetDefaultVal
	Call InitComboBox
    Call SetToolbar("11000000000011")
		frm1.txtYyyymm.focus
		frm1.BtnExe.disabled = True
		frm1.BtnCnc.disabled = True
		
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_ProcurTypeNm Or NewCol <= C_ProcurTypeNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
End Sub


'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ 
    	If lgStrPlantPrevKey <> "" Then                  '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
      	DbQuery
    	End If

    End if
    
End Sub



Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub



'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================


'======================================================================================================
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables                                                      'Initializes local global variables
    															
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
	IF DbQuery = False Then
		Exit Function
	END IF
	       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 '��: ȭ�� ���� 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      '��:ȭ�� ����, Tab ���� 
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
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
    FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery() 

    DbQuery = False

	Dim intRetCd	
	Dim arrTemp
	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		IF frm1.txtPlantCd.value <> "" Then
			intRetCd = CommonQueryRs("plant_nm","b_plant","plant_cd = " & FilterVar(Trim(frm1.txtPlantCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtPlantNm.value = arrTemp(0)
			ELSE
				frm1.txtPlantNm.value = ""
			ENd IF
		ELSE
			frm1.txtPlantNm.value = ""
		ENd IF

		IF frm1.txtCostCd.value <> "" Then
			intRetCd = CommonQueryRs("cost_nm","b_cost_center","cost_cd = " & FilterVar(Trim(frm1.txtCostCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtCostNm.value = arrTemp(0)
			ELSE
				frm1.txtCostNm.value = ""
			ENd IF
		ELSE
			frm1.txtCostNm.value = ""
		ENd IF

		IF frm1.txtTrnsTypeCd.value <> "" Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd = " & FilterVar("I0002", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtTrnsTypeCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtTrnsTypeNm.value = arrTemp(0)
			ELSE
				frm1.txtTrnsTypeNm.value = ""
			ENd IF
		ELSE
			frm1.txtTrnsTypeNm.value = ""
		ENd IF

		IF frm1.txtMovTypeCd.value <> ""  Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd =" & FilterVar("I0001", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtMovTypeCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtMovTypeNm.value = arrTemp(0)
			ELSE
				frm1.txtMovTypeNm.value = ""
			ENd IF
		ELSE
			frm1.txtMovTypeNm.value = ""
		ENd IF

		IF frm1.txtItemAcctCd.value <> "" Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd = " & FilterVar("P1001", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtItemAcctCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtItemAcctNm.value = arrTemp(0)
			ELSE
				frm1.txtItemAcctNm.value = ""
			ENd IF
		ELSE
			frm1.txtItemAcctNm.value = ""
		ENd IF

		IF frm1.txtItemCd.value <> "" Then
			intRetCd = CommonQueryRs("item_nm","b_item","item_cd = " & FilterVar(Trim(frm1.txtItemCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtItemNm.value = arrTemp(0)
			ELSE
				frm1.txtItemNm.value = ""
			ENd IF
		ELSE
			frm1.txtItemNm.value = ""
		ENd IF
		
	ENd If

    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	    
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
	Dim strYear, strMonth, strDay
   

	frm1.BtnExe.disabled = True
	frm1.BtnCnc.disabled = True
	
   
    With frm1
    
    
		Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

		If lgIntFlgMode = Parent.OPMD_UMODE Then
		 '@Query_Hidden     
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'Hidden�� �˻��������� Query
			strVal = strVal & "&txtYyyymm=" & .hYyyymm.value				
			strVal = strVal & "&cboProcurType=" & .hProcurType.value				
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.hCostCd.value)
			strVal = strVal & "&txtTrnsTypeCd=" & Trim(.hTrnsTypeCd.value)
			strVal = strVal & "&txtMovTypeCd=" & Trim(.hMovTypeCd.value)
			strVal = strVal & "&txtItemAcctCd=" & Trim(.hItemAcctCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
			strVal = strVal & "&lgStrPlantPrevKey=" & lgStrPlantPrevKey
			strVal = strVal & "&lgStrTrnsPrevKey=" & lgStrTrnsPrevKey
			strVal = strVal & "&lgStrMovPrevKey=" & lgStrMovPrevKey
			strVal = strVal & "&lgStrCostPrevKey=" & lgStrCostPrevKey
			strVal = strVal & "&lgStrItemPrevKey=" & lgStrItemPrevKey
			strVal = strVal & "&lgStrTrnsPlantPrevKey=" & lgStrTrnsPlantPrevKey
			strVal = strVal & "&lgStrTrnsSlPrevKey=" & lgStrTrnsSlPrevKey
			strVal = strVal & "&lgStrTrnsItemPrevKey=" & lgStrTrnsItemPrevKey						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS_D
		Else
			
		 '@Query_Text     
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'Hidden�� �˻��������� Query
			strVal = strVal & "&txtYyyymm=" & strYear & strMonth
			strVal = strVal & "&cboProcurType=" & .cboProcurType.value				
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.txtCostCd.value)
			strVal = strVal & "&txtTrnsTypeCd=" & Trim(.txtTrnsTypeCd.value)
			strVal = strVal & "&txtMovTypeCd=" & Trim(.txtMovTypeCd.value)
			strVal = strVal & "&txtItemAcctCd=" & Trim(.txtItemAcctCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&lgStrPlantPrevKey=" & lgStrPlantPrevKey
			strVal = strVal & "&lgStrTrnsPrevKey=" & lgStrTrnsPrevKey
			strVal = strVal & "&lgStrMovPrevKey=" & lgStrMovPrevKey
			strVal = strVal & "&lgStrCostPrevKey=" & lgStrCostPrevKey
			strVal = strVal & "&lgStrItemPrevKey=" & lgStrItemPrevKey
			strVal = strVal & "&lgStrTrnsPlantPrevKey=" & lgStrTrnsPlantPrevKey
			strVal = strVal & "&lgStrTrnsSlPrevKey=" & lgStrTrnsSlPrevKey
			strVal = strVal & "&lgStrTrnsItemPrevKey=" & lgStrTrnsItemPrevKey						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS_D
		End If


		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()													'��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
	

    
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field

    Call SetToolbar("11000000000111")
	frm1.BtnExe.disabled = False
	frm1.BtnCnc.disabled = False
	
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	
'======================================================================================================= -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������̹ݿ�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A  href="vbscript:OpenMoveDtlRef()">���һ� |</A>&nbsp;<A  href="vbscript:OpenPopupGL()">ȸ����ǥ���� </A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>�۾����</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c3900ma1_txtYyyymm_txtYyyymm.js'></script>
									</TD>								
									<TD CLASS="TD5" NOWRAP>���ޱ���</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboProcurType" tag="12X" STYLE="WIDTH:82px:" ALT="���ޱ���"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">����</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(1)">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">�ڽ�Ʈ��Ÿ</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ڽ�Ʈ��Ÿ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(2)">
										<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=30 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">���ұ���</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtTrnsTypeCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���ұ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrnsTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(3)">
										<INPUT TYPE=TEXT NAME="txtTrnsTypeNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">��������</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMovTypeCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(4)">
										<INPUT TYPE=TEXT NAME="txtMovTypeNm" SIZE=30 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">ǰ�����</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtItemAcctCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="ǰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(5)">
										<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">ǰ��</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(6)">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14X">
									</TD>
								</TR>
							
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
  							<TR>
 								<TD HEIGHT="100%" NOWRAP>
 								<script language =javascript src='./js/c3900ma1_I907901257_vspdData.js'></script>
 								</TD>
 							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TDT">
									<TD CLASS="TD6">
									<TD CLASS="TD5" NOWRAP>�� ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3900ma1_fpDoubleSingle2_txtSum.js'></script>&nbsp;
	                                </TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
				<TD><BUTTON NAME="BtnExe" CLASS="CLSSBTN" ONCLICK="ExeReflect()" >�ݿ�</BUTTON>&nbsp;<BUTTON NAME="BtnCnc" CLASS="CLSSBTN" ONCLICK="ExeCancel()" >���</BUTTON></TD>
				<TD WIDTH=*>&nbsp;</TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>

</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hProcurType" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTrnsTypeCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovTypeCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

