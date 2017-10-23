<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111ma5
'*  4. Program Name         : ���ſ�û�������-��Ƽ 
'*  5. Program Desc         : ���ſ�û�������-��Ƽ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Han Kwang Soo
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
'********************************************************************************************************* !-->
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "m2111mb5.asp"	
Const BIZ_PGM_ID2 = "m2111mb501.asp"	
Const BIZ_PGM_SAVE_ID = "m2111mb5.asp"	
											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��� ��������	
Dim C_CfmFlg				'���� 
Dim C_PlantCd 	            '���� 
Dim C_PlantNm 	            '����� 
Dim C_ItemCd 	            'ǰ�� 
Dim C_ItemNm 	            'ǰ��� 
Dim C_SpplSpec              'ǰ��԰� 
Dim C_ReqQty 	            '��û�� 
Dim C_Unit 		            '���� 
Dim C_UnitPopup	            '�����˾� 
Dim C_DlvyDt				'�ʿ䳳���� 
Dim C_ORGCd                 '��������			'2003-02-24�߰� - KSH
Dim C_ORGCdPopup            '�������� �˾�		
Dim C_ORGNm                 '����������		
Dim C_ReqNo 	            '��û��ȣ 
Dim C_ReqDt					'��û�� 
Dim C_ReqStateCd			'���ſ�û���� 
Dim C_ReqStateNm            '���ſ�û���¸� 
Dim C_ReqTypeCd				'���ſ�û���� 
Dim C_ReqTypeNm	            '���ſ�û���и� 
Dim C_MrpRunNo	            'MRP run ��ȣ 
Dim C_ReqDept				'��û �μ�	'030107
Dim C_ReqPrsn				'��û��		'030107
Dim C_TrackingNo			'tracking_no 200308
'��� �������� ���� 
Dim C_ProcType				'���ޱ��� 
Dim C_ReqCfmQty				'��ûȮ������ 
Dim C_BaseReqQty			'�⺻��û���� 
Dim C_BaseReqUnit			'�⺻��û���� 
Dim C_OrdQty                '���ּ��� 
Dim C_RcptQty				'�԰� 
Dim C_IvQty					'���Է� 

'�ϴ� �������� 
Dim C_SpplCd				'����ó 
Dim C_SpplPopup              '����ó �˾� 
Dim C_SpplNm 	             '����ó�� 
Dim C_Quota_Rate             '��к��� 
Dim C_ApportionQty            '��η� 
Dim C_PlanDt                 '���ֿ����� 
Dim C_GrpCd 	             '���ű׷� 
Dim C_GrpPopup               '���ű׷��˾� 
Dim C_GrpNm 	             '���ű׷�� 
Dim C_ParentPrNo 	         '���� ��û��ȣ (Ű��)
Dim C_ParentRowNo            '���� row ��ȣ 
Dim C_Flag                   '�ڱ� ��ȣ 

Dim lgSpdHdrClicked	'2003-03-01 Release �߰� 
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim lgIntFlgModeM                 'Variable is for Operation Status

Dim lgStrPrevKeyM			'Multi���� �������� ���� ���� 
Dim lglngHiddenRows		'Multi���� �������� ���� ����	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgSortKey1
Dim lgSortKey2

Dim IsOpenPop
Dim lgCurrRow
Dim strInspClass

Dim lgPageNo1
Dim StartPRDt, StartDlvyDt, EndDt,CurrDate
CurrDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
StartPRDt = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartPRDt = UniConvDateAToB(StartPRDt, parent.gServerDateFormat, parent.gDateFormat)
StartDlvyDt = uniDateAdd("m", 0, "<%=GetSvrDate%>", parent.gServerDateFormat)			'������ 
StartDlvyDt = UniConvDateAToB(StartDlvyDt, parent.gServerDateFormat, parent.gDateFormat)
EndDt = uniDateAdd("m", 1, "<%=GetSvrDate%>", parent.gServerDateFormat)
EndDt   = UniConvDateAToB(EndDt, parent.gServerDateFormat, parent.gDateFormat)
      
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0						'initializes Group View Size
        
    lgStrPrevKey1 = ""						'initializes Previous Key
    lgStrPrevKey2 = ""						'initializes Previous Key
    
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    lgPageNo = 0
    lgPageNo1 = 0
    
    '###�˻�з��� ����κ� Start###
    strInspClass = "R"
	'###�˻�з��� ����κ� End###
    'ggoSpread.ClearSpreadData
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
	frm1.txtPlantCd.Value	= Parent.gPlant
	frm1.txtReqFrDt.Text	= StartPRDt
	frm1.txtReqToDt.Text	= EndDt
	frm1.txtDlvyFrDt.Text	= StartDlvyDt
	frm1.txtDlvyToDt.Text	= EndDt
	
	Call SetToolbar("1100000000001111")
	
    frm1.txtPlantCd.focus 
    Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()
	
	With frm1.vspdData

	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030901",,Parent.gAllowDragDropSpread

	.ReDraw = false
		
    .MaxCols = C_TrackingNo + 1							
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0
   
    Call GetSpreadColumnPos("A")
	ggoSpread.SSSetCheck    C_CfmFlg, "����",10,,,true
    ggoSpread.SSSetEdit 	C_PlantCd,"����",7,,,4,2
    ggoSpread.SSSetEdit 	C_ItemCd, "ǰ��", 10,,,18,2
    ggoSpread.SSSetEdit 	C_ItemNm, "ǰ���", 20
    ggoSpread.SSSetEdit     C_SpplSpec, "ǰ��԰�", 20        'ǰ��԰� �߰� 
    SetSpreadFloatLocal 	C_ReqQty, "��û��", 15, 1,3
    ggoSpread.SSSetEdit 	C_Unit,"����", 9,,,3,2
    ggospread.sssetButton 	C_unitPopup
    ggoSpread.SSSetDate 	C_DlvyDt, "�ʿ���", 12, 2, parent.gDateFormat
	ggoSpread.SSSetEdit		C_ORGCd,"��������",10,,,4,2
    ggoSpread.SSSetButton	C_ORGCdPopup
    ggoSpread.SSSetEdit		C_ORGNm,"����������",20
    ggoSpread.SSSetEdit 	C_ReqNo, "��û��ȣ", 20,,,,2
    ggoSpread.SSSetDate 	C_ReqDt, "��û��", 10, 2, parent.gDateFormat
    ggoSpread.SSSetEdit 	C_ReqStateCd, "���ſ�û����",15,,,5,2
    ggoSpread.SSSetEdit 	C_ReqStateNm, "���ſ�û���¸�",20
    ggoSpread.SSSetEdit 	C_ReqTypeCd, "���ſ�û����",15,,,5,2
    ggoSpread.SSSetEdit 	C_ReqTypeNm, "���ſ�û���и�",18
    ggoSpread.SSSetEdit 	C_MrpRunNo, "MRP Run��ȣ",20
    ggoSpread.SSSetEdit 	C_ReqDept, "��û�μ�",15		
	ggoSpread.SSSetEdit 	C_ReqPrsn, "��û��",15
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.",25
	
	Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)
    Call SetSpreadLock 
    .ReDraw = true

    End With    
End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables2()
    With frm1

		.vspdData2.ReDraw = false
		
		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread
       
	   .vspdData2.MaxCols = C_Flag+1
	   .vspdData2.MaxRows = 0
 
		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit		C_SpplCd, "����ó", 10,,,10,2
		ggoSpread.SSSetButton	C_SpplPopup
		ggoSpread.SSSetEdit 	C_SpplNm, "����ó��", 18
		SetSpreadFloatLocal		C_Quota_Rate, "��к���(%)",15,1,5
		SetSpreadFloatLocal		C_ApportionQty, "��η�",15,1,3
		ggoSpread.SSSetDate		C_PlanDt, "���ֿ�����", 15, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_GrpCd,"���ű׷�",15,,,4,2
		ggoSpread.SSSetButton	C_GrpPopUp
		ggoSpread.SSSetEdit 	C_GrpNm,"���ű׷��",20
		ggoSpread.SSSetEdit 	C_ParentPrNo, "��û��ȣ", 10
		ggoSpread.SSSetEdit		C_ParentRowNo , "C_ParentRowNo", 5
		ggoSpread.SSSetEdit		C_Flag , "C_Flag", 5
		
		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(C_ParentPrNo,	C_ParentPrNo,	True)
		Call ggoSpread.SSSetColHidden(C_ParentRowNo,C_ParentRowNo, True)
 		Call ggoSpread.SSSetColHidden(C_Flag, C_Flag+1, True)

		.vspdData2.ReDraw = True
 
    End With
	Call SetSpreadLock2()
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
 Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock		C_PlantCd , -1
    ggoSpread.SpreadUnLock		C_ReqQty , -1, -1
    ggoSpread.SSSetRequired		C_ReqQty, -1, -1                  '��û�� 
    ggoSpread.SpreadUnLock		C_Unit , -1, C_UnitPopUp, -1   '���� 
	ggoSpread.SSSetRequired		C_Unit, -1, -1    
    ggoSpread.SpreadUnLock		C_DlvyDt , -1, -1
    ggoSpread.SSSetRequired		C_DlvyDt, -1, -1                   '�ʿ��� 
    ggoSpread.SpreadUnLock		C_ORGCd , -1, -1
    ggoSpread.SSSetRequired		C_ORGCd, -1, -1                   '�ʿ��� 
    ggoSpread.SpreadUnLock		C_ORGCdPopup , -1, C_ORGCdPopup, -1   '���� 
    
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()    
    With frm1
    
    .vspdData2.ReDraw = False
    
    ggoSpread.Source = frm1.vspdData2
            
	ggoSpread.SpreadLock		C_SpplCd,		-1,	C_SpplNm,		-1
	ggoSpread.SSSetRequired		C_Quota_Rate, -1, -1	
	ggoSpread.SSSetRequired		C_ApportionQty, -1, -1
	ggoSpread.SSSetRequired		C_PlanDt, -1, -1
	ggoSpread.spreadUnlock		C_GrpCd,		-1,	C_GrpPopup,    -1
	ggoSpread.SSSetRequired		C_GrpCd, -1, -1
	ggoSpread.SpreadLock		C_GrpNm,		-1,	C_GrpNm,		-1

	.vspdData2.ReDraw = True
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
    ggoSpread.SSSetProtected		C_ItemCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ItemNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SpplSpec, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_ReqQty, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_Unit, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_DlvyDt, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_ORGCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ReqNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ReqDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_ReqStateCd, pvStartRow, pvEndRow		'030107
    ggoSpread.SSSetProtected		C_ReqStateNm, pvStartRow, pvEndRow		'030107
    
    ggoSpread.SSSetProtected		C_ReqTypeCd, pvStartRow, pvEndRow		'030107
    ggoSpread.SSSetProtected		C_ReqTypeNm, pvStartRow, pvEndRow		'030107
    ggoSpread.SSSetProtected		C_MrpRunNo, pvStartRow, pvEndRow		'030107
    ggoSpread.SSSetProtected		C_ReqDept, pvStartRow, pvEndRow		'030107
    ggoSpread.SSSetProtected		C_ReqPrsn, pvStartRow, pvEndRow		'030107
    
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData2.ReDraw = False
    
	ggoSpread.SSSetRequired  C_SpplCd,		    pvStartRow,	pvEndRow	
	ggoSpread.SSSetProtected C_SpplNm,		    pvStartRow,	pvEndRow	
	ggoSpread.SSSetRequired  C_Quota_Rate,		pvStartRow,	pvEndRow				
    ggoSpread.SSSetRequired  C_ApportionQty,	pvStartRow,	pvEndRow				
    ggoSpread.SSSetRequired  C_PlanDt,			pvStartRow,	pvEndRow				
    ggoSpread.SSSetRequired  C_GrpCd,			pvStartRow,	pvEndRow				
    ggoSpread.SSSetProtected C_GrpNm,		    pvStartRow,	pvEndRow	
    
   .vspdData2.ReDraw = True
    End With
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_CfmFlg		= 1		'���� 
	C_PlantCd		= 2	 	'���� 
	C_ItemCd		= 3	 	'ǰ�� 
	C_ItemNm		= 4	 	'ǰ��� 
	C_SpplSpec 		= 5		'ǰ��԰� 
	C_ReqQty 		= 6		'��û�� 
	C_Unit 			= 7		'���� 
	C_UnitPopup		= 8		'�����˾� 
	C_DlvyDt		= 9		'�ʿ䳳���� 
	C_ORGCd         = 10    '��������			'2003-02-24�߰� - KSH
	C_ORGCdPopup    = 11    '�������� �˾�		
	C_ORGNm         = 12    '����������		
	C_ReqNo 		= 13	'��û��ȣ 
	C_ReqDt			= 14	'��û�� 
	C_ReqStateCd 	= 15	'���ſ�û���� 
	C_ReqStateNm	= 16	'���ſ�û���¸� 
	C_ReqTypeCd		= 17	'���ſ�û���� 
	C_ReqTypeNm		= 18	'���ſ�û���и� 
	C_MrpRunNo		= 19	'MRP run ��ȣ 
	C_ReqDept		= 20	'��û �μ� 
	C_ReqPrsn		= 21	'��û�� 
	C_TrackingNo	= 22	'Tracking_No		'200308 �߰� 
	
End Sub

Sub InitSpreadPosVariables2()
	C_SpplCd        = 1          '����ó 
	C_SpplPopup     = 2          '����ó �˾� 
	C_SpplNm 	    = 3          '����ó�� 
	C_Quota_Rate    = 4          '��к��� 
	C_ApportionQty  = 5          '��η� 
	C_PlanDt        = 6          '���ֿ����� 
	C_GrpCd 	    = 7          '���ű׷� 
	C_GrpPopup      = 8          '���ű׷��˾� 
	C_GrpNm 	    = 9          '���ű׷�� 
	C_ParentPrNo    = 10	     '���� ��û��ȣ (Ű��)
	C_ParentRowNo   = 11         '���� row ��ȣ 
	C_Flag          = 12         '�ڱ� ��ȣ 
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
			C_CfmFlg			= iCurColumnPos(1)
			C_PlantCd 		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_SpplSpec		= iCurColumnPos(5)
			C_ReqQty		= iCurColumnPos(6)
			C_Unit			= iCurColumnPos(7)
			C_UnitPopup  	= iCurColumnPos(8)
			C_DlvyDt		= iCurColumnPos(9)
			C_ORGCd         = iCurColumnPos(10)    '��������			'2003-02-24�߰� - KSH
			C_ORGCdPopup    = iCurColumnPos(11)    '�������� �˾�		
			C_ORGNm         = iCurColumnPos(12)    '����������		
			C_ReqNo 		= iCurColumnPos(13)
			C_ReqDt			= iCurColumnPos(14)
			C_ReqStateCd	= iCurColumnPos(15)
			C_ReqStateNm	= iCurColumnPos(16)
			C_ReqTypeCd    	= iCurColumnPos(17)
			C_ReqTypeNm		= iCurColumnPos(18)
			C_MrpRunNo		= iCurColumnPos(19)
			C_ReqDept 		= iCurColumnPos(20)
			C_ReqPrsn		= iCurColumnPos(21)
			C_TrackingNo	= iCurColumnPos(22)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SpplCd        =	iCurColumnPos(1)         '����ó 
			C_SpplPopup     =	iCurColumnPos(2)         '����ó �˾� 
			C_SpplNm 	    =	iCurColumnPos(3)         '����ó�� 
			C_Quota_Rate    =	iCurColumnPos(4)         '��к��� 
			C_ApportionQty  =	iCurColumnPos(5)         '��η� 
			C_PlanDt        =	iCurColumnPos(6)         '���ֿ����� 
			C_GrpCd 	    =	iCurColumnPos(7)        '���ű׷� 
			C_GrpPopup      =	iCurColumnPos(8)        '���ű׷��˾� 
			C_GrpNm 	    =	iCurColumnPos(9)        '���ű׷�� 
			C_ParentPrNo    =	iCurColumnPos(10)	     '���� ���� (Ű��)
			C_ParentRowNo   =	iCurColumnPos(11)        '���� row ��ȣ 
			C_Flag          =	iCurColumnPos(12)        '�ڱ� ��ȣ 
	End Select
End Sub	

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp  ���� 
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

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
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
    Dim iCalledAspName
    Dim IntRetCD
    
	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

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

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenPrTypeCd()  ---------------------------------------------
'	Name : OpenPrTypeCd()
'	Description : PR Type PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPrTypeCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ſ�û����"				
	arrParam(1) = "B_MINOR"						
	arrParam(2) = Trim(frm1.txtPrTypeCd.Value)	
'	arrParam(3) = Trim(frm1.txtPrTypeNm.Value)	
	arrParam(4) = "MAJOR_CD = " & FilterVar("M2102", "''", "S") & " "			
	arrParam(5) = "���ſ�û����"				
	
    arrField(0) = "MINOR_CD"					
    arrField(1) = "MINOR_NM"					
        
    arrHeader(0) = "���ſ�û����"			
    arrHeader(1) = "���ſ�û���и�"			
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPrTypeCd.focus
		Exit Function
	Else
		frm1.txtPrTypeCd.Value = arrRet(0)
		frm1.txtPrTypeNm.Value = arrRet(1)
		frm1.txtPrTypeCd.focus
	End If	
End Function

'===========================================================================
' Function Name : OpenMrp
' Function Desc : OpenMrp Reference Popup
'===========================================================================
Function OpenMrp()
    Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
   
    If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If 
		
	IsOpenPop = True

	arrParam(0) = "MRP Run��ȣ"				<%' �˾� ��Ī %>
	arrParam(1) = "(select distinct a.order_no A,a.confirm_dt B," & FilterVar("������������", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_EXPL_HISTORY a, m_pur_req b where a.order_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
    arrParam(1) = arrParam(1) & "union "
    arrParam(1) = arrParam(1) & "select distinct  a.run_no A, a.start_dt B ," & FilterVar("MRP����", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_MRP_HISTORY a, m_pur_req b where a.run_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " ) as g" <%' TABLE ��Ī %>
    

	arrParam(2) = Trim(frm1.txtMRP.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "MRP Run��ȣ"				<%' TextBox ��Ī %>

	arrField(0) = "A"
	arrField(1) = "B"
	arrField(2) = "D"
	
	arrHeader(0) = "MRP Run��ȣ"				<%' Header��(0)%>
	arrHeader(1) = "����"					<%' Header��(1)%>
	arrHeader(2) = "��������"				<%' Header��(2)%>
				
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtMRP.focus
		Exit Function
	Else
		frm1.txtMRP.value = arrRet(0)
		frm1.txtMRP.focus
	End If	
End Function

'------------------------------------------  OpenUnit()  ---------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp ���� 
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_Unit
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "��û����"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "��û����"			
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "��û����"		
    arrHeader(1) = "��û������"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_Unit
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Text = arrRet(0)		
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	End If	
End Function

'------------------------------------------  OpenSSupplier()  ---------------------------------------------
'	Name : OpenSSupplier()
'	Description : SpplCd PopUp ����ó 
'---------------------------------------------------------------------------------------------------------
Function OpenSSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

	arrParam(0) = "����ó"   
	arrParam(1) = "B_Biz_Partner"  
	 
	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow 
	frm1.vspdData2.Col = C_SpplCd 
	arrParam(2) = FilterVar(Trim(frm1.vspdData2.text)," ","SNM")
	 
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
		With frm1.vspdData2 
			.Row = .ActiveRow 
			.Col = C_SpplCd
			.text = arrRet(0) 
			.Row = .ActiveRow 
			.Col = C_SpplNm
			.text = arrRet(1)
			Call SpplChange() 
		End With 
	End If 
End Function

'------------------------------------------  OpenSORG()  -------------------------------------------------
'	Name : OpenSORG()
'	Description : OpenSORG PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"				
	arrParam(1) = "B_Pur_Org"				

	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	frm1.vspdData.Col=C_ORGCd 
	
	arrParam(2) = Trim(frm1.vspdData.Text)
'	arrParam(3) = Trim(frm1.txtORGNm.Value)	

	arrParam(4) = ""						
	arrParam(5) = "��������"				
	
    arrField(0) = "PUR_ORG"					
    arrField(1) = "PUR_ORG_NM"				
    
    arrHeader(0) = "��������"			
    arrHeader(1) = "����������"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_ORGCd
			.text = arrRet(0) 
			.Row = .ActiveRow 
			.Col = C_ORGNm
			.text = arrRet(1) 
		End With 
	End If	
End Function

'------------------------------------------  OpenSGrp()  ---------------------------------------------
'	Name : OpenSGrp()
'	Description : grpCd PopUp ���ű׷� 
'---------------------------------------------------------------------------------------------------------
Function OpenSGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData2.Col=C_GrpCd 
	frm1.vspdData2.Row=frm1.vspdData2.ActiveRow 

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_PUR_GRP"				
	arrParam(2) = Trim(frm1.vspdData2.Text)
	arrParam(3) = ""
	frm1.vspdData.Col=C_ORGCd 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	arrParam(4) = "B_PUR_GRP.PUR_ORG= " & FilterVar(frm1.vspdData.Text, "''", "S") & " "
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData2 
			.Row = .ActiveRow 
			.Col = C_GrpCd
			.text = arrRet(0) 
			.Row = .ActiveRow 
			.Col = C_GrpNm
			.text = arrRet(1) 
			Call vspdData2_Change(C_GrpCd,.ActiveRow)
		End With 
	End If	
End Function
'===========================================================================
' Function Name : OpenTrackingNo				200309
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = ""	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 
    
'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		lgBlnFlgChgValue = True
	End If	

End Function
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select   
End Sub

'====================================== sprRedComColor() ======================================
'	Name : sprRedComColor()
'	Description : �������ڰ� ���� ���ں��� ������ ���� ��ȣ...
'==============================================================================================
Sub sprRedComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData2.Col = Col
		.vspdData2.Col2 = Col
		.vspdData2.Row = Row
		.vspdData2.Row2 = Row2
		.vspdData2.ForeColor = vbRed
    End With
End Sub
'====================================== sprBlackComColor() ======================================
'	Name : sprBlackComColor()
'	Description : �������ڰ� ���� ���ں��� ������ ���� ��ȣ...
'==============================================================================================
Sub sprBlackComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData2.Col = Col
		.vspdData2.Row = Row
        .vspdData2.ForeColor = &H0&
    End With
End Sub
'====================================== checkdt() ======================================
'	Name : checkdt()
'	Description : �������ڿ� ���� ����üũ.
'==============================================================================================
Sub checkdt(ByVal Row)
    With frm1
        .vspdData2.Row = Row
        .vspdData2.Col = C_PlanDt
        If UniConvDateToYYYYMMDD(.vspdData2.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(CurrDate,parent.gDateFormat,"") and Trim(.vspdData2.Text) <> "" Then 
            Call sprRedComColor(C_PlanDt,Row,Row)
		else
		    Call sprBlackComColor(C_PlanDt,Row,Row)
        End If 
    End With
End Sub


'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setreference()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub Setminorcd()
    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iminorcd

    Err.Clear

	Call CommonQueryRs(" minor_cd ", " b_configuration ", " major_cd = " & FilterVar("M2105", "''", "S") & " and reference = " & FilterVar("Y", "''", "S") & "  and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iminorcd = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName
		Err.Clear 
		Exit Sub
	End If

    if Trim(lgF0) <> "" then
        if UCase(Trim(iminorcd(0))) = "D" then
            frm1.rdoAssflg(0).Checked = true
        elseif UCase(Trim(iminorcd(0))) = "R" then
            frm1.rdoAssflg(1).Checked = true
        else
            frm1.rdoAssflg(2).Checked = true
        End If
    End If

End Sub

'==========================================   ApportionQtyChange()  ======================================
'	Name : ApportionQtyChange()
'	Description : 
'=================================================================================================

Sub ApportionQtyChange(ByVal Row)
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
		
		.Col		= C_Quota_Rate
		iquotarate  = Unicdbl(.text)
		
		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)
		
		totalquotarate = 0
		totalApportionQty = 0
		
		for index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0 
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then
				.Col = C_Quota_Rate
				totalquotarate = totalquotarate + Unicdbl(.text)
		        if index <> clng(Row) then
				    .Col = C_ApportionQty
				    totalApportionQty = totalApportionQty + Unicdbl(.text)
		        End If
		    End If
		next 
		
		frm1.vspdData.Row = iparentrow
		frm1.vspdData.Col = C_ReqQty
		iReqQty = Unicdbl(frm1.vspdData.text)
		
		'�հ� ������� 100�̸� ��η� = ��û�� - �����η��� 
		if totalquotarate = 100 then
		    iApportionQty = iReqQty - totalApportionQty
		else
			iApportionQty = (iquotarate * iReqQty)/100
	    End If
	
		.Row  = Row    
		.Col  = C_ApportionQty
		.text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End with
End Sub		

'==========================================   SpplChange()  ======================================
'	Name : SpplChange()
'	Description : 
'=================================================================================================

Sub SpplChange()
    Err.Clear 

    If CheckRunningBizProcess = True Then
		Exit Sub
	End If           
    
    Dim strVal
    Dim strssText1, strssText2
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim iparentrow
    Dim index 
    Dim iRow
	
	with frm1.vspdData2
	    iRow        = .ActiveRow
		.Row		= .ActiveRow    
		.Col		= C_ParentPrNo
		strssText1	= Trim(.text)	
		.Col		= C_SpplCd
		strssText2	= Trim(.text)
		.Col        = C_ParentRowNo
		iparentrow  = Trim(.text)
		if strssText2 = "" then
			Exit Sub
		End If
				
	End with
	
	lngRangeFrom = DataFirstRow(iparentrow)
	lngRangeTo   = DataLastRow(iparentrow)

	for index = lngRangeFrom to lngRangeTo
	    if index <> iRow and strssText2 <> "" then
	        frm1.vspdData2.Row = index     
	        frm1.vspdData2.Col = C_SpplCd
	        if UCase(strssText2) = UCase(Trim(frm1.vspdData2.text)) then
                Call DisplayMsgBox("17A005","X" ,"����ó", "X")	
                frm1.vspdData2.Row = iRow     
	            frm1.vspdData2.Col = C_SpplCd	  
	            frm1.vspdData2.text = ""      
 	            Exit sub
	        End If
	    End If
	next 
		
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookSppl"	
    strVal = strVal & "&txtPrNo=" & strssText1
    strVal = strVal & "&txtBpCd=" & strssText2

    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)				
End Sub		

'=======================================================================================================
'   Sub Name : SheetFocus
'   Sub Desc : 
'=======================================================================================================
Sub SheetFocus(Byval iChildRow)
	Dim iParentRow
	Dim CheckField1 
	Dim CheckField2
	Dim i 
	Dim lngStart
	Dim lngEnd
	Dim strSampleNo
	Dim strFlag
	
	With frm1.vspdData2
		.Row = iChildRow
		.Col = C_ParentRowNo
		iParentRow = CLng(.Text)
		.Col = C_SpplCd
		strSampleNo = .Text
		.Col = C_Flag
		strFlag = .Text						
	End With
	
	Call ParentGetFocusCell(iParentRow, strSampleNo, strFlag)	
End Sub

'=======================================================================================================
'   Event Name : ParentGetFocusCell
'   Event Desc :
'=======================================================================================================
Sub ParentGetFocusCell(ByVal ParentRow, ByVal strSampleNo, Byval strFlag)
	Dim CheckField1 
	Dim CheckField2
	Dim i 
	Dim lngStart
	Dim lngEnd

	With frm1.vspdData
		.Row = ParentRow
		.Col = 1
		.Action = 0		'Active Cell
	End With
	
	With frm1.vspdData2
		.ReDraw = False
		lngStart = ShowFromData(ParentRow, lglngHiddenRows(ParentRow - 1))
		.ReDraw = True
		lngEnd = lngStart + lglngHiddenRows(ParentRow - 1) - 1
		For i = lngStart To lngEnd
			.Row = i
			.Col = C_SpplCd
			CheckField1 = .Text
			.Col = C_Flag
			CheckField2 = .Text
			If CheckField1 = strSampleNo And CheckField2 = strFlag Then
				Exit For
			End If
		Next
					
	End With

	Set gActiveElement = document.activeElement

End Sub

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
'check this !		
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
			.Col = C_Flag
			.Col2 = C_Flag
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

'=======================================================================================================
'   Function Name : DeleteDataForInsertSampleRows
'   Function Desc : 
'=======================================================================================================
Function DeleteDataForInsertSampleRows(ByVal Row, Byval lngShowingRows)
	DeleteDataForInsertSampleRows = False
	
	Dim lngRow
	Dim lngStartRow
	
	With frm1.vspdData2
		
		Call SortSheet()
		
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		lngStartRow = 0
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo                
			If Row = Clng(.Text) Then
				lngStartRow = lngRow
				DeleteDataForInsertSampleRows = True
				Exit For
			End If    
		Next
		
		'------------------------------------
		' Delete Data
		'------------------------------------ 
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = lngStartRow
			.Row2 = lngStartRow + lngShowingRows - 1
			.Action = 5		'5 - Delete Row 	SS_ACTION_DELETE_ROW
			'********** START
			.MaxRows = .MaxRows - lngShowingRows
			'********** END			
			.BlockMode = False
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
        .SortKey(2) = C_Flag
        
        .SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
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
								.Row = i
								.Action = 0
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

'=======================================================================================================
' Function Name : ChangeCheck
' Function Desc : 
'=======================================================================================================
Function ChangeCheck()
	ChangeCheck = False
	
	Dim i
	Dim strInsertMark
	Dim strDeleteMark
	Dim strUpdateMark
	
	ggoSpread.Source = frm1.vspdData2
	strInsertMark = ggoSpread.InsertFlag
	strDeleteMark = ggoSpread.UpdateFlag
	strUpdateMark = ggoSpread.DeleteFlag
	
	If frm1.vspdData.maxrows <= 0 Then Exit Function
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = strInsertMark Or .Text = strDeleteMark Or .Text = strUpdateMark Then
				ChangeCheck = True
				exit for
			End If
		Next
	End With
	
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or ChangeCheck = True Then
        ChangeCheck = True
    End If
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc : 
'=======================================================================================================
Function CheckDataExist()
	CheckDataExist = False
	Dim i
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow()
	ShowDataFirstRow = 0
	Dim i
	
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
' Function Name : ShowDataLastRow
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow()
	ShowDataLastRow = 0
	Dim i
	
	With frm1.vspdData2
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
' Function Name : DataFirstRow
' Function Desc : 
'=======================================================================================================
Function DataFirstRow(ByVal Row)
	DataFirstRow = 0
	Dim i
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
	DataLastRow = 0
	Dim i
	
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
' Function Name : InsertSampleRows
' Function Desc : 
'=======================================================================================================
Sub InsertSampleRows()
	Dim i
	Dim j
	Dim lngMaxRows
	Dim strInspItemCd
	Dim strInspSeries
	Dim lngOldMaxRows
	Dim strMark
	Dim lRow
	
    With frm1
    	If .vspdData.Row < 1 Then
    		Exit Sub
    	End If
    	
   		Call LayerShowHide(1)
    	
    	lRow = .vspdData.ActiveRow
    	' �ش� �˻��׸�/������ ������ �ִ� ����ġ�� ���� 
    	Call DeleteDataForInsertSampleRows(lRow, lglngHiddenRows(lRow - 1))
    	
    	' �� �߰� 
    	lngOldMaxRows = .vspdData2.MaxRows
    	
    	.vspdData.Row = lRow
    	.vspdData.Col = C_ApportionQty
    	lngMaxRows = UNICDbl(.vspdData.Text)
    	.vspdData2.MaxRows = lngOldMaxRows + lngMaxRows 
  
	End With        
	
    ggoSpread.Source = frm1.vspdData2
    strMark = ggoSpread.InsertFlag
    
    With frm1.vspdData2
		.BlockMode = True
		.Row = lngOldMaxRows + 1
		.Row2 = .MaxRows
		.Col = C_ParentRowNo
		.Col2 = C_ParentRowNo
		.Text = lRow
		.BlockMode = False
		
		j = 0
        For i = lngOldMaxRows + 1 To .MaxRows
			j = j + 1
			.Row = i
			.Col = 0
			.Text = strMark
			'********** START
			.Col = C_Flag
			.Text = strMark
			'********** END			
			.Col = C_SupplierCd
			.Text = j
		Next
	End With
	
	frm1.vspdData.Col = C_InspUnitIndctnCd
	
	Call SetSpreadColor2byInspUnitIndctn(lngOldMaxRows + 1, frm1.vspdData2.MaxRows, frm1.vspdData.Text, "I")
	
	frm1.vspdData2.Row = lngOldMaxRows + 1
	frm1.vspdData2.Col = C_SpplCd
	frm1.vspdData2.Action = 0
	lglngHiddenRows(lRow - 1) = lngMaxRows
    Call LayerShowHide(0)
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

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###
 	dim iParPrNo,  iParRowNo 
 	dim i ,j ,iReqNo
 	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData
 	
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0101111111")         'ȭ�麰 ���� 
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		
 		
 				'2006-10 hong  �ص� Ŭ���� ���� ���� ����*************
 		lgSpdHdrClicked = 0		'
 		Call vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
 		
 		frm1.vspdData.vspdData.ReDraw = False
 			for i = 1 to frm1.vspdData.MaxRows
 				iReqNo = Trim(GetSpreadText(frm1.vspdData,C_ReqNo,i,"X","X"))
 				 frm1.vspdData.Row = i 

 				For j = 1 to  frm1.vspdData2.MaxRows
 			        iParPrNo = Trim(GetSpreadText(frm1.vspdData2,C_ParentPrNo,j,"X","X"))
 			        iParRowNo = cdbl(GetSpreadText(frm1.vspdData2,C_ParentRowNo,j,"X","X"))
 			    
 			        If iReqNo = iParPrNo and   iParRowNo <> i then 
 			           frm1.vspdData2.Row = j 
 			           frm1.vspdData2.Col =  C_ParentRowNo
 			           frm1.vspdData2.text = i
 				           
 			        End If
					lglngHiddenRows(i-1) = DataLASTRow(i)
				 
				
				
 			    Next 
 			  frm1.vspdData.Col =  frm1.vspdData.MaxCols
			  frm1.vspdData.text = i
		    '2006-10 hong  �ص� Ŭ���� ���� ���� ����*************
 			   
 		    Next 	  
 		    
 		 frm1.vspdData.vspdData.ReDraw = True
 	    
 	    
	Else
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
 	Dim i,k
 	Dim strFlag,strFlag1
 	gMouseClickStatus = "SP2C"   

 	Set gActiveSpdSheet = frm1.vspdData2

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
	End If

 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2
 		strShowDataFirstRow = Clng(ShowDataFirstRow)
 		strShowDataLastRow = Clng(ShowDataLastRow)
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, strShowDataFirstRow, strShowDataLastRow	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, strShowDataFirstRow, strShowDataLastRow	'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
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
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData2_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
        
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###

	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
'	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")  'Lock  Suitable  Field
	Call InitSpreadSheet 
	Call InitSpreadSheet2
	Call InitVariables
	Call SetDefaultVal
	set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
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
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo	
	Dim lRow
	Dim i
	Dim strFlag
	Dim strParentRowNo

    ggoSpread.Source = gActiveSpdSheet
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet
		Call ggoSpread.ReOrderingSpreadData
		
    ElseIf gActiveSpdSheet.Name = "vspdData2" Then
		For i = 1 To frm1.vspdData2.MaxRows
			frm1.vspdData2.Row = i
			frm1.vspdData2.Col = 0
			strFlag = frm1.vspdData2.Text
			If strFlag = ggoSpread.InsertFlag Then
				frm1.vspdData2.Col = C_ParentRowNo
				strParentRowNo = CInt(frm1.vspdData2.Text)
				lglngHiddenRows(strParentRowNo - 1) = CInt(lglngHiddenRows(strParentRowNo - 1)) - 1
			End If
		Next

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet2
		frm1.vspdData2.Redraw = False
		
		Call ggoSpread.ReOrderingSpreadData("F")
		
		Call DbQuery2(frm1.vspdData.ActiveRow,False)
		
		lngRangeFrom = Clng(ShowDataFirstRow)
		lngRangeTo = Clng(ShowDataLastRow)
		
		lRow = frm1.vspdData.ActiveRow	'###�׸��� ������ ���Ǻκ�###
		frm1.vspdData2.Redraw = True
    End If
    
 	'------ Developer Coding part (Start)	
 	'------ Developer Coding part (End) 	
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	'###�׸��� ������ ���Ǻκ�###
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If

	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - START */
	Dim lRow
	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - END */
	
	Set gActiveSpdSheet = frm1.vspdData

	frm1.vspdData.redraw = false
	If Row <> NewRow And NewRow > 0 Then
		With frm1        
			.vspdData.redraw = false
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - START */
		'	If DefaultCheck = False Then
		'		.vspdData.Row = Row
		'		.vspdData.Col = 1
		'		.vspdData2.focus
    	'		Exit Sub
		'	End If
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - END */
			
			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - START */
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = 1
				Exit Sub
			End If
			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - END */
			lgCurrRow = NewRow	
			.vspdData.redraw = true
		End With

		lgIntFlgModeM = Parent.OPMD_CMODE
	
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
	frm1.vspdData.redraw = true
End Sub

'=======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
		
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	'/��û���� ����Ǹ� ��η��� �����Ѵ�.(��û�� * ��κ���)
	.Row = Row
	Select Case col
		Case C_ReqQty,C_ORGCd
		    ggoSpread.UpdateRow Row
			.Row = Row
			.Col = C_ReqQty
			Call ReqQty_Change(Col, Row, UniCdbl(.text))
			.Row = Row
			.Col = C_CfmFlg		
			.value="1"
		Case C_Unit,C_DlvyDt
		    ggoSpread.UpdateRow Row
			.Row = Row
			.Col = C_CfmFlg		
			.value="1"
	End Select
	
	
    End With
    
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
		.Col = C_Flag 
		.Text = strMark
		
		Select Case col
	        Case C_PlanDt 
				 .Row = Row
				 .Col = Col
				 If UniConvDateToYYYYMMDD(frm1.vspdData2.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(CurrDate,parent.gDateFormat,"") Then 
				     Call sprRedComColor(C_PlanDt,Row,Row)
				 else
				     Call sprBlackComColor(C_PlanDt,Row,Row)
				 End If
	        Case C_SpplCd 
	             Call SpplChange()
	        Case C_Quota_Rate
	             Call ApportionQtyChange(Row)   
	    end select
	    
    End With
   
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow iparentrow
	
	With frm1.vspdData
		.Row = iparentrow
		.Col = C_CfmFlg
		If .value = 0 then
			.value = 1
		End If
    End With
    
End Sub	

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
  
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
		If lgPageNo <> "" Then			'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DbQuery = False Then
				Exit Sub
			End If
		End If	
		
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    Dim lRow
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    With frm1
		 
    	lRow = .vspdData.ActiveRow
    	'/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    	If ShowDataLastRow < NewTop + VisibleRowCnt(.vspdData2, NewTop) Then	        '��: ������ üũ 
		'/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
'    		If lgStrPrevKeyM(lRow - 1) <> "" Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
    		If lgPageNo1 <> "" Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If
				
				Call DisableToolBar(Parent.TBC_QUERY)
				If DbQuery2(lRow, True) = False Then
					Call RestoreToolBar()
					Exit Sub
				End If
			End If
		End If
    End With
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim index
    Dim intSeq
    
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData 
	If Col = C_CfmFlg And Row > 0 Then
		frm1.vspdData.Redraw = false
		
		.Col = C_CfmFlg
		.Row = Row
		if Row <= 0 Then Exit Sub
	    If Trim(.value)="1" Then
			ggoSpread.UpdateRow Row
	    Else
			.Col  = 0
			.Row  = Row
			.text = ""			
	    End If

		frm1.vspdData.Redraw = true
    ElseIf Row > 0 And Col = C_UnitPopup Then       '���� 
        
        .Col = Col
        .Row = Row
        Call OpenUnit()
        
		Elseif Row > 0 And Col = C_ORGCdPopup Then
			Call OpenSORG()
    End If    			    
	End With
	
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData2
	 
		ggoSpread.Source = frm1.vspdData2
		  
		If Row > 0 And Col = C_SpplPopup Then
			Call OpenSSupplier()		
		Elseif Row > 0 And Col = C_GrpPopup Then
			Call OpenSGrp()
		End If 
			    
	End With
End Sub

'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtDlvyFrDt
'   Event Desc : �ʿ��� 
'==========================================================================================
Sub txtDlvyFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtDlvyFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc : �ʿ��� 
'==========================================================================================
Sub txtDlvyToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtDlvyToDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtReqFrDt
'   Event Desc : ��û�� 
'==========================================================================================
 Sub txtReqFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReqFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtReqToDt
'   Event Desc : ��û�� 
'==========================================================================================
 Sub txtReqToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReqToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqToDt.Focus
	End If
End Sub


'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtDlvyFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtDlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###�׸��� ������ ���Ǻκ�###
    FncQuery = False                                                        
	
    Dim IntRetCD     
    '-----------------------
    'Check previous data area
    '-----------------------
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If
 
 	with frm1
		if (UniConvDateToYYYYMMDD(.txtReqFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtReqToDt.text,Parent.gDateFormat,"")) and Trim(.txtReqFrDt.text)<>"" and Trim(.txtReqToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","��û��", "X")			
			Exit Function
		End If   
		
		if (UniConvDateToYYYYMMDD(.txtDlvyFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtDlvyToDt.text,Parent.gDateFormat,"")) and Trim(.txtDlvyFrDt.text)<>"" and Trim(.txtDlvyToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")			
			Exit Function
		End If   
		
	End with
	
	'-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If																		'��: Query db data
          
	Set gActiveElement = document.activeElement
    FncQuery = True	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    FncNew = False                                                          
    
    Dim IntRetCD 
    
	'-----------------------
    'Check previous data area
    '----------------------- 
    If ChangeCheck = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")      
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables                                                      'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call SetDefaultVal
    
	Set gActiveElement = document.activeElement
    FncNew = True
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncDelete() 
    Dim lDelRows, iSelRow1, iSelRow2
    Dim iDelRowCnt, i

	if frm1.vspdData.Maxrows < 1 then exit function
    
    With frm1.vspdData
    	.focus
		ggoSpread.Source = frm1.vspdData
        ggoSpread.DeleteRow
		
    End With    
    
	Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    FncSave = False                                                         
    
    Dim IntRetCD 

	'-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ START
	If DefaultCheck = False Then
		Exit Function
	End If
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ END

'	intRetCd = DisplayMsgBox("900018", VB_YES_NO, "X", "X")   '�� �ٲ�Eκ?
'	If intRetCd = VBNO Then
'		Exit Function
'	End IF


    '-----------------------
    'Check content area
    '-----------------------
'    If Not chkField(Document, "1") Then 
'       		Exit Function
'    End If
   
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
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy() 
	FncCopy = false
	
	Dim IntRetCD
	Dim lRow
	Dim lRow2
	Dim totalQuotaRate,totalApportionQty
	Dim iQuotaRate,iApportionQty,iReqQty

	'ȯ�漳������ ��ü�ڵ������̸� ��ü�� �ڵ��������� �������� ��ü�߰��Ѱű��� �߰��Ϸ��� �õ��ϱ⶧���� 
	'�����߻�. ���� ��ü�ڵ������̸� �������� ��ü�߰����� ���ϵ��� ���� 200309
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strMinorCd

	Call CommonQueryRs(" minor_cd ", " b_configuration ", " major_cd = " & FilterVar("M2104", "''", "S") & " and reference = " & FilterVar("Y", "''", "S") & "  and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    strMinorCd = Split(lgF0, Chr(11))
    	
    If 	Trim(UCase(strMinorCd(0))) = "A" then 
		Call DisplayMsgBox("172143","X", "X","X")
		exit function
	End if
			
	With frm1
		'Check Spread2 Data Exists for the keys
		If CheckDataExist = False Then
			Exit function
		End If
    		
		.vspdData2.ReDraw = False
		
		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow
		
		lRow2 = .vspdData2.ActiveRow
		.vspdData2.Row = lRow2
		
		.vspdData2.Col = C_SpplCd
		.vspdData2.Text = ""
		
		.vspdData2.Col = C_SpplNm
		.vspdData2.Text = ""
		
		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = 0
		
	    .vspdData2.Col = C_ApportionQty
		.vspdData2.Text = 0
		
		.vspdData2.Col = C_Flag
		.vspdData2.Text = ggoSpread.InsertFlag

		Call SetSpreadColor2(lRow2,lRow2)
	    
	    lRow = .vspdData.ActiveRow
	    .vspdData.Row = lRow
	    .vspdData.Col = C_ReqQty
        iReqQty = Unicdbl(.vspdData.text)
	
		'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
		lglngHiddenRows(lRow - 1) = lglngHiddenRows(lRow - 1) + 1
	
	    Dim i 
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k

		lngRangeFrom = ShowDataFirstRow()
		lngRangeTo = ShowDataLastRow()

		k = 0
		totalQuotaRate = 0
		totalApportionQty     = 0


		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			.vspdData2.Row = i
			.vspdData2.Col = 0
			strFlag = .vspdData2.Text
			
			if strFlag <> ggoSpread.DeleteFlag then
			    .vspdData2.Col = C_Quota_Rate
			    totalQuotaRate = totalQuotaRate + Unicdbl(.vspdData2.Text)
			    .vspdData2.Col = C_ApportionQty 
			    totalApportionQty     = totalApportionQty     + Unicdbl(.vspdData2.Text)
		    End If
		Next
		
		iQuotaRate = 100 - totalQuotaRate
		iApportionQty     = iReqQty - totalApportionQty

        if iQuotaRate < 0 then iQuotaRate = 0
        if iApportionQty     < 0 then iApportionQty     = 0

		.vspdData2.Row = lRow2  
		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = UNIFormatNumber(iQuotaRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		
		.vspdData2.Col = C_ApportionQty
    	.vspdData2.Text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		                  
		.vspdData.Row = lRow
		.vspdData.Col = C_CfmFlg
		.vspdData.value = 1	
	
		.vspdData2.ReDraw = True
		.vspdData2.Action = 0
		.vspdData2.focus
	End With
	Set gActiveElement = document.activeElement
	FncCopy = true
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
		Exit function
	End If

	If gActiveSpdSheet.ID = "B" Then

		'Check Spread2 Data Exists for the keys
		If CheckDataExist = False Then
			Exit function
		End If

		ggoSpread.Source = frm1.vspdData2	
		With frm1.vspdData2
			lngRangeFrom = ShowDataFirstRow()
			lngRangeTo = ShowDataLastRow()
			.Redraw = False
			ggoSpread.EditUndo                                                  '��: Protect system from crashing
			Call checkdt(.ActiveRow)
			If lngRangeFrom > 0 Then
				iCnt=1
				lngRangeFrom = ShowDataFirstRow()
				lngRangeTo = ShowDataLastRow()
				For k=lngRangeFrom To lngRangeTo
					.Row=k
					.col=0
					if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
					iCnt = iCnt + 1
				Next
			End If
			.Redraw = True
		End With
	Else

		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo                                                  '��: Protect system from crashing

		if frm1.vspdData2.maxrowS > 0 Then 
			ggoSpread.Source = frm1.vspdData2	
			With frm1.vspdData2
				.Redraw = False
				
				lngRangeFrom = ShowDataFirstRow()
				lngRangeTo = ShowDataLastRow()
				If lngRangeFrom > 0 Then
					For k=lngRangeFrom to lngRangeTo
						.Row=k
						ggoSpread.EditUndo k                                                 '��: Protect system from crashing
						Call checkdt(k)
					Next
					lngRangeFrom = ShowDataFirstRow()
					lngRangeTo = ShowDataLastRow()
					For k=lngRangeFrom To lngRangeTo
						.Row=k-1
						.col=0
						if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
						iCnt = iCnt + 1
					Next
				End If

				.Redraw = True
			End WIth	
		End If
	End If

	lRow = frm1.vspdData.ActiveRow
	lngRangeFrom = ShowDataFirstRow()
	lngRangeTo = ShowDataLastRow()
	If lngRangeTo = 0 Then
		lglngHiddenRows(lRow - 1) = 0
	Else
		lglngHiddenRows(lRow - 1) = lngRangeTo - lngRangeFrom + 1
	End If

	k = 0 
	If lngRangeFrom > 0 Then
		for i = lngRangeFrom to lngRangeTo
		    frm1.vspdData2.Row = i 
		    frm1.vspdData2.Col = 0
		    strFlag = Trim(frm1.vspdData2.Text)
		    If strFlag = ggoSpread.InsertFlag or strFlag = ggoSpread.UpdateFlag or strFlag = ggoSpread.DeleteFlag then 
		        k = 1
		        Exit for
		    End If
		next 
	End If

	if k = 0 then
	    frm1.vspdData.Row = lRow
	    frm1.vspdData.Col = C_CfmFlg
	    frm1.vspdData.value = 0
	End If
	
	Set gActiveElement = document.activeElement
	FncCancel = true
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)	'###�׸��� ������ ���Ǻκ�###
	FncInsertRow = false

	On Error Resume Next
	
	Err.Clear
	
	'ȯ�漳������ ��ü�ڵ������̸� ��ü�� �ڵ��������� �������� ��ü�߰��Ѱű��� �߰��Ϸ��� �õ��ϱ⶧���� 
	'�����߻�. ���� ��ü�ڵ������̸� �������� ��ü�߰����� ���ϵ��� ���� 200309
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strMinorCd

	Call CommonQueryRs(" minor_cd ", " b_configuration ", " major_cd = " & FilterVar("M2104", "''", "S") & " and reference = " & FilterVar("Y", "''", "S") & "  and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    strMinorCd = Split(lgF0, Chr(11))
    	
    If 	Trim(UCase(strMinorCd(0))) = "A" then 
		Call DisplayMsgBox("172143","X", "X","X")
		exit function
	End if
	
	Dim lRow
	Dim lRow2
	Dim lconvRow
	Dim strMark
	Dim iInsertRow
	Dim IntRetCD
	Dim imRow
	Dim strInspUnitIndctnCd
	Dim iparentprno,iparentrow
	Dim totalQuotaRate,totalApportionQty
	Dim iQuotaRate,iApportionQty,iReqQty
	Dim iStrDlvyDt
	
	With frm1
		If .vspdData.MaxRows <= 0 Then
			Exit Function
		End If
		
		.vspdData2.ReDraw = False		
		
		If IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		Else
			imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
				Exit Function
			End If
		End If			
		
		'Insert Row in Spread2
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
		
		lRow = .vspdData.ActiveRow
		.vspdData.Row = lRow
		.vspdData.Col = .vspdData.MaxCols
		lconvRow = CInt(.vspdData.Text)
        
        .vspdData.Col = C_ReqNo
        iparentprno = .vspdData.value

        .vspdData.Col = C_ReqQty
        iReqQty = Unicdbl(.vspdData.text)
		
        .vspdData.Col = C_DlvyDt
        iStrDlvyDt = UNIConvDate(Trim(.vspdData.Text))

		For iInsertRow = 0 To imRow - 1			
			lRow2 = .vspdData2.ActiveRow + iInsertRow

			.vspdData2.Row = lRow2
			.vspdData2.Col = 0
			strMark = .vspdData2.Text

			.vspdData2.Col = C_Flag 
			.vspdData2.Text = strMark
		
			.vspdData2.Col = C_ParentRowNo
			.vspdData2.Text = lconvRow

			.vspdData2.Col = C_ParentPrNo
			.vspdData2.value = iparentprno		

			.vspdData2.Col = C_PlanDt
			.vspdData2.text = UniConvDateAToB(iStrDlvyDt, parent.gServerDateFormat, parent.gDateFormat)

			'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
			lglngHiddenRows(lconvRow - 1) = CInt(lglngHiddenRows(lconvRow - 1)) + 1
			Call SetSpreadColor2(lRow2, lRow2)
		Next
			
		'/* ���� : ����� �� �ѹ��� ���� �߰� START */
		Dim i 
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k
		
		lngRangeFrom = ShowDataFirstRow()
		lngRangeTo = ShowDataLastRow()
		k = 0
		totalQuotaRate = 0
		totalApportionQty     = 0

		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			.vspdData2.Row = i
			.vspdData2.Col = 0
			strFlag = .vspdData2.Text
			If strFlag <> ggoSpread.InsertFlag and strFlag <> ggoSpread.UpdateFlag and strFlag <> ggoSpread.DeleteFlag then
				.vspdData2.Text = CStr(k)
			End If
			if strFlag <> ggoSpread.DeleteFlag then
			    .vspdData2.Col = C_Quota_Rate
			    totalQuotaRate = totalQuotaRate + Unicdbl(.vspdData2.Text)
			    .vspdData2.Col = C_ApportionQty 
			    totalApportionQty     = totalApportionQty     + Unicdbl(.vspdData2.Text)
		    End If 
		Next
		
		iQuotaRate = 100 - totalQuotaRate
		iApportionQty     = iReqQty - totalApportionQty

        if iQuotaRate < 0 then iQuotaRate = 0
        if iApportionQty     < 0 then iApportionQty     = 0

		.vspdData2.Row = lRow2  
		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = UNIFormatNumber(iQuotaRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		
		.vspdData2.Col = C_ApportionQty
    	.vspdData2.Text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		                  
		.vspdData.Row = lRow
		.vspdData.Col = C_CfmFlg
		.vspdData.value = 1
		
		'/* ���� END */
		.vspdData2.Action = 0
		.vspdData2.focus
		.vspdData2.ReDraw = True
	End With
	FncInsertRow = true
	
	Set gActiveElement = document.activeElement
	Call SetSpreadLock()
	
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()		'###�׸��� ������ ���Ǻκ�###
	FncDeleteRow = false
	
	Dim lDelRows
	Dim iDelRowCnt, i,j
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim iparentrow
    
	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End If
		
	'Check Spread2 Data Exists for the keys
'	If CheckDataExist = False Then
'		Exit function
'	End If

	If gActiveSpdSheet.ID = "B" Then
	
		With frm1.vspdData2
			.Redraw = False
			.Focus
		    lngRangeFrom = .SelBlockRow
		    .Row = lngRangeFrom
			lngRangeFrom = ShowDataFirstRow()
			lngRangeTo = ShowDataLastRow()
			
		    ggoSpread.Source = frm1.vspdData2 
		     '----------  Coding part  -------------------------------------------------------------   
			.Row = lngRangeFrom
			lDelRows = ggoSpread.DeleteRow
			.Col = C_ParentRowNo
			iparentrow = .text 
			
			frm1.vspddata.row = iparentrow
			frm1.vspddata.Col = C_CfmFlg
			frm1.vspddata.value=1

			.Redraw = True
		End With
	Else
		With frm1.vspdData
			.Redraw = False
			.row = .activerow	
			ggoSpread.Source = frm1.vspdData 
			ggoSpread.DeleteRow
			
			for   i =0 to frm1.vspddata.maxrows 
			
				  frm1.vspdData.row = i
				  frm1.vspdData.Col = 0
		
				If frm1.vspddata.text =ggoSpread.DeleteFlag then
				 .Col = C_CfmFlg
				 .value=1
				 
						With frm1.vspdData2
							.Redraw = False
							.Focus
							'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START	    
						    lngRangeFrom = .SelBlockRow
						    .Row = lngRangeFrom
							lngRangeFrom = ShowDataFirstRow()
							lngRangeTo = .SelBlockRow2
							.Row = lngRangeTo
							lngRangeTo = ShowDataLastRow()

						    ggoSpread.Source = frm1.vspdData2 
							For j=lngRangeFrom To lngRangeTo
						     '----------  Coding part  -------------------------------------------------------------   
									.Row = j
								If .RowHidden = False Then
									ggoSpread.DeleteRow j

									.Row = lngRangeFrom
									.Col = C_ParentRowNo
									iparentrow = .text 
								End If
							Next
							
							.Redraw = True
						End With
		
		
				 
				 
				     
				End if
				
			next 
			
			

		
		End WIth
	End If
		
	Set gActiveElement = document.activeElement
	FncDeleteRow = true
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
	Call Parent.FncPrint()
	Set gActiveElement = document.activeElement
	FncPrint = True
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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
	Set gActiveElement = document.activeElement
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
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

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	DbQuery = False                                                             
	
	Dim strVal
	
	Call LayerShowHide(1)
	
	with frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    
	    strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlant.value)            '���� 
	    strVal = strVal & "&txtItemCd=" & Trim(.hdnItem.value)              'ǰ�� 
	    strVal = strVal & "&txtDlvyFrDt=" & Trim(.hdnDFrDt.value)           '��û�� 
		strVal = strVal & "&txtDlvyToDt=" & Trim(.hdnDToDt.value)
		strVal = strVal & "&txtReqFrDt=" & Trim(.hdnRFrDt.value)            '�ʿ��� 
		strVal = strVal & "&txtReqToDt=" & Trim(.hdnRToDt.value)
	    strVal = strVal & "&txtPrTypeCd=" & Trim(.hdnPrTypeCd.value)  '��û���� 
	    strVal = strVal & "&txtMRP=" & Trim(.hdnMrp.value)                  'mrp run ��ȣ 
	    strVal = strVal & "&txtTrackingNo=" & Trim(.hdnTrackingNo.value)		'200309
	    strVal = strVal & "&lgPageNo=" & lgPageNo                  '��: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
		strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
		strVal = strVal & "&txtReqFrDt=" & Trim(.txtReqFrDt.text)
		strVal = strVal & "&txtReqToDt=" & Trim(.txtReqToDt.text)
	    strVal = strVal & "&txtPrTypeCd=" & Trim(.txtPrTypeCd.value)  
	    strVal = strVal & "&txtMRP=" & Trim(.txtMRP.value)
	    strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)		'200309
	    strVal = strVal & "&lgPageNo=" & lgPageNo                  '��: Next key tag
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End If
		
	End with

	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)
	DbQueryOk = False

	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows
	
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("11001111001011")				'��ư ���� ���� 
			
	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows

		i=0
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then 
				ReDim lgStrPrevKeyM(intARow - 1)	
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows
				
				ReDim lgStrPrevKeyM(intTRow+intARow - 1)	
				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
				For i = 0 To intTRow-1
					lgStrPrevKeyM(i) = TmpArrPrevKey(i)
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next 
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next 

			if lgIntFlgModeM = Parent.OPMD_CMODE then
			    If DbQuery2(1, False) = False Then	Exit Function
		    End If
		    lgIntFlgModeM = Parent.OPMD_UMODE
		    lgIntFlgMode = Parent.OPMD_UMODE
		End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
    DbQueryOk = true
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	DbQuery2 = False
	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim pRow
		
	'/* 9�� ������ġ: ���� ���������� �ణ �̵� �� �̹� ��ȸ�� �ڷᳪ �Էµ� �ڷḦ �о� ���� ������ '' â ���� - START */
	Call LayerShowHide(1)
	
	With frm1
		.vspdData.redraw = false
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		pRow  = CInt(.vspdData.Text)	

		If lglngHiddenRows(pRow - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			lngRet = ShowFromData(pRow, lglngHiddenRows(pRow - 1))	'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.
			Call SetToolBar("11001111001011")				'��ư ���� ���� 
			Call LayerShowHide(0)
			.vspdData2.ReDraw = True
			DbQuery2 = True
			.vspdData.redraw = True
			Exit Function
		End If
		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		.vspdData.Row = Row
		.vspdData.Col = C_ReqNo		    

		strVal = strVal & "&txtPrNo=" & Trim(.vspdData.text)
		strVal = strVal & "&lgStrPrevKeyM=" & lgStrPrevKeyM(Row - 1)		    
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag 
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(pRow)
	
		.vspdData.redraw = True

	End With
	Call RunMyBizASP(MyBizASP, strVal)
	DbQuery2 = True
End Function

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
		lngRangeFrom = .MaxRows - DataCount + 1
		lngRangeTo = .MaxRows
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_Flag
		
		.Col2 = C_Flag
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With
	
	For Index = lngRangeFrom to lngRangeTo
    	frm1.vspdData2.Row = Index
    	Call checkdt(Index)
    	If Index = lngRangeTo Then
				frm1.vspdData2.Row = Index
				frm1.vspdData2.Col = 1
				frm1.vspdData2.Action = 0
				frm1.vspdData2.focus
		End if    		
	Next

	DbQueryOk2 = true

End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	DbSave = False                                                          '��: Processing is NG
	Dim lRow
	Dim lGrpCnt     
	Dim strVal,strIU, strDel
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim parentRow
	Dim iReqQty,totalQty,totalRate
	Dim lgTransSep
	Dim lgHdDtlSep
	Dim strValUp, strReqNo, strDlvyDt, strModifyChk, iRowMode 
	Dim iStrPurOrg
	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	Dim iColSep,iRowSep
	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size

	Dim intRetCd
	
	Dim chknum
	chknum = 0
	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002
	End With	    

	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 1
	strVal = ""
    strDel = ""
    strIU  = ""
    lgTransSep = "��"
    lgHdDtlSep = "��"
    iRowMode = ""

	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	'-----------------------
	'Data manipulate area
	'-----------------------
	With frm1
	    For parentRow = 1 To .vspdData.MaxRows
		    iStrPurOrg=""
			If Trim(GetSpreadValue(.vspdData,C_CfmFlg,parentRow,"X","X")) = 1 Then
			    lngRangeFrom = DataFirstRow(parentRow)
			    lngRangeTo   = DataLastRow(parentRow)
			    
			    iReqQty = Unicdbl(GetSpreadText(.vspdData,C_ReqQty,parentRow,"X","X"))
				strReqNo = Trim(GetSpreadText(.vspdData,C_ReqNo,parentRow,"X","X"))
				strDlvyDt = UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDt,parentRow,"X","X")))
				iStrPurOrg=Trim(GetSpreadText(.vspdData,C_ORGCd,parentRow,"X","X"))
			   
			    '-----��� �������� ���� ��´�. -------------------------
			    iRowMode = Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
			    If iRowMode = ggoSpread.UpdateFlag Then
					strValUp = "UPDATE" & iColSep
				ElseIf iRowMode = ggoSpread.DeleteFlag Then
					strValUp = "DELETE" & iColSep
				End If
					    
				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_ReqNo,parentRow,"X","X")) & iColSep
				If Trim(GetSpreadText(.vspdData,C_ReqQty,parentRow,"X","X"))="" Then
					strValUp = strValUp & "0" & iColSep
				Else
					strValUp = strValUp & UNIConvNum(Trim(GetSpreadText(.vspdData,C_ReqQty,parentRow,"X","X")),0) & iColSep
				End If
					       
				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_Unit,parentRow,"X","X")) & iColSep

				If iRowMode = ggoSpread.UpdateFlag AND _
					CDate(UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDt,parentRow,"X","X")))) < CDate(UNIConvDate(Trim(CurrDate))) Then
				    Call DisplayMsgBox("172120","X", parentRow & "�� ","X")	
					Call LayerShowHide(0)
					Call RemovedivTextArea
'msg modify 20040506 by kjt
					.vspdData.Row = ParentRow
					.vspdData.Col = C_DlvyDt
					.vspdData.Action = 0		'Active Cell
					Call DbQuery2(ParentRow,False) 
					Exit Function
				End If
'' 2004 04 13 update by kjt
				If iRowMode = ggoSpread.DeleteFlag Then
					For lRow = lngRangeFrom To lngRangeTo
						frm1.vspddata2.Row = lRow
						frm1.vspddata2.Col = 0
						if frm1.vspddata2.text = ggoSpread.InsertFlag Then
							intRetCd = DisplayMsgBox("900038", Parent.VB_YES_NO, "X", "X")   '�� �ٲ�Eκ?
							If intRetCd = VBNO Then
								Call LayerShowHide(0)
								frm1.vspdData.Row = parentRow
								frm1.vspdData.Col = 0
								frm1.vspdData.text = ggoSpread.UpdateFlag
								frm1.vspdData.Col = 1
								frm1.vspdData.Action = 0
								frm1.vspdData.focus
								Call DbQuery2(parentRow,False)								
								Exit Function
							End IF
						End if
					Next
				End If

				strValUp = strValUp & strDlvyDt & iColSep
				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_ORGCd,parentRow,"X","X")) & iColSep
				strValUp = strValUp & parentRow & lgHdDtlSep	'7 ���� 
				strVal = strValUp
				'----------------------------------------------------------
			    totalQty  = 0
			    totalRate = 0

			    If lngRangeTo > 0 Then
					For lRow = lngRangeFrom To lngRangeTo
						chknum = chknum + 1
						If CheckDuplSppl(lRow) = False Then
							DbSave = False
							Call LayerShowHide(0)
							Call RemovedivTextArea
							Exit Function
						End If
						.vspddata2.row = lRow
						.vspddata2.col = C_SpplCd
					    If Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) <> ggoSpread.DeleteFlag Then
							totalQty = totalQty + Unicdbl(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))
							totalRate = totalRate + Unicdbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))
					    End If
					    
					    Select Case GetSpreadText(.vspdData2,0,lRow,"X","X")
						
							Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
							    If GetSpreadText(.vspdData2,0,lRow,"X","X")=ggoSpread.InsertFlag then
									strIU = strIU & "C" & iColSep	
								Else
									strIU = strIU & "U" & iColSep
								End If    
									
					            strIU = strIU & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep

							    If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
									strIU = strIU & "0" & iColSep
								Else
									strIU = strIU & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
								End If
					    
								If Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))="" Then
									strIU = strIU & "0" & iColSep
								Else
									strIU = strIU & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X")),0) & iColSep
								End If

								If CDate(UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X")))) < CDate(UNIConvDate(Trim(CurrDate))) Then
								    Call DisplayMsgBox("172140","X", strReqNo & " - " & chknum  & chr(32) & " �� ","X")
								    Call LayerShowHide(0)
								    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = 1
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = lRow 
								    .vspdData2.Col = C_PlanDt
									.vspdData2.Action = 0		'Active Cell
								    Exit Function
								End If
								
								If CDate(UNIConvDate(Trim(strDlvyDt))) <  CDate(UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X")))) Then
								    Call DisplayMsgBox("172125","X", strReqNo & " - " & chknum  & chr(32) & " �� ","X")	
								    Call LayerShowHide(0)
								    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = C_DlvyDt
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = lRow 
								    .vspdData2.Col = C_PlanDt
									.vspdData2.Action = 0		'Active Cell
								    Exit Function
								End If
								strIU = strIU & UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X"))) & iColSep
								strIU = strIU & Trim(iStrPurOrg) & iColSep
								strIU = strIU & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
								strIU = strIU & "" & iColSep
								strIU = strIU & parentRow & iRowSep
									
							Case ggoSpread.DeleteFlag				'��: ���� 
								strDel = strDel & "D" & iColSep			'��: D=Delete
					            strDel = strDel & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep
          
							    If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
									strDel = strDel & "0" & iColSep
								Else
									strDel = strDel & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
								End If
					    
								If Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))="" Then
									strDel = strDel & "0" & iColSep
								Else
									strDel = strDel & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X")),0) & iColSep
								End If
									
								strDel = strDel & UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X"))) & iColSep
								strDel = strDel & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
								strDel = strDel & parentRow & iRowSep
					                
						End Select
						
					Next
				Else
					totalRate=100
					totalQty=iReqQty
				End If
							    
			   	If iRowMode = ggoSpread.UpdateFlag Then
			   		If totalRate <> 100 Then
					    Call DisplayMsgBox("171325", "X", parentRow & "�� ", "X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = 1
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = 1
								    .vspdData2.Col = C_Quota_Rate
									.vspdData2.Action = 0		'Active Cell
					    
					    Exit Function
					End If

			   		If totalQty <> iReqQty Then
					    Call DisplayMsgBox("172420","X",strReqNo, "X")	
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
					    Exit Function
					End If
				
				End If
							
				strVal =  strVal & strDel & strIU & lgTransSep			
				Select Case Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag,ggoSpread.DeleteFlag
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
			strDel  = ""
			strIU   = ""
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

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData2.MaxRows = 0

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If
 
 	with frm1
		if (UniConvDateToYYYYMMDD(.txtReqFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtReqToDt.text,Parent.gDateFormat,"")) and Trim(.txtReqFrDt.text)<>"" and Trim(.txtReqToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","��û��", "X")			
			Exit Function
		End If   
		
		if (UniConvDateToYYYYMMDD(.txtDlvyFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtDlvyToDt.text,Parent.gDateFormat,"")) and Trim(.txtDlvyFrDt.text)<>"" and Trim(.txtDlvyToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")			
			Exit Function
		End If   
		
	End with
	
	'-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If				
	
End Function

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : ���ó����� ��ư�� Ŭ���� ��� �߻� 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CfmFlg
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_CfmFlg, i, 0)
		Next	
	End If
End Sub

'======================================================================================================
' Area Name   : ReqQty_Change
' Description : ��û�� ������ ��η� �ڵ� ���� ó�� 
'=======================================================================================================
Sub ReqQty_Change(ByVal Col, ByVal Row,ByVal iReqQty)
	Dim i
	Dim strReqNo
	Dim strReqNo2
	Dim isExists
	
	isExists = False
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.Row = Row
		.Col = C_ReqNo
		strReqNo = .Text
	End With
	
	If frm1.vspdData2.Maxrows > 0 then
		Dim iSumReqQty
		Dim iQuotaRate 
		Dim iAppQty
		Dim strMark
		
		iSumReqQty = 0
				
		ggoSpread.Source = frm1.vspdData2
		With frm1.vspdData2
			
			For i = 1 to .Maxrows
				.Row = i
				.Col = C_ParentPrNo
				If strReqNo = Trim(.Text) Then
					isExists = True
					.Row = i
					.Col = 0 
					Select Case .Text					
						Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag, ggoSpread.DeleteFlag
							
						Case Else
							ggoSpread.UpdateRow i		
							
							.Row = i
							.Col = 0
							strMark = .Text
							.Col = C_Flag 
							.Text = strMark
					End Select
						
					.Row = i
					.Col = C_Quota_Rate
					iQuotaRate = Unicdbl(.Text)	'��κ��� 
					iAppQty = (iReqQty * iQuotaRate)/100	'��η� ��� 
					
					.Col = 0
					If Trim(.Text) <> ggoSpread.DeleteFlag  Then
						
						If iSumReqQty <= iReqQty Then	'�߰� �հ�� �� 
							
							If iSumReqQty+iAppQty >= iReqQty Then
								iAppQty= UNIFormatNumber((iReqQty-iSumReqQty),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
								iAppQty = UniCdbl(iAppQty)					
							End If
						Else	
							iAppQty = 0 
						End If
					Else
						iAppQty = 0
					End If
					
					.Col = C_ApportionQty
					.Text = UNIFormatNumber(iAppQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					iSumReqQty = iSumReqQty + UniCdbl(.Text)
					
					Call checkdt(i)
										
				End If
			Next
			
		End With
	End If
End Sub

'======================================================================================================
' Area Name   : CheckDuplSppl
' Description : ����ó �ߺ� üũ 
'=======================================================================================================
Function CheckDuplSppl(ByVal iChildRow)
	
	CheckDuplSppl = false
	
	Dim strVal
    Dim strssText1, strssText2
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim iparentrow
    Dim index 
    Dim iRow
	
	With frm1.vspdData2
	    .Row		= iChildRow    
		.Col		= C_ParentPrNo
		strssText1	= Trim(.text)	
		.Col		= C_SpplCd
		strssText2	= Trim(.text)
		.Col        = C_ParentRowNo
		iparentrow  = Trim(.text)
		If strssText2 = "" Then
			CheckDuplSppl = False
			Exit Function
		End If
				
	End with
	
	lngRangeFrom = DataFirstRow(iparentrow)
	lngRangeTo   = DataLastRow(iparentrow)

	For index = lngRangeFrom To lngRangeTo
	    If index <> iChildRow And strssText2 <> "" Then
	        frm1.vspdData2.Row = index     
	        frm1.vspdData2.Col = C_SpplCd
	        If UCase(strssText2) = UCase(Trim(frm1.vspdData2.text)) Then
				Call DisplayMsgBox("17A005","X",iChildRow,"����ó")	
                frm1.vspdData2.Row = iChildRow     
	            frm1.vspdData2.Col = C_SpplCd	  
	            frm1.vspdData2.text = ""      
	            CheckDuplSppl = False
 	            Exit Function
	        End If
	    End If
	Next 
	
	CheckDuplSppl = true
End Function

'======================================================================================================
' Area Name   : DeleteDownRowsAll
' Description : ��� Row ������ �ϴ��� ��� Row ���� 
'=======================================================================================================
Function DeleteDownRowsAll()
	
	Dim parentRow, lngRangeFrom, lngRangeTo, strTemp, index, strMark
	
		ggoSpread.Source = frm1.vspdData
		With frm1.vspdData
	
			.Row = parentRow
			.col = C_CfmFlg
			If .value = 0 Then
				.value = 1
			End If	
		End With

		ggoSpread.DeleteRow iSelRow1
				
		lngRangeFrom = DataFirstRow(parentRow)
		lngRangeTo   = DataLastRow(parentRow)
		
		frm1.vspdData2.Redraw = False
		ggoSpread.Source = frm1.vspdData2
		
		With frm1.vspdData2		
	
		For index = lngRangeFrom To lngRangeTo
			ggoSpread.DeleteRow index
		    .Row = index
		    .Col = 0
		    strMark = .Text
		    
			.Col = C_Flag 
			.Text = strMark
		Next
	
		.Redraw = True
	
		End With
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################## -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ſ�û�������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;<label id="lblT" name="lblTest"></label></TD>
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
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemcd" SIZE=10 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
													   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm" SIZE=20 tag="14X"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>��û��</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="��û��" NAME="txtReqFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="��û��" NAME="txtReqToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtDlvyFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtDlvyToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
							</TR>
							<TR><TD CLASS="TD5" NOWRAP>��û����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="��û����" NAME="txtPrTypeCd" SIZE=10 MAXLENGTH=18  MAXLENGTH=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrTypeCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPrTypeCd()">
													   <INPUT TYPE=TEXT NAME="txtPrTypeNm" SIZE=20 tag="14"></TD>
								</TD>
								<TD CLASS="TD5" NOWRAP>MRP Run��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="MRP Run��ȣ" NAME="txtMRP" SIZE=32 MAXLENGTH=12 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMrp"></TD>
							</TR>
							<TR><!--200309-->
								<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
								<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No." TYPE="Text" MAXLENGTH=26 SiZE=32  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
								<TD CLASS="TD5">
								<TD CLASS="TD6">
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
			<TR HEIGHT= 30%>
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
		</TABLE>
		
		</TD>
	</TR>
	
    <tr>
      <td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT="<%=BizSize%>"><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT="<%=BizSize%>" FRAMEBORDER="0" SCROLLING="NO" noresize framespacing="0"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnState" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
