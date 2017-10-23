<%@ LANGUAGE="VBSCRIPT" %>
<%
Response.Expires = -1
%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111ma4
'*  4. Program Name         : ���ſ�û���(��Ƽ)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/09/22
'*  9. Modifier (First)     : Kim Ji Hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                           
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID = "m2111mb4.asp"	
Const BIZ_PGM_JUMP_ID = "m2111qa1"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_ReqNo 																'��: Spread Sheet�� Column�� ��� 
Dim C_PlantCd		
Dim C_PlantPopUp	
'Dim C_PlantNm	= 
Dim C_ItemCd 		
Dim C_ItemPopUp	
Dim C_ItemNm		
Dim C_ItemSpec	
Dim C_ReqQty	    
Dim C_ReqUnit	    
Dim C_ReqUnitPopUp 
Dim C_DlvyDt 		
Dim C_ReqDt 		
Dim C_PurOrg 		
Dim C_PurOrgPopUp	
Dim C_DeptCd		
Dim C_DeptPopUp	
'Dim C_DeptNm		= 
Dim C_ReqPrsn		
Dim C_StorageCd 	
Dim C_StoragePopUp 
'Dim C_StorageNm 	= 
Dim C_Tracking	
Dim C_TrackingPopUp 
''Dim C_SpplCd	'20031008 ���� ������	
''Dim C_SpplPopUp	
'Dim C_SpplNm		= 
''Dim C_GrpCd 		
''Dim C_GrpPopUp	
'Dim C_GrpNm		= 
Dim C_PlanDt		
Dim C_ReqStateCd 	
Dim C_ReqStateNm 	
Dim C_ReqTypeCd 	
Dim C_ReqTypeNm 	

Dim C_HdnTrackingflg 
Dim C_HdnProcurType 
Dim C_HdnMrpNo 	

Dim C_OrderLT		
Dim C_DlvyLT		
Dim C_SpplCd	'20040525 ���� ������	
Dim lgPageNo
Dim lgNextKey
Const C_SHEETMAXROWS = 100
Dim lgSortKey


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/incSvrVariables.inc" -->

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim IsOpenPop          
Dim StartDate,EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################

Function changeDlvy()
	
	Dim  dlvy,orderlt,dlvylt
		
	with frm1.vspdData	
		.Row	= .ActiveRow    
		.Col	= C_DlvyDt
		dlvy	= .text	
		.Col	= C_OrderLT
		orderlt		= .text	
		.Col	= C_DlvyLT
		dlvylt		= .text	
		'.Col	= C_SpplCd
		'if Trim(.text) = "" and orderlt <> 0 then		
		'	.Col	= C_PlanDt
		'	.text	= uniDateAdd("d", CInt(unicdbl(orderlt)) * -1, dlvy, gDateFormat)
		'else
		'	if CInt(dlvylt) > 0 then
		'		.Col	= C_PlanDt
		'		.text	= uniDateAdd("d", CInt(unicdbl(dlvylt)) * -1, dlvy, gDateFormat)		
		'	else 
		'		.Col	= C_PlanDt
		'		.text	= ""		
		'	end if 
		'end if			
		
	End with
	
End Function


Function changeTagTracking()
	
	ggoSpread.Source = frm1.vspdData
	
	with frm1.vspdData	
		.Row		= .ActiveRow    
		.Col		= C_HdnTrackingflg
		
		if UCase(Trim(.text)) = "Y" then			
			ggoSpread.SpreadUnLock	C_Tracking , .ActiveRow, C_TrackingPopUp, .ActiveRow
			ggoSpread.SSSetRequired	C_Tracking, .ActiveRow, .ActiveRow	
		else
			ggoSpread.SpreadLock	C_Tracking , .ActiveRow, C_TrackingPopUp, .ActiveRow	
		end if
	
	End with

End Function


'========================================================================================
' Function Name : changeItemPlant()
' Function Desc : 
'========================================================================================
 Function changeItemPlant()
    Dim strRow 		
    If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If                               
    
    Dim strssText1, strssText2

	with frm1.vspdData
	
		.Row		= .ActiveRow  
		strRow 		= .ActiveRow   
		.Col		= C_ItemCd
		strssText1	= Trim(.text)		
		.Col		= C_PlantCd
		strssText2	= Trim(.text)
						
	End with
    
	if Trim(strssText2) = "" or Trim(strssText1) = "" then
		exit Function
	End if
    
	changeItemPlant = False                 
    
	If LayerShowHide(1) = False Then Exit Function
		    
	Dim strVal    
    
	strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemPlant"
	strVal = strVal & "&txtItemCd=" & Trim(strssText1)
	strVal = strVal & "&txtPlantCd=" & Trim(strssText2)
	strVal = strVal & "&txtRow=" & strRow 		

	Call RunMyBizASP(MyBizASP, strVal)
	
    changeItemPlant = True                  

End Function


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
    Dim strssText1, strssText2, strssText3

	with frm1.vspdData
	
		.Row		= .ActiveRow    
		.Col		= C_ReqNo
		strssText1	= Trim(.text)		
		'.Col		= C_SpplCd
		'strssText2	= Trim(.text)
		.Col		= C_DlvyDt
		strssText3	= Trim(.text)
				
	End with
	
	if Trim(strssText2) = "" or Trim(strssText3) = "" then
		exit Sub
	End if
		
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookSppl"	
    strVal = strVal & "&txtPrNo=" & strssText1
    strVal = strVal & "&txtBpCd=" & strssText2
        
    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)				
	
End Sub									

'========================================================================================
' Function Name : CookiePage
' Function Desc : 
'========================================================================================
Function WriteCookiePage()

	Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Call WriteCookie("m2111ma1_plantcd", frm1.txtPlantCd.Value)
	Call WriteCookie("m2111ma1_itemcd", frm1.txtItemCd.Value)
	
	Call PgmJump(BIZ_PGM_JUMP_ID)
	
End Function

Sub ReadCookiePage()

	Dim strTemp

	strTemp = ReadCookie("ReqNo")
	
	If strTemp = "" then Exit sub
	
	frm1.txtReqNo.value = ReadCookie("ReqNo")
	
	Call WriteCookie("ReqNo" , "")
	
	Call MainQuery()

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
    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    frm1.vspdData.MaxRows = 0
    lgPageNo         = ""
    lgNextKey	     = ""
    
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
	'frm1.txtORGCd.Value = gPurOrg
	frm1.txtPlantCd.Value = parent.gPlant
    frm1.txtItemCd.focus 

	Set gActiveElement = document.activeElement
	Call SetToolbar("1110111100101111")

	frm1.txtDlvyFrDt.Text	= StartDate
	frm1.txtDlvyToDt.Text	= EndDate
	frm1.txtReqFrDt.Text	= StartDate
	frm1.txtReqToDt.Text	= EndDate
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	
	C_ReqNo 		= 1
	C_PlantCd		= 2
	C_PlantPopUp	= 3
	C_ItemCd 		= 4
	C_ItemPopUp		= 5
	C_ItemNm		= 6
	C_ItemSpec		= 7
	C_ReqQty	    = 8
	C_ReqUnit	    = 9
	C_ReqUnitPopUp  = 10
	C_DlvyDt 		= 11
	C_ReqDt 		= 12
	C_PurOrg 		= 13
	C_PurOrgPopUp	= 14
	C_DeptCd		= 15
	C_DeptPopUp		= 16
	C_ReqPrsn		= 17
	C_StorageCd 	= 18
	C_StoragePopUp	= 19
	C_Tracking		= 20
	C_TrackingPopUp = 21
	C_PlanDt		= 22
	C_ReqStateCd 	= 23
	C_ReqStateNm 	= 24
	C_ReqTypeCd 	= 25
	C_ReqTypeNm 	= 26
	C_HdnTrackingflg = 27
	C_HdnProcurType = 28
	C_HdnMrpNo 		= 29
	C_OrderLT		= 30
	C_DlvyLT		= 31
	C_SpplCd		= 32
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

			C_ReqNo 		= iCurColumnPos(1)
			C_PlantCd		= iCurColumnPos(2)
			C_PlantPopUp	= iCurColumnPos(3)
			C_ItemCd 		= iCurColumnPos(4)
			C_ItemPopUp		= iCurColumnPos(5)
			C_ItemNm		= iCurColumnPos(6)
			C_ItemSpec		= iCurColumnPos(7)
			C_ReqQty	    = iCurColumnPos(8)
			C_ReqUnit	    = iCurColumnPos(9)
			C_ReqUnitPopUp  = iCurColumnPos(10)
			C_DlvyDt 		= iCurColumnPos(11)
			C_ReqDt 		= iCurColumnPos(12)
			C_PurOrg 		= iCurColumnPos(13)
			C_PurOrgPopUp	= iCurColumnPos(14)
			C_DeptCd		= iCurColumnPos(15)
			C_DeptPopUp		= iCurColumnPos(16)
			C_ReqPrsn		= iCurColumnPos(17)
			C_StorageCd 	= iCurColumnPos(18)
			C_StoragePopUp	= iCurColumnPos(19)
			C_Tracking		= iCurColumnPos(20)
			C_TrackingPopUp = iCurColumnPos(21)
			C_PlanDt		= iCurColumnPos(22)
			C_ReqStateCd 	= iCurColumnPos(23)
			C_ReqStateNm 	= iCurColumnPos(24)
			C_ReqTypeCd 	= iCurColumnPos(25)
			C_ReqTypeNm 	= iCurColumnPos(26)
			C_HdnTrackingflg = iCurColumnPos(27)
			C_HdnProcurType = iCurColumnPos(28)
			C_HdnMrpNo 		= iCurColumnPos(29)
			C_OrderLT		= iCurColumnPos(30)
			C_DlvyLT		= iCurColumnPos(31)
			C_SpplCd		= iCurColumnPos(32)
	End Select

End Sub	

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

 Sub InitSpreadSheet()

	Call InitSpreadPosVariables
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030923",,Parent.gAllowDragDropSpread  
	
	.ReDraw = false

    .MaxCols = C_DlvyLT + 1
	.Col = .MaxCols:	  .ColHidden = True
	
    Call GetSpreadColumnPos("A")

    .MaxRows = 0
    
    '.OperationMode = 5
    ggoSpread.SSSetEdit C_ReqNo, "��û��ȣ", 20,,,18,2
    ggoSpread.SSSetEdit C_PlantCd, "����", 10,,,4,2
    ggoSpread.SSSetButton 	C_PlantPopUp
    'ggoSpread.SSSetEdit C_PlantNm, "�����",20
    ggoSpread.SSSetEdit C_ItemCd, "ǰ��", 10,,,18,2
    ggoSpread.SSSetButton 	C_ItemPopUp
    ggoSpread.SSSetEdit C_ItemNm, "ǰ���",20,,,,2
	ggoSpread.SSSetEdit C_ItemSpec, "ǰ��԰�",20,,,,2
	SetSpreadFloat		C_ReqQty, "��û��", 15,1,3
	ggoSpread.SSSetEdit C_ReqUnit, "��û����",10,,,3,2
	ggoSpread.SSSetButton 	C_ReqUnitPopUp
    ggoSpread.SSSetDate C_DlvyDt, "�ʿ���", 10,2,gDateFormat
    ggoSpread.SSSetDate C_ReqDt, "��û��", 10,2,gDateFormat
    ggoSpread.SSSetEdit 	C_PurOrg,"��������",15,,,4,2
    ggoSpread.SSSetButton 	C_PurOrgPopUp
    ggoSpread.SSSetEdit C_DeptCd, "��û�μ�",10,,,10,2
	ggoSpread.SSSetButton 	C_DeptPopUp
    'ggoSpread.SSSetEdit C_DeptNm, "��û�μ���",20
	ggoSpread.SSSetEdit C_ReqPrsn, "��û��",20,,,,2
	ggoSpread.SSSetEdit C_StorageCd, "�԰�â��", 10,,,7,2
	ggoSpread.SSSetButton 	C_StoragePopUp
    'ggoSpread.SSSetEdit C_StorageNm, "�԰�â���", 20
	ggoSpread.SSSetEdit C_Tracking, "Tracking No.",15,,,25,2
	ggoSpread.SSSetButton 	C_TrackingPopUp
	'ggoSpread.SSSetEdit C_SpplCd, "����ó",10,,,10,2
	'ggoSpread.SSSetButton 	C_SpplPopUp
	'ggoSpread.SSSetEdit C_SpplNm, "����ó��",20
'	ggoSpread.SSSetEdit C_GrpCd, "���ű׷�",10,,,4,2
'	ggoSpread.SSSetButton 	C_GrpPopUp
	'ggoSpread.SSSetEdit C_GrpNm, "���ű׷��",20
	ggoSpread.SSSetDate C_PlanDt,"���ֿ�����", 15,2,gDateFormat
	ggoSpread.SSSetEdit C_ReqStateCd, "��û�������",15,,,5,2
	ggoSpread.SSSetEdit C_ReqStateNm, "��û������¸�",15,,,,2
	ggoSpread.SSSetEdit C_ReqTypeCd, "��û����",15,,,5,2
	ggoSpread.SSSetEdit C_ReqTypeNm, "��û���и�",15,,,,2
	ggoSpread.SSSetEdit C_HdnTrackingflg, "",5
	ggoSpread.SSSetEdit C_HdnProcurType, "",20
	ggoSpread.SSSetEdit C_HdnMrpNo, "",20
	ggoSpread.SSSetEdit C_OrderLT, "",5
	ggoSpread.SSSetEdit C_DlvyLT, "",5
	ggoSpread.SSSetEdit C_SpplCd, "",15
	Call ggoSpread.SSSetColHidden(C_HdnTrackingflg,C_SpplCd, True)
	
    Call SetSpreadLock()
	.ReDraw = true
	
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

 Sub SetSpreadLock()
   
    With frm1
    
		.vspdData.ReDraw = False
		ggoSpread.SpreadUnLock		C_PlantCd ,		-1,		C_PlantPopUp,		-1
		ggoSpread.SSSetRequired		C_PlantCd,		-1,		-1
		ggoSpread.SpreadUnLock		C_ItemCd ,		-1,		C_ItemPopUp,		-1
		ggoSpread.SSSetRequired		C_ItemCd,		-1,		-1
		ggoSpread.SSSetProtected	C_ItemNm,		-1,		-1
		ggoSpread.SSSetProtected	C_ItemSpec,		-1,		-1
		ggoSpread.SSSetRequired		C_ReqQty,		-1,		-1
		ggoSpread.SpreadUnLock		C_ReqUnit ,		-1,		C_ReqUnitPopUp,		-1
		ggoSpread.SSSetRequired		C_ReqUnit,		-1,		-1    
		ggoSpread.SSSetRequired		C_DlvyDt,		-1,		-1
		ggoSpread.SSSetRequired		C_ReqDt,		-1,		-1    
		ggoSpread.SpreadUnLock		C_PurOrg ,		-1,		C_PurOrgPopUp,		-1
		ggoSpread.SSSetRequired		C_PurOrg,		-1,		-1
		ggoSpread.SpreadLock		C_Tracking ,	-1,		C_TrackingPopUp,	-1
		ggoSpread.SSSetProtected	C_PlanDt, 	-1, 		-1
		ggoSpread.SSSetProtected	C_ReqStateCd,	-1,		-1
		ggoSpread.SSSetProtected	C_ReqStateNm,	-1,		-1
		ggoSpread.SSSetProtected	C_ReqTypeCd,	-1,		-1
		ggoSpread.SSSetProtected	C_ReqTypeNm,	-1,		-1
		.vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_PlantCd,	pvStartRow,		pvEndRow
    ggoSpread.SSSetRequired		C_ItemCd,	pvStartRow,		pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm,	pvStartRow,		pvEndRow
    ggoSpread.SSSetProtected	C_ItemSpec, pvStartRow,		pvEndRow
    ggoSpread.SSSetRequired		C_ReqQty,	pvStartRow,		pvEndRow
    ggoSpread.SSSetRequired		C_ReqUnit,  pvStartRow,		pvEndRow    
    ggoSpread.SSSetRequired		C_DlvyDt,	pvStartRow,		pvEndRow
    ggoSpread.SSSetRequired		C_ReqDt,	pvStartRow,		pvEndRow    
    ggoSpread.SSSetRequired		C_PurOrg,	pvStartRow,		pvEndRow
    ggoSpread.SSSetProtected		C_PlanDt, 	pvStartRow, 		pvEndRow
    ggoSpread.SSSetProtected	C_ReqStateCd,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ReqStateNm,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ReqTypeCd,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ReqTypeNm,	pvStartRow, pvEndRow    
    .vspdData.ReDraw = True
    
    End With
End Sub

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
Function OpenReqNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��û��ȣ"     
	arrParam(1) = "M_PURCHASE_REQUISITION"  
	 
	arrParam(2) = Trim(frm1.txtReqNo.Value)  
	 
	arrParam(4) = ""       
	arrParam(5) = "��û��ȣ"     
	 
	arrField(0) = "Pr_No"     
	arrField(1) = "F2" & Parent.gColSep & "Convert(varchar(10), req_qty)" 
	arrField(2) = "req_unit"          
	    
	arrHeader(0) = "��û��ȣ"         
	arrHeader(1) = "����"          
	arrHeader(2) = "����"     
	    
	'arrRet = window.showModalDialog("m2111pa1.asp", Array(window.parent,""), _
	'"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	iCalledAspName = AskPRAspName("M2111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReqNo(arrRet)
	End If 
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
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
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)	
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
	
End Function

Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.vspdData.Text)
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
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
		Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(Parent.UCN_PROTECTED) then Exit Function
	 
	if Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		Exit Function
	end if
	 
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
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement
	End If					

End Function

Function OpenItemCd()
	Dim arrRet, iCalledAspName
	Dim arrParam(5), arrField(2)
	
	If IsOpenPop = True Then Exit Function
	
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	if  Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		Exit Function
	End if

	IsOpenPop = True
	
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(0) = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col=C_ItemCd
	arrParam(1) = Trim(frm1.vspdData.Text)
	
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 
	arrField(2) = 3 ' -- Spec	
	
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
	
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItem(arrRet)
	End If	
	
End Function

'------------------------------------------  OpenORG()  -------------------------------------------------
'	Name : OpenORG()
'	Description : OpenORG PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"						<%' �˾� ��Ī %>
	arrParam(1) = "B_Pur_Org"						<%' TABLE ��Ī %>
	
	arrParam(2) = Trim(frm1.txtORGCd.Value)	<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtORGNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = "usage_flg = " & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "��������"							<%' TextBox ��Ī %>
	
    arrField(0) = "PUR_ORG"					<%' Field��(0)%>
    arrField(1) = "PUR_ORG_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "��������"						<%' Header��(0)%>
    arrHeader(1) = "����������"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtORGCd.Value    = arrRet(0)		
		frm1.txtORGNm.Value    = arrRet(1)	
	End If	
End Function

Function OpenORGCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_PurOrg
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "��������"						<%' �˾� ��Ī %>
	arrParam(1) = "B_Pur_Org"						<%' TABLE ��Ī %>
	
	arrParam(2) = Trim(frm1.vspdData.Text)		<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtORGNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = "usage_flg = " & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "��������"							<%' TextBox ��Ī %>
	
    arrField(0) = "PUR_ORG"					<%' Field��(0)%>
    arrField(1) = "PUR_ORG_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "��������"						<%' Header��(0)%>
    arrHeader(1) = "����������"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetORG(arrRet)
	End If	
End Function

'------------------------------------------  OpenMrp()  -------------------------------------------------
'	Name : OpenMrp()
'	Description : OpenMrp PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenMrp()
    Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
   
    If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if 
		
	IsOpenPop = True

	arrParam(0) = "MRP Run��ȣ"				<%' �˾� ��Ī %>
	arrParam(1) = "(select distinct a.order_no A,a.confirm_dt B," & FilterVar("������������", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_EXPL_HISTORY a, m_pur_req b where a.order_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "
    arrParam(1) = arrParam(1) & "union "
    arrParam(1) = arrParam(1) & "select distinct  a.run_no A, a.start_dt B ," & FilterVar("MRP����", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_MRP_HISTORY a, m_pur_req b where a.run_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ") as g" <%' TABLE ��Ī %>
    

	arrParam(2) = Trim(frm1.txtMRP.value)		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
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

'------------------------------------------  OpenDept()  -------------------------------------------------
'	Name : OpenDept()
'	Description :  OpenDept PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenDept()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_DeptCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "��û�μ�"						<%' �˾� ��Ī %>
	arrParam(1) = "B_ACCT_DEPT"						<%' TABLE ��Ī %>
	
	arrParam(2) = Trim(frm1.vspdData.Text)		<%' Code Condition%>
	
	
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(parent.gChangeOrgId, "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "��û�μ�"							<%' TextBox ��Ī %>
	
    arrField(0) = "DEPT_CD"					<%' Field��(0)%>
    arrField(1) = "DEPT_NM"
    
    arrHeader(0) = "��û�μ�"						<%' Header��(0)%>
    arrHeader(1) = "��û�μ���"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
	
End Function


Function OpenTracking()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_Tracking
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "TRACKINGNO"	
	arrParam(1) = "s_so_tracking"				
	
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	
	arrParam(4) = ""			
	arrParam(5) = "Tracking No"			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking No"		
    arrHeader(1) = "ǰ��"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTracking(arrRet)
	End If	
	
End Function


Function OpenBP()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	'frm1.vspdData.Col=C_SpplCd
	'frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "����ó"	
	arrParam(1) = "M_PUR_REQ, B_BIZ_PARTNER"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = "M_PUR_REQ.SPPL_CD = B_BIZ_PARTNER.BP_CD"			
	arrParam(5) = "����ó"			
	
    arrField(0) = "M_PUR_REQ.SPPL_CD"	
    arrField(1) = "B_BIZ_PARTNER.BP_NM"	
    
    arrHeader(0) = "����ó"		
    arrHeader(1) = "����ó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBP(arrRet)
	End If	
	
End Function


Function OpenGrp()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_GrpCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_pur_grp"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	
	frm1.vspdData.Col=C_PurOrg
	
	arrParam(4) = "Usage_flg=" & FilterVar("Y", "''", "S") & "  and PUR_ORG =  " & FilterVar(UCase(frm1.vspdData.Text), "''", "S") & " "
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
		Call SetGrp(arrRet)
	End If	
	
End Function


Function OpenReqUnit()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_ReqUnit
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
		Call SetReqUnit(arrRet)
	End If	
	
End Function


Function OpenStorage()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData.Col=C_StorageCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "�԰�â��"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�԰�â��"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "�԰�â��"		
    arrHeader(1) = "�԰�â���"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetStorage(arrRet)
	End If	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  Set?????()  --------------------------------------------------
' Name : SetPlant()
' Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReqNo(byval arrRet)
	frm1.txtReqNo.Value= arrRet(0) 
	frm1.txtReqNo.focus	
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetItem()  --------------------------------------------------
'	Name : SetItem()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItem(byval arrRet)
	
	frm1.vspdData.Col = C_ItemCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)	
	frm1.vspdData.Col  = C_ItemNm
	frm1.vspdData.Text = arrret(1)	
	frm1.vspdData.Col  = C_ItemSpec
	frm1.vspdData.Text = arrret(2)	
	
	lgBlnFlgChgValue = True
	
	Call changeItemPlant()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	
	lgBlnFlgChgValue = True
	Call changeItemPlant()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
			
End Function


Function SetOrg(byval arrRet)
	
	frm1.vspdData.Col = C_PurOrg
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_SpplNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
End Function


Function SetDept(byval arrRet)
	
	frm1.vspdData.Col = C_DeptCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_DeptNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
End Function


Function SetTracking(byval arrRet)
	
	frm1.vspdData.Col = C_Tracking
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_SpplNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
End Function


Function SetBP(byval arrRet)
	
	frm1.vspdData.Col = C_SpplCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_SpplNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
	Call SpplChange()
	
End Function


Function SetGrp(byval arrRet)
	
	frm1.vspdData.Col = C_GrpCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_GrpNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
	
End Function


Function SetReqUnit(byval arrRet)

	frm1.vspdData.Col = C_ReqUnit
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
		
End Function


Function SetStorage(byval arrRet)
	
	frm1.vspdData.Col = C_StorageCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)		
	'frm1.vspdData.Col  = C_SpplNm
	'frm1.vspdData.Text = arrret(1)
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################%>
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
 Sub Form_Load()

    Call LoadInfTB19029    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call SetDefaultVal
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    
    Call InitVariables                                                      '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    'Call SetToolbar("1100000000001111")										'��: ��ư ���� ����	
    Call ReadCookiePage()
    
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub


'==========================================================================================
'   Event Name : txtDlvyFrDt
'   Event Desc :
'==========================================================================================
 Sub txtDlvyFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyFrDt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : txtDlvyToDt
'   Event Desc :
'==========================================================================================
 Sub txtDlvyToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyToDt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : txtReqFrDt
'   Event Desc :
'==========================================================================================

 Sub txtReqFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReqFrDt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : txtReqFrDt
'   Event Desc :
'==========================================================================================

 Sub txtReqToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReqToDt.Action = 7
	End if
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


'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
   gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	Call SetPopupMenuItemInf("1101111111")
	

	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = Frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    
End Function

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	
	if Col = C_ItemCd then
		Call changeItemPlant()
	End if
	
	if Col = C_PlantCd then
		Call changeItemPlant()
	End if
	
	'if Col = C_DlvyDt then
	'	Call changeDlvy()
	'End if
	
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col

	'If Frm1.vspdData.CellType = SS_CELL_TYPE_FLOAT Then
	'	If uniCDbl(Frm1.vspdData.text) < uniCDbl(Frm1.vspdData.TypeFloatMin) Then
	'		Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
	'	End If
	'End If		
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

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
'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
 Sub vspdData_DblClick(ByVal Col , ByVal Row)
  	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
 End Sub



'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

' Sub vspdData_GotFocus()
'    ggoSpread.Source = frm1.vspdData
'End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_PlantPopUp Then
        .Col = Col
        .Row = Row
        Call OpenPlantCd()
    
    elseif Row > 0 And Col = C_ItemPopUp Then
        .Col = Col
        .Row = Row
	    Call OpenItemCd()
    
    elseif Row > 0 And Col = C_ReqUnitPopUp Then
        .Col = Col
        .Row = Row
        Call OpenReqUnit()
    
    elseif Row > 0 And Col = C_PurOrgPopUp Then
        .Col = Col
        .Row = Row
        Call OpenORGCd()
    
    elseif Row > 0 And Col = C_DeptPopUp Then
        .Col = Col
        .Row = Row
        Call OpenDept()
    
    elseif Row > 0 And Col = C_StoragePopUp Then
        .Col = Col
        .Row = Row
        Call OpenStorage()
        
    elseif Row > 0 And Col = C_TrackingPopUp Then
        .Col = Col
        .Row = Row
        Call OpenTracking()
        
    elseif Row > 0 And Col = C_SpplPopUp Then
        .Col = Col
        .Row = Row
        Call OpenBP()
    
    elseif Row > 0 And Col = C_GrpPopUp Then
        .Col = Col
        .Row = Row
        Call OpenGrp()
           
    End If
    
    End With
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
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLock()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
       
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	           
    	If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
	End If
    End if

End Sub

'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### %>

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
'######################################################################################################### %>
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* %>
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
 Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    ggoSpread.Source = frm1.vspdData
    Err.Clear                                                               '��: Protect system from crashing
                                 
  	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
 
    '-----------------------
    'Check condition area
    '-----------------------
	with frm1
	     If CompareDateByFormat(.txtDlvyFrDt.text,.txtDlvyToDt.text,.txtDlvyFrDt.Alt,.txtDlvyToDt.Alt, _
                 "970025",.txtDlvyFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtDlvyFrDt.text) <> "" And Trim(.txtDlvyToDt.text) <> "" Then
	''	'if (UniCDate(.txtDlvyFrDt.text) > UniCDate(.txtDlvyToDt.text)) and trim(.txtDlvyFrDt.text)<>"" and trim(.txtDlvyToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")			
			Exit Function
		End if 
	
	    If CompareDateByFormat(.txtReqFrDt.text,.txtReqToDt.text,.txtReqFrDt.Alt,.txtReqToDt.Alt, _
                   "970025",.txtReqFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtReqFrDt.text) <> "" And Trim(.txtReqToDt.text) <> "" Then
		'if (UniCDate(.txtReqFrDt.text) > UniCDate(.txtReqToDt.text)) and trim(.txtReqFrDt.text)<>"" and trim(.txtReqToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","��û��", "X")			
			Exit Function
		end if     
			
	End with	

	'----------------------------------------------------------------
    'Set Parameter to Hidden area (Added By Lee Sung Yong 2005/01/28)
    '----------------------------------------------------------------
    
    With frm1
        
		.hdnPlant.value = Trim(.txtPlantcd.value)
		.hdnItem.value  = Trim(.txtItemcd.value)
		.hdnDFrDt.Value = Trim(.txtDlvyFrDt.text)
		.hdnDToDt.Value = Trim(.txtDlvyToDt.text)
		.hdnRFrDt.Value = Trim(.txtReqFrDt.text)
		.hdnRToDt.Value = Trim(.txtReqToDt.text)
		.hdnORGCd.value = Trim(.txtORGCd.value)
		.hdnMRP.value 	= Trim(.txtMRP.value)
	
	End with
	
    '-----------------------
    'Erase contents area
    '-----------------------
  '  Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True												
														'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
 Function FncNew() 

    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    

    Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '��: Clear Contents  Field
    Call ggoOper.ClearField(Document, "Q")
    
    Call InitVariables
    Call SetDefaultVal
    frm1.vspdData.MaxRows = 0
   
    FncNew = True                                                           '��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
 Function FncDelete() 

	Dim IntRetCD

    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData  
    
    IntRetCD = DisplayMsgBox("900003", VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    						
    If lgIntFlgMode <> OPMD_UMODE Then 
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True    
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

 Function FncSave() 
 
    Dim IntRetCD 
	    
    FncSave = False 

    Err.Clear       

    ggoSpread.Source = frm1.vspdData 
         
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  
        IntRetCD = DisplayMsgBox("900001","X","X","X")            
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then               
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData                  
    If Not ggoSpread.SSDefaultCheck Then              
       Exit Function
    End If
    
    'If Trim(UniCdbl(frm1.txtReqQty.Text)) = "" Or Trim(UniCdbl(frm1.txtReqQty.Text)) = "0" then
	'	Call DisplayMsgBox("970021", "X","��û��", "X")
	'	frm1.txtReqQty.focus
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	'End if

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False Then Exit Function
    
    FncSave = True                                    
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

 Function FncCancel() 

	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                
    
    if frm1.vspdData.MaxRows < 1 then
    	Call ChangeTag(False)
    End if
    
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
'msgbox 1
    Dim IntRetCD
    Dim imRow
    Dim inti
    inti=1
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If
	
	With frm1
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
    
		.vspdData.ReDraw = False
		ggoSpread.InsertRow, imRow
		
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow+imRow - 1
    
		.vspdData.Col = C_PlantCd
		.vspdData.Row = frm1.vspdData.ActiveRow
		.vspdData.Text = parent.gPlant		
    
		.vspdData.Col = C_ReqDt
		.vspdData.Row = frm1.vspdData.ActiveRow
		.vspdData.Text =  EndDate
    
		.vspdData.Col = C_DeptCd
		.vspdData.Row = frm1.vspdData.ActiveRow
		.vspdData.Text = parent.gDepart		    

		.vspdData.Col = C_ReqPrsn	
		.vspdData.Row = frm1.vspdData.ActiveRow
		.vspdData.Text = parent.gUsrID
		
		.vspdData.ReDraw = True
		    
    End With
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
 Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i
    
    if frm1.vspdData.Maxrows < 1	then exit function
	
    With frm1.vspdData 
    
		.focus
		ggoSpread.Source = frm1.vspdData 

		.Row = .ActiveRow
		.Col = C_ReqStateCd    
		if Trim(.text) <> "RQ" then 
			call DisplayMsgBox("172126","X","X","X")
			exit function
		end if
		
		lDelRows = ggoSpread.DeleteRow
    
    End With
    
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                              
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

 Function FncNext() 
    On Error Resume Next                              
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(C_MULTI)					
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(C_MULTI , False)           
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
 Function FncExit()
	
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    

	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    
End Function

'=======================================================================================================
'=	Event Name : FncCopy																				=
'=	Event Desc : This function is related to Copy Button of Main ToolBar								=
'========================================================================================================

Function FncCopy()
	frm1.vspdData.ReDraw = False

	if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = 1
    frm1.vspdData.text = ""
    
    frm1.vspdData.Col = C_HdnTrackingflg
	if UCase(Trim(frm1.vspdData.text)) = "Y" then			
		ggoSpread.SpreadUnLock	C_Tracking, frm1.vspdData.ActiveRow, C_TrackingPopUp, frm1.vspdData.ActiveRow		
		ggoSpread.SSSetRequired	C_Tracking, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow	
	else
		ggoSpread.SpreadLock	C_Tracking, frm1.vspdData.ActiveRow, C_TrackingPopUp, frm1.vspdData.ActiveRow	
	end if
	
    'frm1.vspdData.Col = C_ReqDt
    'frm1.vspdData.Text = EndDate 
    frm1.vspdData.Col = C_PlanDt
    frm1.vspdData.text = ""
 
     frm1.vspdData.Col = C_ReqStateCd 
     frm1.vspdData.text = ""
     frm1.vspdData.Col = C_ReqStateNm
     frm1.vspdData.text = ""
     frm1.vspdData.Col = C_ReqTypeCd 	
     frm1.vspdData.text = ""
     frm1.vspdData.Col = C_ReqTypeNm
     frm1.vspdData.text = ""
     frm1.vspdData.Col = C_DeptCd
     frm1.vspdData.Text = parent.gDepart		    
     frm1.vspdData.Col = C_ReqPrsn	
     frm1.vspdData.Text = parent.gUsrID
	
     frm1.vspdData.ReDraw = True

End Function


'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* %>

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
    Dim strVal
    'Dim pP21018         'As New P21018ListIndReqSvr

    DbQuery = False
    Err.Clear                                                               '��: Protect system from crashing
   
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
	    strVal = strVal & "&txtItemCd=" & .hdnItem.value
	    strVal = strVal & "&txtDlvyFrDt=" & .hdnDFrDt.Value
		strVal = strVal & "&txtDlvyToDt=" & .hdnDToDt.Value
		strVal = strVal & "&txtReqFrDt=" & .hdnRFrDt.Value
		strVal = strVal & "&txtReqToDt=" & .hdnRToDt.Value
		'strVal = strVal & "&txtStateCd=" & .hdnState.value		
		'strVal = strVal & "&txtDeptCd=" & .hdnDept.Value
	    strVal = strVal & "&txtORGCd=" & Trim(.hdnORGCd.value)
	    strVal = strVal & "&txtReqNo=" & Trim(.hdnReqNo.value)
	    strVal = strVal & "&txtMRP=" & Trim(.hdnMRP.value)
    Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantcd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemcd.value)
	    strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
		strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
		strVal = strVal & "&txtReqFrDt=" & Trim(.txtReqFrDt.text)
		strVal = strVal & "&txtReqToDt=" & Trim(.txtReqToDt.text)
		'strVal = strVal & "&txtStateCd=" & Trim(.txtStateCd.value)
	    'strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.Value)	
	    strVal = strVal & "&txtORGCd=" & Trim(.txtORGCd.value)
	    strVal = strVal & "&txtReqNo=" & Trim(.hdnReqNo.value)
	    strVal = strVal & "&txtMRP=" & Trim(.txtMRP.value)
   End If
		strVal = strVal & "&lgPageNo=" & lgPageNo 
		strVal = strVal & "&lgNextKey=" & lgNextKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    
    DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Dim index,strPrSts,strSpplCd
    '-----------------------
    'Reset variables area
    '-----------------------											
    lgBlnFlgChgValue = False

    'Call ggoOper.LockField(Document, "Q")								

    if frm1.vspdData.MaxRows > 0 then
    	Call SetToolbar("1110111100101111")
		lgIntFlgMode = parent.OPMD_UMODE
    else

		'frm1.txtPlantCd2.value = frm1.txtPlantCd.value
		'frm1.txtPlantNm2.value = frm1.txtPlantNm.value
		'frm1.txtItemCd2.value = frm1.txtItemCd.value
		'frm1.txtItemNm2.value = frm1.txtItemNm.value
		'frm1.txtSpplCd2.value = frm1.txtSpplCd.value
		'frm1.txtSpplNm2.value = frm1.txtSpplNm.value
		Call ggoOper.LockField(Document, "N")
    	Call SetToolbar("1110111100101111")
		lgIntFlgMode = parent.OPMD_CMODE   
		'��ȸ �� �ڷᰡ ������ ��ȸ ������ �ڷḦ ������ �ο� �����ִ� �κ� 
		'queryok�� ���� ������ ���� �ű��� ���� �̹Ƿ� �ű� ���� ������ �� 
		'-> ����� �����ϸ� update�� �Ǿ���� 2001.08.02 Ever
    end if
    
    frm1.vspddata.ReDraw = False
    
    ggoSpread.Source = frm1.vspdData
    
      
    For Index = 1 to frm1.vspdData.MaxRows
    	
    	frm1.vspdData.Row = Index   
    	frm1.vspdData.Col = C_ReqStateCd 
    	strPrSts = Trim(frm1.vspdData.Text)

		frm1.vspdData.Col = C_ReqStateCd 
    	strPrSts = Trim(frm1.vspdData.Text)
    	
    	frm1.vspdData.Col = C_SpplCd
    	strSpplCd = Trim(frm1.vspdData.Text)
    	
    	if UCase(strPrSts) <> "RQ"  OR strSpplCd <> "" then
			ggoSpread.SSSetProtected -1, Index, Index    	
		end if
		
		frm1.vspdData.Col = C_ReqNo
		if Trim(frm1.vspdData.Text) <> "" then
    		ggoSpread.SSSetProtected 1, Index, Index    	
		end if
		
	Next
	
	frm1.vspdData.ReDraw = True
	    	
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 

    Err.Clear																<%'��: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
	'<!-- �߰��κ� ���� -->
	Dim PvArr
	Dim strReqNo,strPntCd,strItemCd, strReqQTy,strReqUnit,strDlvyDt,strReqDt,strPurOrg
	Dim strDeptCd,strReqPrsn, strSlCd,strTrackingNo,strSpplCd, strGrpCd, strPlanDt,strReqStsCd
	Dim strReqTypeCd, strHdnProcurType, strHdnMrpNo
	Dim iSpdCount,ColSep,RowSep
	Dim lValCnt
	Dim strHTML
	Dim iArrStrVal
	Dim iTempTxt
	Dim i
	'<!-- �߰��κ� ��   -->

	DbSave = False                                                          '��: Processing is NG
	Call LayerShowHide(1)

	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep		

	With frm1
		.txtMode.value = parent.UID_M0002
		
		'.txtFlgMode.value = parent.lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
	
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
		strHTML = ""
		iSpdCount = 0

		ReDim iArrStrVal(iSpdCount)
		ReDim PvArr(500) 
		'-----------------------
		'Data manipulate area
		'-----------------------
    
		ggoSpread.Source = .vspdData
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
			Select Case .vspdData.Text

				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
					if .vspdData.Text=ggoSpread.InsertFlag then		
						strVal = strVal & "C" & ColSep				'��: �ű� 
					Else
						strVal = strVal & "U" & ColSep				'��: ���� 
					End if      
					
					.vspdData.Col = C_DlvyDt 
					 if Trim(.vspdData.Text) = ""  then
	   					 Call DisplayMsgBox("970021","X","�ʿ���","X")
	   					 Exit Function
					End if	
									
					 .vspdData.Col = C_ReqDt 
					 if Trim(.vspdData.Text) = ""  then
	   					 Call DisplayMsgBox("970021","X","��û��","X")
	   					 Exit Function
					End if	
					
    
					'--- ��û��ȣ 
		            .vspdData.Col = C_ReqNo 		            
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strReqNo = Trim(.vspdData.Text)
		            '--- ���� 
		            .vspdData.Col = C_PlantCd 
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strPntCd = Trim(.vspdData.Text)
					'--- ǰ�� 
		            .vspdData.Col = C_ItemCd 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strItemCd = Trim(.vspdData.Text)
					'--- ��û�� 
		            .vspdData.Col = C_ReqQty 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strReqQTy = Trim(.vspdData.Text)
					'--- ��û���� 
		            .vspdData.Col = C_ReqUnit 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strReqUnit = Trim(.vspdData.Text)
					'--- �ʿ��� 
		            .vspdData.Col = C_DlvyDt 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strDlvyDt = UNIConvDate(Trim(.vspdData.Text))
					'--- ��û��	
				'msgbox strDlvyDt 
		            .vspdData.Col = C_ReqDt 		
		          ' strVal = strVal & Trim(.vspdData.Text) & ColSep
					strReqDt = UNIConvDate(Trim(.vspdData.Text))
	'msgbox strReqDt 
		            '--- �������� 
                    .vspdData.Col = C_PurOrg 		
		           ' strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strPurOrg = Trim(.vspdData.Text)

					'--- ��û�μ� 
                    .vspdData.Col = C_DeptCd 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strDeptCd = Trim(.vspdData.Text)

					'--- ��û�� 
                    .vspdData.Col = C_ReqPrsn 		
		           ' strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strReqPrsn = Trim(.vspdData.Text)
					'--- �԰�â�� 
                    .vspdData.Col = C_StorageCd 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strSlCd = Trim(.vspdData.Text)
		            '--- Tracking No
                    .vspdData.Col = C_Tracking 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
			if Trim(.vspdData.Text) <> "" Then
		            strTrackingNo =  Trim(.vspdData.Text)
			else
			     strTrackingNo = "*"
			end if
					'--- ����ó 
		     '        .vspdData.Col = C_SpplCd 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		      '      strSpplCd =  Trim(.vspdData.Text)
					'--- ���ű׷� 
                  '  .vspdData.Col = C_GrpCd 		
		   '        ' strVal = strVal & Trim(.vspdData.Text) & ColSep
		  '--- ���ֿ����� 
                   ' .vspdData.Col = C_PlanDt 		
		    '        'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strPlanDt= UNIConvDate("1990-01-01")
		            '--- ��û������� 
                    .vspdData.Col = C_ReqStateCd 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep		            		            		            
		            strReqStsCd = Trim(.vspdData.Text)
		            '--- ��û���� 
                            strReqTypeCd = "E"
		            '---C_HdnProcurType
					.vspdData.Col = C_HdnProcurType 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strHdnProcurType = Trim(.vspdData.Text)
		            '---C_HdnMrpNo
					.vspdData.Col = C_HdnMrpNo 		
		            'strVal = strVal & Trim(.vspdData.Text) & ColSep
		            strHdnMrpNo = Trim(.vspdData.Text)
				
		            '--- gUsrID
					'strVal = strVal & Trim(parent.gUsrID) & RowSep
					PvArr(lValCnt) = strVal & strReqNo & ColSep & strPntCd & ColSep & strItemCd & ColSep & strReqQTy & ColSep & strReqUnit & ColSep & strDlvyDt & ColSep & strReqDt & ColSep & strPurOrg & ColSep & _
									strDeptCd & ColSep & strReqPrsn & ColSep & strSlCd & ColSep & strTrackingNo & ColSep & "" & ColSep & "" & ColSep & strPlanDt & ColSep & strReqStsCd & ColSep & _
									strReqTypeCd & ColSep & strHdnProcurType & ColSep & strHdnMrpNo & ColSep & Trim(parent.gUsrID) & ColSep & lRow & RowSep
'msgbox PvArr(lValCnt)
		            strVal = ""
                    
                    '<!-- �߰��κ� ���� -->
					lGrpCnt = lGrpCnt + 1
					lValCnt = lValCnt + 1
                    If lValCnt = 500 Then	'���� ���� '1000'�� MB�� �ѱ�� STRINGũ�⿡ ���� �޶���.
                        iArrStrVal(iSpdCount) = Join(PvArr,"")
                        iSpdCount = iSpdCount + 1
                        ReDim Preserve iArrStrVal(iSpdCount)
                        strHTML = strHTML & "<TEXTAREA CLASS=hidden Name=txtSpread" & iSpdCount & " Width=100% tag=""24"" TABINDEX=""-1""></TEXTAREA>"
                        'strVal = ""
                        ReDim PvArr(500)
                        lValCnt = 0
                    End If       
                    '<!-- �߰��κ� ��   -->
                    
		        Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & ColSep
					 .vspdData.Col = C_ReqNo 		            
		            strDel = strDel & Trim(.vspdData.Text) & ColSep & lRow & RowSep
		          ' strDel = strDel & Trim(gUsrID) & RowSep
		            lGrpCnt = lGrpCnt + 1 
		    End Select       

		Next
    End With
	
    '<!-- �߰��κ� ���� -->
    If lValCnt <> 0 Then
        iArrStrVal(iSpdCount) = Join(PvArr,"")
        iSpdCount = iSpdCount + 1
	    strHTML = strHTML & "<TEXTAREA CLASS=hidden Name=txtSpread" & iSpdCount & " Width=100% tag=""24"" TABINDEX=""-1""></TEXTAREA>"
    End If
    '<!-- �߰��κ� ��   -->
    
    	
	frm1.txtMaxRows.value = lGrpCnt
	frm1.txtSpread.value = strDel 
	'msgbox frm1.txtSpread.value
	'<!-- �߰��κ� ���� -->
	divTextArea.innerHTML = strHTML
	frm1.SpdCount.value = iSpdCount

    i=0
    For Each iTempTxt in divTextArea.childNodes
'msgbox typename(iTempTxt)
        iTempTxt.value = iArrStrVal(i)
        i=i+1
    Next
	'<!-- �߰��κ� ��   -->
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			
    DbSave = True                                 '��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
        Call MainQuery()
End Function


Function DbInsertOk()													'��: ���� ������ ���� ���� 
'msgbox "DbSaveOk"
    lgIntFlgMode = OPMD_UMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0  
'	Call InitVariables
	'frm1.txtConSoNo.value = frm1.txtHSoNo.value
'	frm1.vspdData.MaxRows = 0
    'Call MainQuery()
End Function

Function DbUpdateOk()													'��: ���� ������ ���� ���� 
'msgbox "DbSaveOk"
    lgIntFlgMode = OPMD_UMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0  
'	Call InitVariables
	'frm1.txtConSoNo.value = frm1.txtHSoNo.value
'	frm1.vspdData.MaxRows = 0
    'Call MainQuery()
End Function

<!--
'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
-->
Function DbDeleteOk()												
	lgBlnFlgChgValue = False
	'Call MainQuery()
End Function

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
%>

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<% '#########################################################################################################
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
'######################################################################################################### %>
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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ſ�û</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
														   <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
														   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm" SIZE=20 tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtDlvyFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtDlvyToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>��û��</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="��û��" NAME="txtReqFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="��û��" NAME="txtReqToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtORGCd" ALT="��������" SIZE=10 MAXLENGTH=4  tag="11NXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG()">
													   <INPUT TYPE=TEXT ID="txtORGNm" ALT="��������" NAME="arrCond" tag="14X"></TD>
								    <TD CLASS="TD5" NOWRAP>MRP Run��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="MRP Run��ȣ" NAME="txtMRP" SIZE=26 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMrp()"></TD>					   
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<!--<table <%=LR_SPACE_TYPE_30%>>
				<tr>	 
					<td WIDTH="*" ALIGN="RIGHT"><a href="VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:WriteCookiePage()">���ſ�û��ȸ</a></td>
					<td WIDTH="20"></td>
				</tr>
			</table>-->
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="<%=BIZ_PGM_ID%>" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<!--	�߰��κ� ����	-->
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="SpdCount" tag="24" TABINDEX="-1">
<!--	�߰��κ� ��	    -->

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnState" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrsn" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDept" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnOrgCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnReqNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMRP" tag="24">

</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
