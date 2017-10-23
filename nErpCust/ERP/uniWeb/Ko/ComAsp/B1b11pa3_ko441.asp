<%@ LANGUAGE="VBSCRIPT" %>
<!--********************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        : Popup Item By Plant														*	
'*  3. Program ID           : b1b11pa3.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Item by Plant Popup														*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2004/12/06																*
'*  9. Modifier (First)     : Im Hyun Soo																*
'* 10. Modifier (Last)      : Chen, Jae Hyun																*
'* 11. Comment              : New Item Account Version																		*
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../inc/incSvrCcm.inc" -->
<!-- #Include file="../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "b1b11pb3.asp"							'☆: Asp name of Biz logic

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_BasicUnit
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ItemGroupCd
Dim C_ItemClass
Dim C_ProcurType
Dim C_ProcurTypeNm
Dim C_ProdtEnv
Dim C_ProdtEnvNm
Dim C_PhantomFlg
Dim C_LotFlg
Dim C_MajorSlCd
Dim C_IssuedSlCd
Dim C_ValidFlg
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_FormalNm
Dim C_BItemItemAcct
Dim C_HsCd
Dim C_HsUnit
Dim C_BaseItemCd
Dim C_TrackingFlg
Dim C_OrderUnitMfg
Dim C_OrderUnitPur
Dim C_OrderLtMfg
Dim C_OrderLtPur
Dim C_OrderType
Dim C_OrderRule
Dim C_FixedMrpQty
Dim C_MinMrpQty
Dim C_MaxMrpQty
Dim C_RoundQty
Dim C_RoundPerd
Dim C_MpsFlag
Dim C_IssueMthd
Dim C_PurOrg
Dim C_OptionFlg
Dim C_CycleCntPerd
Dim C_IssuedUnit
Dim C_RecvInspecFlg
Dim C_ProdInspecFlg
Dim C_FinalInspecFlg
Dim C_ShipInspecFlg
Dim C_InspecLtMfg
Dim C_InspecLtPur
Dim C_InspecMgr
Dim C_BItemValidFlg
Dim C_CollectiveFlg
Dim C_MajorWorkCenter
Dim C_AbcFlg
Dim C_ItemAcctGrp

<!-- #Include file="../inc/lgVariables.inc" -->

Dim strReturn
Dim lgCurDate
Dim gblnWinEvent
Dim IsOpenPop

Dim arrReturn
Dim arrParam					
Dim arrField
Dim PlantCd
Dim arrParent
Dim strNextKey	'item_nm Next Key Value	2003-09-02

Dim lgItemAcctGrp
Dim lgProcType
Dim lgDefItemAcct
Dim lgDefProcType
Dim lgWhere
Dim PopupParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName
	
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCd				= 1
	C_ItemNm				= 2
	C_Spec					= 3
	C_BasicUnit				= 4
	C_ItemAcct				= 5
	C_ItemAcctNm			= 6
	C_ItemGroupCd			= 7
    C_ItemClass				= 8
	C_ProcurType			= 9
	C_ProcurTypeNm			= 10
	C_ProdtEnv				= 11
	C_ProdtEnvNm			= 12
	C_PhantomFlg			= 13
	C_LotFlg				= 14
	C_MajorSlCd				= 15
	C_IssuedSlCd			= 16
	C_ValidFlg				= 17
	C_ValidFromDt			= 18
	C_ValidToDt				= 19
	C_FormalNm				= 20
	C_BItemItemAcct			= 21
	C_HsCd					= 22
	C_HsUnit				= 23
	C_BaseItemCd			= 24
	C_TrackingFlg			= 25
	C_OrderUnitMfg			= 26
	C_OrderUnitPur			= 27
	C_OrderLtMfg			= 28
	C_OrderLtPur			= 29
	C_OrderType				= 30
	C_OrderRule				= 31
	C_FixedMrpQty			= 32
	C_MinMrpQty				= 33
	C_MaxMrpQty				= 34
	C_RoundQty				= 35
	C_RoundPerd				= 36
	C_MpsFlag				= 37
	C_IssueMthd				= 38
	C_PurOrg				= 39
	C_OptionFlg				= 40
	C_CycleCntPerd			= 41
	C_IssuedUnit			= 42
	C_RecvInspecFlg			= 43
	C_ProdInspecFlg			= 44
	C_FinalInspecFlg		= 45
	C_ShipInspecFlg			= 46
	C_InspecLtMfg			= 47
	C_InspecLtPur			= 48
	C_InspecMgr				= 49
	C_BItemValidFlg			= 50
	C_CollectiveFlg			= 51
	C_MajorWorkCenter		= 52 
	C_AbcFlg				= 53 
	C_ItemAcctGrp			= 54
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()

	lgStrPrevKeyIndex = ""
	strNextKey = ""
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	gblnWinEvent = False
	
    lgSortKey = 1                                       '⊙: initializes sort direction
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
	
'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "PA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
	Dim i, iProcurTypeArr, iProcurTypeNmArr
	Dim iWhereItemAcct  

    On Error Resume Next
    Err.Clear

	'------------------------------------------------------------
	' Setting Item Account Combo
	'------------------------------------------------------------
	If Ucase(lgItemAcctGrp(0)) = "Y" Then
		iWhereItemAcct = " AND B.ITEM_ACCT_GROUP > " & Filtervar(lgItemAcctGrp(1), "''", "S") _
				&  " AND B.ITEM_ACCT_GROUP < " & Filtervar(Trim(lgItemAcctGrp(2)) + 1, "''", "S")
	Else
		iWhereItemAcct = ""
	End If	 
	
	Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM ", _
				" B_MINOR A, B_ITEM_ACCT_INF B ", _
				" A.MINOR_CD = B.ITEM_ACCT AND A.MAJOR_CD = " & Filtervar("P1001", "''", "S") & "  " & iWhereItemAcct &  "  ORDER BY A.MINOR_CD ", _
				 lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    
    Call SetCombo2(cboItemAccount, lgF0, lgF1, Chr(11))
    
	   '------------------------------------------------------------
	   ' Setting Default Value in Item Account Combo
	   '------------------------------------------------------------ 
	
	If lgDefItemAcct <> "" Then
		cboItemAccount.value = UCase(lgDefItemAcct)
	Else
		cboItemAccount.SelectedIndex=0
	End If	

	   '------------------------------------------------------------
	   ' Setting Item Class Combo
	   '------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1002' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(cboItemClass, lgF0, lgF1, Chr(11))
	
	   '------------------------------------------------------------
	   ' Setting Procur Type Combo
	   '------------------------------------------------------------ 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1003' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    iProcurTypeArr = Split(lgF0, Chr(11))
    iProcurTypeNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iProcurTypeArr) - 1
		IF UCase(lgProcType(0)) = "Y" Then
			If (UCase(iProcurTypeArr(i)) = UCase(lgProcType(1))) Or (UCase(iProcurTypeArr(i)) = UCase(lgProcType(2))) Then	
				Call SetCombo(cboProcurType, UCase(iProcurTypeArr(i)), iProcurTypeNmArr(i))
			End If
		Else
			Call SetCombo(cboProcurType, UCase(iProcurTypeArr(i)), iProcurTypeNmArr(i))
		End If
	Next
	
	   '------------------------------------------------------------
	   ' Setting Default Value in Procur Type Combo
	   '------------------------------------------------------------  
	If lgDefProcType <> "" Then
		cboProcurType.Value = UCase(lgDefProcType)
	Else
		cboProcurType.SelectedIndex=0		
	End If	

	   '------------------------------------------------------------
	   ' Setting Prodt Env Combo
	   '------------------------------------------------------------  
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1004' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(cboProdtEnv, lgF0, lgF1, Chr(11))

	   '------------------------------------------------------------
	   ' Setting Prodt Env Combo
	   '------------------------------------------------------------ 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Q0001' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(cboInspType, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim IStr1
	Dim IStr2
	Dim ILen
	Dim i
	Dim j
	Dim IPos
		
	PlantCd = arrParam(0)
	txtItemCd.value = arrParam(1)		
		
	If arrParam(2) <> "" Then
		IPos = InStr(1, arrParam(2), "!")
		If IPos <> 0 Then
			IStr1 = Left(arrParam(2), IPos - 1)
			If IPos < Len(arrParam(2)) Then
				IStr2 = Mid(arrParam(2),IPos+1,Len(arrParam(2)))
			End If
		Else
			IStr1 = arrParam(2)
		End If
	End If
		
	   '-------------------------------
	   ' 품목계정 Setting
	   '-------------------------------  
	If IStr1 = "" Then
		ReDim lgItemAcctGrp(0) 
		lgItemAcctGrp(0) = "N"
	Else
		j=0
		ILen = Len(IStr1)

		ReDim lgItemAcctGrp(ILen) 
		lgItemAcctGrp(0) = "Y"

        For i = 1 To ILen
		    j=j+1
		    lgItemAcctGrp(j) = UCase(Mid(IStr1, i, 1))
        Next    
    End If
    
	   '-------------------------------
	   ' 조달구분 Setting
	   '------------------------------- 
        
    If IStr2 <> "" Then
		ILen = Len(IStr2)

        ReDim lgProcType(ILen)
        lgProcType(0) = "Y"

        j = 0
        For i = 1 To ILen Step 1
			j = j + 1
            lgProcType(j) = UCase(Mid(IStr2, i, 1))
        Next 
    Else
		ReDim lgProcType(0)
		lgProcType(0) = "N"
    End If
		
	'-------------------------------
	' 품목계정,조달구분 Default Setting
	'-------------------------------
	If arrParam(3) <> "" Then
		IPos = InStr(1, arrParam(3), "!")
		If IPos <> 0 Then
			lgDefItemAcct = Left(arrParam(3), IPos - 1)
			
			If IPos < Len(arrParam(3)) Then
				lgDefProcType = Mid(arrParam(3),IPos+1,Len(arrParam(3)))
			End If
		Else
			lgDefItemAcct = arrParam(3)
		End If
		
	End If                      
	
	'-------------------------------
	' Date Default Setting
	'-------------------------------   
    If arrParam(4) <> "" Then
		txtBaseDt.text = arrParam(4)
    Else
		txtBaseDt.text = StartDate
    End If    
	
	'-------------------------------
	' Where Condition Setting
	'-------------------------------
	If Ubound(arrParam) > 4 Then
		If Trim(arrParam(5)) <> "" Then
			lgWhere = arrParam(5)
		Else
			lgWhere = ""	
		End If
	Else 
		lgWhere = ""		
	End If	
	
	hPlantCd.value = arrParam(0)
	hItemCd.value = arrParam(1)
	hBaseDt.value = txtBaseDt.text
	hTrackingFlg.value = "%"
	

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i
	    
	Call InitSpreadPosVariables()

    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20050101",, PopupParent.gAllowDragDropSpread

    vspdData.ReDraw = False

'    vspdData.OperationMode = 3

    vspdData.MaxCols = C_ItemAcctGrp + 1
    vspdData.MaxRows = 0
	    
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_ItemCd,			"품목",			16												
	ggoSpread.SSSetEdit C_ItemNm,			"품목명",		25												
	ggoSpread.SSSetEdit C_Spec,				"규격",			25											
	ggoSpread.SSSetEdit C_BasicUnit,		"단위",			8											
	ggoSpread.SSSetEdit C_ItemAcct,			"품목계정",		10											
	ggoSpread.SSSetEdit C_ItemAcctNm,		"품목계정",		14											
	ggoSpread.SSSetEdit C_ItemGroupCd,		"품목그룹",		12											
	ggoSpread.SSSetEdit C_ItemClass,		"집계용품목클래스",14
	ggoSpread.SSSetEdit C_ProcurType,		"조달구분",		10										
	ggoSpread.SSSetEdit C_ProcurTypeNm,		"조달구분",		14										
	ggoSpread.SSSetEdit C_ProdtEnv,			"생산전략",		10											
	ggoSpread.SSSetEdit C_ProdtEnvNm,		"생산전략",		10											
	ggoSpread.SSSetEdit C_PhantomFlg,		"팬텀",			8,2										
	ggoSpread.SSSetEdit C_LotFlg,			"LOT관리",		10,2					
	ggoSpread.SSSetEdit C_MajorSlCd,		"입고창고",		10										
	ggoSpread.SSSetEdit C_IssuedSlCd,		"출고창고",		10										
	ggoSpread.SSSetEdit C_ValidFlg,			"유효구분",		10,2
	ggoSpread.SSSetEdit C_ValidFromDt,		"시작일",		10,2
	ggoSpread.SSSetEdit C_ValidToDt,		"종료일",		10,2
	ggoSpread.SSSetEdit C_FormalNm,			"품목정식명칭",	25										                        
	ggoSpread.SSSetEdit C_BItemItemAcct,	"품목계정",		10									
	ggoSpread.SSSetEdit C_HsCd,				"HS코드",		10												
	ggoSpread.SSSetEdit C_HsUnit,			"HS단위",		10												
	ggoSpread.SSSetEdit C_BaseItemCd,		"기준품목",		16										
	ggoSpread.SSSetEdit C_TrackingFlg,		"Tracking 구분",10									
	ggoSpread.SSSetEdit C_OrderUnitMfg,		"제조오더단위",	10										
	ggoSpread.SSSetEdit C_OrderUnitPur,		"구매오더단위",	10
	ggoSpread.SSSetEdit C_OrderLtMfg,		"제조오더L/T",	10											
	ggoSpread.SSSetEdit C_OrderLtPur,		"구매오더L/T",	10
	ggoSpread.SSSetEdit C_OrderType,		"오더타입",		10											
	ggoSpread.SSSetEdit C_OrderRule,		"Lot Sizing",		10			
	ggoSpread.SSSetEdit C_FixedMrpQty,		"고정수배수",	10	
	ggoSpread.SSSetEdit C_MinMrpQty,		"최소수배수",	10										
	ggoSpread.SSSetEdit C_MaxMrpQty,		"최대수배수",	10										
	ggoSpread.SSSetEdit C_RoundQty,			"올림수",		10											
	ggoSpread.SSSetEdit C_RoundPerd,		"올림기간",		10											
	ggoSpread.SSSetEdit C_MpsFlag,			"MPS구분",		10					
	ggoSpread.SSSetEdit C_IssueMthd,		"출고방법",		10				
	ggoSpread.SSSetEdit C_PurOrg,			"구매조직",		10				
	ggoSpread.SSSetEdit C_OptionFlg,		"OPTION 구분",	10			
	ggoSpread.SSSetEdit C_CycleCntPerd,		"실사주기",		10			
	ggoSpread.SSSetEdit C_IssuedUnit,		"출고단위",		10			
	ggoSpread.SSSetEdit C_RecvInspecFlg,	"수입검사구분",	10				    							
	ggoSpread.SSSetEdit C_ProdInspecFlg,	"공정검사구분",	10					    							
	ggoSpread.SSSetEdit C_FinalInspecFlg,	"최종검사구분",	10												
	ggoSpread.SSSetEdit C_ShipInspecFlg,	"출하검사구분",	10					    
	ggoSpread.SSSetEdit C_InspecLtMfg,		"제조검사L/T",	10				
	ggoSpread.SSSetEdit C_InspecLtPur,		"구매검사L/T",	10
	ggoSpread.SSSetEdit C_InspecMgr,		"검사담당자",	10			
	ggoSpread.SSSetEdit C_BItemValidFlg,	"품목유효구분",	10
	ggoSpread.SSSetEdit C_CollectiveFlg,	"단공정여부",	10
	ggoSpread.SSSetEdit C_MajorWorkCenter,	"주작업장",		10
	ggoSpread.SSSetEdit C_AbcFlg,			"ABC구분",		10		
	ggoSpread.SSSetEdit C_ItemAcctGrp,		"품목계정그룹", 10												
		
	Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
	Call ggoSpread.SSSetColHidden(C_ProcurType, C_ProcurType, True)
	Call ggoSpread.SSSetColHidden(C_ProdtEnv, C_ProdtEnv, True)
	Call ggoSpread.SSSetColHidden(C_FormalNm, C_ItemAcctGrp, True)
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)

	ggoSpread.SSSetSplit2(1)										'frozen 기능추가 

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

			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Spec				= iCurColumnPos(3)
			C_BasicUnit			= iCurColumnPos(4)
			C_ItemAcct			= iCurColumnPos(5)
			C_ItemAcctNm		= iCurColumnPos(6)
			C_ItemGroupCd		= iCurColumnPos(7)
			C_ItemClass			= iCurColumnPos(8)
			C_ProcurType		= iCurColumnPos(9)
			C_ProcurTypeNm		= iCurColumnPos(10)
			C_ProdtEnv			= iCurColumnPos(11)
			C_ProdtEnvNm		= iCurColumnPos(12)
			C_PhantomFlg		= iCurColumnPos(13)
			C_LotFlg			= iCurColumnPos(14)
			C_MajorSlCd			= iCurColumnPos(15)
			C_IssuedSlCd		= iCurColumnPos(16)
			C_ValidFlg			= iCurColumnPos(17)
			C_ValidFromDt		= iCurColumnPos(18)
			C_ValidToDt			= iCurColumnPos(19)
			C_FormalNm			= iCurColumnPos(20)
			C_BItemItemAcct		= iCurColumnPos(21)
			C_HsCd				= iCurColumnPos(22)
			C_HsUnit			= iCurColumnPos(23)
			C_BaseItemCd		= iCurColumnPos(24)
			C_TrackingFlg		= iCurColumnPos(25)
			C_OrderUnitMfg		= iCurColumnPos(26)
			C_OrderUnitPur		= iCurColumnPos(27)
			C_OrderLtMfg		= iCurColumnPos(28)
			C_OrderLtPur		= iCurColumnPos(29)
			C_OrderType			= iCurColumnPos(30)
			C_OrderRule			= iCurColumnPos(31)
			C_FixedMrpQty		= iCurColumnPos(32)
			C_MinMrpQty			= iCurColumnPos(33)
			C_MaxMrpQty			= iCurColumnPos(34)
			C_RoundQty			= iCurColumnPos(35)
			C_RoundPerd			= iCurColumnPos(36)
			C_MpsFlag			= iCurColumnPos(37)
			C_IssueMthd			= iCurColumnPos(38)
			C_PurOrg			= iCurColumnPos(39)
			C_OptionFlg			= iCurColumnPos(40)
			C_CycleCntPerd		= iCurColumnPos(41)
			C_IssuedUnit		= iCurColumnPos(42)
			C_RecvInspecFlg		= iCurColumnPos(43)
			C_ProdInspecFlg		= iCurColumnPos(44)
			C_FinalInspecFlg	= iCurColumnPos(45)
			C_ShipInspecFlg		= iCurColumnPos(46)
			C_InspecLtMfg		= iCurColumnPos(47)
			C_InspecLtPur		= iCurColumnPos(48)
			C_InspecMgr			= iCurColumnPos(49)
			C_BItemValidFlg		= iCurColumnPos(50)
			C_CollectiveFlg		= iCurColumnPos(51)
			C_MajorWorkCenter	= iCurColumnPos(52)
			C_AbcFlg			= iCurColumnPos(53)
			C_ItemAcctGrp		= iCurColumnPos(54)
    End Select    
End Sub

Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(txtItemGroupCd.className) = UCase(PopUpParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	txtItemGroupCd.focus
 
End Function


'=========================================================================================================
Function SetItemGroup(byval arrRet)
	txtItemGroupCd.Value    = arrRet(0)  
	txtItemGroupNm.Value    = arrRet(1)  
End Function

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
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000111111")

    If vspdData.MaxRows <= 0 Then Exit Sub
   	    
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' 타이틀 cell을 dblclick했거나....
	   Exit Function
	End If

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	If KeyAscii=27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 and vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				'DbQuery
				If DbQuery = False Then
					Exit Sub
				End If
			End If
		End If
	End With
End Sub
	

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			'DbQuery
			If DbQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub

'========================================================================================================
'	Name : FncQuery
'	Desc : 
'========================================================================================================
Function FncQuery()

	FncQuery = False

	Call InitVariables()
			
	vspdData.MaxRows = 0						'Grid 초기화 
		
	lgIntFlgMode = PopupParent.OPMD_CMODE	

	If DbQuery = False Then
		Exit Function
	End If
	
	FncQuery = True

	hPlantCd.value =Trim(UCase(PlantCd))
	hItemCd.value =Trim(UCase(txtItemCd.value))
	hItemNm.value = Trim(txtItemNm.value)
	hItemAccount.value = Trim(cboItemAccount.value)
	hItemClass.value = Trim(cboItemClass.value)
	hProcurType.value = Trim(cboProcurType.value)
	hProdEnv.value = Trim(cboProdtEnv.value)
	hSpec.value = Trim(txtItemSpec.value)
	hBaseDt.value = txtBaseDt.Text
	hInspecFlg.value = Trim(cboInspType.value)
	
	If rdoTrackingItem1.checked = True Then
		hTrackingFlg.value = "Y"
	ElseIf rdoTrackingItem2.checked = True Then
		hTrackingFlg.value = "N"
	Else  
		hTrackingFlg.value = "%"
	End If
		
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    vspdData.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    'Err.Clear                                                               '☜: Protect system from crashing
	
    DbQuery = False                                                         '⊙: Processing is NG
	
	 '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	Call LayerShowHide(1)												<%'⊙: 작업진행중 표시 %>	
	    
    Dim strVal
         
  	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001												'☜: 

	If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		strVal = strVal & "&PlantCd=" & PlantCd									'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)					'☆: 조회 조건 데이타 
		strVal = strVal & "&txtitemNm=" & txtItemNm.value
		strVal = strVal & "&cboItemAccount=" & Trim(cboItemAccount.value)
		strVal = strVal & "&cboItemClass=" & Trim(cboItemClass.value)
		strVal = strVal & "&cboProcurType=" & Trim(cboProcurType.value)
		strVal = strVal & "&cboProdtEnv=" & Trim(cboProdtEnv.value)
		strVal = strVal & "&txtItemSpec=" & txtItemSpec.value
		strVal = strVal & "&cboInspType=" & Trim(cboInspType.value)
	
		If rdoTrackingItem1.checked = True Then
			strVal = strVal & "&rdoTrackingItem=Y" 
		ElseIf rdoTrackingItem2.checked = True Then
			strVal = strVal & "&rdoTrackingItem=N"
		Else  
			strVal = strVal & "&rdoTrackingItem=%"
		End If
		strVal = strVal & "&lgCurDate=" & txtBaseDt.Text
		strVal = strVal & "&pType="
		strVal = strVal & "&txtItemGroupCd=" & Trim(txtItemGroupCd.value)
	Else
		strVal = strVal & "&PlantCd=" & hPlantCd.value
		strVal = strVal & "&txtItemCd=" & hItemCd.value
		strVal = strVal & "&txtitemNm=" & hItemNm.value
		strVal = strVal & "&strNextKey=" & strNextKey
		strVal = strVal & "&cboItemAccount=" & hItemAccount.value
		strVal = strVal & "&cboItemClass=" & hItemClass.value
		strVal = strVal & "&cboProcurType=" & hProcurType.value
		strVal = strVal & "&cboProdtEnv=" & hProdEnv.value
		strVal = strVal & "&txtItemSpec=" & hSpec.value
		strVal = strVal & "&cboInspType=" & hInspecFlg.value
		strVal = strVal & "&rdoTrackingItem=" & hTrackingFlg.value
		strVal = strVal & "&lgCurDate=" & hBaseDt.value
		strVal = strVal & "&pType=" & hpType.value
		strVal = strVal & "&txtItemGroupCd=" & Trim(hItemGroupCd.value)
	End If
		
	If lgItemAcctGrp(0) = "Y" Then
		strVal = strVal & "&FromItemAcctGrp=" & lgItemAcctGrp(1)
		strVal = strVal & "&ToItemAcctGrp=" & lgItemAcctGrp(2)
	Else
		strVal = strVal & "&FromItemAcctGrp=" & ""
		strVal = strVal & "&ToItemAcctGrp=" & "zz"
	End If

	If lgProcType(0) = "Y" Then
		strVal = strVal & "&FromProcType=" & lgProcType(1)
		strVal = strVal & "&ToProcType=" & lgProcType(2)
	Else
		strVal = strVal & "&FromProcType=" & ""
		strVal = strVal & "&ToProcType=" & "zz"
	End If

	strVal = strVal & "&txtWhere=" & lgWhere
	If arrParam(4) <> "" Then
		strVal = strVal & "&txtAssignDtFlag=" & "Y"
	Else
		strVal = strVal & "&txtAssignDtFlag=" & "N"	
	End If	
	strVal = strVal & "&txtMaxRows="         & vspdData.MaxRows
	strVal = strVal & "&lgStrPrevKeyIndex="  & lgStrPrevKeyIndex    '☜: Max fetched data at a time
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
		
    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
	End If
		
	lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
End Function

'========================================================================================================
'	Name : OKClick
'	Desc : 
'========================================================================================================
Function OKClick()
	Dim i, iCurColumnPos
	
	If vspdData.MaxRows > 0 Then
		
		Redim arrReturn(UBound(arrField))

        ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 

msgbox "UBound(arrField) : " & UBound(arrField)			
		For i = 0 To UBound(arrField)
			If arrField(i) <> "" Then
				vspddata.Col = iCurColumnPos(CInt(arrField(i)))
				arrReturn(i) = vspdData.Text
			End If
		Next

		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
					
End Function

'========================================================================================================
'	Name : CancelClick
'	Desc : 
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : MousePointer
'	Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        txtBaseDt.Action = 7
        Call SetFocusToDocument("P")
		txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	
	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	
	Call SetDefaultVal()
	Call InitVariables
	Call InitComboBox()
	
	Call InitSpreadSheet()
	
	If FncQuery = False Then
		Exit Sub
	End If
	
End Sub

</SCRIPT>
<!-- #Include file="../inc/Uni2kCMCom.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
-->
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목</TD>
				<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=45 MAXLENGTH=40 tag="11" ALT="품목명"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>품목계정</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAccount" ALT="품목계정" STYLE="Width: 160px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
				<TD CLASS=TD5 NOWRAP>품목그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5 NOWRAP>조달구분</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcurType" ALT="조달구분" STYLE="Width: 160px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
				<TD CLASS=TD5 NOWRAP>집계용품목클래스</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="집계용품목클래스" STYLE="Width: 160px;" tag="11XXXU"><OPTION VALUE = ""></OPTION></SELECT></TD>
			</TR>		
			<TR>
				<TD CLASS=TD5 NOWRAP>규격</TD>
				<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=40 tag="11" ALT="규격">&nbsp;</TD>
				<TD CLASS=TD5 NOWRAP>생산전략</TD>
				<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdtEnv" ALT="생산전략" STYLE="Width: 160px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>검사분류</TD>
				<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspType" ALT="검사분류" STYLE="WIDTH: 160px" TAG="11"><OPTION VALUE="" selected></OPTION></SELECT></TD>										
				<TD CLASS=TD5 NOWRAP>기준일</TD>
				<TD CLASS=TD6 NOWRAP>
					<script language =javascript src='./js/b1b11pa3_OBJECT1_txtBaseDt.js'></script> 															
				</TD>
			</TR>
				<TD CLASS="TD5" NOWRAP></TD>
				<TD CLASS="TD6" NOWRAP></TD>
				<TD CLASS=TD5 NOWRAP>Tracking여부</TD>
				<TD CLASS=TD6 NOWRAP>
					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="11" ID="rdoTrackingItem1" VALUE="Y"><LABEL FOR="rdoTrackingItem1">예</LABEL>
					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="11" ID="rdoTrackingItem2" VALUE="N"><LABEL FOR="rdoTrackingItem2">아니오</LABEL>
					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="11" CHECKED ID="rdoTrackingItem3" VALUE="%"><LABEL FOR="rdoTrackingItem3">전체</LABEL></TD>
			<TR>
			</TR>
		</TABLE></FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/b1b11pa3_I213456851_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"	SRC="../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemNm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemAccount" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemClass" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hProcurType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hProdEnv" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hSpec" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBaseDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspecFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTrackingFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hpType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>