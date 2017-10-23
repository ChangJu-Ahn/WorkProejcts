<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : p2351ma1.asp
*  4. Program Name         : MRP예시전개전환 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/10/09
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Im Hyun Soo
* 10. Modifier (Last)      : Jung Yu Kyung	
* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
Const BIZ_PGM_ID = "p2351mb1.asp"
Const BIZ_PGM_CONVPAR_ID ="p2351mb2.asp"
'========================================================================================================
Const CookieSplit = 1233

Dim C_Select
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec		
Dim C_StartDt
Dim C_EndDt	
Dim C_PlanQty
Dim C_Unit	
Dim C_ProcTypeNm
Dim C_Status	
Dim C_TrackingNo
Dim C_ProdMgrNm	
Dim C_PurOrg	
Dim C_ProdMgr	
Dim C_ProcType	
Dim C_Seq		

Const COOKIE_SPLIT      = 4877	                                      'Cookie Split String

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop  
Dim lgButtonSelection
Dim lgSelRows
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
    C_Select		= 1
    C_ItemCd		= 2
    C_ItemNm		= 3
    C_Spec			= 4
    C_StartDt		= 5
    C_EndDt			= 6
    C_PlanQty		= 7
    C_Unit			= 8
    C_ProcTypeNm	= 9
    C_Status		= 10
    C_TrackingNo	= 11    
    C_ProdMgrNm		= 12
    C_PurOrg		= 13
    C_ProdMgr		= 14
    C_ProcType		= 15
    C_Seq			= 16

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = ""
    lgStrPrevKeyIndex1 = ""
    lgSortKey         = 1
    lgSelRows		  = 0	
	lgButtonSelection = "DESELECT"
	frm1.btnSelect1.disabled = True
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtBaseFromDt.Text = StartDate
	frm1.txtBaseToDt.Text = UNIDateAdd("M", 1, StartDate, parent.gDateFormat)
	frm1.btnSelect1.disabled = True
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
	If Kubun = 0 Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		
		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""
		
	End If
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
	lgKeyStream = ""
	
	Select Case pRow
		Case "P"
			lgKeyStream = frm1.txtPlantCd.Value & parent.gColSep
		Case "Q"
			lgKeyStream = frm1.txtPlantCd.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.txtItemCd.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.cboProcType.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.txtBaseFromDt.text & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.txtBaseToDt.text & parent.gColSep
		Case "R"
			lgKeyStream = frm1.hPlantCd.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.hItemCd.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.hProcType.Value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.hBaseFromDt.value & parent.gColSep
			lgKeyStream = lgKeyStream & frm1.hBaseToDt.value & parent.gColSep	
	End Select
					
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

    Dim iCodeArr 
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.cboProcType ,iCodeArr, iNameArr,Chr(11))
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()    
	
    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    
    
		.Redraw = False
    
		.MaxCols = C_Seq + 1
		.MaxRows = 0
    
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck	C_Select,		"", 2,,,1 		    
		ggoSpread.SSSetEdit		C_ItemCd, 		"품목"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		, 25
		ggoSpread.SSSetEdit		C_Spec,			"규격"			, 25
		ggoSpread.SSSetDate 	C_StartDt, 		"시작일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetDate 	C_EndDt, 		"완료일"		, 11, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat	C_PlanQty, 		"계획수량"		, 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_Unit, 		"단위"			, 7
		ggoSpread.SSSetEdit 	C_ProcTypeNm, 	"조달구분"		, 10
		ggoSpread.SSSetEdit 	C_Status, 		"Status"		, 6
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No."	, 25
		ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자"	, 10
		ggoSpread.SSSetEdit		C_PurOrg,		"구매조직"		, 10
		ggoSpread.SSSetEdit		C_ProdMgrNm,	"생산담당자"	, 10
	    ggoSpread.SSSetEdit		C_ProdMgr,		"생산담당자코드", 10
	    ggoSpread.SSSetEdit		C_ProcType,		"조달구분코드"	, 6
		ggoSpread.SSSetEdit		C_Seq,			"SEQ"			, 6
		
		Call ggoSpread.SSSetColHidden(C_Seq, C_Seq, True)
		Call ggoSpread.SSSetColHidden(C_ProdMgr, C_ProdMgr, True)
		Call ggoSpread.SSSetColHidden(C_ProcType, C_ProcType, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SSSetSplit2(2)
    
		ggoSpread.Source = frm1.vspdData
		.ReDraw = True
    
   End With
    
    Call SetSpreadLock()
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1

    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock -1, -1
	.vspdData.ReDraw = True
	
   End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of error
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = iPosArr(0)
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0
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

			C_Select		= iCurColumnPos(1)
			C_ItemCd		= iCurColumnPos(2)
			C_ItemNm		= iCurColumnPos(3)    
			C_Spec			= iCurColumnPos(4)
			C_StartDt		= iCurColumnPos(5)
			C_EndDt			= iCurColumnPos(6)
			C_PlanQty		= iCurColumnPos(7)
			C_Unit			= iCurColumnPos(8)
			C_ProcTypeNm	= iCurColumnPos(9)
			C_Status		= iCurColumnPos(10)
			C_TrackingNo	= iCurColumnPos(11)			
			C_ProdMgrNm		= iCurColumnPos(12)
			C_PurOrg		= iCurColumnPos(13)
			C_ProdMgr		= iCurColumnPos(14)
			C_ProcType		= iCurColumnPos(15)
			C_Seq			= iCurColumnPos(16)
	
    End Select    

End Sub
'------------------------------------------  BatchSelect()  -----------------------------------------
Function btnAutoSel_onClick()

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "전체선택"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "전체선택취소"
	End If

	Dim index,Count
	Dim strStatus
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 

	For index = 1 to Count
		
		frm1.vspdData.Row = index
		strStatus = GetSpreadValue(frm1.vspdData,C_ProcType,index,"X","X")
		frm1.vspdData.Col = C_Select

		If lgButtonSelection = "SELECT" and Trim(strStatus) = "P" Then 
			frm1.vspdData.Value = 1
			frm1.vspdData.Col = 0 
			'ggoSpread.UpdateRow Index
		Else
			frm1.vspdData.Value = 0
			frm1.vspdData.Col = 0 
			frm1.vspdData.Text=""
		End if

	Next 
	
	frm1.vspdData.ReDraw = true

End Function
'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    
    Err.Clear
	Call LoadInfTB19029
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
            
    Call InitSpreadSheet

	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("11000000000011")
	Call CookiePage (0)
	
	If parent.gPlant <> "" And Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
	ElseIf Trim(frm1.txtPlantCd.value) <> "" Then
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement
	
    Call InitComboBox
    			
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    															
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtBaseFromDt, frm1.txtBaseToDt)  = False Then		
		Exit Function
	End If 

    Call InitVariables
    Call MakeKeyStream("Q")

    If DbQuery = False Then
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True
End Function
	
'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True     
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	Dim ChkFlg

    FncCancel = False
    Err.Clear
    
    If frm1.vspdData.MaxRows <= 0 Then
		Exit Function
	End If
	 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	
    Set gActiveElement = document.ActiveElement   
    FncCancel = False
    
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
    
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False
    Err.Clear

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False
    Err.Clear

	Call Parent.FncFind(parent.C_MULTI, True)

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

End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False
    Err.Clear
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

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
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtSpreadPos="       & "0"
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
        strVal = strVal     & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		
    End With

    Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery2()
	Dim strVal

    Err.Clear
    DbQuery2 = False
		
    Call LayerShowHide(1)
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtSpreadPos="       & "1"
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream 
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    End With
		
    Call RunMyBizASP(MyBizASP, strVal) 
	
    DbQuery2 = True

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
	Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim iColSep, iRowSep
	Dim arrVal
	ReDim arrVal(0)
	
	iColSep = parent.gColSep
	iRowSep = parent.gRowSep
	
    Err.Clear 
    DbSave = False 
    Call LayerShowHide(1) 
    
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           If GetSpreadText(.vspdData,0,lRow,"X","X") = ggoSpread.UpdateFlag Then
				
			    strVal = ""
				strVal = strVal & "U" & iColSep & lRow & iColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_Seq,lRow,"X","X")) & iColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_StartDt,lRow,"X","X")) & iColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_EndDt,lRow,"X","X")) & iColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlanQty,lRow,"X","X")) & iRowSep
				
				ReDim Preserve arrVal(lGrpCnt)
				arrVal(lGrpCnt) = strVal
				
				lGrpCnt = lGrpCnt + 1
			
			End If           
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = Join(arrVal,"")

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk(ByVal lgLngMaxRow)
	
    Call SetToolbar("11001001000111")
    
    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
	
    With frm1.vspdData

 		Dim lRow
		Dim strStatus

		.ReDraw = False

		For lRow = lgLngMaxRow To .MaxRows

			strStatus = GetSpreadValue(frm1.vspdData,C_ProcType,lRow,"X","X")

			If Trim(strStatus) = "P" Or Trim(strStatus) = "O" Then
				ggoSpread.SpreadUnLock C_Select,	lRow, C_Select, lRow
				ggoSpread.SpreadUnLock C_StartDt,	lRow, C_StartDt,lRow
				ggoSpread.SpreadUnLock C_EndDt,		lRow, C_EndDt,	lRow
				ggoSpread.SpreadUnLock C_PlanQty,	lRow, C_PlanQty,lRow
				ggoSpread.SSSetRequired C_PlanQty,		lRow, lRow
				ggoSpread.SSSetRequired C_StartDt,		lRow, lRow
				ggoSpread.SSSetRequired C_EndDt,		lRow, lRow
			End If
			
		Next
	
		.ReDraw = True
		Set gActiveElement = document.ActiveElement   
		.Focus

    End With
    
	frm1.btnSelect1.disabled = False
	frm1.btnAutoSel.disabled = False
    
End Sub
	
Sub DBQueryNotOk()
    Call SetToolbar("11000000000011")
End Sub

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables
	Call ggoOper.ClearField(Document, "2")

    Call DisableToolBar(parent.TBC_QUERY)
    If DBQuery = False Then 
       Call RestoreToolBar()
       Exit Sub
    End If 
    Set gActiveElement = document.ActiveElement   
    
End Sub

Function Transfer()

    Dim lRow        
	Dim strVal
	Dim IntRetCD
	Dim arrVal
	ReDim arrVal(0)

	If lgSelRows = 0 Then
		IntRetCD = DisplayMsgBox("181216", "X", "X", "X")
		Exit Function
	Else
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If
	End If
	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	Else
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
        
    Call LayerShowHide(1)
    
	With frm1
		.txtMode.value = parent.UID_M0002
    
		For lRow = 1 To .vspdData.MaxRows
			
		    Select Case GetSpreadValue(.vspdData,C_Select,lRow,"X","X")

		        Case 1
					strVal = ""
		            strVal = strVal & Trim(GetSpreadText(.vspdData,C_Seq,lRow,"X","X")) & parent.gRowSep
		            
		            ReDim Preserve arrVal(lRow)
		            arrVal(lRow) = strVal
		
		    End Select
		            
		Next
	
		.txtSpread.value = Join(arrVal,"")

		Call ExecMyBizASP(frm1, BIZ_PGM_CONVPAR_ID)

	End With

End Function

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""			' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
End Function


 '------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If lgIsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	lgIsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0):"ITEM_CD"
    arrField(1) = 2 							' Field명(1):"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet)
	End If	

End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(ByRef arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus
    Set gActiveElement = document.activeElement		
End Function

 '------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement		
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row )
    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	If Col <> C_Select Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
	End If
	
End Sub
   	
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If
	
	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
  
    With frm1.vspdData
		
		If Row <= 0 Then Exit Sub

		If Col = C_Select Then
			.Row = Row
			.Col = C_Select
			
			If Buttondown = 1 Then
				lgSelRows = lgSelRows + 1
			Else
				If lgSelRows - 1 < 0 Then
					lgSelRows = 0 
				Else
					lgSelRows = lgSelRows - 1
				End If
			End If

		End If
	End With
		
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then 
    		Call MakeKeyStream("R")                      
      		Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery2 = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
    	End If
    End if
End Sub


'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtBaseFromDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtBaseFromDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtBaseToDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtBaseToDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
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

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()

	With frm1.vspdData

 	Dim lRow
	Dim strStatus

	.ReDraw = False

	For lRow = 1 To .MaxRows

		strStatus = GetSpreadValue(frm1.vspdData,C_ProcType,lRow,"X","X")
		
		If Trim(strStatus) = "P" Then
			ggoSpread.SpreadUnLock C_Select, lRow, C_Select,lRow
			ggoSpread.SpreadUnLock C_StartDt, lRow, C_StartDt, lRow
			ggoSpread.SpreadUnLock C_EndDt, lRow, C_EndDt, lRow
			ggoSpread.SpreadUnLock C_PlanQty, lRow, C_PlanQty, lRow
			ggoSpread.SSSetRequired C_PlanQty,		lRow, lRow
			ggoSpread.SSSetRequired C_StartDt,		lRow, lRow
			ggoSpread.SSSetRequired C_EndDt,		lRow, lRow
		End If
		
	Next
	
	.ReDraw = True
	Set gActiveElement = document.ActiveElement   
	
    End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP예시전개전환</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>시작일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p2351ma1_fpDateTime3_txtBaseFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2351ma1_fpDateTime3_txtBaseToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>													
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcType" ALT="조달구분" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>기준일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2351ma1_fpDateTime3_txtFixExecFromDt.js'></script>
								</TD>
								<TD CLASS=TD5 NOWRAP>가용재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoAvailInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg2 VALUE="N"><LABEL FOR=rdoAvailInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>확정전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2351ma1_fpDateTime4_txtFixExecToDt.js'></script>
								</TD>
								<TD CLASS=TD5 NOWRAP>안전재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoSafeInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg2 VALUE="N"><LABEL FOR=rdoSafeInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>예시전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2351ma1_fpDateTime4_txtPlanExecToDt.js'></script>
								</TD>
								<TD CLASS=TD5 NOWRAP>MPS 확정여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMpsConfirmFlg ID=rdoMpsConfirmFlg1 VALUE="%" tag="24X" CHECKED><LABEL FOR=rdoMpsConfirmFlg1>전체</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMpsConfirmFlg ID=rdoMpsConfirmFlg2 VALUE="Y" tag="24X"><LABEL FOR=rdoMpsConfirmFlg2>예</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMpsConfirmFlg ID=rdoMpsConfirmFlg3 VALUE="N" tag="24X"><LABEL FOR=rdoMpsConfirmFlg3>아니오</LABEL></TD>								
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN=4>
									<script language =javascript src='./js/p2351ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
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
					<TD><BUTTON NAME="btnSelect1" CLASS="CLSMBTN" onclick="Transfer()">전환</BUTTON>&nbsp;<BUTTON NAME="btnAutoSel" CLASS="CLSMBTN">전체선택</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtMode"    tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd"    tag="24">
<INPUT TYPE=HIDDEN NAME="hProcType" tag="24"><INPUT TYPE=HIDDEN NAME="hBaseFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hBaseToDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
