<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		: INTERFACE 
'*  2. Function Name		: 
'*  3. Program ID		: xi218ma1_KO441
'*  4. Program Name		: 작업지시전송현황 
'*  5. Program Desc		:
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2006/04/18
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First) 	: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				:
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "XI218MB1_KO441.asp"                               'Biz Logic ASP
Const BIZ_PGM_JUMP_ID = "Q2317MA1"									  '☆: Jump ASP명  
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop          
Dim lsConcd
Dim IsOpenPop     

Dim lsConPlantcd
Dim lsConItemCd
Dim lsConInspReqNo

Dim C_ItemCd
Dim C_ItemNm
Dim C_ProdtOrderNo
Dim C_LotNo
Dim C_RequestSeq
Dim C_IfSeq
Dim C_InspDt
Dim C_Decision
Dim C_InspQty
Dim C_DefectQty
Dim C_CreateType
Dim C_InspReqNo
Dim C_SendDt
Dim C_MesReceiveFlag
Dim C_ErrDesc
Dim C_MesReceiveDt

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
	C_ProdtOrderNo			= 1
	C_ItemCd				= 2
	C_ItemNm				= 3
	C_LotNo					= 4
	C_RequestSeq			= 5
	C_IfSeq					= 6
	C_InspDt				= 7
	C_Decision				= 8
	C_InspQty				= 9
	C_DefectQty				= 10
	C_CreateType			= 11
	C_InspReqNo				= 12
	C_SendDt				= 13
	C_MesReceiveFlag		= 14
	C_ErrDesc				= 15
	C_MesReceiveDt			= 16
	
end sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtStart_dt.Focus	
	frm1.txtStart_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtStart_dt.Month = strMonth 
	frm1.txtStart_dt.Day = strDay
	frm1.txtEnd_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtEnd_dt.Month = strMonth 
	frm1.txtEnd_dt.Day = strDay
	frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("M", 0, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("M", 0, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)

End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

Sub InitComboBox()
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))
    
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	Call vspdData_Click(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)

	If flgs = 1 Then
		WriteCookie "txtPlantcd", lsConPlantcd
		WriteCookie "txtItemcd", lsConItemCd
		WriteCookie "txtInspReqNo", lsConInspReqNo
	End If

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	Dim changeTxtEnd_dt
	changeTxtEnd_dt = Trim(Frm1.txtEnd_dt.text)
	changeTxtEnd_dt = UniConvDateAToB(UNIDateAdd ("D", 1, changeTxtEnd_dt, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)

	Dim changeTxtProdtODt
	changeTxtProdtODt = Trim(Frm1.txtProdtODt.text)
	changeTxtProdtODt = UniConvDateAToB(UNIDateAdd ("D", 1, changeTxtProdtODt, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)


    lgKeyStream = Trim(Frm1.txtPlantCd.Value) & parent.gColSep              
    lgKeyStream = lgKeyStream & Trim(Frm1.txtStart_dt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEnd_dt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtItemCd.Value) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtProdFromDt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtProdToDt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtProdOrderNo.Value) & parent.gColSep

	If frm1.txtMesReceivetFlg(0).checked Then 
		lgKeyStream = lgKeyStream & "A" & parent.gColSep
	Elseif frm1.txtMesReceivetFlg(1).checked Then 
		lgKeyStream = lgKeyStream & "Y" & parent.gColSep
	Elseif frm1.txtMesReceivetFlg(2).checked Then 
		lgKeyStream = lgKeyStream & "N" & parent.gColSep
	Else
		lgKeyStream = lgKeyStream & "A" & parent.gColSep		
	End if


    'msgbox(lgKeyStream)

End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  
		
	With frm1.vspdData
	
        ggoSpread.Source = Frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false

       .MaxCols = C_MesReceiveDt + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
    
       .MaxRows = 0	
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData        
		Call GetSpreadColumnPos("A")       
        Call  AppendNumberPlace("6","7","2")


		ggoSpread.SSSetEdit		C_ProdtOrderNo,		"제조오더번호",			15,,,18
		ggoSpread.SSSetEdit		C_ItemCd,			"품목",					15,,,18,2
		ggoSpread.SSSetEdit		C_ItemNm,			"품목명",				20,,,40
		ggoSpread.SSSetEdit		C_LotNo,			"LOT번호",				15,,,25
		ggoSpread.SSSetEdit		C_RequestSeq,		"검사의뢰순번",			10,2,,5
		ggoSpread.SSSetEdit		C_IfSeq,			"I/F순번",				 6,2,,5
		
		ggoSpread.SSSetDate		C_InspDt,			"검사일",				10,2,parent.gDateFormat
		ggoSpread.SSSetEdit		C_Decision,			"검사결과",				 8,2,,1
		ggoSpread.SSSetFloat	C_InspQty,			"검사수량",				10,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_DefectQty,		"불량수량",				10,Parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_CreateType,		"생성구분",				 8,2,,1
		ggoSpread.SSSetEdit		C_InspReqNo,		"검사의뢰번호",			15,,,18
		ggoSpread.SSSetEdit		C_SendDt,			"ERP 최종송신일시",		20, , ,30
		ggoSpread.SSSetEdit		C_MesReceiveFlag,	"MES 반영여부",			12,2,,1
		ggoSpread.SSSetEdit		C_ErrDesc,			"MES 수신 에러내역",	25,,,40
		ggoSpread.SSSetEdit		C_MesReceiveDt,		"MES 최종수신일시",		20, , ,30


		ggoSpread.SSSetSplit2(1)										'frozen 기능추가 


	   .ReDraw = true    
	
       Call SetSpreadLock 
    
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
            
			C_ProdtOrderNo		= iCurColumnPos(1)
			C_ItemCd			= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_LotNo				= iCurColumnPos(4)
			C_RequestSeq		= iCurColumnPos(5)
			C_IfSeq				= iCurColumnPos(6)
			C_InspDt			= iCurColumnPos(7)
			C_Decision			= iCurColumnPos(8)
			C_InspQty			= iCurColumnPos(9)
			C_DefectQty			= iCurColumnPos(10)
			C_CreateType		= iCurColumnPos(11)
			C_InspReqNo			= iCurColumnPos(12)
			C_SendDt			= iCurColumnPos(13)
			C_MesReceiveFlag	= iCurColumnPos(14)
			C_ErrDesc			= iCurColumnPos(15)
			C_MesReceiveDt		= iCurColumnPos(16)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
 	
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
 	Call InitComboBox

    Call InitVariables                                                              'Initializes local global variables
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
   
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
    Call CookiePage (0)                                                             '☜: Check Cookie
	
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

End Sub



'------------------------------------------  OpenPlantCd()  --------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_Plant"						
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "공장"						

    arrField(0) = "Plant_Cd"					
    arrField(1) = "Plant_NM"					
    
    arrHeader(0) = "공장"					
    arrHeader(1) = "공장명"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -----------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
    
	'arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(window.parent, arrParam, arrField, arrHeader), _
	'	"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function


'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
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
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	'arrParam(3) = frm1.cboOrderStatus.value
	'arrParam(4) = frm1.cboOrderStatus.value
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	'arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	'arrParam(8) = Trim(frm1.cboOrderType.value)
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'=======================================================================================================
'   Event Name : txtProdFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtProdToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtProdFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtProdToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    If  ValidDateCheck(frm1.txtStart_dt, frm1.txtEnd_dt)=False Then
        Exit Function
    End If
    If  ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt)=False Then
        Exit Function
    End If
     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
                                                         '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
 
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
  
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function


'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
 	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
	Frm1.vspdData.focus	
End Function
 

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim intIndex
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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

    frm1.vspdData.Row = Row

    lsConPlantcd	  = frm1.txtPlantCd.value
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 2
    lsConItemCd		  = frm1.vspdData.Text  

    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 12
    lsConInspReqNo  = frm1.vspdData.Text      		
       
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
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


'========================================================================================================
' Name : txtStart_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtStart_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtStart_dt.Action = 7 
		frm1.txtStart_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtEnd_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtEnd_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtEnd_dt.Action = 7 
        frm1.txtEnd_dt.focus
    End If
End Sub


'==========================================================================================
'   Event Name : txtStart_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtStart_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : txtEnd_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtEnd_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			                <TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ERP 송신기간</TD>
							    	<TD CLASS="TD6" NOWRAP>
							    		<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD  name=txtStart_dt classid=<%=gCLSIDFPDT%> ALT="시작일" tag="12X1X" VIEWASTEXT></OBJECT>');</SCRIPT>
	                                    &nbsp;~&nbsp;
	                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEnd_dt classid=<%=gCLSIDFPDT%>ALT="종료일" tag="12X1X" VIEWASTEXT></OBJECT>');</SCRIPT>
							    	</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="품목명"></TD>
									<TD CLASS=TD5 NOWRAP>검사기간</TD>
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT classid="CLSID:DD55D13A-EBF7-11D0-8810-0000C0E5948C" name=txtProdFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작일" id=fpDateTime1></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT classid="CLSID:DD55D13A-EBF7-11D0-8810-0000C0E5948C" name=txtProdtODt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료일" id=fpDateTime1></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
									<TD CLASS=TD5 NOWRAP>MES 수신여부</TD>
									<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg1" Value="A" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoValidFlg1">전체</LABEL>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg2" Value="Y" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg2">성공</LABEL>
												<INPUT TYPE="RADIO" NAME="txtMesReceivetFlg" ID="rdoValidFlg3" Value="N" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg3">실패</LABEL></TD>
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
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">최종검사Release</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>  		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



