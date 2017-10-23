<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3161MA1
'*  4. Program Name         : 재고할당 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*							  
'*  7. Modified date(First) : 2002/11/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho inkuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :     
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit															'☜: Turn on the Option Explicit option.

Const BIZ_PGM_ID				= "s3161mb1.asp"

Dim C_ItemCd					'품목		
Dim C_ItemName					'품목명 
Dim C_ItemSpec					'규격 
Dim C_TrackingNo				'Tracking No
Dim C_SoUnit					'단위 
Dim C_SoQty						'수주량 
Dim C_PreAllocQty				'기할당량 
Dim C_AllocQty					'할당량 
Dim C_BonusQty					'덤수량	
Dim C_PreAllocBonusQty			'기할당덤수량	
Dim C_AllocBonusQty				'할당덤수량 
Dim C_PromiseDt					'출고예정일 
Dim C_DlvyDt					'납기일	
Dim C_SlCd						'창고코드 
Dim C_SlNm						'창고명 
Dim C_PlantCd					'공장코드 
Dim C_PlantNm					'공장명 
Dim C_SoNo						'수주번호 
Dim C_SoSeq						'수주순번 
Dim C_SchdNo					'납품순번 
Dim C_PrePurReqQty				'기구매요청량(Hidden)
Dim C_GiQty     				'출고수량(Hidden)

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop						' Popup
Dim iDBSYSDate
Dim EndDate, StartDate
iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Dim arrValue(3)
Dim lsItemCode
Dim lsSoUnit
Dim lsSoQty
Dim lsPriceQty
Dim lsAPSHost
Dim lsAPSPort
Dim lsCTPTimes
Dim lsCTPCheckFlag
Dim arrCollectVatType

'================================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCd					= 1		'품목		
	C_ItemName					= 2		'품목명 
	C_ItemSpec					= 3		'규격 
	C_TrackingNo				= 4		'Tracking No
	C_SoUnit					= 5		'단위 
	C_SoQty						= 6		'수주량 
	C_PreAllocQty				= 7		'기할당량 
	C_AllocQty					= 8		'할당량 
	C_BonusQty					= 9		'덤수량	
	C_PreAllocBonusQty			= 10	'기할당덤수량	
	C_AllocBonusQty				= 11	'할당덤수량 
	C_PromiseDt					= 12	'출고예정일 
	C_DlvyDt					= 13	'납기일	
	C_SlCd						= 14	'창고코드 
	C_SlNm						= 15	'창고명 
	C_PlantCd					= 16	'공장코드 
	C_PlantNm					= 17	'공장명 
	C_SoNo						= 18	'수주번호 
	C_SoSeq						= 19	'수주순번 
	C_SchdNo					= 20	'납품순번 
	C_PrePurReqQty				= 21	'기구매요청량(Hidden)
	C_GiQty						= 22    '출고수량 
End Sub

'================================================================================================================
Sub InitVariables()

	lgIntFlgMode      = Parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed    
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'================================================================================================================	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iReferArr

    Err.Clear

	Call CommonQueryRs(" REFERENCE ", " B_CONFIGURATION ", " MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF len(lgF0) = 0 Then
		Call ggoOper.SetReqAttr(frm1.txtFromConSoNo, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtToConSoNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtShipToParty, "Q")
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtPlant, "Q")
		Call ggoOper.SetReqAttr(frm1.txtItem, "Q")
		Call ggoOper.SetReqAttr(frm1.txtFromDate, "Q")
		Call ggoOper.SetReqAttr(frm1.txtToDate, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagAll, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagN, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagY, "Q")
		Call SetToolbar("1000000000011111")		
		Exit Sub
	End If
	
    iReferArr = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	If iReferArr(0) = "N" Or len(lgF0) = 0 Then		
		Call ggoOper.SetReqAttr(frm1.txtFromConSoNo, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtToConSoNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtShipToParty, "Q")
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtPlant, "Q")
		Call ggoOper.SetReqAttr(frm1.txtItem, "Q")
		Call ggoOper.SetReqAttr(frm1.txtFromDate, "Q")
		Call ggoOper.SetReqAttr(frm1.txtToDate, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagAll, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagN, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoAllocFlagY, "Q")
		Call SetToolbar("1000000000011111")		
		Exit Sub
	End If
	
	lgBlnFlgChgValue = False
	
	frm1.txtFromConSoNo.focus

	frm1.txtFromDate.text = StartDate
	frm1.txtToDate.text = EndDate	
	'------ Developer Coding part (End )   --------------------------------------------------------------	
End Sub

'================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'================================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread
		.ReDraw = false
		
	    .MaxCols = C_GiQty + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    .MaxRows = 0
	    
	    Call GetSpreadColumnPos("A")	
	    
	    					   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		ggoSpread.SSSetEdit		C_ItemCd,			"품목",			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit		C_ItemName,			"품목명",		25,		,					,	  40
		ggoSpread.SSSetEdit		C_ItemSpec,			"규격",			20
		ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No",	15,		,					,	  25,	  2				
		ggoSpread.SSSetEdit		C_SoUnit,			"단위",			8,		,					,	  3,	  2  
							   'ColumnPosition      Header              Width	Grp		  IntegeralPart			DeciPointpart       Align		 Sep			PZ  Min Max 
		ggoSpread.SSSetFloat	C_SoQty,			"수주량",		15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000, parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat	C_PreAllocQty,		"기할당량",		15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000, parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat	C_AllocQty,			"할당량",		15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat	C_BonusQty,			"덤수량",		15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000, parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat	C_PreAllocBonusQty,	"기할당덤수량",	15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000, parent.gComNumDec,	,	,	"Z"		
	    ggoSpread.SSSetFloat	C_AllocBonusQty,	"할당덤수량",	15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
							   'ColumnPosition      Header				Width	Align(0:L,1:R,2:C)  Format         Row
		ggoSpread.SSSetDate		C_PromiseDt,		"출고예정일",	15,		2,					parent.gDateFormat
		ggoSpread.SSSetDate		C_DlvyDt,			"납기일",		15,		2,					parent.gDateFormat
							   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		ggoSpread.SSSetEdit		C_SlCd,				"창고",			8,		,					,	  7,	  2
		ggoSpread.SSSetEdit		C_SlNm,				"창고명",		8
		ggoSpread.SSSetEdit		C_PlantCd,			"공장",			8,		,					,	  4,	  2
		ggoSpread.SSSetEdit		C_PlantNm,			"공장명",		8
		ggoSpread.SSSetEdit		C_SoNo,				"수주번호",		18,		,					,	  18,	  2
		ggoSpread.SSSetEdit		C_SoSeq,			"수주순번",		10,		1
		ggoSpread.SSSetEdit		C_SchdNo,			"납품순번",		10,		1
		ggoSpread.SSSetFloat	C_PrePurReqQty,		"기구매요청량",	15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,	,   ,	"Z"
		ggoSpread.SSSetFloat	C_GiQty,			"출고수량",		15,		parent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,	,   ,	"Z"
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
		
		.ReDraw = true
   
    End With
    
End Sub


'================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)  

							'Col					Row         Row2
	ggoSpread.SSSetProtected C_ItemCd,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemName,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemSpec,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TrackingNo,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SoUnit,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SoQty,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PreAllocQty,			pvStartRow, pvEndRow		
	ggoSpread.SSSetProtected C_BonusQty,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PreAllocBonusQty,	pvStartRow, pvEndRow		
	ggoSpread.SSSetProtected C_PromiseDt,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_DlvyDt,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SlCd,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SlNm,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlantCd,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlantNm,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SoNo,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SoSeq,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SchdNo,				pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PrePurReqQty,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_GiQty,				pvStartRow, pvEndRow
		
    With frm1
		.vspdData.Col = C_ItemCd 
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Action = 0
		.vspdData.EditMode = True    
    End With
    
End Sub

'================================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.parent.gColSep)
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


'================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_ItemCd			 = iCurColumnPos(1)   '품목  
			C_ItemName			 = iCurColumnPos(2)   '품목명		
			C_ItemSpec			 = iCurColumnPos(3)   '규격 
			C_TrackingNo		 = iCurColumnPos(4)   'Tracking No
			C_SoUnit			 = iCurColumnPos(5)   '단위 
			C_SoQty			  	 = iCurColumnPos(6)   '수주량 
			C_PreAllocQty		 = iCurColumnPos(7)   '기할당량 
			C_AllocQty			 = iCurColumnPos(8)   '할당량 
			C_BonusQty			 = iCurColumnPos(9)   '덤수량 
			C_PreAllocBonusQty	 = iCurColumnPos(10)  '기할당덤수량 
			C_AllocBonusQty		 = iCurColumnPos(11)  '할당덤수량 
			C_PromiseDt			 = iCurColumnPos(12)  '출고예정일 
			C_DlvyDt			 = iCurColumnPos(13)  '납기일 
			C_SlCd				 = iCurColumnPos(14)  '창고코드 
			C_SlNm				 = iCurColumnPos(15)  '창고명       
			C_PlantCd			 = iCurColumnPos(16)  '공장코드 
			C_PlantNm			 = iCurColumnPos(17)  '공장명 
			C_SoNo				 = iCurColumnPos(18)  '수주번호 
			C_SoSeq				 = iCurColumnPos(19)  '수주순번 
			C_SchdNo			 = iCurColumnPos(20)  '납품순번 
			C_PrePurReqQty		 = iCurColumnPos(21)  '기구매요청량(Hidden)
			C_GiQty				 = iCurColumnPos(22)  '출고수량 
		
    End Select    
End Sub

'================================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	
	Dim SoSts, BillQty
	
    With frm1
		.vspdData.ReDraw = False    
 		ggoSpread.SSSetProtected C_ItemCd, -1, -1
		ggoSpread.SSSetProtected C_ItemName, -1, -1
		ggoSpread.SSSetProtected C_ItemSpec, -1, -1
		ggoSpread.SSSetProtected C_TrackingNo, -1, -1
		ggoSpread.SSSetProtected C_SoUnit, -1, -1
		ggoSpread.SSSetProtected C_SoQty, -1, -1
		ggoSpread.SSSetProtected C_PreAllocQty, -1, -1
		ggoSpread.SSSetProtected C_BonusQty, -1, -1
		ggoSpread.SSSetProtected C_PreAllocBonusQty, -1, -1
		ggoSpread.SSSetProtected C_PromiseDt, -1, -1
		ggoSpread.SSSetProtected C_DlvyDt, -1, -1
		ggoSpread.SSSetProtected C_SlCd, -1, -1
		ggoSpread.SSSetProtected C_SlNm, -1, -1
		ggoSpread.SSSetProtected C_PlantCd, -1, -1
		ggoSpread.SSSetProtected C_PlantNm, -1, -1
		ggoSpread.SSSetProtected C_SoNo, -1, -1
		ggoSpread.SSSetProtected C_SoSeq, -1, -1
		ggoSpread.SSSetProtected C_SchdNo, -1, -1
		ggoSpread.SSSetProtected C_PrePurReqQty, -1, -1	
		ggoSpread.SSSetProtected C_GiQty, -1, -1		 	
		.vspdData.ReDraw = True
    End With
    
End Sub


'================================================================================================================
Sub Form_Load()
	Err.Clear                                                                '☜: Clear err status
    Call LoadInfTB19029														 '☜: Load table , B_numeric_format
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 	    
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                    

    Call InitSpreadSheet		
    Call SetToolbar("11000000000011")										 '⊙: 버튼 툴바 제어	
    Call SetDefaultVal
	Call InitVariables														 
    '----------  Coding part  -------------------------------------------------------------	
End Sub

'================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear																	  '☜: Clear error status
    
    FncQuery = False															  '☜: Processing is NG    

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")								  '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then								 '☜: This function check required field
       Exit Function
    End If
	
	'------ Developer Coding part (Start ) --------------------------------------------------------------  
	If ValidDateCheck(frm1.txtFromDate, frm1.txtToDate) = False Then Exit Function
   
    Call InitVariables															<%%>

    If DbQuery = False Then												
        Exit Function
    End If														

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                               '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
          
End Function


'================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False																  '☜: Processing is NG
           
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then								  '☜:match pointer
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")				  '☜:There is no changed data.  
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData      
    IF ggoSpread.SSDefaultCheck = False Then								  '☜: Check contents area
		Exit Function
    End If    
    
    '------ Developer Coding part (Start ) --------------------------------------------------------------        
    If Not chkField(Document, "2")  Then                          
       Exit Function
    End If
    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbSave = False Then                                                        
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
       
End Function

'================================================================================================================
Function FncCancel() 
	Dim iDx
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
	
End Function
'================================================================================================================
Function FncPrint() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'================================================================================================================
Function FncExcel() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
End Function


'================================================================================================================
Function FncFind() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG
	
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement  
End Function


'================================================================================================================
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub


'================================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		  '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'================================================================================================================
Function DbQuery() 

	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
 
    Call LayerShowHide(1)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    arrValue(0) = ""   
    
    If frm1.rdoAllocFlagAll.checked = True Then
		frm1.txtAllocFlagRadio.value = frm1.rdoAllocFlagAll.value 
	ElseIf frm1.rdoAllocFlagN.checked = True Then
		frm1.txtAllocFlagRadio.value = frm1.rdoAllocFlagN.value 
	ElseIf frm1.rdoAllocFlagY.checked = True Then
		frm1.txtAllocFlagRadio.value = frm1.rdoAllocFlagY.value 	
	End IF	

    If lgIntFlgMode = parent.OPMD_UMODE Then        
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001										
		strVal = strVal & "&txtFromConSoNo=" & Trim(frm1.txtHFromConSoNo.value)
		strVal = strVal & "&txtToConSoNo=" & Trim(frm1.txtHToConSoNo.value)
		strVal = strVal & "&txtShipToParty=" & Trim(frm1.txtHShipToParty.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value)
		strVal = strVal & "&txtItem=" & Trim(frm1.txtHItem.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtHPlant.value)		
		strVal = strVal & "&txtFromDate=" & Trim(frm1.txtHFromDate.value)
		strVal = strVal & "&txtToDate=" & Trim(frm1.txtHToDate.value)
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtHAllocFlagRadio.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		strVal = strVal & "&txtFromConSoNo=" & Trim(frm1.txtFromConSoNo.value)			
		strVal = strVal & "&txtToConSoNo=" & Trim(frm1.txtToConSoNo.value)
		strVal = strVal & "&txtShipToParty=" & Trim(frm1.txtShipToParty.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value)
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)		
		strVal = strVal & "&txtFromDate=" & Trim(frm1.txtFromDate.text)
		strVal = strVal & "&txtToDate=" & Trim(frm1.txtToDate.text)
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtAllocFlagRadio.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If	
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement	
End Function

'================================================================================================================
Function DbSave() 
	
	Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                   '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message

	Frm1.txtMode.value        = Parent.UID_M0002                                  '☜: Delete		
	
	'------ Developer Coding part (Start)  --------------------------------------------------------------    
	lGrpCnt = 0    
	strVal = ""
		
	With frm1	
    
		ggoSpread.Source = .vspdData

		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.UpdateFlag							'☜: 수정 
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep'☜: U=Update			
			End Select

			Select Case .vspdData.Text
				Case ggoSpread.UpdateFlag
				
					.vspdData.Col = C_ItemCd				'--- 품목		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_ItemName				'--- 품목명 
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col = C_ItemSpec				'--- 규격		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_TrackingNo			'--- Tracking No
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_SoUnit				'--- 단위		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_SoQty					'--- 수주량 
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
					.vspdData.Col = C_PreAllocQty			'--- 기할당량		            
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                    
		            .vspdData.Col = C_AllocQty				'--- 할당량 
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
					.vspdData.Col = C_BonusQty				'--- 덤수량		            
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                    
		            .vspdData.Col = C_PreAllocBonusQty		'--- 기할당덤수량 
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep
		            .vspdData.Col = C_AllocBonusQty			'--- 할당덤수량		            
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gColSep                    
		            .vspdData.Col = C_PromiseDt				'--- 출고예정일 
		            strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
		            .vspdData.Col = C_DlvyDt				'--- 납기일 
		            strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
		            .vspdData.Col = C_SlCd					'--- 창고코드		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_SlNm					'--- 창고명 
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col = C_PlantCd				'--- 공장코드		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_PlantNm				'--- 공장명 
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col = C_SoNo					'--- 수주번호		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
		            .vspdData.Col = C_SoSeq					'--- 수주순번 
		            strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep
		            .vspdData.Col = C_SchdNo				'--- 납품순번		            
		            strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep                    
		            .vspdData.Col = C_PrePurReqQty			'--- 기구매요청량 
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1 
		    End Select      

		Next

		If frm1.rdoAvaInvY.checked = True Then
			frm1.txtHAvaInvRadio.value = frm1.rdoAvaInvY.value 
		Else
			frm1.txtHAvaInvRadio.value = frm1.rdoAvaInvN.value 	
		End IF	
		
		If frm1.rdoPurReqAutoY.checked = True Then
			frm1.txtHPurReqAutoRadio.value = frm1.rdoPurReqAutoY.value 
		Else
			frm1.txtHPurReqAutoRadio.value = frm1.rdoPurReqAutoN.value 	
		End IF

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strVal
	
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)											

    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement     
    
End Function
'================================================================================================================
Function DbDelete() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'================================================================================================================
Function DbQueryOk()												
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
 
	Call SetToolbar("11001001000111")

	Call SetQuerySpreadColor(1)	   
	
	lgBlnFlgChgValue = False
    
    frm1.vspdData.Focus   
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
	
    Set gActiveElement = document.ActiveElement      

End Function


'================================================================================================================
Function DbSaveOk()														
	On Error Resume Next                                                   '☜: If process fails
    Err.Clear                                                              '☜: Clear error status

    Call InitVariables													   
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    frm1.txtFromConSoNo.value = frm1.txtHFromConSoNo.value
	frm1.vspdData.MaxRows = 0
    Call MainQuery()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement
End Function


'================================================================================================================
Function DbDeleteOk()            
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement 
End Function

'================================================================================================================
Sub	SetNm()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim lgH0,lgH1,lgH2,lgH3,lgH4,lgH5,lgH6
	Dim lgJ0,lgJ1,lgJ2,lgJ3,lgJ4,lgJ5,lgJ6
	Dim lgK0,lgK1,lgK2,lgK3,lgK4,lgK5,lgK6
	
	Dim iBpNmArr
	Dim iSalesGrpArr
	Dim iPlantArr
	Dim iItemArr

    Err.Clear
    
    If frm1.txtShipToParty.value <> "" Then
		Call CommonQueryRs(" BP_NM ",		 " B_BIZ_PARTNER ", " BP_CD =  " & FilterVar(frm1.txtShipToParty.value, "''", "S") & "  " , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Len(lgF0) = 0 Then
		Else
			iBpNmArr = Split(lgF0, Chr(11))	
			frm1.txtShipToPartyNm.value = iBpNmArr(0) 
		End If
	End If
	
	If Trim(frm1.txtSalesGrp.value) <> "" Then
		Call CommonQueryRs(" SALES_GRP_NM ", " B_SALES_GRP ",	" SALES_GRP =  " & FilterVar(frm1.txtSalesGrp.value, "''", "S") & "  ", lgH0,lgH1,lgH2,lgH3,lgH4,lgH5,lgH6)
		If Len(lgH0) = 0 Then
		Else
			iSalesGrpArr = Split(lgH0, Chr(11))
			frm1.txtSalesGrpNm.value = iSalesGrpArr(0) 
		End If
	End If
	
	If Trim(frm1.txtPlant.value) <> "" Then
		Call CommonQueryRs(" PLANT_NM ",	 " B_PLANT ",		" PLANT_CD =  " & FilterVar(frm1.txtPlant.value, "''", "S") & "  ", lgJ0,lgJ1,lgJ2,lgJ3,lgJ4,lgJ5,lgJ6)
		If Len(lgJ0) = 0 Then
		Else
			iPlantArr = Split(lgJ0, Chr(11))
			frm1.txtPlantNm.value = iPlantArr(0) 
		End If
	End If
	
	If Trim(frm1.txtItem.value) <> "" Then
		Call CommonQueryRs(" ITEM_NM ",		 " B_ITEM ",		" ITEM_CD =  " & FilterVar(frm1.txtItem.value, "''", "S") & "  ", lgK0,lgK1,lgK2,lgK3,lgK4,lgK5,lgK6)
		If Len(lgK0) = 0 Then
		Else
			iItemArr = Split(lgK0, Chr(11))
			frm1.txtItemNm.value	= iItemArr(0) 	
		End If
	End If	
	
	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If	
		
End Sub

'================================================================================================================
Sub txtFromDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDate.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFromDate.Focus
	End If
End Sub

Sub txtToDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDate.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDate.Focus
	End If
End Sub
	
'================================================================================================================
Sub txtFromDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub



'================================================================================================================
Function OpenPurReqRef()		
	Dim iCalledAspName	
	Dim arrRet
			
	On Error Resume Next		

	If IsOpenPop = True Then Exit Function
			
	Call vspdData_Click(frm1.vspdData.ActiveCol , frm1.vspdData.ActiveRow)
			
	If arrValue(0) = "" Then
			Call DisplayMsgBox("203056", "x", "x", "x")	<% '⊙: "Will you destory previous data" %>
			Exit Function
	End If	
			
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3161pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3161pa1", "x")
		IsOpenPop = False
		exit Function
	end if
	IsOpenPop = True			
			
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrValue), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False			

End Function


'================================================================================================================
Function OpenFromSoDtl()
	Dim iCalledAspName
	Dim strRet
	
	If frm1.txtFromConSoNo.readOnly = True Then Exit Function

	If IsOpenPop = True Then Exit Function
			
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if
	IsOpenPop = True
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, "ALLOCATION"), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtFromConSoNo.value = strRet
		frm1.txtFromConSoNo.focus
	End If	

End Function	


'================================================================================================================
Function OpenToSoDtl()
	Dim iCalledAspName
	Dim strRet
	
	If frm1.txtToConSoNo.readOnly = True Then Exit Function
	
	If IsOpenPop = True Then Exit Function
			
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if
	IsOpenPop = True	
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, "ALLOCATION"), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtToConSoNo.value = strRet
		frm1.txtToConSoNo.focus
	End If	

End Function	

'================================================================================================================
Function OpenSoDtl(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 0	'납품처 
		If frm1.txtShipToParty.readOnly = True Then Exit Function
		
		arrParam(1) = "B_BIZ_PARTNER BP, B_BIZ_PARTNER_FTN BP_FTN"	
		arrParam(2) = Trim(frm1.txtShipToParty.value)										
		arrParam(4) = "BP.bp_cd = BP_FTN.partner_bp_cd AND BP_FTN.partner_ftn = " & FilterVar("SSH", "''", "S") & " AND BP_FTN.usage_flag = " & FilterVar("Y", "''", "S") & " "		
		arrParam(5) = "납품처"						
	
		arrField(0) = "BP_FTN.partner_bp_cd"			
		arrField(1) = "BP.bp_nm"						
    
		arrHeader(0) = "납품처"						
		arrHeader(1) = "납품처명"	

	Case 1 '영업그룹 
		If frm1.txtSalesGrp.readOnly = True Then Exit Function
		
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGrp.value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"
		
	Case 2	'공장 
		If frm1.txtPlant.readOnly = True Then Exit Function	
		
		arrParam(1) = "B_PLANT"							
		arrParam(2) = Trim(frm1.txtPlant.value)		
		arrParam(4) = ""							
		arrParam(5) = "공장"				
	
		arrField(0) = "PLANT_CD"				
		arrField(1) = "PLANT_NM"				
    
		arrHeader(0) = "공장"					
		arrHeader(1) = "공장명"		
		
	Case 3	'품목			
		If frm1.txtItem.readOnly = True Then Exit Function	
		
		arrParam(1) = "B_ITEM ITEM, B_PLANT PLANT, B_ITEM_BY_PLANT ITEM_PLANT"			
		arrParam(2) = Trim(frm1.txtItem.value)																
		
		If Trim(frm1.txtPlant.value) = "" Then
			arrParam(4) = "ITEM.item_cd = ITEM_PLANT.item_cd AND PLANT.plant_cd = ITEM_PLANT.plant_cd "			
		Else
			arrParam(4) = "ITEM.item_cd = ITEM_PLANT.item_cd AND PLANT.plant_cd = ITEM_PLANT.plant_cd AND ITEM_PLANT.plant_cd = " + FilterVar(Trim(frm1.txtPlant.value), "''", "S") + ""
		End If			
		
		arrParam(5) = "품목"					
	
		arrField(0) = "item.item_cd"				
		arrField(1) = "item.item_nm"				
		arrField(2) = "plant.plant_cd"				
		arrField(3) = "plant.plant_nm"				
    
		arrHeader(0) = "품목"						
		arrHeader(1) = "품목명"						
		arrHeader(2) = "공장"						
		arrHeader(3) = "공장명"				
		
	End Select
	
	arrParam(0) = arrParam(5)	
							
	Select Case iWhere
	Case 3
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select 	
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSoDtl(arrRet, iWhere)
	End If	
	
End Function


'================================================================================================================
Function SetSoDtl(Byval arrRet,ByVal iWhere)

	With frm1

		Select Case iWhere
		Case 0	'납품처 
			.txtShipToParty.value	= arrRet(0) 
			.txtShipToPartyNm.value = arrRet(1)
			.txtShipToParty.focus
			
		Case 1 '영업그룹			
			.txtSalesGrp.value		= arrRet(0) 
			.txtSalesGrpNm.value	= arrRet(1)
			.txtSalesGrp.focus
			
		Case 2	'공장 
			.txtPlant.value			= arrRet(0) 
			.txtPlantNm.value		= arrRet(1)			
			.txtPlant.focus
			
		Case 3	'품목 
			.txtItem.value			= arrRet(0) 
			.txtItemNm.value		= arrRet(1)	
			.txtItem.focus
			
		Case Else
			Exit Function			
		End Select
	
	End With

	lgBlnFlgChgValue = True
	
End Function

'================================================================================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD

    ggoSpread.Source = frm1.vspdData	

	<% '-- 멀티일때 -- %>
	<% '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
	If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900017", VB_YES_NO, "X", "X")                
	If IntRetCD = vbNo Then Exit Function
	End If

	<% '변경이 없을때 작업진행여부 체크 %>
	If ggoSpread.SSCheckChange = False Then
	IntRetCD = DisplayMsgBox("900018", VB_YES_NO, "X", "X")                
	If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function


'================================================================================================================
Function BizProcessCheck()

	BizProcessCheck = False

	If window.document.all("MousePT").style.visibility = "visible" Then Exit Function

	BizProcessCheck = True

End Function


'================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData
	
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
    End If
            
    If Row <> 0 Then
		With frm1.vspdData
			.Row = Row
			.Col = C_SoNo
			arrValue(0) = .text		
			
			.Col = C_SoSeq
			arrValue(1) = .text

			.Col = C_ItemCd
			arrValue(2) = .text  

			.Col = C_SchdNo
			arrValue(3) = .text  
			
		End With
	End If 
	   
End Sub

'================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True	

End Sub

'================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgStrPrevKey <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고할당</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>
						<A href="vbscript:OpenPurReqRef">구매요청현황</A></TD>					
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
					<TD WIDTH=100% HEIGHT=60 >
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>수주번호</TD>
									<TD CLASS="TD6" colspan = 3>
									<INPUT NAME="txtFromConSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromSoDtl()"> ~									
									<INPUT NAME="txtToConSoNo"	 ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSSoDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenToSoDtl()">
									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>납품처</TD>
									<TD CLASS="TD6"><INPUT NAME="txtShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSoDtl 0">&nbsp;<INPUT NAME="txtShipToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>														
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSoDtl 1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>														
								</TR>
								<TR>									
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6"><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSoDtl 2">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>														
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSoDtl 3">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>														
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출고예정일</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/s3161ma1_fpDateTime1_txtFromDate.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/s3161ma1_fpDateTime2_txtToDate.js'></script></TD>														
									<TD CLASS="TD5" NOWRAP>할당완료여부</TD>
									<TD CLASS="TD6">
										<input type=radio CLASS="RADIO" name="rdoAllocFlag" id="rdoAllocFlagAll" value="ALL" tag = "11" checked>
										<label for="rdoAllocFlagAll">전체</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAllocFlag" id="rdoAllocFlagN" value="N" tag = "11" >
										<label for="rdoAllocFlagN">미완료</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoAllocFlag" id="rdoAllocFlagY" value="Y" tag = "11" >
										<label for="rdoAllocFlagY">완료</label></TD>														
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>					
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>가용재고확인</TD>
								<TD CLASS="TD6">
									<input type=radio CLASS = "RADIO" name="rdoAvaInv" id="rdoAvaInvY" value="Y" tag = "11" >
									<label for="rdoAvaInvY">예</label>
									<input type=radio CLASS = "RADIO" name="rdoAvaInv" id="rdoAvaInvN" value="N" tag = "11" checked>
									<label for="rdoAvaInvN">아니오</label></TD>									
								<TD CLASS="TD5" NOWRAP>구매요청자동생성</TD>
								<TD CLASS="TD6">
									<input type=radio CLASS = "RADIO" name="rdoPurReqAuto" id="rdoPurReqAutoY" value="Y" tag = "11" >
									<label for="rdoPurReqAutoY">예</label>
									<input type=radio CLASS = "RADIO" name="rdoPurReqAuto" id="rdoPurReqAutoN" value="N" tag = "11" checked>
									<label for="rdoPurReqAutoN">아니오</label></TD>	
							</TR>											
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s3161ma1_I481701946_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAllocFlagRadio" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHFromConSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHToConSoNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHShipToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHItem" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHFromDate" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHToDate" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHAllocFlagRadio" tag="24" TABINDEX="-1">
	
<INPUT TYPE=HIDDEN NAME="txtHAvaInvRadio" tag="24" TABINDEX="-1">	
<INPUT TYPE=HIDDEN NAME="txtHPurReqAutoRadio" tag="24" TABINDEX="-1">	
		

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
