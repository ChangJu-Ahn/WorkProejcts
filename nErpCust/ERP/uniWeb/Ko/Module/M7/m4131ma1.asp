<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase 
'*  2. Function Name        : Inspection Result
'*  3. Program ID           : m4131ma1		
'*  4. Program Name         : 검사결과 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/09
'*  8. Modified date(Last)  : 2003/06/03
'*  9. Modifier (First)     : EverForever
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  :
'* 13. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				

'==============================================================================================================================
Const BIZ_PGM_ID = "m4131mb1.asp"						
Const BIZ_PGM_JUMP_ID	= "M4111MA1"
'==============================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'==============================================================================================================================
Dim IsOpenPop          
Dim lblnWinEvent
Dim interface_Account

Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec	
Dim C_RcptQty
Dim C_NormalQty
Dim C_AbnormalQty
Dim C_Unit	
Dim C_ProcSts	
Dim C_ProcStsNm	
Dim C_GRMeth	
Dim C_GRMethNm	
Dim C_NmlSlCd	
Dim C_NmlSlCdPop	
Dim C_NmlSlNm		
Dim C_AbnSlCd		
Dim C_AbnSlCdPop	
Dim C_AbnSlNm		
Dim C_GRNo		
Dim C_GRSeqNo	
Dim C_InspReqNo	
Dim C_MvmtNo	
Dim C_RetOrdQty

Dim C_LotNo
Dim C_LotSeqNo

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

'==============================================================================================================================
Function ChangeTag(Byval Changeflg)
    Dim index

	If Changeflg = true then
		ggoOper.SetReqAttr	frm1.txtRsRegNo1, "Q"
		
		frm1.vspdData.ReDraw = false
		For index = C_PlantCd to frm1.vspdData.MaxCols
			ggoSpread.SpreadLock index , -1, index , -1
		Next
		frm1.vspdData.ReDraw = true
	Else
		Call ggoOper.LockField(Document, "N")			
		ggoOper.SetReqAttr	frm1.txtRsRegNo1, "D"
		Call SetSpreadLock 
	End if 
End Function 
'==============================================================================================================================
Sub InitSpreadPosVariables()
	
	C_PlantCd		= 1
	C_PlantNm		= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_RcptQty		= 6
	C_NormalQty		= 7
	C_AbnormalQty	= 8
	C_Unit			= 9
	C_ProcSts		= 10
	C_ProcStsNm		= 11
	C_GRMeth		= 12
	C_GRMethNm		= 13
	C_NmlSlCd		= 14
	C_NmlSlCdPop	= 15
	C_NmlSlNm		= 16
	C_AbnSlCd		= 17
	C_AbnSlCdPop	= 18
	C_AbnSlNm		= 19
	C_GRNo			= 20
	C_GRSeqNo		= 21
	C_InspReqNo		= 22
	C_MvmtNo		= 23
	C_RetOrdQty		= 24
	
	C_LotNo  		= 25
	C_LotSeqNo		= 26

End Sub
'==============================================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                
    lgBlnFlgChgValue = False                 
    lgIntGrpCount = 0                        
    lgStrPrevKey = ""                        
    lgLngCurRows = 0                         
    frm1.vspdData.MaxRows = 0
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtReDt.Text = EndDate
    Call SetToolBar("1110100000001111")
    frm1.txtRsRegNo.focus 
    Set gActiveElement = document.activeElement
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true 
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.Spreadinit "V20030904",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		'.MaxCols = C_RetOrdQty+1
		.MaxCols = C_LotSeqNo+1
		.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetEdit 		C_PlantCd,	"공장", 10
		ggoSpread.SSSetEdit 		C_PlantNm,	"공장명", 20
		ggoSpread.SSSetEdit 		C_ItemCd,	"품목", 15
		ggoSpread.SSSetEdit 		C_ItemNm,	"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,	    "품목규격", 20 
		SetSpreadFloatLocal 		C_RcptQty,	"입고수량",15,1, 3
		SetSpreadFloatLocal 		C_NormalQty,	"양품판정수량",15,1, 3
		SetSpreadFloatLocal 		C_AbnormalQty,	"불량품판정수량",15,1, 3
		ggoSpread.SSSetEdit 		C_Unit,		"단위", 10
		ggoSpread.SSSetEdit 		C_ProcSts,	"검사상태", 10
		ggoSpread.SSSetEdit 		C_ProcStsNm,	"검사상태명", 20
		ggoSpread.SSSetEdit 		C_GRMeth,	"납입시검사방법", 20 
		ggoSpread.SSSetEdit 		C_GRMethNm,	"납입시검사방법명", 20 
		ggoSpread.SSSetEdit 		C_NmlSlCd,	"양품창고", 10, , , ,2
		ggoSpread.SSSetButton 		C_NmlSlCdPop
		ggoSpread.SSSetEdit 		C_NmlSlNm,	"양품창고명", 20    
		ggoSpread.SSSetEdit 		C_AbnSlCd,	"불량품창고", 10,,,,2
		ggoSpread.SSSetButton 		C_AbnSlCdPop
		ggoSpread.SSSetEdit 		C_AbnSlNm,	"불량품창고명", 20    
		ggoSpread.SSSetEdit 		C_GRNo,		"재고처리번호", 20
		ggoSpread.SSSetFloat 		C_GRSeqNo,	"재고처리순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_InspReqNo,"검사요청번호", 20
		ggoSpread.SSSetEdit 		C_MvmtNo,	"", 20    
		SetSpreadFloatLocal 		C_RetOrdQty, "반품수량",15,1, 3
		
		ggoSpread.SSSetEdit 		C_LotNo,	"Lot No.", 20, , , 25, 2    
		SetSpreadFloatLocal 		C_LotSeqNo, "LOT NO 순번", 20,1,6

		Call ggoSpread.MakePairsColumn(C_NmlSlCd,C_NmlSlCdPop)
		Call ggoSpread.MakePairsColumn(C_AbnSlCd,C_AbnSlCdPop)
		
		Call ggoSpread.SSSetColHidden(C_InspReqNo,C_InspReqNo,True)	
		Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)	
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)	
		Call ggoSpread.SSSetColHidden(C_RetOrdQty,C_RetOrdQty,True)
		
		.ReDraw = true
		
		Call SetSpreadLock()
    End With
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
    
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
	With ggoSpread
		.SpreadLock C_PlantCd, -1, C_PlantCd, -1
		.SpreadLock C_PlantNm, -1, C_PlantNm, -1
		.SpreadLock C_ItemCd, -1, C_ItemCd, -1
		.SpreadLock C_ItemNm, -1, C_ItemNm, -1
		.SpreadLock C_Spec, -1, C_Spec, -1
		.SpreadLock C_RcptQty, -1, C_RcptQty, -1
		.SpreadLock C_NormalQty, -1, C_NormalQty, -1
		.SpreadLock C_AbnormalQty, -1, C_AbnormalQty, -1
		.SpreadLock C_Unit, -1, C_Unit, -1
		.SpreadLock C_ProcSts, -1, C_ProcSts, -1
		.SpreadLock C_ProcStsNm, -1, C_ProcStsNm, -1
		.SpreadLock C_GRMeth, -1, C_GRMeth, -1
		
		.SpreadLock C_NmlSlCd, -1, C_NmlSlCd, -1
		.SpreadLock C_NmlSlCdPop, -1, C_NmlSlCdPop, -1
		.SpreadLock C_NmlSlNm, -1, C_NmlSlNm, -1
		.SpreadLock C_AbnSlCd, -1, C_AbnSlCd, -1
		.SpreadLock C_AbnSlCdPop, -1, C_AbnSlCdPop, -1
		.SpreadLock C_AbnSlNm, -1, C_AbnSlNm, -1
		.SpreadLock C_GRNo, -1, C_GRNo, -1
		.SpreadLock C_GRSeqNo, -1, C_GRSeqNo, -1
		.SpreadLock C_InspReqNo, -1, C_InspReqNo, -1
		.SSSetProtected frm1.vspdData.MaxCols, -1
	End With
	frm1.vspdData.ReDraw = True	
End Sub
'==============================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
	With frm1
        
		ggoSpread.SSSetProtected 	C_PlantCd ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_PlantNm ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_ItemCd ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_ItemNm ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_Spec ,		pvStartRow, pvEndRow	
		ggoSpread.SSSetProtected 	C_RcptQty ,		pvStartRow, pvEndRow
		
		ggoSpread.SpreadUnlock		C_NormalQty , pvStartRow, C_NormalQty, pvEndRow
		ggoSpread.SSSetRequired 	C_NormalQty ,	pvStartRow, pvEndRow
		ggoSpread.SpreadUnlock		C_AbnormalQty , pvStartRow, C_AbnormalQty, pvEndRow
		ggoSpread.SSSetRequired 	C_AbnormalQty , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_Unit ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_ProcSts ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_GRMeth ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_GRMethNm ,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_ProcStsNm ,	pvStartRow, pvEndRow
		
		If UNICDbl(GetSpreadText(frm1.vspdData,C_NormalQty,pvStartRow,"X","X")) = 0 Then
			ggoSpread.SSSetProtected 	C_NmlSlCd ,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected 	C_NmlSlCdpop ,	pvStartRow, pvEndRow		
		
			Call .vspdData.SetText(C_NmlSlCd,	pvStartRow, "")
			Call .vspdData.SetText(C_NmlSlNm,	pvStartRow, "")
		Else
			ggoSpread.SpreadUnlock		C_NmlSlCd , pvStartRow, C_NmlSlCdPop, pvEndRow
			ggoSpread.SSSetRequired 	C_NmlSlCd , pvStartRow, pvEndRow
		End if
		
		ggoSpread.SSSetProtected 		C_NmlSlNm , pvStartRow, pvEndRow
		
		'-- Modify by Byun Jee Hyun
		'If UCase(Trim(GetSpreadText(frm1.vspdData,C_GRMeth,pvStartRow,"X","X"))) = "B" or UNICDbl(GetSpreadText(frm1.vspdData,C_AbnormalQty,pvStartRow,"X","X")) = 0 then
		If UNICDbl(GetSpreadText(frm1.vspdData,C_AbnormalQty,pvStartRow,"X","X")) = 0 then
			ggoSpread.SSSetProtected 	C_AbnSlCd ,		pvStartRow, pvEndRow
			ggoSpread.SSSetProtected 	C_AbnSlCdpop ,	pvStartRow, pvEndRow
		
			Call .vspdData.SetText(C_AbnSlCd,	pvStartRow, "")
			Call .vspdData.SetText(C_AbnSlNm,	pvStartRow, "")
		Else
			ggoSpread.SpreadUnlock		C_AbnSlCd ,		pvStartRow, C_AbnSlCdPop, pvEndRow
			ggoSpread.SSSetRequired 	C_AbnSlCd ,		pvStartRow, pvEndRow
		End if
		
		ggoSpread.SSSetProtected 		C_AbnSlNm ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 		C_GRNo ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 		C_GRSeqNo ,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 		C_InspReqNo ,	pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected 		C_LotNo ,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 		C_LotSeqNo ,	pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStartRow,	pvEndRow
    End With
End Sub
'==============================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_RcptQty		= iCurColumnPos(6)
			C_NormalQty		= iCurColumnPos(7)
			C_AbnormalQty	= iCurColumnPos(8)
			C_Unit			= iCurColumnPos(9)
			C_ProcSts		= iCurColumnPos(10)
			C_ProcStsNm		= iCurColumnPos(11)
			C_GRMeth		= iCurColumnPos(12)
			C_GRMethNm		= iCurColumnPos(13)
			C_NmlSlCd		= iCurColumnPos(14)
			C_NmlSlCdPop	= iCurColumnPos(15)
			C_NmlSlNm		= iCurColumnPos(16)
			C_AbnSlCd		= iCurColumnPos(17)
			C_AbnSlCdPop	= iCurColumnPos(18)
			C_AbnSlNm		= iCurColumnPos(19)
			C_GRNo			= iCurColumnPos(20)
			C_GRSeqNo		= iCurColumnPos(21)
			C_InspReqNo		= iCurColumnPos(22)
			C_MvmtNo		= iCurColumnPos(23)
			C_RetOrdQty		= iCurColumnPos(24)
			
			C_LotNo 		= iCurColumnPos(25)
			C_LotSeqNo 		= iCurColumnPos(26)
	End Select
End Sub	
'==============================================================================================================================
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)
	arrParam(1) = ""
	
    If frm1.hdnGlType.Value = "A" Then               '회계전표팝업 
	   strRet = window.showModalDialog("../../comasp/a5120ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
	   strRet = window.showModalDialog("../../comasp/a5130ra1.asp", Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
	
End Function
'==============================================================================================================================
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function
	
	iCurRow = frm1.vspdData.ActiveRow
	
	IsOpenPop = True

	arrParam(0) = "양품창고"				
	arrParam(1) = "B_STORAGE_LOCATION"	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_NmlSlCd,iCurRow,"X","X"))
	arrParam(4) = "PLANT_CD= " & FilterVar(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X"), "''", "S") & " "
	arrParam(5) = "양품창고"				
	
    arrField(0) = "SL_CD"				
    arrField(1) = "SL_NM"				
    
    arrHeader(0) = "양품창고"			
    arrHeader(1) = "양품창고명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_NmlSlCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_NmlSlNm,	iCurRow, arrRet(1))
		
		Call vspdData_Change(C_NmlSlCd, frm1.vspdData.ActiveRow)
	End If
End Function
'==============================================================================================================================
Function OpenSL2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function
	
	iCurRow = frm1.vspdData.ActiveRow
	
	IsOpenPop = True

	arrParam(0) = "불량품창고"				
	arrParam(1) = "B_STORAGE_LOCATION"	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_AbnSlCd,iCurRow,"X","X"))
	arrParam(4) = "PLANT_CD= " & FilterVar(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X"), "''", "S") & " "
	arrParam(5) = "불량품창고"				
	
    arrField(0) = "SL_CD"				
    arrField(1) = "SL_NM"				
    
    arrHeader(0) = "불량품창고"			
    arrHeader(1) = "불량품창고명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_AbnSlCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_AbnSlNm,	iCurRow, arrRet(1))
	
		Call vspdData_Change(C_AbnSlCd, frm1.vspdData.ActiveRow)
	End If
End Function
'==============================================================================================================================
Function OpenPurRcptRef()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","입고참조" )
		Exit Function
	End if 
			
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = ""	'Trim(frm1.txtSupplierCd.value)
	arrParam(1) = ""	'Trim(frm1.txtSupplierNm.value)
	arrParam(2) = ""	'Trim(frm1.txtGroupCd.value)
	'arrParam(3) = FilterVar(Trim(frm1.txtRegNo.value),"","SNM")
	arrParam(3) = Trim(frm1.txtRegNo.value)
	arrParam(4) = "N"		'Clsflg
	arrParam(5) = "Y"		'Releaseflg
	arrParam(8) = "GR"		'Rcptflg
	arrParam(9) = ""	'Trim(frm1.txtMvmtType.Value)
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""
	
	iCalledAspName = AskPRAspName("M4131RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4131RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetPurRcptRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetPurRcptRef(strRet)

	Dim Index1,Count1,Index3, intCnt2,intEndRow
	Dim boolExist
	dim temp
	Dim iCurRow
	Dim iMaxRows
	Dim ilngRow, TempRow, strMessage
	
	Const C_PlantCd_Ref		= 0
	Const C_PlantNm_Ref		= 1
	Const C_ItemCd_Ref		= 2
	Const C_ItemNm_Ref		= 3
	Const C_Spec_Ref		= 4	
	Const C_RcptQty_Ref		= 5
	Const C_NormalQty_Ref	= 6
	Const C_AbnormalQty_Ref	= 7
	Const C_GRQty_Ref		= 8
	Const C_Unit_Ref		= 9
	Const C_GRDt_Ref		= 10
	Const C_InspStsCd_ref	= 11
	Const C_InspStsNm_ref	= 12
	Const C_InspMethCd_ref	= 13
	Const C_InspMethNm_ref	= 14
	Const C_GRNo_Ref		= 15
	Const C_GRSeqNo_Ref		= 16
	Const C_RsRegNo_Ref		= 17
	Const C_InspReqNo_Ref	= 18
	Const C_RcptNo_Ref		= 19
	Const C_MvmtNo_Ref		= 20
	Const C_RcptSlCd_Ref	= 21
	Const C_RcptSlNm_Ref	= 22
	
	Const C_LotNo_Ref	    = 23
	Const C_LotSeqNo_Ref	= 24

	Count1 = Ubound(strRet,1)
	
	boolExist = False
	ilngRow = 0
	intCnt2 = 0
	
	With frm1.vspdData
	
		frm1.vspdData.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False
		TempRow = .MaxRows					'리스트 max값				
		
		'----------------------------------------------------		
		For index1 = 0 to Count1
		
			boolExist = False	
			'------------------/	
			If TempRow <> 0 Then
				' Modify Byun Jee Hyun
				'for Index3=1 to TempRow			     'count1		'같은 ReqNo가 있으면 Row를 추가하지 않는다.
				for Index3=0 to TempRow			     'count1		'같은 ReqNo가 있으면 Row를 추가하지 않는다.
					.Row = index3
					.Col=C_MvmtNo
					If Trim(.Text) = Trim(strRet(index1,C_MvmtNo_Ref)) then
						strMessage = strMessage & strRet(index1,C_MvmtNo_Ref) & ";"
						boolExist = True
						Exit for
					End If
				Next
			End If
			'----------------/
			
			If boolExist <> True then
				
				'참조시 같은 번호가 있는 것이 포함 되었을때 같지 않은 것은 추가되어야 한다.
				intCnt2 = intCnt2 + 1	
				.MaxRows = CLng(TempRow) + CLng(intCnt2)
				.Row = CLng(TempRow) + CLng(intCnt2)
				
				iCurRow = .Row
				
				Call .SetText(0       ,iCurRow, ggoSpread.InsertFlag)				
				Call .SetText(C_PlantCd,	iCurRow, strRet(index1,C_PlantCd_Ref))
				Call .SetText(C_PlantNm,	iCurRow, strRet(index1,C_PlantNm_Ref))
				Call .SetText(C_itemCd,		iCurRow, strRet(index1,C_ItemCd_Ref))		
				Call .SetText(C_itemNm,		iCurRow, strRet(index1,C_ItemNm_Ref))				
				Call .SetText(C_Spec,		iCurRow, strRet(index1,C_Spec_Ref))							
				temp = UNICDbl(strRet(index1,C_RcptQty_Ref))
				Call .SetText(C_RcptQty,	iCurRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))							
				temp = UNICDbl(strRet(index1,C_NormalQty_Ref))
				Call .SetText(C_NormalQty,	iCurRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))			
				temp = UNICDbl(strRet(index1,C_AbnormalQty_Ref))
				Call .SetText(C_AbnormalQty,	iCurRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))				
				Call .SetText(C_Unit,		iCurRow, strRet(index1,C_Unit_Ref))					
				Call .SetText(C_ProcSts,	iCurRow, strRet(index1,C_InspStsCd_ref))				
				Call .SetText(C_ProcStsNm,	iCurRow, strRet(index1,C_InspStsNm_ref))				
				Call .SetText(C_GRMeth,		iCurRow, strRet(index1,C_InspMethCd_ref))				
				Call .SetText(C_GRMethNm,	iCurRow, strRet(index1,C_InspMethNm_ref))				
				Call .SetText(C_GRNo,		iCurRow, strRet(index1,C_GRNo_Ref))				
				Call .SetText(C_GRSeqNo,	iCurRow, strRet(index1,C_GRSeqNo_Ref))				
				Call .SetText(C_InspReqNo,	iCurRow, strRet(index1,C_InspReqNo_Ref))				
				Call .SetText(C_MvmtNo,		iCurRow, strRet(index1,C_MvmtNo_Ref))		
				
				If UniCDbl(GetSpreadText(frm1.vspdData,C_NormalQty,iCurRow,"X","X")) <> 0 Then
					Call .SetText(C_NmlSlCd,	iCurRow, strRet(index1,C_RcptSlCd_Ref))	
					Call .SetText(C_NmlSlNm,	iCurRow, strRet(index1,C_RcptSlNm_Ref))	
				End If
				
				If UniCDbl(GetSpreadText(frm1.vspdData,C_AbnormalQty,iCurRow,"X","X")) <> 0 Then
					Call .SetText(C_AbnSlCd,	iCurRow, strRet(index1,C_RcptSlCd_Ref))	
					Call .SetText(C_AbnSlNm,	iCurRow, strRet(index1,C_RcptSlNm_Ref))		
				End If
			    
			    Call .SetText(C_LotNo,	iCurRow, strRet(index1,C_LotNo_Ref))
			    Call .SetText(C_LotSeqNo,	iCurRow, strRet(index1,C_LotSeqNo_Ref))		
			    
			    '@@참조되는 입고번호 가져오기[060704]
				If frm1.txtRegNo.Value = "" Then
					frm1.txtRegNo.Value = strRet(index1,C_RcptNo_Ref)
				End If
				
			End if 
		Next
		
		intEndRow = .MaxRows
		
		Call SetSpreadColor(TempRow+1,intEndRow)
		
		if strMessage<>"" then
			Call displaymsgbox("17a005","X",strmessage,"입고번호")
			.ReDraw = True
			Exit Function
		End if
		
		.ReDraw = True
		
		Call SetToolBar("11101001000111")
	End with
End Function
'==============================================================================================================================
Function OpenRsRegNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
		
	If lblnWinEvent = True Or UCase(frm1.txtRsRegNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True

	arrParam(0) = ""	'Trim(frm1.hdnSupplierCd.Value)
	arrParam(1) = ""	'Trim(frm1.hdnGroupCd.Value)
	arrParam(2) = ""	'Trim(frm1.hdnMvmtType.Value)		
		
	iCalledAspName = AskPRAspName("M4131PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4131PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	lblnWinEvent = False
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0) = "" Then
		frm1.txtRsRegNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtRsRegNo.value = strRet(0)
		frm1.txtRsRegNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function
'==============================================================================================================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    arrHeader(2) = "구매조직"		
    arrHeader(3) = "구매조직명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement		
		lgBlnFlgChgValue = True
	End If	
	
End Function
'==============================================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"			
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	arrParam(3) = ""								
	
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "공급처"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"				

	arrHeader(0) = "공급처"			
	arrHeader(1) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'Lot 순번 Maker Lot 순번 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6"				  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub
'==============================================================================================================================
Function CookiePage(Byval Kubun)

	Dim strTemp

	If Kubun = 1 Then
	
	    WriteCookie "MvmtNo" , Trim(frm1.txtRegNo.value)				
		Call PgmJump(BIZ_PGM_JUMP_ID)
	Else
		strTemp = ReadCookie("MvmtNo")
	
		If strTemp = "" then Exit Function
	
		frm1.txtRegNo.value = ReadCookie("MvmtNo")
		Call WriteCookie("MvmtNo" , "")
	End if
	
End Function
'==============================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                    
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                   
    Call SetDefaultVal
    Call InitVariables
    Call CookiePage(0)
End Sub
'==============================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	IF lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows <= 0 Then
		Call SetPopupMenuItemInf("0000111111")
	ElseIf lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows > 0 Then	'참조시 
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
	End If
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
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
'==============================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
End Sub
'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
    
		If Col = C_NmlSlCdPop then
			Call OpenSl()
		Elseif Col = C_AbnSlCdPop then
			Call OpenSl2()
		End if
    End With
    
End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'==============================================================================================================================
Sub txtReDt_DblClick(Button)
	if Button = 1 then
		frm1.txtReDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtReDt.focus
	End if
End Sub
'==============================================================================================================================
Sub txtReDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'==============================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim InspQty, rcptQty
	
    ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = 0
	
	If Frm1.vspdData.text = ggoSpread.DeleteFlag Then Exit Sub

    ggoSpread.UpdateRow Row
    
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	'frm1.vspdData.ReDraw = false
	'SetSpreadColor frm1.vspdData.ActiveRow	, frm1.vspdData.ActiveRow
	'frm1.vspdData.ReDraw = true
	
	Frm1.vspdData.Col = Col
	
	select case Col
		case C_NormalQty
			frm1.vspdData.Col = C_RcptQty
			rcptQty = UNICDbl(frm1.vspdData.Text)
			
			frm1.vspdData.Col = C_NormalQty
			If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
				InspQty = rcptQty
			Else
				InspQty = rcptQty - UNICDbl(frm1.vspdData.Text)
			End If
			
			if InspQty < 0 then
				Call DisplayMsgBox("174127","X", "X", "X")
				frm1.vspdData.Col = C_AbnormalQty
				frm1.vspdData.Text= UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				exit sub				
			end if
			
			frm1.vspdData.Col = C_AbnormalQty
			frm1.vspdData.Text= UNIFormatNumber(InspQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			
		case C_AbnormalQty
			frm1.vspdData.Col = C_RcptQty
			rcptQty = UNICDbl(frm1.vspdData.Text)
			
			frm1.vspdData.Col = C_AbnormalQty
			If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
				InspQty = rcptQty
			Else
				InspQty = rcptQty - UNICDbl(frm1.vspdData.Text)
			End If
			
			if InspQty < 0 then
				Call DisplayMsgBox("174127","X", "X", "X")
				frm1.vspdData.Col = C_NormalQty
				frm1.vspdData.Text= UNIFormatNumber(0,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				exit sub				
			end if
			
			frm1.vspdData.Col = C_NormalQty
			frm1.vspdData.Text= UNIFormatNumber(InspQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	end select
    '-- Modify by Byun Jee Hyun
    Call SetSpreadColor(row, row)
End Sub
'==============================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
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
    End if
    
End Sub
'==============================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==============================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'==============================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
   
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call ggoSpread.ReOrderingSpreadData()
End Sub
'==============================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    On Error Resume Next                                                
    Err.Clear                                               
    
	ggoSpread.Source = frm1.vspdData
	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
	On Error Resume Next 
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ChangeTag(false)
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
	Call ggoOper.LockField(Document, "N")                   
	Call SetDefaultVal
	Call InitVariables

    FncNew = True                                           
	Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncDelete() 
    
	Dim IntRetCD

    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                      
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                        
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim intIndex
    
    FncSave = False                                                        
    
    On Error Resume Next  
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
    If Not ggoSpread.SSDefaultCheck         Then  
       Exit Function
    End If
    
    If frm1.vspdData.Maxrows < 1 then
    	Exit Function
    End if
    
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
    
    If DbSave = False Then Exit Function
    
    FncSave = True  
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncCopy() 
	if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo  
    Set gActiveElement = document.ActiveElement                  
End Function
'==============================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow

	On Error Resume Next
	Err.Clear
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End IF
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow , imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
    End With
	
	If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
    
End Function
'==============================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
	Call Parent.FncPrint()
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLEMULTI)
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call Parent.FncFind(Parent.C_MULTI , False)
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncExit()
	
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")               
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement   
    
End Function
'==============================================================================================================================
Function DbDelete() 
    Err.Clear                                                               
 	    
    DbDelete = False														
    
    Dim strVal
    frm1.txtMode.value = Parent.UID_M0003
    
    If LayerShowHide(1) = False Then
         Exit Function
    End If    
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
   
    DbDelete = True                                                         
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function DbDeleteOk()														
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
'==============================================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey   
       
    Err.Clear      
	     
    If LayerShowHide(1) = False Then
         Exit Function
    End If
  
    DbQuery = False                                                            
	Dim strVal
 
    With frm1    
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtRsRegNo=" & .hdnRsRegNo.value
    else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtRsRegNo=" & Trim(.txtRsRegNo.value)
	End if
    
	Call RunMyBizASP(MyBizASP, strVal)										
    
    End With
    
    DbQuery = True
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function DbQueryOk()														
	
    Call ggoOper.LockField(Document, "Q")									
	
	Call ChangeTag(True)
	lgBlnFlgChgValue = False	
	
	Call SetToolBar("11101011000111")
	lgIntFlgMode = Parent.OPMD_UMODE	
	
	Call RemovedivTextArea
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if											
	frm1.vspdData.focus
End Function
'==============================================================================================================================
Function DbSave() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Dim lRow        
    Dim strVal, strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
	
    DbSave = False                                                          
    
	Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar

    If LayerShowHide(1) = False Then
		Exit Function
	End If 
	
	iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	
	strVal = ""
	strDel = ""
	
	With frm1
		
		For lRow = 1 To .vspdData.MaxRows
		
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag
					
					strVal = "C"																					& iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemNm,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Spec,lRow, "X","X"))						& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_RcptQty,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_NormalQty,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_AbnormalQty,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProcSts,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProcStsNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRMeth,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRMethNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_NmlSlCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_NmlSlCdPop,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_NmlSlNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_AbnSlCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_AbnSlCdPop,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_AbnSlNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRNo,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GRSeqNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_InspReqNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))						& iColSep 
					strVal = strVal & lRow & iRowSep
		        Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PlantNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ItemNm,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_Spec,lRow, "X","X"))						& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_RcptQty,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_NormalQty,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_AbnormalQty,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ProcSts,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ProcStsNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GRMeth,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GRMethNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_NmlSlCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_NmlSlCdPop,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_NmlSlNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_AbnSlCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_AbnSlCdPop,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_AbnSlNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GRNo,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GRSeqNo,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_InspReqNo,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))						& iColSep 
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_RetOrdQty,lRow, "X","X"))) >  "0" then 
						Call DisplayMsgBox("172126","X",lRow & "행","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Exit Function
					End if
					strDel = strDel & lRow & iRowSep
		   	End Select 
		
			.vspdData.Row = lRow
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
					                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
					 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
					       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
					      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
					          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
					       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
					         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select
		Next
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If

	'------ Developer Coding part (End ) -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then	 
	   DbSave = True                                                             '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement                            

End Function
'==============================================================================================================================
Function DbSaveOk()										
	Call InitVariables
	Call MainQuery()
End Function

'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================

</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	

</HEAD>

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사결과</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPurRcptRef()">입고참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>결과등록번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="결과등록번호" NAME="txtRsRegNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRsRegNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRsRegNo()"></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							
							<TR>	
								<TD CLASS="TD5" NOWRAP>입고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtRegNo" MAXLENGTH=18 SIZE=34 tag="24XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>결과등록일</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtReDt" ALT="결과등록일" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="23N1" Title="FPDATETIME" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>결과등록번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="결과등록번호" NAME="txtRsRegNo1" MAXLENGTH=18 SIZE=34 tag="11XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>
								<TD CLASS="TD6" NOWRAP>								
							</TR>														
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td>						
		         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
					</td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">구매입고등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRsRegNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
