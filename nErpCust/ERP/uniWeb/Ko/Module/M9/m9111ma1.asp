<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m9111ma1.asp
'*  4. Program Name         : 재고이동요청등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : OH Chang Won
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'*                            
'*                            
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **************************************** !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =====================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ====================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

'******************************************  1.2 Global 변수/상수 선언  ***********************************
Const BIZ_PGM_ID 					= "m9111mb1.asp"	
Const BIZ_PGM_JUMP_ID 				= "M9211MA1"	
Const BIZ_PGM_JUMP_ID_PUR_CHARGE	= "M6111MA2"
'==========================================  1.2.1 Global 상수 선언  ======================================
Dim C_Po_Seq_No
Dim C_PlantNm
Dim C_itemCd
Dim C_Popup2
Dim C_itemNm
Dim C_SpplSpec
Dim C_OrderQty
Dim C_OrderUnit
Dim C_Popup3
Dim C_DlvyDT
Dim C_SLCd	
Dim C_Popup6
Dim C_SLNm
Dim C_Over
Dim C_Under	
Dim C_TrackingNo
Dim C_TrackingNoPop
Dim C_PrNo
Dim C_Cls_Flg
Dim C_Sto_So_No
Dim C_Sto_So_Seq
Dim C_So_No 
Dim C_So_Seq_No

Const C_SHEETMAXROWS	= 100

Dim StartDate,EndDate
EndDate = UNIConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)

'==========================================  1.2.2 Global 변수 선언  =====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lblnWinEvent
Dim releaseFlg

Dim IsOpenPop          


'==========================================   Release()  ======================================
'	Name : Release()
'===================================================================================================
Sub Release()

    Err.Clear
    
    If CheckRunningBizProcess = True Then	
		Exit Sub
	End If                
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Trim(frm1.hdnMode.Value)	
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.Value)
    strVal = strVal & "&txtStoSoNo=" & Trim(frm1.txtStoSoNo.Value)
    strVal = strVal & "&txtUpdtUserId=" & Parent.gUsrID   
    
    If LayerShowHide(1) = False Then Exit Sub
	Call RunMyBizASP(MyBizASP, strVal)								
	
End Sub
'==========================================   btnCfm()  ======================================
Sub Cfm()
    Dim IntRetCD 
    
    Err.Clear                                                       
    
    if ggoSpread.SSCheckChange = True then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if
	
	if Trim(frm1.hdnReleaseflg.Value) = "N" then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		frm1.hdnMode.Value = "Release"
					                                                
	elseif Trim(frm1.hdnReleaseflg.Value) = "Y" then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		
		frm1.hdnMode.Value = "UnRelease"
		
	End if
	
	Call Release()
	
End Sub
'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)

	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then

		strTemp = ReadCookie("PoNo")
		
		If strTemp = "" then Exit Function

		frm1.txtPoNo.value =  strTemp
	    
	    if Trim(frm1.txtPoNo.value) <> "" then
			frm1.txtQuerytype.value = "Auto"
			frm1.txthdnPoNo.Value = frm1.txtPoNo.value
			Call dbquery()
	    end if
	    
		WriteCookie "PoNo" , " "
			  
	elseIf Kubun = 1 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                          
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End if
	    
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
		WriteCookie "PoNo" , frm1.txtPoNo.value
		
		Call PgmJump(BIZ_PGM_JUMP_ID)

	elseIf Kubun = 2 Then
	
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                          
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End if
	    	
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtPoNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtGroupCd.Value)

		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
				
	End IF
	
End Function

'--------------------------------------------------------------------
'		Name        : SetState()
'		Description : Spread의 Row상태를 "R","C"로 Setting
'					  R-reference 참조      C-InsertRow
'--------------------------------------------------------------------

Sub SetState(byval strState,byval IRow)	
	frm1.vspdData.Row=IRow
	frm1.vspdData.Col=C_Stateflg
	frm1.vspdData.Text=strState
End Sub

'==========================================   ChangeItemPlant()  ======================================
Sub ChangeItemPlant(byVal iRow)

    Err.Clear                                                       
	
	Dim strVal
    Dim strssText1, strssText2, strssText3
    Dim intIndex
    Dim lGrpCnt

	with frm1
	    strVal = ""
		strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpItemByPlant"
		.vspdData.row = iRow
		.vspdData.col = C_ItemCd
		strVal = strVal & "&txtItemCd=" & Trim(.vspdData.text)
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtSupplierCd.Value)
		strVal = strVal & "&SpreadActiveRow=" & iRow
	End with
	
    If LayerShowHide(1) = False Then Exit Sub
    
    Call RunMyBizASP(MyBizASP, strVal)

End Sub

Sub changeItemPlantOK()

	if Trim(frm1.hdnTrackingflg.Value) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
		ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if
	
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE          
    lgBlnFlgChgValue = False
    lgPageNo         = ""           
    lgIntGrpCount = 0                  
    lgStrPrevKey = ""                  
    lgLngCurRows = 0
    ggoSpread.ClearSpreadData
     'frm1.vspdData.MaxRows = 0
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
    call SetToolBar("1110110100001111")
    frm1.btnCfmSel.disabled = true
    frm1.btnCfm.value = "확정"
    frm1.rdoReleaseflg(1).Checked= true
    frm1.txtPoDt.Text = EndDate
    frm1.txtPoNo.focus 
	Set gActiveElement = document.activeElement
End Sub

'=========================================  LoadInfTB19029()  ============================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
    
	C_Po_Seq_No     = 1
	C_itemCd 		= 2
	C_Popup2 		= 3
	C_itemNm 		= 4
	C_SpplSpec      = 5   
	C_OrderQty		= 6
	C_OrderUnit		= 7
	C_Popup3		= 8
	C_DlvyDT		= 9
	C_SLCd			= 10
	C_Popup6		= 11
	C_SLNm			= 12
	C_Over			= 13
	C_Under			= 14
	C_TrackingNo	= 15
	C_TrackingNoPop	= 16
	C_PrNo          = 17
	C_Cls_Flg       = 18
    C_Sto_So_No     = 19
    C_Sto_So_Seq    = 20
    C_So_No 		= 21
    C_So_Seq_No     = 22

End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
	
	.ReDraw = false
	
	ggoSpread.Spreadinit "V20030701",,Parent.gAllowDragDropSpread  
    
    .MaxCols = C_So_Seq_No+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0
	
	
    Call GetSpreadColumnPos("A")
    ggoSpread.SSSetEdit 	C_Po_Seq_No, "순번", 10   
    ggoSpread.SSSetEdit 	C_ItemCd, "품목", 18,,,18,2
    ggoSpread.SSSetButton 	C_Popup2
    ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 20    
    ggoSpread.SSSetEdit		C_SpplSpec, "품목규격", 20        '품목규격 추가 
    SetSpreadFloatLocal		C_OrderQty, "수량",15,1,3
    ggoSpread.SSSetEdit 	C_OrderUnit, "단위", 6,,,3,2
    ggoSpread.sssetButton 	C_Popup3
    ggoSpread.SSSetDate 	C_DlvyDt, "납기일", 10, 2, Parent.gDateFormat
    ggoSpread.SSSetEdit 	C_SLCd, "입고창고", 10,,,7,2
    ggoSpread.SSSetButton 	C_Popup6
    ggoSpread.SSSetEdit 	C_SLNm, "입고창고명", 20
    SetSpreadFloatLocal 	C_Over, "과부족허용율(+)(%)",20,1,6
    SetSpreadFloatLocal 	C_Under,"과부족허용율(-)(%)",20,1,6
    ggoSpread.SSSetEdit 	C_TrackingNo, "Tracking No.",  15,,,25,2
    ggoSpread.SSSetButton 	C_TrackingNoPop
    ggoSpread.SSSetEdit 	C_PrNo, "요청번호", 20
    ggoSpread.SSSetEdit 	C_Cls_Flg, "C_Cls_Flg", 5  
    ggoSpread.SSSetEdit 	C_Sto_So_No, "STO수주번호", 20
    ggoSpread.SSSetEdit 	C_Sto_So_Seq, "STO수주순번", 10
    ggoSpread.SSSetEdit 	C_So_No, "수주번호", 10
    ggoSpread.SSSetEdit 	C_So_Seq_No, "C_So_Seq_No", 20
    
            
	call ggoSpread.MakePairsColumn(C_ItemCd,C_Popup2)
	call ggoSpread.MakePairsColumn(C_OrderUnit,C_Popup3)
	call ggoSpread.MakePairsColumn(C_SLCd,C_Popup6)
	call ggoSpread.MakePairsColumn(C_TrackingNo,C_TrackingNoPop)
    
    call ggoSpread.SSSetColHidden(C_So_No,C_So_No,True)	    
    call ggoSpread.SSSetColHidden(C_PrNo,C_PrNo,True)
    call ggoSpread.SSSetColHidden(C_Cls_Flg,C_Cls_Flg,True)
    call ggoSpread.SSSetColHidden(C_So_Seq_No,C_So_Seq_No,True)
    call ggoSpread.SSSetColHidden(C_Po_Seq_No,C_Po_Seq_No,True)
    
    Call SetSpreadLock
    
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    ggoSpread.SpreadLock C_So_No , -1
    'ggoSpread.sssetrequired C_ItemCd, -1
    ggoSpread.SpreadUnLock C_ItemCd, -1
    ggoSpread.sssetrequired C_ItemCd, -1
    ggoSpread.spreadlock C_SpplSpec,-1         '품목규격 추가 
    ggoSpread.spreadlock C_ItemNm , -1
    ggoSpread.SpreadUnLock C_OrderQty, -1
    ggoSpread.sssetrequired C_OrderQty, -1
    ggoSpread.SpreadUnLock C_OrderUnit , -1
    ggoSpread.sssetrequired C_OrderUnit, -1
    ggoSpread.SpreadUnLock C_Popup3 , -1
    
    ggoSpread.SpreadUnLock C_DlvyDT, -1
    ggoSpread.sssetrequired C_DlvyDT, -1

    ggoSpread.SpreadUnLock C_SLCd , -1
    ggoSpread.sssetrequired C_SLCd, -1
    ggoSpread.SpreadUnLock C_Popup6 , -1
    ggoSpread.spreadlock C_SLNm, -1
    ggoSpread.spreadUnlock C_Under, -1
	ggoSpread.spreadUnlock C_Over, -1

    ggoSpread.spreadlock C_TrackingNo , -1   
    ggoSpread.spreadlock C_Sto_So_No, -1 
    ggoSpread.spreadlock C_Sto_So_Seq, -1 
    End With
    
       
End Sub

Sub SetSpreadLockAfterQuery()

Dim index,Count,index1 

    With frm1
    
    .vspdData.ReDraw = False
    
    if .vspdData.MaxRows < 1 then
		if .txtRelease.Value <> "Y" then
			'call SetToolBar("1110111111101")
		End if
		Exit sub
	end if
	
	'index1 = Cint(.hdnmaxrow.value) + 1
	
    if .txtRelease.Value = "Y" then
		For index = C_Po_Seq_No to C_So_Seq_No
			ggoSpread.SpreadLock index , -1
		Next
	Else
		For index1 = Cint(.hdnmaxrow.value) + 1 to .vspdData.MaxRows
		    ggoSpread.SpreadLock C_So_No , index1,C_So_No,index1
			ggoSpread.spreadlock C_PlantNm , index1,C_PlantNm,index1

			ggoSpread.spreadlock C_ItemCd, index1,C_Popup2,index1
			ggoSpread.spreadlock C_ItemNm , index1,C_ItemNm,index1
			ggoSpread.spreadlock C_SpplSpec,index1,C_SpplSpec,index1         '품목규격 추가 
			
			ggoSpread.SpreadUnLock C_OrderQty,index1,C_OrderQty,index1
			ggoSpread.sssetrequired C_OrderQty, index1,index1
            
            .vspdData.Row = index1
			.vspdData.Col = C_PrNo
			if Trim(.vspdData.Text) = "" then			
			    ggoSpread.SpreadUnLock C_OrderUnit , index1,C_OrderUnit,index1
			    ggoSpread.sssetrequired C_OrderUnit, index1,index1
			    ggoSpread.SpreadUnLock C_Popup3 , index1,C_Popup3,index1
		        ggoSpread.SpreadUnLock C_DlvyDT , index1,C_DlvyDT,index1
		        ggoSpread.sssetrequired C_DlvyDT, index1,index1
		    else
		        ggoSpread.spreadlock C_OrderUnit, index1,C_Popup3,index1
		        ggoSpread.spreadlock C_DlvyDT, index1,C_DlvyDT,index1
		    end if
		    
		    ggoSpread.SpreadUnLock C_SLCd , index1,C_SLCd,index1
		    ggoSpread.sssetrequired C_SLCd, index1,index1
		    ggoSpread.SpreadUnLock C_Popup6 , index1,C_Popup6,index1
		    ggoSpread.spreadlock C_SLNm, index1,C_SLNm,index1

            .vspdData.Row = index1
			.vspdData.Col = C_TrackingNo
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			else
				ggoSpread.spreadUnlock C_TrackingNo, index1, C_TrackingNoPop, index1
				ggoSpread.sssetrequired C_TrackingNo, index1, index1
			end if

            ggoSpread.SpreadLock C_Sto_So_No , index1,C_Sto_So_No,index1
            ggoSpread.SpreadLock C_Sto_So_Seq , index1,C_Sto_So_Seq,index1
	    next
	End if
	
    .vspdData.ReDraw = True
    
    End With
End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    ggoSpread.SSSetProtected	C_So_No		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_OrderUnit	, pvStartRow, pvEndRow

    ggoSpread.SSSetRequired		C_DlvyDt    , pvStartRow, pvEndRow
	
    ggoSpread.SSSetRequired		C_SLCd	    , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SLNm	    , pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Sto_So_No	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Sto_So_Seq, pvStartRow, pvEndRow
 	ggoSpread.SSSetProtected	C_TrackingNo ,pvStartRow, pvEndRow 
 	ggoSpread.SSSetProtected	C_TrackingNoPop ,pvStartRow, pvEndRow 
 	'******************************************
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_Po_Seq_No     = iCurColumnPos(1)
			C_itemCd 		= iCurColumnPos(2)
			C_Popup2 		= iCurColumnPos(3)
			C_itemNm 		= iCurColumnPos(4)
			C_SpplSpec      = iCurColumnPos(5)
			C_OrderQty		= iCurColumnPos(6)
			C_OrderUnit		= iCurColumnPos(7)
			C_Popup3		= iCurColumnPos(8)
			C_DlvyDT		= iCurColumnPos(9)
			C_SLCd			= iCurColumnPos(10)
			C_Popup6		= iCurColumnPos(11)
			C_SLNm			= iCurColumnPos(12)
			C_Over			= iCurColumnPos(13)
			C_Under			= iCurColumnPos(14)
			C_TrackingNo	= iCurColumnPos(15)
			C_TrackingNoPop	= iCurColumnPos(16)
            C_PrNo          = iCurColumnPos(17)
            C_Cls_Flg       = iCurColumnPos(18)
            C_Sto_So_No     = iCurColumnPos(19)
            C_Sto_So_Seq    = iCurColumnPos(20)
            C_So_No 		= iCurColumnPos(21)
            C_So_Seq_No     = iCurColumnPos(22)
            
	End Select

End Sub	
'---------------------------------  OpenSupplier()  ----------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = " IN_OUT_FLAG = " & FilterVar("I", "''", "S") & "  "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus
	End If	
	
End Function

'------------------------------------------  OpenGroup()  ------------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = " B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
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
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtGroupCd.focus
	End If	
	
End Function
'------------------------------------------  OpenPoType()  ----------------------------------------------
Function OpenPoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPoTypeCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "이동요청형태"	
	arrParam(1) = "m_config_process"
	
	arrParam(2) = Trim(frm1.txtPoTypeCd.Value)
	
	arrParam(4) = "sto_flg = " & FilterVar("Y", "''", "S") & "  "
	arrParam(5) = "이동요청형태"			
	
    arrField(0) = "po_type_cd"
    arrField(1) = "po_type_nm"
    
    arrHeader(0) = "이동요청형태"		
    arrHeader(1) = "이동요청형태명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPoTypeCd.focus
		Exit Function
	Else 
		frm1.txtPoTypeCd.Value	= arrRet(0)		
		frm1.txtPoTypeCdNm.Value= arrRet(1)
		'Call changeMvmtType()
		lgBlnFlgChgValue = True
		frm1.txtPoTypeCd.focus
	End If	

End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
			
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True

	iCalledAspName = AskPRAspName("M9111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M9111PA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(Window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus
	End If	
		
End Function
'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	if  Trim(Trim(frm1.txtSupplierCd.Value)) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtSupplierCd.focus
		Exit Function
	End if

	IsOpenPop = True

	' -- 그리드에 있는 값을 참조하기에 추가하였음	
	frm1.vspdData.Col=C_itemCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	

	arrParam(0) = Trim(frm1.txtSupplierCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.vspdData.text)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec	
    
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ItemNm
		frm1.vspdData.Text = arrRet(1)
		Call ChangeItemPlant(frm1.vspdData.ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_OrderQty,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위"					
	arrParam(1) = "B_Unit_OF_MEASURE"		
	
	frm1.vspdData.Col=C_OrderUnit
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)	
	
	arrParam(4) = ""						
	arrParam(5) = "단위"					
	
    arrField(0) = "Unit"					
    arrField(1) = "Unit_Nm"					
    
    arrHeader(0) = "단위"				
    arrHeader(1) = "단위명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_OrderUnit,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col=C_OrderUnit
		frm1.vspdData.text= arrRet(0)
		Call SetActiveCell(frm1.vspdData,C_DlvyDT,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function

'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	if Trim(frm1.txtSupplierCd.Value)="" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtSupplierCd.focus
		Exit function	
	End if 
	'같은 공장의 창고만 사용가능함. 즉 TO창고와 FROM창고의  공장이 일치해야함. 200310
	arrParam(4) = "PLANT_CD= " & FilterVar(frm1.txtSupplierCd.Value, "''", "S") & ""

	IsOpenPop = True

	frm1.vspdData.Col=C_SLCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	
	arrParam(0) = "입고창고"					
	arrParam(1) = "B_STORAGE_LOCATION"		
	
	'arrParam(2) = Trim(frm1.vspdData.Text)	
	arrParam(5) = "입고창고"					
	
    arrField(0) = "SL_CD"					
    arrField(1) = "SL_NM"					
    
    arrHeader(0) = "입고창고"				
    arrHeader(1) = "입고창고명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_SLCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_SLCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_SLNm
		frm1.vspdData.Text = arrRet(1)
	
		Call vspdData_Change(C_SLCd, frm1.vspdData.ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_Over,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function


'------------------------------------------  OpenTrackingNo()  -------------------------------------------
Function OpenTrackingNo()

	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD
		
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
    arrParam(0) = ""
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""
	
	frm1.vspdData.Col=C_So_No
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	
	arrParam(4) = Trim(frm1.vspdData.Text)
	arrParam(5) = " and A.tracking_no not in (" & FilterVar("*", "''", "S") & " ) " 
	arrParam(6) = "M" 
    
    iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet = "" Then
		Call SetActiveCell(frm1.vspdData,C_TrackingNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_TrackingNo
		frm1.vspdData.Text = arrRet
	
		Call vspdData_Change(C_TrackingNo, frm1.vspdData.ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_So_No,frm1.vspdData.ActiveRow,"M","X","X")
	End If	

End Function

'------------------------------------------  OpenReqRef()  -------------------------------------------------
'	Name : OpenReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------

Function OpenReqRef()

	Dim strRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD	
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
    If Trim(frm1.txtSupplierCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "공급처", "X")
		frm1.txtSupplierCd.focus()
    	Exit Function
    End IF

	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = ""
	arrParam(1) = ""
'	arrParam(2) = Trim(frm1.txtGroupCd.value)
'	arrParam(3) = Trim(frm1.txtGroupNm.value)
	arrParam(4) = "P"
	arrParam(5) = "Y"
	arrParam(6) = Trim(frm1.txtSupplierCd.Value)

	iCalledAspName = AskPRAspName("M2111RA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If
		
End Function

Function SetReqRef(strRet)

Dim Index1,index2,Index3,Count1,Count2
Dim IntIflg
Dim strMessage
Dim intstartRow,intEndRow

Const C_ReqNo_Ref		= 0
Const C_PlantCd_Ref		= 1
Const C_PlantNm_Ref		= 2
Const C_ItemCd_Ref		= 3
Const C_ItemNm_Ref		= 4
Const C_SpplSpec_Ref    = 5                         '품목 규격 추가 
Const C_Qty_Ref			= 6
Const C_Unit_Ref		= 7
Const C_DlvyDt_Ref		= 8
Const C_Pr_Type_Ref		= 9 
Const C_Pr_Type_Nm_Ref	= 10
Const C_SoNo_Ref		= 11
Const C_SoSeqNo_Ref		= 12
Const C_TrackingNo_Ref	= 13
Const C_SLCd_Ref		= 14
Const C_SLNm_Ref		= 15 
Const C_HSCd_Ref		= 16
Const C_HSNm_Ref		= 17


	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	with frm1
	
	intStartRow = .vspdData.MaxRows + 1
	
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
	
		for Index3=0 to .vspdData.MaxRows
			.vspdData.Row = index3+1
			.vspdData.Col=C_PrNo
			if .vspdData.Text = strRet(index1,C_ReqNo_Ref) then
				strMessage = strMessage & strRet(Index1,C_ReqNo_Ref) & ";"
				intIflg=False
				Exit for
			End if
		Next
		
		if IntIflg <> False then
		    ggoSpread.Source = .vspdData
	         .vspdData.ReDraw = False
	         ggoSpread.InsertRow
		    Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow)
			.vspdData.Row=.vspdData.ActiveRow 

			'Call SetState("C",.vspdData.ActiveRow)
			
			for index2 = 0 to Count2 - 1 
		
				Select Case Index2
				Case C_ItemCd_Ref
					.vspdData.Col=C_itemCd
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_ItemCd,.vspdData.ActiveRow,C_ItemCd,.vspdData.ActiveRow
					ggoSpread.spreadlock C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow
				Case C_ItemNm_Ref
					.vspdData.Col=C_itemNm
					.vspdData.Text=strRet(index1,index2)
					
				Case C_SpplSpec_Ref                              '품목규격 추가 
				    .vspdData.Col=C_SpplSpec
					.vspdData.Text=strRet(index1,index2)			
					
				Case C_Qty_Ref
					.vspdData.Col=C_OrderQty
					.vspdData.Text=strRet(index1,index2)
				Case C_Unit_Ref
					.vspdData.Col=C_OrderUnit
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_OrderUnit,.vspdData.ActiveRow,C_Popup3,.vspdData.ActiveRow
					'ggoSpread.spreadlock C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow
				Case C_DlvyDt_Ref
					.vspdData.Col=C_DlvyDT
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_DlvyDT, .vspdData.ActiveRow, C_DlvyDT ,.vspdData.ActiveRow
				Case C_SLCd_Ref
					.vspdData.Col=C_SLCd
					.vspdData.Text=strRet(index1,index2)
				Case C_SLNm_Ref
					.vspdData.Col=C_SLNm
					.vspdData.Text=strRet(index1,index2)
				Case C_TrackingNo_Ref
					.vspdData.Col=C_TrackingNo
					.vspdData.Text=strRet(index1,index2)
				     ggoSpread.spreadlock C_TrackingNo, .vspdData.ActiveRow, C_TrackingNoPop ,.vspdData.ActiveRow
				Case C_ReqNo_Ref
					.vspdData.Col=C_PrNo
					.vspdData.Text=strRet(index1,index2)	
                Case C_SoNo_Ref
					.vspdData.Col=C_So_No
					.vspdData.Text=strRet(index1,index2)					
			    Case C_SoSeqNo_Ref
					.vspdData.Col=C_So_Seq_No
					.vspdData.Text=strRet(index1,index2)		
				End Select
				
			next
				
		Else
			IntIFlg=True
		End if 
	next
	
	intEndRow = .vspdData.ActiveRow
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	'.vspdData.Col 	= C_Stateflg
	'.vspdData.Text = "C"
	
	.vspdData.ReDraw = True
	
	End with

			
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	     
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999.9999"
    End Select
         
End Sub

'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setreference()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub Setreference()
    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim ireference

    Err.Clear

	call CommonQueryRs(" reference ", " b_configuration ", " major_cd = " & FilterVar("M9016", "''", "S") & " and minor_cd = " & FilterVar("CH", "''", "S") & " and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    ireference = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

    if Trim(lgF0) <> "" then
        frm1.hdnreference.value = UCase(Trim(ireference(0)))
    end if

End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029
    call ggoOper.LockField(Document, "N")   
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    'call AppendNumberPlace("6", "3", "0")
    Call InitSpreadSheet                    
    Call SetDefaultVal
    Call InitVariables                      
    Call CookiePage(0)
    
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'==========================================================================================
'   Event Name : txtPoDt
'==========================================================================================
Sub txtPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtPoDt.focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtPoDt
'==========================================================================================
Sub txtPoDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoReleaseflg1_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoReleaseflg2_OnClick()
	lgBlnFlgChgValue = true	
End Sub

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
	Call SetPopupMenuItemInf("1101111111") 
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    ggoSpread.Source = frm1.vspdData
    call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
   
    ggoSpread.Source = frm1.vspdData
    call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				

    If Row <= 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Function FncSplitColumn()
   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLockAfterQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim Qty, Price, DocAmt, VatAmt, VatRate, chk_vat_flg,orgNetAmt
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row 
    
    with frm1.vspdData 		
		.Row = Row
		.Col = 0
		
		if Trim(.Text) = ggoSpread.DeleteFlag  then
		    Exit Sub
		end if    
		
		if Col = C_itemCd then 
			.Col = C_ItemCd
			strssTemp1 = Trim(.Text)
			strssTemp2 = Trim(frm1.txtSupplierCd.Value)
			
			if strssTemp1 = "" or strssTemp2 = "" then
				ggoSpread.spreadlock C_TrackingNo, Row, C_TrackingNoPop, Row
				.Col 	= C_TrackingNo
				.Text   = ""
				Exit Sub
			End if
			
			Call ChangeItemPlant(Row)
		End if
		
    End With
    
	Select Case col
	Case C_OrderQty
		
		frm1.vspdData.Col = C_OrderQty
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			Qty = 0
		Else
			Qty = UNICDbl(frm1.vspdData.Text)
		End If
		
	Case C_DlvyDt                       '납기일 
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_DlvyDt
		strsstemp1 = frm1.vspdData.Text
		if strsstemp1 = "" then Exit Sub
		strsstemp2 = frm1.txtPoDt.text
		if UniConvDateToYYYYMMDD(strsstemp2,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(strsstemp1,Parent.gDateFormat,"") then
			Call DisplayMsgBox("970023", "X", "납기일", frm1.txtPoDt.Alt)
		end if
	end select
      
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
    call CheckMinNumSpread(frm1.vspdData, Col, Row) 
	
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
 
End Sub
'==========================================================================================
'   Event Name : vspdData_DblClick
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 Then
        .Col = Col
        .Row = Row
        
		Select Case Col 
			
		Case C_Popup2
			Call OpenItem()
		Case C_Popup3
			Call OpenUnit()
		Case C_Popup6
			Call OpenSL()
		Case C_TrackingNoPop
			Call OpenTrackingNo()
		End Select
        
    End If
    
    End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    If NewRow = .MaxRows Then
        'DbQuery
    End if    

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If		

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If		 
End Sub
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim intIndex
    
    FncQuery = False                        
    
    Err.Clear                               

	ggoSpread.Source = frm1.vspdData
	

    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    call ggoOper.ClearField(Document, "2")	
    
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
	
    ggoSpread.ClearSpreadData

    Call InitVariables
    										
    If Not chkField(Document, "1") Then		
       Exit Function
    End If
   
    frm1.txtQuerytype.value = "Query"
    If DbQuery = False Then Exit Function
       
    FncQuery = True	
    
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                  
    
    Err.Clear                       
    
    ggoSpread.Source = frm1.vspdData
    

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If

    call ggoOper.ClearField(Document, "A") 
    call ggoOper.LockField(Document, "N")  
    Call SetDefaultVal()
    Call InitVariables                     
    FncNew = True                          

End Function

'========================================================================================
' Function Name : FncNew1
'========================================================================================
Function FncNew1() 
    Dim IntRetCD 
    
    FncNew1 = False                  
    
    Err.Clear                       
    
    ggoSpread.Source = frm1.vspdData
    
    call ggoOper.ClearField(Document, "A") 
    call ggoOper.LockField(Document, "N")  
    Call SetDefaultVal()
    Call InitVariables                     
    FncNew1 = True                          

End Function


'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                      
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then Exit Function
    														
    Err.Clear                              
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then             
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    call ggoOper.ClearField(Document, "A")         
    
    FncDelete = True                               
    
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                
    
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                             
    
    ggoSpread.Source = frm1.vspdData
    

    If  lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then		
		Exit Function
	End If

    If DbSave = False Then Exit Function
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 
    Dim SumTotal,tmpGrossAmt
    if frm1.vspdData.Maxrows < 1	then exit function
    ggoSpread.Source = frm1.vspdData	
    
    ggoSpread.CopyRow
    
    frm1.vspdData.ReDraw = False
    
    Call SetSpreadColor(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
    
    frm1.vspdData.ReDraw = True
    
    frm1.vspdData.Col = C_Po_Seq_No
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_So_No
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_So_Seq_No
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_PrNo
    frm1.vspdData.Text = ""
    
  	frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_TrackingNo
	
	
 	if Trim(frm1.vspdData.Text) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
	    ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if
	    
    frm1.vspdData.ReDraw = True
   
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                     
End Function

'========================================================================================
' Function Name : FncInsertRow
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
    Dim IntRetCD
    Dim imRow,index
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False
        
    If Trim(frm1.txtSupplierCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "공급처", "X")
		frm1.txtSupplierCd.focus()
    	Exit Function
    End IF
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If
	'툴바설정수정-복사버튼 활성화 되도록(2003.07)
	If frm1.vspdData.MaxRows = 0 Then call SetToolBar("11101101001011")
    
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        for index = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1 
		'공장 기본값 추가 
		    .vspdData.Row=index
		    .vspdData.Text=frm1.txtSupplierCd.Value

		    .vspdData.Col = C_DlvyDT
		    .vspdData.Text = .hdnDlvydt.value
    
  
		    .vspdData.Col = C_TrackingNo
		    .vspdData.Text = "*"
		'---------------------------------------------------------
		 next     
		    ggoSpread.spreadUnlock C_ItemCd,.vspdData.ActiveRow,C_ItemCd,.vspdData.ActiveRow + imRow - 1
		    ggoSpread.sssetrequired C_ItemCd,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		    ggoSpread.spreadUnlock C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow + imRow - 1
        
        .vspdData.ReDraw = True
    End With

	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK
    
    
    
    Set gActiveElement = document.ActiveElement   
        
End Function

'========================================================================================
' Function Name : FncDeleteRow
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim index,SumTotal,idel
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    End With
    
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                   
End Function

'========================================================================================
' Function Name : FncNext
'========================================================================================
Function FncNext() 
    On Error Resume Next                                   
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)							
End Function
		
'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)                    
End Function
'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    
    DbQuery = False
    
    Err.Clear           

	Dim strVal
    
    call SetToolBar("11100000000111") '조회버튼 누르자마자 행추가 누르는 것을 방지 

    With frm1    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtPoNo=" & .txthdnPoNo.value
	    strVal = strVal & "&txtQuerytype=" & .txtQuerytype.value
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS             '☜: 한번에 가져올수 있는 데이타 건수  
	 
    else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
	    strVal = strVal & "&txtQuerytype=" & .txtQuerytype.value
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS             '☜: 한번에 가져올수 있는 데이타 건수  
    end if 
    
    .hdnmaxrow.value = .vspdData.MaxRows
    If LayerShowHide(1) = False Then Exit Function

    Call RunMyBizASP(MyBizASP, strVal)				
   
   
    End With
    
    DbQuery = True
    
End Function
'--------------------------------  ToolBarCtrl()  ---------------------------------
Function ToolBarCtrl()
    if frm1.txtRelease.Value <> "Y" then
		call SetToolBar("11101111001111")'200309 헤더삭제는 디테일삭제가 완료되며 자동으로 삭제됨.
    else
		call SetToolBar("11100000000111")
    end if
    
End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()								
	
    lgIntFlgMode = Parent.OPMD_UMODE
    lgBlnFlgChgValue = False	
	Call SetSpreadLockAfterQuery

	Call ToolBarCtrl()
	Call RemovedivTextArea

	if Trim(UCase(frm1.hdnReleaseflg.Value)) = "Y" then
        call ggoOper.LockField(Document, "Q")
        ggoOper.SetReqAttr	frm1.txtPoNo1, "Q"
        ggoOper.SetReqAttr	frm1.txtSuppPrsn, "Q"
        ggoOper.SetReqAttr	frm1.txtTel, "Q"
        ggoOper.SetReqAttr	frm1.txtRemark, "Q"
        if frm1.hdnClsflg.value = "Y" then
		    frm1.btnCfmSel.disabled = true
		else
		    frm1.btnCfmSel.disabled = False
		end if
		frm1.btnCfm.value = "확정취소"
	else
		call ggoOper.LockField(Document, "N")
		ggoOper.SetReqAttr	frm1.txtPoNo1, "Q"
		ggoOper.SetReqAttr	frm1.txtPoTypeCd, "Q"
		ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
		frm1.btnCfm.value = "확정"
		frm1.btnCfmSel.disabled = False	
	end if
	if frm1.vspdData.MaxRows > 0 then	
		frm1.vspdData.focus
	else
		frm1.txtPoNo.Focus
	End if

End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim intIndex				

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
    
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '초기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '초기 버퍼의 설정[삭제]
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
    'On Error Resume Next                           

	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep															

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""
    intIndex = 0
    
    if .vspdData.MaxRows < 1 then
        Call DisplayMsgBox("173133", "X","내역", "X")
        Call LayerShowHide(0)
		Exit Function
    end if
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag	
				if .vspdData.Text=ggoSpread.InsertFlag then
					strVal = "C" & ColSep	
				Else
					strVal = "U" & ColSep
				End if      
				
				.vspdData.Col = C_OrderQty
				If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
					Call DisplayMsgBox("970021", "X",lRow & "행:"&"발주수량", "X")
					Call LayerShowHide(0)
					Exit Function
				End if
          
          			.vspdData.Col = C_Po_Seq_No       '1
                    strVal = strVal & Trim(.vspdData.Text) & ColSep
          		          
                    '2
                    strVal = strVal & Trim("" & frm1.txtSupplierCd.Value) & ColSep
                    
                    .vspdData.Col = C_itemCd   '3
                    strVal = strVal & Trim("" & .vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_OrderQty '4
                    If Trim(.vspdData.Text)="" Then
						strVal = strVal & "0" & ColSep
					Else
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If
                   
                    .vspdData.Col = C_OrderUnit '5
                    strVal = strVal & Trim("" & .vspdData.Text) & ColSep
                    
                   
                    .vspdData.Col = C_DlvyDT     '6
                    strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & ColSep
                    
                    .vspdData.Col = C_SLCd       '7
                    strVal = strVal & Trim("" & .vspdData.Text) & ColSep
    
                    .vspdData.Col = C_Over       '8
                     If Trim(.vspdData.Text)="" Then
						strVal = strVal & "0" & ColSep
					Else
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If
					      
                    .vspdData.Col = C_Under      '9
                     If Trim(.vspdData.Text)="" Then
						strVal = strVal & "0" & ColSep
					Else
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If

                    .vspdData.Col = C_TrackingNo '10
                    strVal = strVal & Trim("" & .vspdData.Text) & ColSep

                    .vspdData.Col = C_PrNo       '11
                    strVal = strVal & Trim(.vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_So_No    '12
                    strVal = strVal & Trim(.vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_So_Seq_No    '13
                    strVal = strVal & Trim(.vspdData.Text) & ColSep
                    
                    strVal = strVal & lRow & RowSep
                    lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag
            
                    intIndex = intIndex + 1
                    strDel = "D" & ColSep
				
          			.vspdData.Col = C_Po_Seq_No       '1
                    strDel = strDel & Trim(.vspdData.Text) & ColSep
          
                    '2
                    strDel = strDel & Trim("" & frm1.txtSupplierCd.Value) & ColSep
                    
                    .vspdData.Col = C_itemCd   '3
                    strDel = strDel & Trim("" & .vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_OrderQty '4
                    If Trim(.vspdData.Text)="" Then
						strDel = strDel & "0" & ColSep
					Else
						strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If
                   
                    .vspdData.Col = C_OrderUnit '5
                    strDel = strDel & Trim("" & .vspdData.Text) & ColSep
                    
                   
                    .vspdData.Col = C_DlvyDT     '6
                    strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & ColSep
                    
                    .vspdData.Col = C_SLCd       '7
                    strDel = strDel & Trim("" & .vspdData.Text) & ColSep
        
                    .vspdData.Col = C_Over       '8
                     If Trim(.vspdData.Text)="" Then
						strDel = strDel & "0" & ColSep
					Else
						strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If
					      
                    .vspdData.Col = C_Under      '9
                     If Trim(.vspdData.Text)="" Then
						strDel = strDel & "0" & ColSep
					Else
						strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
					End If

                    .vspdData.Col = C_TrackingNo '10
                    strDel = strDel & Trim("" & .vspdData.Text) & ColSep

                    .vspdData.Col = C_PrNo       '11
                    strDel = strDel & Trim(.vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_So_No    '12
                    strDel = strDel & Trim(.vspdData.Text) & ColSep
                    
                    .vspdData.Col = C_So_Seq_No    '13
                    strDel = strDel & Trim(.vspdData.Text) & ColSep
    
                    strDel = strDel & lRow & RowSep
                    lGrpCnt = lGrpCnt + 1
                
        End Select
        
        '=====================
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

			'=====================
                
    Next
    
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
	

    if intIndex = .vspdData.MaxRows then
        .txtMode.value = Parent.UID_M0003
    end if

	.txtMaxRows.value = lGrpCnt-1
	'.txtSpread.value = strDel & strVal

	If LayerShowHide(1) = False Then Exit Function

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					

	End With

    DbSave = True                                           
    
End Function
'======================================  RemovedivTextArea()  =================================
Function RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()											

   
	Call InitVariables
	
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim intIndex				
	Dim iOrderQty
	Dim iCost
	Dim iOrderAmt
    DbDelete = False                                  
    
    'On Error Resume Next                           

	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep															

	With frm1
		.txtMode.value = Parent.UID_M0003
		.txtFlgMode.value = lgIntFlgMode
    
    lGrpCnt = 1
    
    strVal = ""
 
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0

        strDel = strDel & "D" & ColSep
				
        .vspdData.Col = C_Po_Seq_No       '1
        strDel = strDel & Trim(.vspdData.Text) & ColSep
          
        '2
        strDel = strDel & Trim("" & frm1.txtSupplierCd.Value) & ColSep
                    
        .vspdData.Col = C_itemCd   '3
        strDel = strDel & Trim("" & .vspdData.Text) & ColSep
                    
        .vspdData.Col = C_OrderQty '4
        If Trim(.vspdData.Text)="" Then
			strDel = strDel & "0" & ColSep
		Else
			strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
		End If
                   
        .vspdData.Col = C_OrderUnit '5
        strDel = strDel & Trim("" & .vspdData.Text) & ColSep
                    
                   
        .vspdData.Col = C_DlvyDT     '6
        strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & ColSep
                    
        .vspdData.Col = C_SLCd       '7
        strDel = strDel & Trim("" & .vspdData.Text) & ColSep
        
        .vspdData.Col = C_Over       '8
         If Trim(.vspdData.Text)="" Then
			strDel = strDel & "0" & ColSep
		Else
			strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
		End If
					      
        .vspdData.Col = C_Under      '9
         If Trim(.vspdData.Text)="" Then
			strDel = strDel & "0" & ColSep
		Else
			strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
		End If

        .vspdData.Col = C_TrackingNo '10
        strDel = strDel & Trim("" & .vspdData.Text) & ColSep

        .vspdData.Col = C_PrNo       '11
        strDel = strDel & Trim(.vspdData.Text) & ColSep
                    
        .vspdData.Col = C_So_No    '12
        strDel = strDel & Trim(.vspdData.Text) & ColSep
                    
        .vspdData.Col = C_So_Seq_No    '13
        strDel = strDel & Trim(.vspdData.Text) & ColSep
                        
        strDel = strDel & lRow & RowSep
        lGrpCnt = lGrpCnt + 1
                
               
    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel

	If LayerShowHide(1) = False Then Exit Function

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					

	End With

    DbDelete = True                                  
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고이동요청</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReqRef">구매요청참조</A></TD>
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
							<TABLE  <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>재고이동요청번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=29 MAXLENGTH=18 ALT="재고이동요청번호" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
							    <TD CLASS="TD5" NOWRAP>재고이동요청번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="재고이동요청번호" NAME="txtPoNo1" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>확정여부</TD>
							    <TD CLASS=TD6 NOWRAP>
								<input type=radio CLASS="RADIO" name="rdoReleaseflg" id="rdoReleaseflg1" tag = "24"><label for="rdoConfirmFlg_Yes">확정</label>&nbsp;&nbsp;
								<input type=radio CLASS = "RADIO" name="rdoReleaseflg" id="rdoReleaseflg2"   tag = "24"><label for="rdoConfirmFlg_No">미확정</label>
							</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>이동유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="이동유형" NAME="txtPoTypeCd" SIZE=10 MAXLENGTH=5 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="이동유형" NAME="txtPoTypeCdNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>등록일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m9111ma1_fpDateTime1_txtPoDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
							</TR>
							<TR>
				                <TD CLASS="TD5" NOWRAP>공급처담당</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처담당" NAME="txtSuppPrsn" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>긴급연락처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="긴급연락처" NAME="txtTel" MAXLENGTH=18 SIZE=35 tag="21XXXU"></TD>
							</TR>
							<TR>
				                <TD CLASS="TD5" NOWRAP>비고</TD>
				                <TD CLASS="TD6" NOWRAP Colspan=3><INPUT TYPE=TEXT ALT="비고" NAME="txtRemark" MAXLENGTH=70 SIZE=91 tag="21XXXU"></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m9111ma1_OBJECT1_vspdData.js'></script>
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
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">확정</button></a></td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">재고이동입고</a><!-- | <a href="VBSCRIPT:CookiePage(2)">경비등록</a>--></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m9111mb1.asp" FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden"  NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRelease" tag="14">
<INPUT TYPE=HIDDEN NAME="txthdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtQuerytype" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnDlvyDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubContraFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnXch" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMode" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingflg" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMaintNo" tag="14">
<!-- 2002.2.14 VAT Append -->
<INPUT TYPE=HIDDEN NAME="hdnVATType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATRate" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVATINCFLG" tag="1">
<!-- 2002.2.19 환율 연산자 -->
<INPUT TYPE=HIDDEN NAME="hdnXchRateOp"  tag="14">
<!-- 2002.3.14 매입여부 -->
<INPUT TYPE=HIDDEN NAME="hdnIVFlg"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnreference"  tag="14">
<INPUT TYPE=HIDDEN NAME="txtStoSoNo"  tag="14">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
