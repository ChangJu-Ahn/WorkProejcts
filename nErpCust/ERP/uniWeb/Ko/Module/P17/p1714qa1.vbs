Const BIZ_PGM_QRY_ID = "p1714qb1.asp"                          '☆: 비지니스 로직 ASP명 

Dim C_Seq               '순서 
Dim C_ReqTransNo        '이관의뢰번호 
Dim C_Status            '이관상태 
Dim C_ChildItemCd       '자품목 
Dim C_ChildItemNm       '자품목명 
Dim C_Spec              '규격 
Dim C_ReqTransDt        '이관요청일 
Dim C_TransDt           '이관일 
Dim C_ItemAcctNm        '품목계정 
Dim C_ProcTypeNm        '조달구분 
Dim C_ChildItemBaseQty  '자품목기준수 
Dim C_ChildBasicUnit    '자품목기준단위 
Dim C_PrntItemBaseQty   '모품목기준수 
Dim C_PrntBasicUnit     '모품목기준단위 
Dim C_SafetyLT          '안전L/T
Dim C_LossRate          'LOSS RATE
Dim C_SupplyFlgNm       '유무상구분 
Dim C_ValidFromDt       '시작일 
Dim C_ValidToDt         '종료일 
Dim C_ReasonNm          '설계변경근거 
Dim C_ECNNo             '설계변경번호 
Dim C_ECNDesc           '설계변경내용 
Dim C_DrawingPath       '도면경로 
Dim C_Row				

Dim isClicked
Dim iCol
Dim iRow
Dim IsOpenPop
Dim lgStrBOMHisFlg
Dim iStrFree


'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_Seq               = 1			'순서 
	C_ReqTransNo        = 2			'이관의뢰번호 
	C_Status            = 3			'이관상태 
	C_ChildItemCd       = 4			'자품목 
	C_ChildItemNm       = 5			'자품목명 
	C_Spec              = 6			'규격 
	C_ReqTransDt        = 7			'이관요청일 
	C_TransDt           = 8			'이관일 
	C_ItemAcctNm        = 9			'품목계정 
	C_ProcTypeNm        = 10		'조달구분 
	C_ChildItemBaseQty  = 11		'자품목기준수 
	C_ChildBasicUnit    = 12		'자품목기준단위 
	C_PrntItemBaseQty   = 13		'모품목기준수 
	C_PrntBasicUnit     = 14		'모품목기준단위 
	C_SafetyLT          = 15		'안전L/T
	C_LossRate          = 16		'LOSS RATE
	C_SupplyFlgNm       = 17		'유무상구분 
	C_ValidFromDt       = 18		'시작일 
	C_ValidToDt         = 19		'종료일 
	C_ReasonNm          = 20		'설계변경근거 
	C_ECNNo             = 21		'설계변경번호 
	C_ECNDesc           = 22		'설계변경내용 
	C_DrawingPath       = 23		'도면경로 
	C_Row				= 24								

                
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'       Name : InitVariables()
'       Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '[ Coding part ]======================================================================================
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                               '⊙: initializes sort direction
    '[ Coding part End ]==================================================================================
    
End Sub

Sub SetDefaultVal()
    '[ Coding part ]======================================================================================
    
    '[ Coding part End ]==================================================================================
End Sub

'=============================================== 2.2.3 InitSpreadSheet() =================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables

    '============================================================================================
    '☜: Spreadsheet vspdData
    '============================================================================================
    ggoSpread.Source = frm1.vspdData
        
    With frm1.vspdData
        
		ggoSpread.Spreadinit "V20050204", , Parent.gAllowDragDropSpread
		.ReDraw = False
		.MaxCols = C_Row
		.MaxRows = 0
        
		Call GetSpreadColumnPos()
		
        ggoSpread.SSSetFloat C_Seq, 			 "순서", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
        ggoSpread.SSSetEdit	 C_ReqTransNo, 		 "이관의뢰번호", 12
        ggoSpread.SSSetEdit	 C_Status, 	 		 "이관상태", 10
        ggoSpread.SSSetEdit	 C_ChildItemCd, 	 "자품목", 20, , , 18, 2
		ggoSpread.SSSetEdit	 C_ChildItemNm, 	 "자품목명", 30
		ggoSpread.SSSetEdit	 C_Spec, 			 "규격", 30
        ggoSpread.SSSetDate	 C_ReqTransDt, 		 "이관요청일", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetDate	 C_TransDt, 		 "이관일", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetEdit	 C_ItemAcctNm, 		 "품목계정", 10
        ggoSpread.SSSetEdit	 C_ProcTypeNm, 		 "조달구분", 12
        ggoSpread.SSSetFloat C_ChildItemBaseQty, "자품목기준수", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        ggoSpread.SSSetEdit  C_ChildBasicUnit,   "단위", 6, , , 3, 2
        ggoSpread.SSSetFloat C_PrntItemBaseQty,  "모품목기준수", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        ggoSpread.SSSetEdit  C_PrntBasicUnit, 	 "단위", 6, , , 3, 2
        ggoSpread.SSSetFloat C_SafetyLT, 		 "안전L/T", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
		ggoSpread.SSSetFloat C_LossRate, 		 "Loss율", 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
		ggoSpread.SSSetEdit C_SupplyFlgNm, 	 	 "유무상구분", 10                        
        ggoSpread.SSSetDate  C_ValidFromDt, 	 "시작일", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetDate  C_ValidToDt, 		 "종료일", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetEdit  C_ReasonNm, 		 "설계변경근거명", 14                
        ggoSpread.SSSetEdit  C_ECNNo, 			 "설계변경번호", 18, , , 18, 2
        ggoSpread.SSSetEdit  C_ECNDesc, 		 "설계변경내용", 30, , , 100
        ggoSpread.SSSetEdit  C_DrawingPath, 	 "도면경로", 30, , , 100
        ggoSpread.SSSetEdit  C_Row, 			 "순서", 5

        'ggoSpread.SSSetSplit2 (C_ChildItemPopUp)     'frozen 기능 추가(?)
        
		'====================================================================================
		'관련 Column은 묶어준다.
		'====================================================================================
        'Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnitPopup)
        '====================================================================================
		
		'====================================================================================
		'Hidden Column 지정 
		'====================================================================================
		Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
		'====================================================================================
        .ReDraw = True
                       
	End With
    
	Call SetSpreadLock
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
        ggoSpread.Source = frm1.vspdData
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected -1, -1
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'UnLock		
'		ggoSpread.SpreadUnLock C_Remark, -1, C_Remark
		
		'필수입력설정 
'		ggoSpread.SSSetRequired C_ChildItemBaseQty, -1, -1
		                
		.vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc :
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal QueryStatus)

	ggoSpread.Source = frm1.vspdData
	
	frm1.vspdData.ReDraw = False
	
	If QueryStatus = 1 Then         'When Query is OK
'		ggoSpread.SSSetProtected 	-1, pvStartRow, pvEndRow
'		ggoSpread.SpreadUnLock 		C_Remark, pvStartRow, C_Remark, pvEndRow
'		ggoSpread.SSSetRequired 	C_Seq, pvStartRow, pvEndRow
	Else
'		ggoSpread.SSSetProtected 	-1, pvStartRow, pvEndRow
'		ggoSpread.SpreadUnLock 		C_Remark, pvStartRow, C_Remark, pvEndRow
'		ggoSpread.SSSetRequired 	C_Seq, pvStartRow, pvEndRow
	End If
	
	frm1.vspdData.ReDraw = True
        
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos()
	
	Dim iCurColumnPos
   
	ggoSpread.Source = frm1.vspdData
	
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_Seq               = iCurColumnPos(1)		'순서                   
	C_ReqTransNo        = iCurColumnPos(2)		'이관의뢰번호           
	C_Status            = iCurColumnPos(3)		'이관상태               
	C_ChildItemCd       = iCurColumnPos(4)		'자품목                 
	C_ChildItemNm       = iCurColumnPos(5)		'자품목명               
	C_Spec              = iCurColumnPos(6)		'규격                   
	C_ReqTransDt        = iCurColumnPos(7)		'이관요청일             
	C_TransDt           = iCurColumnPos(8)		'이관일                 
	C_ItemAcctNm        = iCurColumnPos(9)		'품목계정               
	C_ProcTypeNm        = iCurColumnPos(10)		'조달구분               
	C_ChildItemBaseQty  = iCurColumnPos(11)		'자품목기준수           
	C_ChildBasicUnit    = iCurColumnPos(12)		'자품목기준단위         
	C_PrntItemBaseQty   = iCurColumnPos(13)		'모품목기준수           
	C_PrntBasicUnit     = iCurColumnPos(14)		'모품목기준단위         
	C_SafetyLT          = iCurColumnPos(15)		'안전L/T                
	C_LossRate          = iCurColumnPos(16)		'LOSS RATE              
	C_SupplyFlgNm       = iCurColumnPos(17)		'유무상구분             
	C_ValidFromDt       = iCurColumnPos(18)		'시작일                 
	C_ValidToDt         = iCurColumnPos(19)		'종료일                 
	C_ReasonNm          = iCurColumnPos(20)		'설계변경근거           
	C_ECNNo             = iCurColumnPos(21)		'설계변경번호           
	C_ECNDesc           = iCurColumnPos(22)		'설계변경내용           
	C_DrawingPath       = iCurColumnPos(23)		'도면경로               
	C_Row				= iCurColumnPos(24)								

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   :
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   :
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim iIntCnt
	
	ggoSpread.Source = gActiveSpdSheet
	
	Call ggoSpread.RestoreSpreadInf
	Call InitSpreadSheet
	
	frm1.vspdData.ReDraw = False
		Call InitComboBox
		Call ggoSpread.ReOrderingSpreadData
		Call SetSpreadColor(1, 1, 0, 1)
		
		With frm1
			.vspdData.Col = C_Row
			If .vspdData.Text <> "" Then
				For iIntCnt = 2 To .vspdData.MaxRows
					.vspdData.Col = C_HdrProcType
					.vspdData.Row = iIntCnt
					
					If UCase(Trim(.vspdData.Text)) = "O" Then
						Call SetFieldProp(iIntCnt, "D", "O")
					Else
						Call SetFieldProp(iIntCnt, "D", "P")
					End If
				Next
			End If
		End With

	frm1.vspdData.ReDraw = True

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'       Name : InitComboBox()
'       Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
End Sub

'==========================================  2.2.6 InitData()  =======================================
'       Name : InitData()
'       Description : Combo Display
'=====================================================================================================
Sub InitData(ByVal lngStartRow)
End Sub

'------------------------------------------  OpenConBasePlant()  -----------------------------------------
'       Name : OpenConBasePlant()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConBasePlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "기준공장팝업"             						' 팝업 명칭 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"                  ' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBasePlantCd.Value)						' Code Condition
	arrParam(3) = ""            										' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"    ' Where Condition
	arrParam(5) = "기준공장"											' TextBox 명칭 
	
	arrField(0) = "A.PLANT_CD"                      					' Field명(0)
	arrField(1) = "A.PLANT_NM"                      					' Field명(1)
	
	arrHeader(0) = "공장"    										' Header명(0)
	arrHeader(1) = "공장명"  										' Header명(1)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
	        Call SetConBasePlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
	frm1.txtDestPlantCd.focus

End Function


'------------------------------------------  OpenCondDestPlant()  ----------------------------------------
'       Name : OpenCondDestPlant()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConDestPlant()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "대상공장팝업"             	' 팝업 명칭 
	arrParam(1) = "B_PLANT"                      	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtDestPlantCd.Value)	' Code Condition
	arrParam(3) = ""            					' Name Cindition
	arrParam(4) = ""            					' Where Condition
	arrParam(5) = "대상공장"						' TextBox 명칭 
	
	arrField(0) = "PLANT_CD"                      	' Field명(0)
	arrField(1) = "PLANT_NM"                      	' Field명(1)
	
	arrHeader(0) = "공장"    					' Header명(0)
	arrHeader(1) = "공장명"  					' Header명(1)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
	        Call SetConDestPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
    frm1.txtItemCd.focus
        
End Function


'------------------------------------------  OpenReqTransNo()  -------------------------------------------------
'       Name : OpenReqTransNo()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenReqTransNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strPlantCd
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	strPlantCd = Trim(frm1.txtDestPlantCd.Value)
	
	' 팝업 명칭 
	arrParam(0) = "이관의뢰번호"
	' TABLE 명칭 
	arrParam(1) = "P_EBOM_TO_PBOM_MASTER A, B_ITEM B, B_PLANT C"
	' Code Condition
	arrParam(2) = Trim(frm1.txtReqTransNo.Value)
	' Name Cindition
	arrParam(3) = ""
	' Where Condition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
	' TextBox 명칭 
	arrParam(5) = "이관의뢰번호"
	
	arrField(0) = "A.REQ_TRANS_NO"             	' Field명(0)
	arrField(1) = "A.PLANT_CD"                 	' Field명(1)
	arrField(2) = "C.PLANT_NM"                 	' Field명(2)
	arrField(3) = "A.ITEM_CD"                  	' Field명(3)
	arrField(4) = "B.ITEM_NM"                  	' Field명(4)
	arrField(5) = "dbo.ufn_GetCodeName('Y4001', A.STATUS) STATUS"                   	' Field명(5)
	
	arrHeader(0) = "이관의뢰번호"          	' Header명(0)
	arrHeader(1) = "대상공장"              	' Header명(1)
	arrHeader(2) = "대상공장명"            	' Header명(2)
	arrHeader(3) = "품목"                  	' Header명(3)
	arrHeader(4) = "품목명"                	' Header명(4)
	arrHeader(5) = "이관상태"              	' Header명(5)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReqTransNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
	frm1.txtReqTransNo.focus
        
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'       Name : OpenItemCd()
'       Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtDestPlantCd.Value = "" Then
		Call DisplayMsgBox("971012", "X", frm1.txtDestPlantCd.Alt, "X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtDestPlantCd.Value)   ' Plant Code
	arrParam(1) = Trim(str) ' Item Code
	
	arrField(0) = 1         'ITEM_CD
	arrField(1) = 2     	'ITEM_NM
	arrField(2) = 5         'ITEM_ACCT
	arrField(3) = 9     	'PROC_TYPE
	arrField(4) = 4     	'BASIC_UNIT
	arrField(5) = 51    	'SINGLE_ROUT_FLG
	arrField(6) = 52    	'Major_Work_Center
	arrField(7) = 13    	'Phantom_flg
	arrField(8) = 18    	'valid_from_dt
	arrField(9) = 19    	'valid_to_dt
	arrField(10) = 3    	'Field명(1) : "SPECIFICATION"
	
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
	IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "B1B11PA4", "X")
	IsOpenPop = False
	Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, arrParam, arrField), _
										"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
        
End Function

'------------------------------------------  SetConBasePlant()  ----------------------------------------------
'       Name : SetConBasePlant()
'       Description : Condition Base Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConBasePlant(ByVal arrRet)
	frm1.txtBasePlantCd.Value = arrRet(0)
	frm1.txtBasePlantNm.Value = arrRet(1)
	
	Call txtBasePlantCd_OnChange
End Function

'------------------------------------------  SetConDestPlant()  ----------------------------------------------
'       Name : SetConDestPlant()
'       Description : Condition Destination Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConDestPlant(ByVal arrRet)
	frm1.txtDestPlantCd.Value = arrRet(0)
	frm1.txtDestPlantNm.Value = arrRet(1)
	
	Call txtDestPlantCd_OnChange
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
'       Name : SetItemCd()
'       Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(ByVal arrRet, ByVal iPos)
	frm1.txtItemCd.Value = arrRet(0)
	frm1.txtItemNm.Value = arrRet(1)
End Function

'------------------------------------------  SetReqTransNo()  --------------------------------------------------
'       Name : SetReqTransNo()
'       Description : SetReqTransNo
'---------------------------------------------------------------------------------------------------------
Function SetReqTransNo(ByVal arrRet)
    frm1.txtReqTransNo.Value 	= arrRet(0)
    frm1.txtDestPlantCd.Value 	= arrRet(1)
    frm1.txtDestPlantNm.Value 	= arrRet(2)
    frm1.txtItemCd.Value 		= arrRet(3)
    frm1.txtItemNm.Value 		= arrRet(4)
End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :여러 Case에 따른 Field들의 속성을 변경한다.
'==========================================================================================
Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)

End Function


'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :선택한 품목의 Item Acct를 읽는다.
'==========================================================================================
Sub LookUpItemByPlant(ByVal strItemCd, ByVal iRow)

	Err.Clear                                                                                                                   '☜: Protect system from crashing
	
	Dim strSelect
	If strItemCd = "" Then Exit Sub
	
	frm1.vspdData.Col = C_ChildItemCd
	frm1.vspdData.Row = iRow
	
	strSelect = " b.ITEM_NM, a.ITEM_ACCT, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, a.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, b.SPEC, b.BASIC_UNIT, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "
	
	If CommonQueryRs2by2(strSelect, " B_ITEM_BY_PLANT a, B_ITEM b ", " a.ITEM_CD = b.ITEM_CD AND a.PLANT_CD = " & _
								FilterVar(frm1.txtDestPlantCd.Value, "''", "S") & " AND a.ITEM_CD = " & FilterVar(strItemCd, "''", "S"), lgF0) = False Then
		Call DisplayMsgBox("122700", "X", strItemCd, "X")
		Call LookUpItemByPlantNotOk
		Exit Sub
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	Call LookUpItemByPlantOk(lgF0(1), lgF0(2), lgF0(3), lgF0(4), lgF0(5), lgF0(6), lgF0(7), iRow, lgF0(8))
End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlantOk
'   Function Desc :선택한 품목의 존재여부를 Check함를 읽는다.
'==========================================================================================
Function LookUpItemByPlantOk(ByVal strItemNm, ByVal strItemAcct, ByVal strItemAcctNm, ByVal strProcType, ByVal strProcTypeNm, ByVal strSpec, ByVal strBasicUnit, ByVal iRow, ByVal strItemAcctGrp)
End Function

Function LookUpItemByPlantNotOk()
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Call SetPopupMenuItemInf("0000110111")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
	

End Sub	


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos()
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
        If Button = 2 And gMouseClickStatus = "SPC" Then
                gMouseClickStatus = "SPCR"
        End If
End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed jslee
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
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
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	'----------  Coding part  -------------------------------------------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
	
		If lgStrPrevKeyIndex <> "" Then            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar
				Exit Sub
			End If
		End If
	End If
End Sub

Sub txtBasePlantCd_OnChange()
	Dim arrVal
	
	If Trim(frm1.txtBasePlantCd.Value) <> "" Then
		If CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtBasePlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtBasePlantNm.Value = Trim(arrVal(0)) 	
		Else
			frm1.txtBasePlantNm.Value = ""
		End If		
	End If
End Sub
 
Sub txtDestPlantCd_OnChange()
	Dim arrVal
	
	If Trim(frm1.txtDestPlantCd.Value) <> "" Then
		If CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtDestPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtDestPlantNm.Value = Trim(arrVal(0)) 	
		Else
			frm1.txtDestPlantNm.Value = ""
		End If		
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD
    
    FncQuery = False                                                        '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing

    '스프레드 초기화 
    Call ggoSpread.ClearSpreadData
    
    Call InitVariables                                                                                                                  '⊙: Initializes local global variables
                                                                                                                                                        
    '-----------------------
    'Check condition area
    'TAG = '12'인 오브젝트에 값이 있는지 체크 
    '-----------------------
    If Not chkField(document, "1") Then                                                                                 '⊙: This function check indispensable field
       Exit Function
    End If

	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then                                                                                             '☜: Query db data (설계BOM)
		Exit Function
	End If
	
	FncQuery = True                                                                                                                             '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()

End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	Dim strLevel, strChildLevel
	Dim TempChildLevel
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Level
		strLevel = CLng(Replace(.Text, ".", ""))
		
		Do
			ggoSpread.EditUndo
			
			If .MaxRows = 0 Then Exit Do
			
			.Col = C_Level
			.Row = .ActiveRow
			If Trim(.Text) = "" Then
			    strChildLevel = CLng(0)
			Else
			    strChildLevel = CLng(Replace(Trim(.Text), ".", ""))
			End If
		Loop While (strLevel < strChildLevel)
	End With
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
   Call Parent.FncPrint
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call Parent.FncExport(Parent.C_SINGLEMULTI)                                                 '☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call Parent.FncFind(Parent.C_SINGLEMULTI, False)                       '☜:화면 유형, Tab 유무 
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
    ggoSpread.SSSetSplit (gActiveSpdSheet.ActiveCol)
    
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
	    IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                  '⊙: "Will you destory previous data"
	    If IntRetCD = vbNo Then
            Exit Function
	    End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display(상단부)
'========================================================================================
Function DbQuery()

    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim StrNextKey
    Dim iStrBasePlantCd, iStrDestPlantCd, iStrItemCd, iStrReqTransNo
    Dim strQueryType

    DbQuery = False

    LayerShowHide (1)
                
    Err.Clear                                                               '☜: Protect system from crashing

    Dim strVal

    iStrBasePlantCd	= UCase(Trim(frm1.txtBasePlantCd.Value))
    iStrDestPlantCd	= UCase(Trim(frm1.txtDestPlantCd.Value))
    iStrItemCd 		= UCase(Trim(frm1.txtItemCd.Value))
    iStrReqTransNo 	= UCase(Trim(frm1.txtReqTransNo.Value))
    
    With frm1
        strVal = BIZ_PGM_QRY_ID & "?txtBasePlantCd=" & iStrBasePlantCd      '☆: 조회 조건 데이타 
        strVal = strVal & "&txtDestPlantCd=" & iStrDestPlantCd              '☆: 조회 조건 데이타 
        strVal = strVal & "&txtItemCd=" & iStrItemCd                        '☆: 조회 조건 데이타 
        strVal = strVal & "&txtReqTransNo=" & iStrReqTransNo                '☆: 조회 조건 데이타 
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
  
		Call RunMyBizASP(MyBizASP, strVal)                                                                      '☜: 비지니스 ASP 를 가동 

    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)                                                                           '☆: 조회 성공후 실행로직 
        
    Dim lRow
    Dim i
    '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.vspdData.MaxRows > 0 Then
        lgIntFlgMode = Parent.OPMD_UMODE                                                                '⊙: Indicates that current mode is Update mode
    End If

    frm1.vspdData.ReDraw = False

        With frm1
        
			.vspdData.Col = C_Row
			
			If .vspdData.Text <> "" Then
			
				For i = LngMaxRow To frm1.vspdData.MaxRows
					frm1.vspdData.Col = C_HdrProcType
					frm1.vspdData.Row = i
					
					If UCase(Trim(frm1.vspdData.Text)) = "O" Then
					    Call SetFieldProp(i, "D", "O")
					Else
					    Call SetFieldProp(i, "D", "P")
					End If
				Next
			
			End If
                        
        End With
        
        
        frm1.vspdData.focus
        lgBlnFlgChgValue = False
	
	frm1.vspdData.ReDraw = True

End Function
        
Function DbQueryNotOk()
         
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()                                                                                                     '☆: 저장 성공후 실행 로직 

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()                                                                                           '☆: 삭제 성공후 실행 로직 
    Call InitVariables
    Call FncNew
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()
	
	Dim ii
	        
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild (divTextArea.children(0))
	Next

End Function

'========================================================================================
' Function Name : CheckPlant
' Function Desc : 생산Configuration에 설계공장으로 설정이 되었는지 Check
'========================================================================================
Function CheckPlant(ByVal sPlantCd)	
														
    Err.Clear																

    CheckPlant = False
    
	Dim arrVal, strWhere, strFrom

	If Trim(sPlantCd) <> "" Then
	
		strFrom = "B_PLANT A, P_PLANT_CONFIGURATION B"
		strWhere = 				" A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND"
		strWhere = strWhere & 	" A.PLANT_CD = " & FilterVar(sPlantCd, "''", "S")

		If Not CommonQueryRs("A.PLANT_NM", strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    		Exit Function
		End If
	End If

	CheckPlant = True
	
End Function
