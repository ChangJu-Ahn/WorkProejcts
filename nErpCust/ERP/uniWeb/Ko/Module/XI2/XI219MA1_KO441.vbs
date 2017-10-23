
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID1 = "XI219MB1_KO441.asp"                                      'Biz Logic ASP
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================

Dim C_ITEM_CD          '품목코드(PK)
Dim C_ITEM_CD_POPUP    '품목PopUp Button
Dim C_ITEM_NM          '품목명
Dim C_MAT_LOT_NO       'LOT번호(PK)
Dim C_SELLER_CD        '납품처코드(PK)
Dim C_SELLER_CD_POPUP  '납품처PopUp Button
Dim C_SELLER_NM        '납품처명
Dim C_PRODUCT_DT       '생산일자(Hidden)
Dim C_PRINT_DT         '발행일자(PK)
Dim C_RCPT_DT          '납입일자(PK)
Dim C_RCPT_TM          '납입시간(PK)
Dim C_RCPT_QTY         '납입수량
Dim C_BP_ISSUE_NO      '납품처발행번호
Dim C_ISSUE_FLAG       '발행구분
Dim C_ISSUE_FLAG_NM    '발행구분명
Dim C_PLANT_FLAG       '공장구분
Dim C_PLANT_CD         '공장코드
Dim C_PLANT_CD_POPUP   '공장PopUp Button
Dim C_GATE_CD          'GATE
Dim C_SNP              'SNP
Dim C_BOX_QTY          'Box수량
Dim C_SEPARATE_FLAG    '분할구분
Dim C_DELIVERY_NO      '납품번호
Dim C_ISSUE_DT         '출하일자
Dim C_ISSUE_TIME       '출하시간
Dim C_DELETE_FLAG      '삭제여부
Dim C_IF_SEQ			'전송순번
Dim C_MES_RECEIVE_FLAG 'MES수신여부
Dim C_MES_RECEIVE_DT   'MES최종수신일시
Dim C_ERP_SEND_DT	   'ERP최종송신일시
Dim C_ERR_DESC         '에러내역
DIm C_CREATE_TYPE      '생성구분(PK)


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_ITEM_CD          = 1  '품목코드(PK)
	C_ITEM_CD_POPUP    = 2  '품목PopUp Button
	C_ITEM_NM          = 3  '품목명
	C_MAT_LOT_NO       = 4  'LOT번호(PK)
	C_SELLER_CD        = 5  '납품처코드(PK)
	C_SELLER_CD_POPUP  = 6  '납품처PopUp Button
	C_SELLER_NM        = 7  '납품처명
	C_PRODUCT_DT       = 8  '생산일자(Hidden)
	C_PRINT_DT         = 9  '발행일자(PK)
	C_RCPT_DT          = 10 '납입일자(PK)
	C_RCPT_TM          = 11 '납입시간(PK)
	C_RCPT_QTY         = 12 '납입수량
	C_BP_ISSUE_NO      = 13 '납품처발행번호
	C_ISSUE_FLAG       = 14 '발행구분
	C_ISSUE_FLAG_NM    = 15 '발행구분명
	C_PLANT_FLAG       = 16 '공장구분
	C_PLANT_CD         = 17 '공장코드
	C_PLANT_CD_POPUP   = 18 '공장PopUp Button
	C_GATE_CD          = 19 'GATE
	C_SNP              = 20 'SNP
	C_BOX_QTY          = 21 'Box수량
	C_SEPARATE_FLAG    = 22 '분할구분
	C_DELIVERY_NO      = 23 '납품번호
	C_ISSUE_DT         = 24 '출하일자
	C_ISSUE_TIME       = 25 '출하시간
	C_DELETE_FLAG      = 26 '삭제여부
	C_IF_SEQ		   = 27 '전송순번
	C_MES_RECEIVE_FLAG = 28 'MES수신여부
	C_MES_RECEIVE_DT   = 29 'MES수신일시
	C_ERP_SEND_DT	   = 30	'ERP최종송신일시
	C_ERR_DESC         = 31 '에러내역
	C_CREATE_TYPE      = 32 '생성구분(PK)

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
   		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtPrintFrDt.Text = StartDate
	frm1.txtPrintToDt.Text = EndDate
	frm1.rdoDelFlagNomal.checked = True
	frm1.rdoMesRcvFlagAll.checked = True
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
End Sub
	

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Select Case pOpt
		Case "MQ"
			lgKeyStream = UNIConvDate(frm1.txtPrintFrDt.Text)               & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIConvDate(frm1.txtPrintToDt.Text) & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtItemCd.Value)          & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtBpCd.Value)            & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtLotNo.Value)           & Parent.gColSep
			If frm1.rdoDelFlagAll.checked Then
				lgKeyStream = lgKeyStream & ""  & Parent.gColSep
			ElseIf frm1.rdoDelFlagNomal.checked Then
				lgKeyStream = lgKeyStream & "N"  & Parent.gColSep
			Else'rdoDelFlagDel
				lgKeyStream = lgKeyStream & "Y"  & Parent.gColSep
			End If
			If frm1.rdoMesRcvFlagAll.checked Then
				lgKeyStream = lgKeyStream & ""  & Parent.gColSep
			ElseIf frm1.rdoMesRcvFlagNomal.checked Then
				lgKeyStream = lgKeyStream & "Y"  & Parent.gColSep
			Else'rdoMesRcvFlagFail
				lgKeyStream = lgKeyStream & "N"  & Parent.gColSep
			End If

		Case "MN"
			lgKeyStream = UNIConvDate(frm1.hPrintFrDt.value)               & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIConvDate(frm1.hPrintToDt.value) & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.hItemCd.Value)           & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.hBpCd.Value)             & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.hLotNo.Value)            & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.hDelFlag.Value)          & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.hMesRcvFlag.Value)       & Parent.gColSep

	End Select                 
	       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBoxf
'========================================================================================================
Sub InitComboBox()    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadComboBox
' Function Desc :
'========================================================================================================
Sub InitSpreadComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
    ggoSpread.Source = frm1.vspdData
                       'Data        Seperator            Column position 
    ggoSpread.SetCombo "G"    & vbTab & "J"   & vbTab & "E"    & vbTab & "C" ,   C_ISSUE_FLAG
    ggoSpread.SetCombo "정상" & vbTab & "JIT" & vbTab & "긴급" & vbTab & "취소", C_ISSUE_FLAG_NM
   
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()

	Dim intRow
	Dim intIndex 

	With frm1.vspdData
	
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_ISSUE_FLAG     : intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
			.Col = C_ISSUE_FLAG_NM  : .Value = intindex					
		Next	
		
	End With

End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050519",, parent.gAllowDragDropSpread
		.ReDraw  = False
		.MaxCols = C_CREATE_TYPE + 1                                                  ' ☜:☜: Add 1 to Maxcols

		Call ggoSpread.ClearSpreadData()

'		Call AppendNumberPlace("6","4","2")
		Call AppendNumberPlace("6", "18", "0")

		Call GetSpreadColumnPos("A")
       
		ggoSpread.SSSetEdit    C_ITEM_CD,         "품목코드",        15, 0,,  18, 2
		ggoSpread.SSSetButton  C_ITEM_CD_POPUP,    -1
		ggoSpread.SSSetEdit    C_ITEM_NM,          "품명",           20, 0,,  40, 2
		ggoSpread.SSSetEdit    C_MAT_LOT_NO,       "LOT번호",        20, 0,,  25, 2
		ggoSpread.SSSetEdit    C_SELLER_CD,        "납품처",		  8, 0,,  10, 2
		ggoSpread.SSSetButton  C_SELLER_CD_POPUP,  -1
		ggoSpread.SSSetEdit    C_SELLER_NM,        "납품처명",       20, 0,,  50, 2
		ggoSpread.SSSetDate    C_PRODUCT_DT,       "생산일자",       10, 2, parent.gDateFormat
		ggoSpread.SSSetDate    C_PRINT_DT,         "발행일자",       10, 2, parent.gDateFormat
		ggoSpread.SSSetDate    C_RCPT_DT,          "납입일자",       10, 2, parent.gDateFormat
		ggoSpread.SSSetTime    C_RCPT_TM,          "납입시간",        8, 2,    1, 0
		ggoSpread.SSSetFloat   C_RCPT_QTY,         "납입수량",       12, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit    C_BP_ISSUE_NO,      "납품처발행번호", 15, 0,, 18, 2
		ggoSpread.SSSetCombo   C_ISSUE_FLAG,       "발행구분",        8, 2, False
		ggoSpread.SSSetCombo   C_ISSUE_FLAG_NM,    "발행구분명",     10, 2, False
		ggoSpread.SSSetEdit    C_PLANT_FLAG,       "공장구분",        8, 0,,   4, 2
		ggoSpread.SSSetEdit    C_PLANT_CD,         "공장",            5, 0,,   4, 2
		ggoSpread.SSSetButton  C_PLANT_CD_POPUP,   -1
		ggoSpread.SSSetEdit    C_GATE_CD,          "GATE",            5, 0,,   4, 2
		ggoSpread.SSSetFloat   C_SNP,              "SNP",             6, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat   C_BOX_QTY,          "Box수량",         6, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetCheck   C_SEPARATE_FLAG,    "분할구분",        8,,, True
		ggoSpread.SSSetEdit    C_DELIVERY_NO,      "납입지시번호",   18, 0,,  18, 2
		ggoSpread.SSSetDate    C_ISSUE_DT,         "출하일자",       10, 2, parent.gDateFormat
		ggoSpread.SSSetTime    C_ISSUE_TIME,       "출하시간",        8, 2,  1, 0
		ggoSpread.SSSetCheck   C_DELETE_FLAG,      "삭제여부",        8,,, True
		
		ggoSpread.SSSetEdit    C_IF_SEQ,           "최종전송순번",   10, 0,,   3
		ggoSpread.SSSetEdit    C_MES_RECEIVE_FLAG, "MES수신여부",    10, 2,,   1, 2
		ggoSpread.SSSetEdit    C_MES_RECEIVE_DT,   "MES최종수신일시",16, 2,,  20, 2
		ggoSpread.SSSetEdit    C_ERP_SEND_DT,	   "ERP최종송신일시",16, 2,,  20, 2
		ggoSpread.SSSetEdit    C_ERR_DESC,         "에러내역",       30, 0,, 500, 2
		ggoSpread.SSSetEdit    C_CREATE_TYPE,      "생성구분",       10, 2,,   1, 2

		Call ggoSpread.MakePairsColumn(C_ITEM_CD,    C_ITEM_CD_POPUP)
		Call ggoSpread.MakePairsColumn(C_SELLER_CD,  C_SELLER_CD_POPUP)
		Call ggoSpread.MakePairsColumn(C_RCPT_DT,    C_RCPT_TM)
'		Call ggoSpread.MakePairsColumn(C_ISSUE_FLAG, C_ISSUE_FLAG_NM)
		Call ggoSpread.MakePairsColumn(C_PLANT_CD,   C_PLANT_CD_POPUP)
		Call ggoSpread.MakePairsColumn(C_ISSUE_DT,   C_ISSUE_TIME)

		Call ggoSpread.SSSetColHidden(C_PRODUCT_DT,  C_PRODUCT_DT,  True)
		Call ggoSpread.SSSetColHidden(C_ISSUE_FLAG,  C_ISSUE_FLAG,  True)
		Call ggoSpread.SSSetColHidden(C_CREATE_TYPE, C_CREATE_TYPE, True)
		Call ggoSpread.SSSetColHidden(.MaxCols,      .MaxCols,      True)

		.ReDraw = True
		Call SetSpreadLock()
		
	End With    
	
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()

    With frm1

		.vspdData.ReDraw = False 
		                           'Col-1              Row-1       Col-2           Row-2   
		ggoSpread.SpreadLock       C_ITEM_CD,          -1,         C_ITEM_CD
		ggoSpread.SpreadLock       C_ITEM_CD_POPUP,    -1,         C_ITEM_CD_POPUP
		ggoSpread.SpreadLock       C_ITEM_NM,          -1,         C_ITEM_NM
		ggoSpread.SpreadLock       C_MAT_LOT_NO,       -1,         C_MAT_LOT_NO
		ggoSpread.SpreadLock       C_SELLER_CD,        -1,         C_SELLER_CD
		ggoSpread.SpreadLock       C_SELLER_CD_POPUP,  -1,         C_SELLER_CD_POPUP
		ggoSpread.SpreadLock       C_SELLER_NM,        -1,         C_SELLER_NM
		ggoSpread.SpreadLock       C_PRINT_DT,         -1,         C_PRINT_DT
		ggoSpread.SpreadLock       C_RCPT_DT,          -1,         C_RCPT_DT
		ggoSpread.SpreadLock       C_RCPT_TM,          -1,         C_RCPT_TM
		ggoSpread.SSSetRequired    C_RCPT_QTY,         -1,         -1
		ggoSpread.SSSetRequired    C_SNP,              -1,         -1
		ggoSpread.SSSetRequired    C_BOX_QTY,          -1,         -1
		
		ggoSpread.SpreadLock       C_IF_SEQ,		   -1,         C_IF_SEQ
		ggoSpread.SpreadLock       C_MES_RECEIVE_FLAG, -1,         C_MES_RECEIVE_FLAG
		ggoSpread.SpreadLock       C_MES_RECEIVE_DT,   -1,         C_MES_RECEIVE_DT
		ggoSpread.SpreadLock       C_ERP_SEND_DT,	   -1,         C_ERP_SEND_DT
		ggoSpread.SpreadLock       C_ERR_DESC,         -1,         C_ERR_DESC

		ggoSpread.SSSetProtected   .vspdData.MaxCols,  -1,         -1

		.vspdData.ReDraw = True

    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadSetLock
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadSetLock(ByVal pvRow)

    With frm1    
		.vspdData.Col = C_MES_RECEIVE_FLAG
		.vspdData.Row = pvRow
		If Ucase(Trim(.vspdData.Text)) = "Y" Then
			ggoSpread.SpreadLock C_RCPT_QTY,       pvRow, C_RCPT_QTY,       pvRow
			ggoSpread.SpreadLock C_BP_ISSUE_NO,    pvRow, C_BP_ISSUE_NO,    pvRow
			ggoSpread.SpreadLock C_ISSUE_FLAG,     pvRow, C_ISSUE_FLAG,     pvRow
			ggoSpread.SpreadLock C_ISSUE_FLAG_NM,  pvRow, C_ISSUE_FLAG_NM,  pvRow
			ggoSpread.SpreadLock C_PLANT_FLAG,     pvRow, C_PLANT_FLAG,     pvRow
			ggoSpread.SpreadLock C_PLANT_CD,       pvRow, C_PLANT_CD,       pvRow
			ggoSpread.SpreadLock C_PLANT_CD_POPUP, pvRow, C_PLANT_CD_POPUP, pvRow
			ggoSpread.SpreadLock C_GATE_CD,        pvRow, C_GATE_CD,        pvRow
			ggoSpread.SpreadLock C_SNP,            pvRow, C_SNP,            pvRow
			ggoSpread.SpreadLock C_BOX_QTY,        pvRow, C_BOX_QTY,        pvRow
			ggoSpread.SpreadLock C_SEPARATE_FLAG,  pvRow, C_SEPARATE_FLAG,  pvRow
			ggoSpread.SpreadLock C_DELIVERY_NO,    pvRow, C_DELIVERY_NO,    pvRow
			ggoSpread.SpreadLock C_ISSUE_DT,       pvRow, C_ISSUE_DT,       pvRow
			ggoSpread.SpreadLock C_ISSUE_TIME,     pvRow, C_ISSUE_TIME,     pvRow
			ggoSpread.SpreadLock C_CREATE_TYPE,    pvRow, C_CREATE_TYPE,    pvRow
			
			ggoSpread.SpreadLock C_IF_SEQ,		   pvRow, C_IF_SEQ,			pvRow
			
		End If
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1    
       .vspdData.ReDraw = False
		ggoSpread.SSSetRequired    C_ITEM_CD,          pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_ITEM_NM,          pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_MAT_LOT_NO,       pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_SELLER_CD,        pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_SELLER_NM,        pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_PRINT_DT,         pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_RCPT_DT,          pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_RCPT_TM,          pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_RCPT_QTY,         pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_SNP,              pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_BOX_QTY,          pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_DELETE_FLAG,      pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected   C_IF_SEQ,		   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_MES_RECEIVE_FLAG, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_MES_RECEIVE_DT,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_ERP_SEND_DT,	   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_ERR_DESC,         pvStartRow, pvEndRow
       .vspdData.ReDraw = True
    End With
    
End Sub

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
			C_ITEM_CD          = iCurColumnPos(1)
			C_ITEM_CD_POPUP    = iCurColumnPos(2)
			C_ITEM_NM          = iCurColumnPos(3)
			C_MAT_LOT_NO       = iCurColumnPos(4)
			C_SELLER_CD        = iCurColumnPos(5)
			C_SELLER_CD_POPUP  = iCurColumnPos(6)
			C_SELLER_NM        = iCurColumnPos(7)
			C_PRODUCT_DT       = iCurColumnPos(8)
			C_PRINT_DT         = iCurColumnPos(9)
			C_RCPT_DT          = iCurColumnPos(10)
			C_RCPT_TM          = iCurColumnPos(11)
			C_RCPT_QTY         = iCurColumnPos(12)
			C_BP_ISSUE_NO      = iCurColumnPos(13)
			C_ISSUE_FLAG       = iCurColumnPos(14)
			C_ISSUE_FLAG_NM    = iCurColumnPos(15)
			C_PLANT_FLAG       = iCurColumnPos(16)
			C_PLANT_CD         = iCurColumnPos(17)
			C_PLANT_CD_POPUP   = iCurColumnPos(18)
			C_GATE_CD          = iCurColumnPos(19)
			C_SNP              = iCurColumnPos(20)
			C_BOX_QTY          = iCurColumnPos(21)
			C_SEPARATE_FLAG    = iCurColumnPos(22)
			C_DELIVERY_NO      = iCurColumnPos(23)
			C_ISSUE_DT         = iCurColumnPos(24)
			C_ISSUE_TIME       = iCurColumnPos(25)
			C_DELETE_FLAG      = iCurColumnPos(26)
			C_IF_SEQ		   = iCurColumnPos(27)
			C_MES_RECEIVE_FLAG = iCurColumnPos(28)
			C_MES_RECEIVE_DT   = iCurColumnPos(29)
			C_ERP_SEND_DT	   = iCurColumnPos(30)
			C_ERR_DESC         = iCurColumnPos(31)
			C_CREATE_TYPE      = iCurColumnPos(32)
    End Select

End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
'========================================================================================================


'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False															  '☜: Processing is NG
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										  '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
        															
    If Not chkField(Document, "1") Then									          '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables()                                                          '⊙: Initializes local global variables

	If DbQuery("MQ") = False Then                                                 '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()

    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "1")
	Call ggoOper.ClearField(Document, "2")										  '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
	Call ggoOper.LockField(Document, "N")        
	Call SetDefaultVal()
	Call SetToolBar("11101101001111")
	Call InitVariables()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()

    Dim intRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                       '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                          '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            
            

			Call SetSpreadColor(.ActiveRow, .ActiveRow)

            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_DELETE_FLAG
	frm1.vspdData.Text = "0"
	frm1.vspdData.Col = C_MES_RECEIVE_FLAG
	frm1.vspdData.Text = "N"
	frm1.vspdData.Col = C_MES_RECEIVE_DT
	frm1.vspdData.Text = UNIDateClientFormat("1900-01-01")
	frm1.vspdData.Col = C_ERR_DESC
	frm1.vspdData.Text = ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

    Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
   	 Frm1.vspdData.Col = C_ISSUE_FLAG    : iDx = Frm1.vspdData.value
     Frm1.vspdData.Col = C_ISSUE_FLAG_NM : Frm1.vspdData.value = iDx
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

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
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_DELETE_FLAG
	frm1.vspdData.Text = "0"
	frm1.vspdData.Col = C_MES_RECEIVE_FLAG
	frm1.vspdData.Text = "N"
	frm1.vspdData.Col = C_MES_RECEIVE_DT
	frm1.vspdData.Text = UNIDateClientFormat("1900-01-01")
	frm1.vspdData.Col = C_ERR_DESC
	frm1.vspdData.Text = ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()

    Dim lDelRows

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then	 
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

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
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	
End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call MakeKeyStream(pDirect)
    With Frm1
		strVal = BIZ_PGM_ID1 & "?txtMode="       & Parent.UID_M0001						         
        strVal = strVal      & "&txtKeyStream="  & lgKeyStream         '☜: Query Key
        strVal = strVal      & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal      & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim Indx
    Dim strVal, strDel
    Dim ColSep, RowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size
	
	On Error Resume Next
	Err.Clear

    DbSave = False                                     '⊙: Processing is NG
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value         = parent.UID_M0002      '☜: 저장 상태 
		.txtFlgMode.value      = lgIntFlgMode          '☜: 신규입력/수정 상태 
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	'한번에 설정한 버퍼의 크기 설정 
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    iFormLimitByte       = parent.C_FORM_LIMIT_BYTE     '102399byte
    
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			'버퍼의 초기화 
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	ColSep = parent.gColSep : RowSep = parent.gRowSep 
	iTmpCUBufferCount = -1  : iTmpDBufferCount = -1
	strCUTotalvalLen  = 0   : strDTotalvalLen  = 0

	With frm1.vspdData
		For Indx = 1 To .MaxRows
			.Row = Indx
			.Col = 0
			Select Case .Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
			    	strVal = ""
			    	If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & ColSep              '⊙: C=Create, Sheet가 2개 이므로 구별 
					Else
						strVal = strVal & "U" & ColSep              '⊙: U=Update
			    	End If
			    	                            strVal = strVal & Indx & ColSep
					.Col = C_ITEM_CD          : strVal = strVal & UCase(Trim(.Text))         & ColSep '품목코드(PK)
					.Col = C_ITEM_NM          : strVal = strVal & UCase(Trim(.Text))         & ColSep '품목명
					.Col = C_MAT_LOT_NO       : strVal = strVal & UCase(Trim(.Text))         & ColSep 'LOT번호(PK)
					.Col = C_SELLER_CD        : strVal = strVal & UCase(Trim(.Text))         & ColSep '납품처코드(PK)
					.Col = C_PRODUCT_DT       : strVal = strVal & UNIConvDate(Trim(.Text))   & ColSep '생산일자(Hidden)
					.Col = C_PRINT_DT         : strVal = strVal & UNIConvDate(Trim(.Text))   & ColSep '발행일자(PK)
					.Col = C_RCPT_DT          : strVal = strVal & UNIConvDate(Trim(.Text))   & ColSep '납입일자(PK)
					.Col = C_RCPT_TM          : strVal = strVal & UCase(Trim(.Text))         & ColSep '납입시간(PK)
					.Col = C_RCPT_QTY         : strVal = strVal & UNIConvNum(Trim(.Text), 0) & ColSep '납입수량
					.Col = C_BP_ISSUE_NO      : strVal = strVal & UCase(Trim(.Text))         & ColSep '납품처발행번호
					.Col = C_ISSUE_FLAG       : strVal = strVal & UCase(Trim(.Text))         & ColSep '발행구분
					.Col = C_PLANT_FLAG       : strVal = strVal & UCase(Trim(.Text))         & ColSep '공장구분
					.Col = C_PLANT_CD         : strVal = strVal & UCase(Trim(.Text))         & ColSep '공장코드
					.Col = C_GATE_CD          : strVal = strVal & UCase(Trim(.Text))         & ColSep 'GATE
					.Col = C_SNP              : strVal = strVal & UNIConvNum(Trim(.Text), 0) & ColSep 'SNP
					.Col = C_BOX_QTY          : strVal = strVal & UNIConvNum(Trim(.Text), 0) & ColSep 'Box수량
					.Col = C_SEPARATE_FLAG    
					If UCase(Trim(.Text)) = "1" Then
						strVal = strVal & "Y" & ColSep '분할구분
					Else
						strVal = strVal & "N" & ColSep '분할구분
					End If
					.Col = C_DELIVERY_NO      : strVal = strVal & UCase(Trim(.Text))         & ColSep '납품번호
					.Col = C_ISSUE_DT         : strVal = strVal & UNIConvDate(Trim(.Text))   & ColSep '출하일자
					.Col = C_ISSUE_TIME       : strVal = strVal & UCase(Trim(.Text))         & ColSep '출하시간
					.Col = C_DELETE_FLAG      
					If UCase(Trim(.Text)) = "1" Then
						strVal = strVal & "Y" & ColSep '삭제여부
					Else
						strVal = strVal & "N" & ColSep '삭제여부
					End If
					.Col = C_MES_RECEIVE_FLAG : strVal = strVal & UCase(Trim(.Text))         & ColSep 'MES수신여부
					.Col = C_MES_RECEIVE_DT   : strVal = strVal & UCase(Trim(.Text))         & ColSep 'MES수신일시
					.Col = C_ERR_DESC         : strVal = strVal & UCase(Trim(.Text))         & ColSep '에러내역
					.Col = C_CREATE_TYPE      : strVal = strVal & UCase(Trim(.Text))         & ColSep '생성구분(PK)
					                            strVal = strVal & Indx & RowSep
			    Case ggoSpread.DeleteFlag
					strDel = ""
					                            strDel = strDel & "D"  & ColSep                '⊙: D=Delete
					                            strDel = strDel & Indx & ColSep
					.Col = C_ITEM_CD          : strDel = strDel & UCase(Trim(.Text))         & ColSep '품목코드(PK)
					.Col = C_ITEM_NM          : strDel = strDel & UCase(Trim(.Text))         & ColSep '품목명
					.Col = C_MAT_LOT_NO       : strDel = strDel & UCase(Trim(.Text))         & ColSep 'LOT번호(PK)
					.Col = C_SELLER_CD        : strDel = strDel & UCase(Trim(.Text))         & ColSep '납품처코드(PK)
					.Col = C_PRODUCT_DT       : strDel = strDel & UNIConvDate(Trim(.Text))   & ColSep '생산일자(Hidden)
					.Col = C_PRINT_DT         : strDel = strDel & UNIConvDate(Trim(.Text))   & ColSep '발행일자(PK)
					.Col = C_RCPT_DT          : strDel = strDel & UNIConvDate(Trim(.Text))   & ColSep '납입일자(PK)
					.Col = C_RCPT_TM          : strDel = strDel & UCase(Trim(.Text))         & ColSep '납입시간(PK)
					.Col = C_RCPT_QTY         : strDel = strDel & UNIConvNum(Trim(.Text), 0) & ColSep '납입수량
					.Col = C_BP_ISSUE_NO      : strDel = strDel & UCase(Trim(.Text))         & ColSep '납품처발행번호
					.Col = C_ISSUE_FLAG       : strDel = strDel & UCase(Trim(.Text))         & ColSep '발행구분
					.Col = C_PLANT_FLAG       : strDel = strDel & UCase(Trim(.Text))         & ColSep '공장구분
					.Col = C_PLANT_CD         : strDel = strDel & UCase(Trim(.Text))         & ColSep '공장코드
					.Col = C_GATE_CD          : strDel = strDel & UCase(Trim(.Text))         & ColSep 'GATE
					.Col = C_SNP              : strDel = strDel & UNIConvNum(Trim(.Text), 0) & ColSep 'SNP
					.Col = C_BOX_QTY          : strDel = strDel & UNIConvNum(Trim(.Text), 0) & ColSep 'Box수량
					.Col = C_SEPARATE_FLAG    
					If UCase(Trim(.Text)) = "1" Then
						strDel = strDel & "Y" & ColSep '분할구분
					Else
						strDel = strDel & "N" & ColSep '분할구분
					End If
					.Col = C_DELIVERY_NO      : strDel = strDel & UCase(Trim(.Text))         & ColSep '납품번호
					.Col = C_ISSUE_DT         : strDel = strDel & UNIConvDate(Trim(.Text))   & ColSep '출하일자
					.Col = C_ISSUE_TIME       : strDel = strDel & UCase(Trim(.Text))         & ColSep '출하시간
					.Col = C_DELETE_FLAG      
					If UCase(Trim(.Text)) = "1" Then
						strDel = strDel & "Y" & ColSep '삭제여부
					Else
						strDel = strDel & "N" & ColSep '삭제여부
					End If
					.Col = C_MES_RECEIVE_FLAG : strDel = strDel & UCase(Trim(.Text))         & ColSep 'MES수신여부
					.Col = C_MES_RECEIVE_DT   : strDel = strDel & UCase(Trim(.Text))         & ColSep 'MES수신일시
					.Col = C_ERR_DESC         : strDel = strDel & UCase(Trim(.Text))         & ColSep '에러내역
					.Col = C_CREATE_TYPE      : strDel = strDel & UCase(Trim(.Text))         & ColSep '생성구분(PK)
					                            strDel = strDel & Indx & RowSep
			End Select
			
			.Col = 0
			Select Case .Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then
			            Set objTEXTAREA = document.createElement("TEXTAREA")
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then
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
			         If iTmpDBufferCount > iTmpDBufferMaxCount Then
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID1)

    If Err.number = 0 Then	 
       DbSave = True
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
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

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	
	Dim Indx
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolBar("11101101001111")                                              '☆: Developer must customize
	Frm1.vspdData.Focus
    Call InitData()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.ReDraw = False
		For Indx = 1 To frm1.vspdData.MaxRows
			Call SetSpreadSetLock(Indx)
		Next
		frm1.vspdData.ReDraw = True
	End If 
	
    Set gActiveElement = document.ActiveElement   
    ggospread.source = frm1.vspdData
    
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Call InitVariables()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	
	Call RemovedivTextArea() 
	
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
	                                              '☆: Developer must customize
    If DbQuery("MQ") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement   
    
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : OpenPopUp()
' Desc : Item PopUp
'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0, 2
			arrParam(0) = "품  목"                                   ' 팝업    명칭
			arrParam(1) = "B_ITEM"                                   ' TABLE   명칭
			arrParam(2) = Trim(strCode)                              ' Code    Condition
			arrParam(3) = ""                                         ' Name    Cindition
			arrParam(4) = ""                                         ' Where   Condition
			arrParam(5) = "품  목"                                   ' TextBox 명칭
		
			arrField(0) = "ITEM_CD"                                  ' Field명(0)
			arrField(1) = "ITEM_NM"                                  ' Field명(1)
	    
			arrHeader(0) = "품목코드"                                ' Header명(0)
			arrHeader(1) = "품 목 명"                                ' Header명(1)
		
		Case 1, 3
			arrParam(0) = "납품처"
			arrParam(1) = "B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B"
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = " A.BP_CD = B.PARTNER_BP_CD"
			arrParam(4) = arrParam(4) & " AND B.PARTNER_FTN = " & FilterVar("SSH", "''", "S")
			arrParam(4) = arrParam(4) & " AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S")
			arrParam(5) = "납품처"
		
			arrField(0) = "B.PARTNER_BP_CD"
			arrField(1) = "A.BP_NM"
	    
			arrHeader(0) = "납품처코드"
			arrHeader(1) = "납품처명"
	
		Case 4
			arrParam(0) = "공장"						
			arrParam(1) = "B_PLANT"
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""							
			arrParam(5) = "공장"						

			arrField(0) = "PLANT_CD"
			arrField(1) = "PLANT_NM"
    
			arrHeader(0) = "공장"					
			arrHeader(1) = "공장명"					

	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCdNmValue(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================
' Name : SetCdNmValue()
' Desc : Popup Return Value Setting
'========================================================================================================
Function SetCdNmValue(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0
				.txtItemCD.value = arrRet(0) 
				.txtItemNM.value = arrRet(1)   
				.txtItemCD.focus
				
			Case 1
				.txtBpCd.value = arrRet(0) 
				.txtBpNm.value = arrRet(1)
				.txtBpCd.focus
				
			Case 2
				.vspdData.Col  = C_ITEM_CD
				.vspdData.Text = arrRet(0) 
				
				.vspdData.Col  = C_ITEM_NM
				.vspdData.Text = arrRet(1)
				
				Call SetActiveCell(.vspdData, C_ITEM_CD, .vspdData.ActiveRow,"M","X","X")
				
			Case 3
				.vspdData.Col  = C_SELLER_CD
				.vspdData.Text = arrRet(0) 
				
				.vspdData.Col  = C_SELLER_NM
				.vspdData.Text = arrRet(1)
				
				Call SetActiveCell(.vspdData, C_SELLER_CD, .vspdData.ActiveRow,"M","X","X")
			
			Case 4
				.vspdData.Col  = C_PLANT_CD
				.vspdData.Text = arrRet(0) 
				
				Call SetActiveCell(.vspdData, C_PLANT_CD, .vspdData.ActiveRow,"M","X","X")
				
		End Select
	End With

End Function

'===============================================================================================
'   by Shin hyoung jae 
'	Name : GetOpenFilePath()
'	Description : GetTextFilePath	
'================================================================================================= 
Function GetOpenFilePath()

	Dim dlg
    Dim sPath
 
	On Error Resume Next
	Err.Clear 
	
	Set dlg = CreateObject("uni2kCM.SaveFile")
	
	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If
	
    sPath = dlg.GetOpenFilePath()

	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If

	lgFilePath2 = sPath
	frm1.txtFileName.Value = ExtractFileName(sPath)

    Set dlg = Nothing
	frm1.hFilePath.value = sPath
	
End Function

Function ExtractFileName(byVal strPath)

	strPath = StrReverse(strPath)
	strPath = Left(strPath, InStr(strPath, "\") - 1)
	ExtractFileName = StrReverse(strPath)
	
End Function

Function FncImportExcel()

	On Error Resume Next                                                                 '☜: Protect system from crashing
	Err.Clear                                                                            '☜: Clear Error status

	FncImportExcel = False
	    
    Const IG_I1_file_path = 0
    Const IG_I1_excel_sheet = 1
    Const IG_I1_sheet_start_row = 2
    Const IG_I1_sheet_start_col = 3
    
	'EXPORT GROUP
	Const C1_print_dt         = 1 '발행일자(pk)
	Const C1_rcpt_dt          = 2 '납입일자(pk)
	Const C1_rcpt_tm          = 3 '납입시간(pk)
	Const C1_bp_issue_no      = 4 '납품처발행번호
	Const C1_issue_flag       = 5 '발행구분
	Const C1_plant_flag       = 6 '공장구분
	Const C1_gate_cd          = 7 'gate
	Const C1_seller_cd        = 8 '납품처코드(pk)
	Const C1_item_cd          = 9 '품목코드(pk)
	Const C1_item_nm          = 10'품목명
	Const C1_snp              = 11'snp
	Const C1_rcpt_qty         = 12'납입수량
	Const C1_box_qty          = 13'box수량
	Const C1_separate_flag    = 14'분할구분
	Const C1_delivery_no      = 15'납품번호
	Const C1_issue_dt_tm      = 16'출하일시
	Const C1_mat_lot_no       = 17'lot번호(pk)
'	Const C1_plant_cd         = 18'공장코드
	
	Dim lgLngMaxRow
	Dim IG1Array
	Dim EG1Data
	Dim lgstrData
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim arrVal
	Dim iLngRow

	lgLngMaxRow = frm1.vspdData.MaxRows
	
	ReDim IG1Array(IG_I1_sheet_start_col)	
    'Key 값을 읽어온다	
    IG1Array(IG_I1_file_path)       = frm1.hFilePath.Value
'    IG1Array(IG_I1_excel_sheet)	    = "Sheet1"
    IG1Array(IG_I1_sheet_start_row)	= "2"
    IG1Array(IG_I1_sheet_start_col)	= C1_print_dt
	
    If Err.Number <> 0 Then
	    Call DisplayMsgBox("X", "X", Err.Number & " : " & Err.Description, "X")
	    Exit Function
	End If	
    
    If frm1.cFLkUpExcel.F_LOOKUP_EXCEL(IG1Array, EG1Data) = False Then
		If Err.Number <> 0 Then
			Call DisplayMsgBox("X", "X", Err.Number & " : " & Err.Description, "X")
		Else
			Call DisplayMsgBox("X", "X", "Excel 파일을 정상적으로 읽지 못했습니다.", "X")
		End If
	    Exit Function
	End If
	
	If Err.Number <> 0 Then
		Call DisplayMsgBox("X", "X", Err.Number & " : " & Err.Description, "X")
	    Exit Function
	End If	

	If IsEmpty(EG1Data) = False Then
		lgstrData = ""
		For iLngRow = CLng(IG1Array(IG_I1_sheet_start_row)) To UBound(EG1Data, 1)
			If FncInsertRow(1) Then
				With frm1.vspdData
					.Row = .ActiveRow
					.Col = C_ITEM_CD       : .Text = EG1Data(iLngRow, C1_item_cd)                        '품목코드(PK)
					.Col = C_ITEM_NM       : .Text = EG1Data(iLngRow, C1_item_nm)                        '품목명
					.Col = C_MAT_LOT_NO    : .Text = EG1Data(iLngRow, C1_mat_lot_no)                     'LOT번호(PK)
					.Col = C_SELLER_CD     : .Text = EG1Data(iLngRow, C1_seller_cd)                      '납품처코드(PK)
					.Col = C_SELLER_NM     : .Text = EG1Data(iLngRow, C1_seller_cd)                      '납품처코드(PK)
					strSelect = "A.BP_NM"
					strFrom = "B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B"
					strWhere = "A.BP_CD = B.PARTNER_BP_CD"
					strWhere = strWhere & " AND B.PARTNER_FTN = " & FilterVar("SSH", "''", "S")
					strWhere = strWhere & " AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S")
					strWhere = strWhere & " AND B.PARTNER_BP_CD = " & FilterVar(EG1Data(iLngRow, C1_seller_cd), "''", "S")
					If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						.Text= Trim(arrVal(0))'납품처명 
					Else
						.Text= ""
					End If
					.Col = C_PRODUCT_DT       : .Text = ""'생산일자(Hidden)
					.Col = C_PRINT_DT         : .Text = UNIDateClientFormat(EG1Data(iLngRow, C1_print_dt)) '발행일자(PK)
					.Col = C_RCPT_DT          : .Text = UNIDateClientFormat(EG1Data(iLngRow, C1_rcpt_dt))  '납입일자(PK)
					.Col = C_RCPT_TM          : .Text = ConvToSSSSetTime(EG1Data(iLngRow, C1_rcpt_tm))     '납입시간(PK)
					.Col = C_RCPT_QTY         : .Text = UNICDbl(EG1Data(iLngRow, C1_rcpt_qty))             '납입수량
					.Col = C_BP_ISSUE_NO      : .Text = EG1Data(iLngRow, C1_bp_issue_no)                   '납품처발행번호
					.Col = C_ISSUE_FLAG       : .Text = EG1Data(iLngRow, C1_issue_flag)                    '발행구분
					.Col = C_PLANT_FLAG       : .Text = EG1Data(iLngRow, C1_plant_flag)                    '공장구분
'					.Col = C_PLANT_CD         : .Text = EG1Data(iLngRow, C1_plant_cd)                      '공장코드
					.Col = C_PLANT_CD         : .Text = ""                                                 '공장코드
					.Col = C_GATE_CD          : .Text = EG1Data(iLngRow, C1_gate_cd)                       'GATE
					.Col = C_SNP              : .Text = UNICDbl(EG1Data(iLngRow, C1_snp))                  'SNP
					.Col = C_BOX_QTY          : .Text = UNICDbl(EG1Data(iLngRow, C1_box_qty))              'Box수량
					If Ucase(Trim(EG1Data(iLngRow, C1_separate_flag))) = "Y" Then
						.Col = C_SEPARATE_FLAG    : .Text = "1"                                            '분할구분
					Else
						.Col = C_SEPARATE_FLAG    : .Text = "0"
					End If
					.Col = C_DELIVERY_NO      : .Text = EG1Data(iLngRow, C1_delivery_no)                   '납품번호
					.Col = C_ISSUE_DT         : .Text = UNIDateClientFormat(Mid(EG1Data(iLngRow, C1_issue_dt_tm), 1, 10))'출하일자
					.Col = C_ISSUE_TIME       : .Text = ConvToSSSSetTime(Mid(EG1Data(iLngRow, C1_issue_dt_tm), 12))'출하시간
					.Col = C_DELETE_FLAG      : .Text = "0"                                                '삭제여부
					.Col = C_MES_RECEIVE_FLAG : .Text = "N"                                                'MES수신여부
					.Col = C_MES_RECEIVE_DT   : .Text = ""                                                 'MES수신일시
					.Col = C_ERR_DESC         : .Text = ""                                                 '에러내역
					.Col = C_CREATE_TYPE      : .Text = "A"                                                '생성구분(PK)
				End With
			End If
		Next
    End If
    
    Call InitData()
	
	FncImportExcel = True
	
End Function

'============================================================================================================
' Name : ConvToSSSSetTime(iVal)
' Desc : 
'============================================================================================================
Function ConvToSSSSetTime(iVal)

	Dim TempTime
	
	On Error Resume Next
	Err.Clear 
	
	If Trim(IVal) = ":" Or Trim(IVal) = "00:" Or Trim(IVal) = ":00" Then
		ConvToSSSSetTime = "00:00"
	Else
		TempTime = Split(IVal, ":")
		If IsArray(TempTime) Then
			ConvToSSSSetTime = Right("0" & TempTime(0), 2) & ":" & Right("0" & TempTime(1), 2)
		Else
			ConvToSSSSetTime = "00:00"
		End If
	End If
	
End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 

    Dim strVal
    Dim IntRetCD
    
    On Error Resume Next
    Err.Clear 
    
    ExeReflect = False

	If trim(frm1.txtFileName.value) = "" Then
		Call DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X")
		frm1.txtFileName.focus 	
		Exit Function
	Else
		If (ggoSaveFile.fileExists(frm1.hFilePath.value) = 0) = False  Then
			IntRetCD = DisplayMsgBox("115191", "X", "X", "X")
			Exit Function
		End If
	End If		    
    
    If LayerShowHide(1) = False Then
        Exit Function
    End If

    If FncImportExcel = False Then
		If LayerShowHide(0) = False Then
		    Exit Function
		End If
		Exit Function
	End If

	If LayerShowHide(0) = False Then
        Exit Function
    End If
	
    ExeReflect = True

End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim Indx
		
	For Indx = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
  
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
				Case C_ITEM_CD_POPUP
					.Col = Col - 1
					.Row = Row
					Call OpenPopUp(.Text, 2)
					
				Case C_SELLER_CD_POPUP
					.Col = Col - 1
					.Row = Row
					Call OpenPopUp(.Text, 3)
					
				Case C_PLANT_CD_POPUP
					.Col = Col - 1
					.Row = Row
					Call OpenPopUp(.Text, 4)
					
			End Select
		End If    
		
	End With
	
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col
         Case  C_ISSUE_FLAG_NM
                iDx = frm1.vspdData.value
   	            frm1.vspdData.Col = C_ISSUE_FLAG
                frm1.vspdData.value = iDx
                
         Case Else
         
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("1101111111")    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)		
		
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub


'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()

    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
    
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


'=======================================================================================================
'   Event Name : txtPrintFrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPrintFrDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtPrintFrDt.Action = 7
       Call SetFocusToDocument("M")
       Frm1.txtPrintFrDt.Focus
	End If
End Sub

Sub txtPrintFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
	End If
End Sub

Sub txtPrintToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtPrintToDt.Action = 7
       Call SetFocusToDocument("M")
       Frm1.txtPrintToDt.Focus
	End If
End Sub

Sub txtPrintToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
	End If
End Sub
