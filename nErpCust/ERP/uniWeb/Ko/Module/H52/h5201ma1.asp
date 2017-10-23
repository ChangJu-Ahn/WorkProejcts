<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        :
*  3. Program ID           : h5201ma1
*  4. Program Name         : 저축사항 등록 
*  5. Program Desc         : 저축사항 조회,등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const BIZ_PGM_ID      = "h5201mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow
Dim lgBlnDataChgFlg
Dim lsInternal_cd

Dim C_HDC010T_SAVE_CD
Dim C_HDC010T_SAVE_CD_NM
Dim C_HDC010T_SAVE_TYPE
Dim C_HDC010T_SAVE_TYPE_NM
Dim C_HDC010T_BANK_ACCNT
Dim C_HDC010T_BANK_CD
Dim C_HDC010T_BANK_CD_POP
Dim C_FAA090T_BANK_NAME
Dim C_HDC010T_SCRIPT_AMT
Dim C_HDC010T_TOT_SCRIPT_CNT
Dim C_EXPIRE_AMT
Dim C_HDC010T_NEW_DT
Dim C_HDC010T_EXPIR_DT
Dim C_HDC010T_REVOKE_DT
Dim C_HDC010T_TAX_RATE

'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_HDC010T_SAVE_CD			= 1
	 C_HDC010T_SAVE_CD_NM		= 2
	 C_HDC010T_SAVE_TYPE		= 3
	 C_HDC010T_SAVE_TYPE_NM		= 4
	 C_HDC010T_BANK_ACCNT		= 5
	 C_HDC010T_BANK_CD			= 6
	 C_HDC010T_BANK_CD_POP		= 7
	 C_FAA090T_BANK_NAME		= 8
	 C_HDC010T_SCRIPT_AMT		= 9
	 C_HDC010T_TOT_SCRIPT_CNT	= 10
	 C_EXPIRE_AMT				= 11
	 C_HDC010T_NEW_DT			= 12
	 C_HDC010T_EXPIR_DT			= 13
	 C_HDC010T_REVOKE_DT		= 14
	 C_HDC010T_TAX_RATE			= 15
End Sub

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
	lgOldRow = 0
	lsInternal_cd     = ""

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
	lgBlnDataChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
        With frm1
		        .txtHdc010t_save_cd.value = ""
		        .txtHdc010t_script_amt.value = 0
		        .txtHdc010t_new_dt.text = ""
		        .txtHdc010t_save_type.value = ""
		        .txtHdc010t_tot_script_cnt.value = 0
		        .txtHdc010t_expir_dt.text = ""
		        .txtHdc010t_bank_accnt.value = ""
		        .txtHdc010t_expir_amt.value = 0
		        .txtHdc010t_revoke_dt.text = ""
		        .txtHdc010t_bank_cd.value = ""
		        .txtHdc010t_tax_rate.value = 0
		        .txtFaa090t_bank_name.value = ""
		End With
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   lgKeyStream       = Trim(Frm1.txtEmp_no.value) & parent.gColSep                   'You Must append one character( parent.gColSep)
    If  lsInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    Else
        lgKeyStream = lgKeyStream & lsInternal_cd & parent.gColSep
    End If
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
    Dim iDx

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0041", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '저축코드 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_HDC010T_SAVE_CD
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_HDC010T_SAVE_CD_NM

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0042", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '저축구분 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_HDC010T_SAVE_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_HDC010T_SAVE_TYPE_NM
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex

		With frm1.vspdData

		    For intRow = 1 To .MaxRows
		    	.Row = intRow
		    	.Col = C_HDC010T_SAVE_CD
		    	intIndex = .Value
		    	.col = C_HDC010T_SAVE_CD_NM
		    	.Value = intindex

		    	.Col = C_HDC010T_SAVE_TYPE
		    	intIndex = .Value
		    	.col = C_HDC010T_SAVE_TYPE_NM
		    	.Value = intindex
		    Next
        End With
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1 , 1)
		Frm1.vspdData.focus
        Set gActiveElement = document.ActiveElement
	End If
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_HDC010T_TAX_RATE + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True

        .MaxRows = 0
         ggoSpread.ClearSpreadData

			 Call  AppendNumberPlace("6","3","0")
			 Call  AppendNumberPlace("7","2","2")	

			 Call  GetSpreadColumnPos("A")
        
             ggoSpread.SSSetCombo    C_HDC010T_SAVE_CD,         "",8
			 ggoSpread.SSSetCombo    C_HDC010T_SAVE_CD_NM,      "저축코드",12,,False
             ggoSpread.SSSetCombo    C_HDC010T_SAVE_TYPE,       "",8
             ggoSpread.SSSetCombo    C_HDC010T_SAVE_TYPE_NM,    "저축구분",12,,False
             ggoSpread.SSSetEdit     C_HDC010T_BANK_ACCNT,      "계좌번호",12,,,20,2
             ggoSpread.SSSetEdit     C_HDC010T_BANK_CD,         "은행코드",10,,,10,2
             ggoSpread.SSSetButton   C_HDC010T_BANK_CD_POP
             ggoSpread.SSSetEdit     C_FAA090T_BANK_NAME,       "은행명", 20,,,30,2
             ggoSpread.SSSetFloat    C_HDC010T_SCRIPT_AMT,      "월불입액" ,18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_HDC010T_TOT_SCRIPT_CNT,  "총불입횟수" ,12,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetFloat    C_EXPIRE_AMT,              "만기금액" ,18, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             ggoSpread.SSSetDate     C_HDC010T_NEW_DT,          "신규가입일", 10,2,  parent.gDateFormat
             ggoSpread.SSSetDate     C_HDC010T_EXPIR_DT,        "만기일", 10,2,  parent.gDateFormat
             ggoSpread.SSSetDate     C_HDC010T_REVOKE_DT,       "해약일", 10,2,  parent.gDateFormat
             ggoSpread.SSSetFloat    C_HDC010T_TAX_RATE,        "세액공제율", 15,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
             
             Call ggoSpread.MakePairsColumn(C_HDC010T_BANK_CD	,  C_HDC010T_BANK_CD_POP)
             Call ggoSpread.SSSetColHidden(C_HDC010T_SAVE_CD	,  C_HDC010T_SAVE_CD	, True)
             Call ggoSpread.SSSetColHidden(C_HDC010T_SAVE_TYPE	,  C_HDC010T_SAVE_TYPE	, True)
             Call ggoSpread.SSSetColHidden(C_HDC010T_NEW_DT		,  C_HDC010T_NEW_DT		, True)
             Call ggoSpread.SSSetColHidden(C_HDC010T_EXPIR_DT	,  C_HDC010T_EXPIR_DT	, True)
             Call ggoSpread.SSSetColHidden(C_HDC010T_REVOKE_DT	,  C_HDC010T_REVOKE_DT	, True)
             Call ggoSpread.SSSetColHidden(C_HDC010T_TAX_RATE	,  C_HDC010T_TAX_RATE	, True)

	   .ReDraw = true

       Call SetSpreadLock

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	
    With frm1
        .vspdData.ReDraw = False
             ggoSpread.SpreadLock       C_HDC010T_SAVE_CD		, -1, C_HDC010T_SAVE_CD
             ggoSpread.SpreadLock       C_HDC010T_SAVE_CD_NM	, -1, C_HDC010T_SAVE_CD_NM
             ggoSpread.SpreadLock       C_HDC010T_SAVE_TYPE		, -1, C_HDC010T_SAVE_TYPE
             ggoSpread.SpreadLock       C_HDC010T_SAVE_TYPE_NM	, -1, C_HDC010T_SAVE_TYPE_NM
             ggoSpread.SpreadLock       C_HDC010T_BANK_ACCNT	, -1, C_HDC010T_BANK_ACCNT             
             ggoSpread.SpreadLock		C_HDC010T_BANK_CD		, -1, C_HDC010T_BANK_CD
             ggoSpread.SpreadLock       C_HDC010T_BANK_CD_POP	, -1, C_HDC010T_BANK_CD_POP
             ggoSpread.SpreadLock       C_FAA090T_BANK_NAME		, -1, C_FAA090T_BANK_NAME
             ggoSpread.SpreadLock       C_HDC010T_SCRIPT_AMT	, -1, C_HDC010T_SCRIPT_AMT
             ggoSpread.SpreadLock       C_HDC010T_TOT_SCRIPT_CNT, -1, C_HDC010T_TOT_SCRIPT_CNT
             ggoSpread.SpreadLock       C_EXPIRE_AMT			, -1, C_EXPIRE_AMT                 
             ggoSpread.SSSetProtected   .vspdData.MaxCols		, -1, -1
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
         ggoSpread.SSSetRequired    C_HDC010T_SAVE_CD_NM	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_HDC010T_SAVE_TYPE_NM	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_HDC010T_BANK_ACCNT	 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_HDC010T_BANK_CD		 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected   C_FAA090T_BANK_NAME		 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected   C_HDC010T_SCRIPT_AMT	 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected   C_HDC010T_TOT_SCRIPT_CNT , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected   C_EXPIRE_AMT			 , pvStartRow, pvEndRow
    .vspdData.ReDraw = True

    End With
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_HDC010T_SAVE_CD			= iCurColumnPos(1)
			C_HDC010T_SAVE_CD_NM		= iCurColumnPos(2)
			C_HDC010T_SAVE_TYPE			= iCurColumnPos(3)
			C_HDC010T_SAVE_TYPE_NM		= iCurColumnPos(4)
			C_HDC010T_BANK_ACCNT		= iCurColumnPos(5)
			C_HDC010T_BANK_CD			= iCurColumnPos(6)
			C_HDC010T_BANK_CD_POP		= iCurColumnPos(7)
			C_FAA090T_BANK_NAME			= iCurColumnPos(8)
			C_HDC010T_SCRIPT_AMT		= iCurColumnPos(9)
			C_HDC010T_TOT_SCRIPT_CNT	= iCurColumnPos(10)
			C_EXPIRE_AMT				= iCurColumnPos(11)
			C_HDC010T_NEW_DT			= iCurColumnPos(12)
			C_HDC010T_EXPIR_DT			= iCurColumnPos(13)
			C_HDC010T_REVOKE_DT			= iCurColumnPos(14)
			C_HDC010T_TAX_RATE			= iCurColumnPos(15)                        
            
    End Select    
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

	Call  AppendNumberPlace("7", "2", "2")
    Call  AppendNumberPlace("6", "3", "0")

    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

	Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar

    frm1.txtEmp_no.focus
    Call InitComboBox
	Call CookiePage (0)                                                             '☜: Check Cookie

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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt

    FncQuery = False                                                            '☜: Processing is NG

    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if

    FncQuery = True																'☜: Processing is OK

End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field

	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
    Call SetDefaultVal
    Call InitVariables                                                           '⊙: Initializes local global variables

    Set gActiveElement = document.ActiveElement

    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim iDx
    Dim IntRetCD
    Dim iRow
    Dim strNew_dt
    Dim strExpir_dt
    Dim strRevoke_dt

    FncSave = False                                                              '☜: Processing is NG

    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    If  txtEmp_no_Onchange() then
        Exit Function
    End If    
    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case  ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.Col =	C_FAA090T_BANK_NAME
					if Trim(.value) = "" then
						Call  DisplayMsgBox("970000","X","은행코드","X")
                        Set gActiveElement = document.activeElement						
       					exit function
					end if 					

                    If  UNICDbl(frm1.txtHdc010t_script_amt.Value) <= 0  Then
                        Call  DisplayMsgBox("800126","X","X","X")	            '월불입액은 0보다 커야합니다.
                        frm1.txtHdc010t_script_amt.focus()
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End If

                    .Col = C_HDC010T_TOT_SCRIPT_CNT
                    If  UNICDbl(.Value) <= 0  Then
                        Call  DisplayMsgBox("970021","X","총불입횟수","X")	'총불입횟수는 입력항목입니다.
                        frm1.txtHdc010t_tot_script_cnt.focus()
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End If

                    .Col = C_HDC010T_TAX_RATE
                    If .Text = ""  Then
                        .Value = 0
                    End If

   	                .Col = C_HDC010T_NEW_DT
                    strNew_dt = .Text
   	                .Col = C_HDC010T_EXPIR_DT
                    strExpir_dt = .Text
   	                .Col = C_HDC010T_REVOKE_DT
                    strRevoke_dt = .Text
                    If (strExpir_dt <> "" ) And  CompareDateByFormat(strNew_dt,strExpir_dt,"신규일","만기일","800121", parent.gDateFormat, parent.gComDateType,True) = False then	                    
	                    .Row = iRow
  	                    .Col = C_HDC010T_SAVE_CD_NM
  	                    .Action=0
                        Call vspdData_Click(C_HDC010T_NEW_DT , iRow )
                        frm1.txtHdc010t_new_dt.focus()
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    If (strRevoke_dt <> "" ) And  CompareDateByFormat(strNew_dt,strRevoke_dt,"해약일","신규일","800128", parent.gDateFormat, parent.gComDateType,True) = False then	                    
	                    .Row = iRow
  	                    .Col = C_HDC010T_SAVE_CD_NM
  	                    .Action=0
                        Call vspdData_Click(C_HDC010T_EXPIR_DT , iRow )
                        frm1.txtHdc010t_revoke_dt.focus()
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    If (strRevoke_dt <> "" ) And  CompareDateByFormat(strRevoke_dt,strExpir_dt,"해약일","만기일","800127", parent.gDateFormat, parent.gComDateType,True) = False then	                    
	                    .Row = iRow
  	                    .Col = C_HDC010T_SAVE_CD_NM
  	                    .Action=0
                        Call vspdData_Click(C_HDC010T_EXPIR_DT , iRow )
                        frm1.txtHdc010t_revoke_dt.focus()
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if

           End Select

        Next
    End With

    Call MakeKeyStream("X")
    If DbSave = False Then
       Exit Function
    End If
    FncSave = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

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
	
	With Frm1.VspdData			
            .Col  = C_HDC010T_SAVE_CD
            .Text = ""
            .Col  = C_HDC010T_SAVE_CD_NM
            .Text = ""
            .Col  = C_HDC010T_SAVE_TYPE
            .Text = ""
            .Col  = C_HDC010T_SAVE_TYPE_NM
            .Text = ""
            .Col  = C_HDC010T_BANK_ACCNT
            .Text = ""
    End With
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
     ggoSpread.Source = frm1.vspdData
     ggoSpread.EditUndo
    Call  initData()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow, iCnt

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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1            

        For iCnt = 1 To imRow 
            .vspdData.Row = .vspdData.ActiveRow + iCnt - 1
            .vspdData.col = C_HDC010T_NEW_DT
            .vspdData.text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
        Next
        
       .vspdData.ReDraw = True

        Call vspdData_Click(1, .vspdData.ActiveRow)
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
    With Frm1.vspdData
    	.focus()
    	 ggoSpread.Source = frm1.vspdData
    	lDelRows =  ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev()

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz

    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext()
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first.
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area

    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz

    FncNext = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind( parent.C_SINGLE, True)
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
    Call InitComboBox         
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit?
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
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status

	DbSave = False														         '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                                  strVal = strVal & "C" & parent.gColSep
                                                                  strVal = strVal & lRow & parent.gColSep
                                                                  strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_CD  	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_TYPE       	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_BANK_ACCNT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_BANK_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SCRIPT_AMT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_TOT_SCRIPT_CNT    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EXPIRE_AMT                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_NEW_DT            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_EXPIR_DT          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_REVOKE_DT         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_TAX_RATE          : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                                  strVal = strVal & "U" & parent.gColSep
                                                                  strVal = strVal & lRow & parent.gColSep
                                                                  strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_CD  	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_TYPE       	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_BANK_ACCNT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_BANK_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SCRIPT_AMT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_TOT_SCRIPT_CNT    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EXPIRE_AMT                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_NEW_DT            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_EXPIR_DT          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_REVOKE_DT         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_TAX_RATE          : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
					
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                                  strDel = strDel & "D" & parent.gColSep
                                                                  strDel = strDel & lRow & parent.gColSep
                                                                  strDel = strDel & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_CD  	        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_SAVE_TYPE       	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HDC010T_BANK_ACCNT        : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtEmp_no.focus()

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    Call vspdData_Click(1, 1)
    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd        			' Internal_cd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case C_HDC010T_BANK_CD_POP
	        arrParam(0) = "은행코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_bank"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                   			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = ""	                		    	' Where Condition
	    	arrParam(5) = "은행코드" 			            ' TextBox 명칭 

	    	arrField(0) = "bank_cd"						    	' Field명(0)
	    	arrField(1) = "bank_full_nm"    			    	' Field명(1)
	    	arrField(2) = "bank_nm"           					' Field명(2)

	    	arrHeader(0) = "은행코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "은행명"	          		        ' Header명(1)
	    	arrHeader(2) = "은행약어명"	    		        ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_HDC010T_BANK_CD
   	    frm1.vspdData.action = 0
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_HDC010T_BANK_CD_POP
   	            .txtHdc010t_bank_cd.value = arrRet(0)
		    	.vspdData.Col = C_FAA090T_BANK_NAME
		    	.vspdData.text = arrRet(1)
   	            .txtFaa090t_bank_name.value = arrRet(1)
		        .vspdData.Col = C_HDC010T_BANK_CD
		    	.vspdData.text = arrRet(0)
   	            .vspdData.action = 0
        End Select

	End With

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strCode
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col-1
    strCode = frm1.vspdData.text
	Select Case Col
	    Case C_HDC010T_BANK_CD_POP
                    Call OpenCode(strCode, C_HDC010T_BANK_CD_POP, Row)
    End Select
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCD

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With Frm1
        Select Case Col
             Case  C_HDC010T_SAVE_CD_NM
                    iDx = Frm1.vspdData.text
   	                .txtHdc010t_save_cd.value = iDx
             Case  C_HDC010T_SAVE_TYPE_NM
                    iDx = Frm1.vspdData.text
   	                .txtHdc010t_save_type.value = iDx
             Case  C_HDC010T_BANK_ACCNT
                    iDx = Frm1.vspdData.value
   	                .txtHdc010t_bank_accnt.value = iDx
             Case  C_HDC010T_BANK_CD
                    If Trim(.vspdData.Text) = "" Then
                        .vspdData.Col = C_FAA090T_BANK_NAME
                        .vspdData.Text=""
                        .txtHdc010t_bank_cd.value = ""
     	                .txtFaa090t_bank_name.value = ""
                    Else             
                        IntRetCD =  CommonQueryRs(" bank_cd,bank_nm "," b_bank "," bank_cd =  " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                        If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
                            .vspdData.Col = C_FAA090T_BANK_NAME
                            .vspdData.Text= ""
         	                .txtFaa090t_bank_name.value = ""
                        Else
                            .txtHdc010t_bank_cd.value = .vspdData.Text
    	                    .vspdData.Col = C_FAA090T_BANK_NAME
                            .vspdData.Text=Trim(Replace(lgF1,Chr(11),""))
         	                .txtFaa090t_bank_name.value = .vspdData.Text
                        End If
                    End If
             Case  C_FAA090T_BANK_NAME
                    IntRetCD =  CommonQueryRs(" bank_cd,bank_nm "," b_bank "," bank_nm =  " & FilterVar(.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                    Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
                        frm1.vspdData.Text=""
                    ElseIf  CountStrings(lgF0, Chr(11) ) > 1 Then                     ' 같은명일 경우 pop up
                        Call OpenCode(.vspdData.Text, C_FAA090T_BANK_NAME_POP, Row)
                    Else
    	                .vspdData.Col = C_HDC010T_BANK_CD
                        .vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
                    End If

                    iDx = Frm1.vspdData.value
   	                .txtFaa090t_bank_name.value = iDx
             Case  C_HDC010T_SCRIPT_AMT
                    .Col = 0
                    .Text =  ggoSpread.UpdateFlag
             Case  C_HDC010T_TOT_SCRIPT_CNT
                    .Col = 0
                    .Text =  ggoSpread.UpdateFlag
        End Select
    End With

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
 Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData
    
    If lgOldRow <> Row and row <> 0 Then    

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = Row

		lgOldRow = Row

		With frm1
		            .vspdData.Col = 0
		            .vspdData.Col = C_HDC010T_SAVE_CD_NM
		            .txtHdc010t_save_cd.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_SCRIPT_AMT
		            .txtHdc010t_script_amt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_NEW_DT
		            .txtHdc010t_new_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_SAVE_TYPE_NM
		            .txtHdc010t_save_type.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_TOT_SCRIPT_CNT
		            .txtHdc010t_tot_script_cnt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_EXPIR_DT
		            .txtHdc010t_expir_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_BANK_ACCNT
		            .txtHdc010t_bank_accnt.value = .vspdData.Text

		            .vspdData.Col = C_EXPIRE_AMT
		            .txtHdc010t_expir_amt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_REVOKE_DT
		            .txtHdc010t_revoke_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_BANK_CD
		            .txtHdc010t_bank_cd.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_TAX_RATE
		            .txtHdc010t_tax_rate.value = .vspdData.Text

		            .vspdData.Col = C_FAA090T_BANK_NAME
		            .txtFaa090t_bank_name.value = .vspdData.Text
		End With
	End If    

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row     
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
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
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
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow

		With frm1

		            .vspdData.Col = C_HDC010T_SAVE_CD_NM
		            .txtHdc010t_save_cd.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_SCRIPT_AMT
		            .txtHdc010t_script_amt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_NEW_DT
		            .txtHdc010t_new_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_SAVE_TYPE_NM
		            .txtHdc010t_save_type.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_TOT_SCRIPT_CNT
		            .txtHdc010t_tot_script_cnt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_EXPIR_DT
		            .txtHdc010t_expir_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_BANK_ACCNT
		            .txtHdc010t_bank_accnt.value = .vspdData.Text

		            .vspdData.Col = C_EXPIRE_AMT
		            .txtHdc010t_expir_amt.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_REVOKE_DT
		            .txtHdc010t_revoke_dt.Text = .vspdData.Text

		            .vspdData.Col = C_HDC010T_BANK_CD
		            .txtHdc010t_bank_cd.value = .vspdData.Text

		            .vspdData.Col = C_HDC010T_TAX_RATE
		            .txtHdc010t_tax_rate.value = .vspdData.Text

		            .vspdData.Col = C_FAA090T_BANK_NAME
		            .txtFaa090t_bank_name.value = .vspdData.Text
		End With
End Sub


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_HDC010T_SAVE_CD_NM
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_HDC010T_SAVE_CD
                Frm1.vspdData.value = iDx
         Case  C_HDC010T_SAVE_TYPE_NM
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_HDC010T_SAVE_TYPE
                Frm1.vspdData.value = iDx
         Case Else
    End Select

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
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
'   Event Name : txtHdc010t_script_amt_Change
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_script_amt_OnBlur()
    If frm1.vspdData.MaxRows<>0  Then
        If lgBlnDataChgFlg = True Then
            With frm1
                If .txtHdc010t_script_amt.Text="" Then
                    .txtHdc010t_script_amt.Text=0
                End If
                If .txtHdc010t_tot_script_cnt.Text="" Then
                   .txtHdc010t_tot_script_cnt.Text=0
                End If
                .txtHdc010t_expir_amt.Text =  UNIFormatNumber( UNICDbl(.txtHdc010t_script_amt.Text) *  UNICDbl(.txtHdc010t_tot_script_cnt.Text), ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)

                .vspdData.Col=C_HDC010T_SCRIPT_AMT
                .vspdData.Text=.txtHdc010t_script_amt.Text
                .vspdData.Col=C_EXPIRE_AMT
                .vspdData.Text=.txtHdc010t_expir_amt.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
    End If
End Sub
'========================================================================================================
'   Event Name : txtHdc010t_script_amt_Click
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_script_amt_OnFocus()
    If frm1.vspdData.MaxRows<>0     Then
        lgBlnDataChgFlg = True
    End If
End Sub

'========================================================================================================
'   Event Name : txtHdc010t_tot_script_cnt_Change
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_tot_script_cnt_OnBlur()

    If frm1.vspdData.MaxRows<>0      Then
        If lgBlnDataChgFlg = True Then
            With frm1
                If .txtHdc010t_script_amt.Text="" Then
                    .txtHdc010t_script_amt.Text=0
                End If
                If .txtHdc010t_tot_script_cnt.Text="" Then
                    .txtHdc010t_tot_script_cnt.Text=0
                End If
                .txtHdc010t_expir_amt.Text =  UNIFormatNumber( UNICDbl(.txtHdc010t_script_amt.Text) *  UNICDbl(.txtHdc010t_tot_script_cnt.Text), ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
               	.vspdData.Col=C_HDC010T_TOT_SCRIPT_CNT
                .vspdData.Text=.txtHdc010t_tot_script_cnt.Text
                .vspdData.Col=C_EXPIRE_AMT
                .vspdData.Text=.txtHdc010t_expir_amt.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
    End If
End Sub

'========================================================================================================
'   Event Name : txtHdc010t_script_amt_Click
'   Event Desc : This function is data changed
'========================================================================================================

Sub txtHdc010t_tot_script_cnt_OnFocus()
    If frm1.vspdData.MaxRows<>0      Then
        lgBlnDataChgFlg = True
    End If
End Sub


'========================================================================================================
'   Event Name : txtHdc010t_new_dt_Change
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_new_dt_OnBlur()
    If frm1.vspdData.MaxRows<>0      Then
        If lgBlnDataChgFlg = True Then
            With frm1
                .vspdData.Row = .vspdData.ActiveRow
                .vspdData.Col=C_HDC010T_NEW_DT
                .vspdData.Text=.txtHdc010t_new_dt.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
    End If
End Sub
'========================================================================================================
'   Event Name : txtHdc010t_new_dt_OnFocus
'   Event Desc : This function is data changed
'========================================================================================================

Sub txtHdc010t_new_dt_OnFocus()
    If frm1.vspdData.MaxRows<>0      Then
        lgBlnDataChgFlg = True
    End If
End Sub

'========================================================================================================
'   Event Name : txtHdc010t_expir_dt_Change
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_expir_dt_OnBlur()
    If frm1.vspdData.MaxRows<>0      Then
        If lgBlnDataChgFlg = True Then
            With frm1
                .vspdData.Row = .vspdData.ActiveRow
                .vspdData.Col=C_HDC010T_EXPIR_DT
                .vspdData.Text=.txtHdc010t_expir_dt.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
    End If
End Sub
'========================================================================================================
'   Event Name : txtHdc010t_expir_dt_OnFocus
'   Event Desc : This function is data changed
'========================================================================================================

Sub txtHdc010t_expir_dt_OnFocus()
    If frm1.vspdData.MaxRows<>0      Then
        lgBlnDataChgFlg = True
    End If
End Sub


'========================================================================================================
'   Event Name : txtHdc010t_revoke_dt_Change
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_revoke_dt_OnBlur()
    If frm1.vspdData.MaxRows<>0      Then
        If lgBlnDataChgFlg = True Then
            With frm1
                .vspdData.Row = .vspdData.ActiveRow
                .vspdData.Col=C_HDC010T_REVOKE_DT
                .vspdData.Text=.txtHdc010t_revoke_dt.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
     End If
End Sub
'========================================================================================================
'   Event Name : txtHdc010t_revoke_dt_OnFocus
'   Event Desc : This function is data changed
'========================================================================================================

Sub txtHdc010t_revoke_dt_OnFocus()
    If frm1.vspdData.MaxRows<>0      Then
        lgBlnDataChgFlg = True
    End If
End Sub

'========================================================================================================
'   Event Name : txtHdc010t_tax_rate_OnBlur
'   Event Desc : This function is data changed
'========================================================================================================
Sub txtHdc010t_tax_rate_OnBlur()
    If frm1.vspdData.MaxRows<>0      Then
        If lgBlnDataChgFlg = True Then
            With frm1
                .vspdData.Row = .vspdData.ActiveRow
                .vspdData.Col=C_HDC010T_TAX_RATE
                .vspdData.Text=.txtHdc010t_tax_rate.Text
                .vspdData.Col = 0
                If .vspdData.Text <>  ggoSpread.InsertFlag And lgBlnFlgChgValue=True  Then
                    .vspdData.Text =  ggoSpread.UpdateFlag
                End If
            End With
         End If
        lgBlnDataChgFlg = False
     End If
End Sub
'========================================================================================================
'   Event Name : txtHdc010t_tax_rate_OnFocus
'   Event Desc : This function is data changed
'========================================================================================================

Sub txtHdc010t_tax_rate_OnFocus()
    If frm1.vspdData.MaxRows<>0      Then
        lgBlnDataChgFlg = True
    End If
End Sub


'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtHdc010t_new_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtHdc010t_new_dt.Action = 7
        lgBlnDataChgFlg = True
        Call txtHdc010t_new_dt_OnBlur()
		frm1.txtHdc010t_new_dt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtHdc010t_expir_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")        
        frm1.txtHdc010t_expir_dt.Action = 7
        lgBlnDataChgFlg = True
        Call txtHdc010t_expir_dt_OnBlur()
        frm1.txtHdc010t_expir_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtHdc010t_revoke_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")        
        frm1.txtHdc010t_revoke_dt.Action = 7
        lgBlnDataChgFlg = True
        Call txtHdc010t_revoke_dt_OnBlur()
        frm1.txtHdc010t_revoke_dt.focus
    End If
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_Onchange
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtEmp_no.value = ""
        frm1.txtName.value = ""
        frm1.txtDept_nm.value = ""
        frm1.txtRoll_pstn.value = ""
        frm1.txtPay_grd.value = ""
        frm1.txtEntr_dt.text = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            frm1.txtName.value = ""
            frm1.txtDept_nm.value = ""
            frm1.txtRoll_pstn.value = ""
            frm1.txtPay_grd.value = ""
            frm1.txtEntr_dt.Text = ""
            Call  ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text =  UNIDateClientFormat(strEntr_dt)

        End if
    End if

End Function
'========================================================================================================
'   Event Name :
'   Event Desc : Check whether chananges take palase.
'========================================================================================================
Sub txtHdc010t_script_amt_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub
Sub txtHdc010t_new_dt_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub
Sub txtHdc010t_tot_script_cnt_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub
Sub txtHdc010t_expir_dt_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub
Sub txtHdc010t_revoke_dt_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub
Sub txtHdc010t_tax_rate_Change()
    If lgBlnDataChgFlg = True Then
	    lgBlnFlgChgValue = True
	End If
End Sub

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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>저축사항등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
			    	        	    		           <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
			            	    <TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPay_grd" SIZE=20  MAXLENGTH=50 ALT="급호" TYPE="Text"  tag="14XXXU"></TD>
			               	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDept_nm" SIZE=20  MAXLENGTH=40 ALT="부서명" TYPE="Text"  tag="14XXXU"></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtRoll_pstn" SIZE=20  MAXLENGTH=50  ALT="직위" TYPE="Text"  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
			            		<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h5201ma1_fpDateTime2_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP></TD>
			            		<TD CLASS="TD6" NOWRAP></TD>
			            	</TR>
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h5201ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
							<TD HEIGHT=* WIDTH=100%>
                            		<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>저축코드</TD>
	                        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtHdc010t_save_cd" Size="20" MAXLENGTH="50" ALT="저축코드" Tag="24"></TD>
	                        					<TD CLASS="TD5" NOWRAP>월불입액</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_OBJECT_txtHdc010t_script_amt.js'></script>
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>신규가입일</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_fpDateTime12_txtHdc010t_new_dt.js'></script>
	                        					</TD>
	                        			    </TR>
	                        			    <TR>
	                        					<TD CLASS="TD5" NOWRAP>저축구분</TD>
	                        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtHdc010t_save_type" Size="20" MAXLENGTH="50" ALT="저축구분" Tag="24"></TD>
	                        					<TD CLASS="TD5" NOWRAP>총불입횟수</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_OBJECT_txtHdc010t_tot_script_cnt.js'></script>
	                        					</TD>
	                         					<TD CLASS="TD5" NOWRAP>만기일</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_fpDateTime12_txtHdc010t_expir_dt.js'></script>
	                        					</TD>
	                        				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>계좌번호</TD>
	                        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtHdc010t_bank_accnt" Size="20" MAXLENGTH="20" ALT="계좌번호" Tag="24"></TD>
	                        					<TD CLASS="TD5" NOWRAP>만기금액</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_OBJECT_txtHdc010t_expir_amt.js'></script>
	                        					</TD>
	                         					<TD CLASS="TD5" NOWRAP>해약일</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_fpDateTime12_txtHdc010t_revoke_dt.js'></script>
	                        					</TD>
	                        				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>은행코드</TD>
	                        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtHdc010t_bank_cd" Size="20" MAXLENGTH="30" Tag="24"></TD>
	                         					<TD CLASS="TD5" NOWRAP></TD>
	                        					<TD CLASS="TD6" NOWRAP></TD>
	                        					<TD CLASS="TD5" NOWRAP>세액공제율</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <script language =javascript src='./js/h5201ma1_OBJECT_txtHdc010t_tax_rate.js'></script>
	                        					</TD>
                             				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>은행명</TD>
	                        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtFaa090t_bank_name" Size="20" MAXLENGTH="30" Tag="24"></TD>
	                         					<TD CLASS="TD5" NOWRAP></TD>
	                        					<TD CLASS="TD6" NOWRAP></TD>
	                         					<TD CLASS="TD5" NOWRAP></TD>
	                        					<TD CLASS="TD6" NOWRAP></TD>
	                        				</TR>
                            		</TABLE>
                            	</TD>
							<TR>

						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

