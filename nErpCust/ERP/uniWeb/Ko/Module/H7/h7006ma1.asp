<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 상여조회및조정 
*  3. Program ID           : H1a02ma1
*  4. Program Name         : H1a02ma1
*  5. Program Desc         : 상여관리/상여조회및조정 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/26
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : YBI
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H7006mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID  = "h6012ma1"
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row
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
Dim lgSpreadFlg
Dim topleftOK
Dim lgStrPrevKey1

Dim C_SUB_CD
Dim C_SUB_CD_POP
Dim C_SUB_CD_NM
Dim C_SUB_AMT

Dim C_ALLOW_CD_NM
Dim C_ALLOW_AMT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_SUB_CD = 1                                                  'Column constant for Spread Sheet 
        C_SUB_CD_POP = 2
        C_SUB_CD_NM = 3
        C_SUB_AMT = 4    
    ElseIf pvSpdNo = "B" Then
        C_ALLOW_CD_NM = 1
        C_ALLOW_AMT = 2
    End If
    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKey1 = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgSpreadFlg		  = 1
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetSvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtbonus_yymm_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtbonus_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(Byval Kubun)  
    On Error Resume Next
    Const CookieSplit = 4877	
    
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then                                       '☜: h6012ma1.asp 의 쿠기값을 받고 있음.
		strTemp = ReadCookie("EMP_NO")                      '           Kubun = 0 일때 수정금지 요망........!    
		If strTemp = "" then Exit Function
        
        frm1.txtbonus_yymm_dt.text = ReadCookie("PAY_YYMM_DT")
		frm1.txtEmp_no.value = strTemp
    	frm1.txtBonus_type.value = ReadCookie("PROV_TYPE_HIDDEN")
		FncQuery()              
		WriteCookie "PAY_YYMM_DT" , ""
	    WriteCookie "EMP_NO"      , ""
        WriteCookie "PROV_TYPE_HIDDEN"   , ""
        
	ElseIf Kubun = 1 Then 
        WriteCookie "PAY_YYMM_DT" , frm1.txtbonus_yymm_dt.text
	    
	End IF
	
End Function

FUNCTION PgmJumpCheck()         
    PgmJump(BIZ_PGM_JUMP_ID)
End Function
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream  = Frm1.txtEmp_no.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream  = lgKeyStream & Frm1.txtbonus_yymm_dt.Year & Right( "0" & Frm1.txtbonus_yymm_dt.Month , 2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtBonus_type.Value & Parent.gColSep
    lgKeyStream  = lgKeyStream & lgUsrIntCd & Parent.gColSep     
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0040", "''", "S") & " and ((minor_cd >= " & FilterVar("2", "''", "S") & " and minor_cd <= " & FilterVar("9", "''", "S") & ") or minor_cd in (" & FilterVar("C", "''", "S") & " ," & FilterVar("Q", "''", "S") & " )) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtBonus_type, iCodeArr, iNameArr, Chr(11))

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
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

	    With frm1.vspdData
            ggoSpread.Source = frm1.vspdData

            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_SUB_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
           ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk

           ggoSpread.SSSetEdit    C_SUB_CD,      "공제코드",10,,,3,2
           ggoSpread.SSSetButton  C_SUB_CD_POP
           ggoSpread.SSSetEdit    C_SUB_CD_NM,   "공제코드명",31
           ggoSpread.SSSetFloat   C_SUB_AMT    , "공제금액",20, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

           Call ggoSpread.MakePairsColumn(C_SUB_CD,C_SUB_CD_POP)    'sbk

	       .ReDraw = true
	
           Call SetSpreadLock("A")
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With frm1.vspdData1

            ggoSpread.Source = frm1.vspdData1
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols   = C_ALLOW_AMT + 1                                                 ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
           ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("B") 'sbk
	
           ggoSpread.SSSetEdit    C_ALLOW_CD_NM  , "상여기준 수당코드명",  22,,,15
           ggoSpread.SSSetFloat   C_ALLOW_AMT    , "상여기준 수당액",    20, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	       .ReDraw = true

           Call SetSpreadLock("B")
    
        End With    
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
        If pvSpdNo = "A" Then
            ggoSpread.Source = .vspdData
            .vspdData.ReDraw = False
            ggoSpread.SpreadLock      C_SUB_CD, -1, C_SUB_CD, -1
            ggoSpread.SpreadLock      C_SUB_CD_POP, -1, C_SUB_CD_POP, -1
            ggoSpread.SpreadLock      C_SUB_CD_NM, -1, C_SUB_CD_NM, -1
            ggoSpread.SSSetRequired   C_SUB_AMT, -1, -1
            ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
            .vspdData.ReDraw = True
        End If

        If pvSpdNo = "B" Then
            ggoSpread.Source = .vspdData1
            .vspdData1.ReDraw = False
            ggoSpread.SpreadLock      C_ALLOW_CD_NM, -1, C_ALLOW_CD_NM, -1
            ggoSpread.SpreadLock      C_ALLOW_AMT, -1, C_ALLOW_AMT, -1
            ggoSpread.SSSetProtected   .vspdData1.MaxCols   , -1, -1
            .vspdData1.ReDraw = True
        End If
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	With frm1
		If lgSpreadFlg = 1 Then
            ggoSpread.Source = frm1.vspdData
            .vspdData.ReDraw = False
            ggoSpread.SSSetRequired  C_SUB_CD, pvStartRow, pvEndRow
            ggoSpread.SSSetProtected C_SUB_CD_NM, pvStartRow, pvEndRow
            ggoSpread.SSSetRequired  C_SUB_AMT, pvStartRow, pvEndRow
            .vspdData.ReDraw = True
		ElseIf lgSpreadFlg = 2 Then
            ggoSpread.Source = frm1.vspdData1
            .vspdData1.ReDraw = False
            ggoSpread.SSSetRequired  C_ALLOW_CD_NM, pvStartRow, pvEndRow
            ggoSpread.SSSetRequired  C_ALLOW_AMT, pvStartRow, pvEndRow
            .vspdData1.ReDraw = True
		End If
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
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

            C_SUB_CD = iCurColumnPos(1)
            C_SUB_CD_POP = iCurColumnPos(2)
            C_SUB_CD_NM = iCurColumnPos(3)
            C_SUB_AMT = iCurColumnPos(4)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_ALLOW_CD_NM = iCurColumnPos(1)
            C_ALLOW_AMT = iCurColumnPos(2)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call AppendNumberPlace("7", "3", "2")
    Call AppendNumberPlace("8", "4", "2")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call ggoOper.FormatDate(frm1.txtbonus_yymm_dt, Parent.gDateFormat, 2)

    Call InitSpreadSheet("")                                                        'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar

    Frm1.txtName.focus 
   
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
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    Call MakeKeyStream("X")
	topleftOK = false    
	frm1.txtPrevNext.value = ""	
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
       
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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
	Call DisableToolBar(Parent.TBC_DELETE)
    If DbDelete = False Then
        Call RestoreToolBar()
        Exit Function
    End If
    
    
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim strSQL
    Dim strReturn_value
    Dim strTran_flag
    Dim dblSub_tot_amt
    Dim lRow
    Dim dblBonus_amt
    Dim close_Dt, input_Dt
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    close_Dt = UniConvYYYYMMDDToDate(Parent.gDateFormat, frm1.txtbonus_yymm_dt.Year, frm1.txtbonus_yymm_dt.Month, "01")
    close_Dt = UniConvDateToYYYYMM(close_Dt, Parent.gDateFormat, Parent.gServerDateType)

    strReturn_value = "Y"
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & "  AND pay_gubun = " & FilterVar("Z", "''", "S") & "  AND PAY_TYPE =  " & FilterVar(frm1.txtbonus_type.value , "''", "S") & ""
    IntRetCD = CommonQueryRs(" count(*) "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  Replace(lgF0, Chr(11), "") > 0 then
        IntRetCD = CommonQueryRs(" close_type, Convert(char(10),close_dt,20) "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If  IntRetCd = false then
            strReturn_value = "Y"
        else
            input_Dt = UniConvDateToYYYYMM(Replace(lgF1, Chr(11), ""), Parent.gServerDateFormat, Parent.gServerDateType)

            Select Case Replace(lgF0, Chr(11), "")        
               Case "1"    '마감형태 : 정상 
                  if  input_Dt <= close_Dt then
                     strReturn_value = "Y"
                  else
                     strReturn_value = "N"
                  end if
               Case "2"    '마감형태 : 마감 
                  if  input_Dt < close_Dt then
                     strReturn_value = "Y"
                  else
                     strReturn_value = "N"
                  end if
            end Select
        end if
    end if

    if  strReturn_value = "N" then
        Call DisplayMsgBox("800313","X","X","X")
        exit function
    end if

'   자동기표 처리 
    strTran_flag = "N"
    strSQL = " pay_yymm= " & FilterVar(replace(close_Dt, Parent.gServerDateType, ""), "''", "S") & ""   ' 정산년월 
    strSQL = strSQL & " AND emp_no =  " & FilterVar(frm1.txtemp_no.value , "''", "S") & ""
    strSQL = strSQL & " AND prov_type =  " & FilterVar(frm1.txtbonus_type.value , "''", "S") & ""
    IntRetCD = CommonQueryRs(" DISTINCT(tran_flag) "," hdf070t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
        strTran_flag = Replace(lgF0, Chr(11), "")
    end if

    if  strTran_flag = "Y" then
        Call DisplayMsgBox("800408","X","X","X")'이미 자동기표가 처리되었습니다. 회계자동기표처리를 취소한 후 작업하시기 바랍니다.
        exit function
    end if

'*** 공제총액 합계 
    dblSub_tot_amt = 0
	With Frm1
        ggoSpread.Source = frm1.vspdData
        For lRow = 1 To .vspdData.MaxRows
            
            .vspdData.Row = lRow
            .vspdData.Col = 0
            if  .vspdData.Text = ggoSpread.DeleteFlag then

				.vspdData.Col = C_SUB_AMT
                .vspdData.Col = C_SUB_CD
                select case .vspdData.text
                    case "S97"
                        .txtSave_fund.value = 0
                    case "S98"
                        .txtIncome_tax.value = 0
                    case "S99"
                        .txtRes_tax.value = 0
                    case "S01"
                        .txtMed_insur.value = 0
                    case "S02"
                        .txtAnut.value = 0
                    case "S03"
                        .txtEmp_insur.value = 0
                end select
            else
				if   .vspdData.Text = ggoSpread.InsertFlag OR .vspdData.Text = ggoSpread.UpdateFlag then
                    .vspdData.Col = C_SUB_CD_NM
                    If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                        Call DisplayMsgBox("970000", "x","공제코드","x")
                        .vspdData.Col = C_SUB_CD
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                End if
				.vspdData.Col = C_SUB_AMT
                dblSub_tot_amt = dblSub_tot_amt + UNICDbl(.vspdData.text)
				
				.vspdData.Col = C_SUB_CD
                select case .vspdData.text
                    case "S97"
                        .vspdData.Col = C_SUB_AMT
                        .txtSave_fund.value = UNICDbl(.vspdData.text)
                    case "S98"
                        .vspdData.Col = C_SUB_AMT
                        .txtIncome_tax.value = UNICDbl(.vspdData.text)
                    case "S99"
                        .vspdData.Col = C_SUB_AMT
                        .txtRes_tax.value = UNICDbl(.vspdData.text)
                    case "S01"
                        .vspdData.Col = C_SUB_AMT
                        .txtMed_insur.value = UNICDbl(.vspdData.text)
                    case "S02"
                        .vspdData.Col = C_SUB_AMT
                        .txtAnut.value = UNICDbl(.vspdData.text)
                    case "S03"
                        .vspdData.Col = C_SUB_AMT
                        .txtEmp_insur.value = UNICDbl(.vspdData.text)
                end select
            end if
            
        next
     
        .txtSub_tot_amt.text   = UNIFormatNumber(dblSub_tot_amt, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy,ggAmtOfMoney.Rndunit)
  	    .txtReal_prov_amt.text = UNIFormatNumber(UNICDbl(.txtProv_tot_amt.text) - UNICDbl(.txtSub_tot_amt.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy,ggAmtOfMoney.Rndunit)
  	    
  end with
	
'UNIFormatNumber(dblBonus_rate,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)


'    frm1.txtProv_rate.value = CDbl(frm1.txtbonus_rate.value) + CDbl(frm1.txtadd_rate.value) - CDbl(frm1.txtminus1_rate.value) - CDbl(frm1.txtminus2_rate.value)

'   상여금(지급율 * 상여기준금액)
'    dblBonus_amt = frm1.txtProv_rate.value / 100 * frm1.txtbonus_bas.value
'   끝전처리(f_round(dblBonus_amt))
'    frm1.txtBonus.value = f_round("000", dblBonus_amt)'dblBonus_amt

'   생상장려금(생산장려율 * 상여기준금액)
'    dblBonus_amt = frm1.txtSplendor_rate.value / 100 * frm1.txtbonus_bas.value
'   끝전처리(f_round(dblBonus_amt))
'    frm1.txtSplendor_amt.value = f_round("000", dblBonus_amt)'dblBonus_amt

'    frm1.txtProv_tot_amt.value = CDbl(frm1.txtBonus.value) + CDbl(frm1.txtSplendor_amt.value)

'    frm1.txtReal_prov_amt.value = CDbl(frm1.txtProv_tot_amt.value) - CDbl(frm1.txtSub_tot_amt.value)

    Call MakeKeyStream("X")
	Call DisableToolBar(Parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : F_round
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function F_round(allow_cd, allow_in)

    Dim IntRetCD
    Dim ls_proc_bas
    Dim ldc_bas_amt
    Dim allow_out 
    Dim strNum

    ' 입력값이 0이면 0을 return 
    IF allow_in = 0 THEN
    	F_round = 0 
    	exit function
    END IF

    IntRetCD = CommonQueryRs(" bas_amt, proc_bas "," hda040t ", "allow_cd= " & FilterVar(allow_cd, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
        ldc_bas_amt = Replace(lgF0, cHR(11), "")
        ls_proc_bas = Replace(lgF1, cHR(11), "")
    else
        allow_out = allow_in
    end if

    select case ls_proc_bas
        case "1"    ' 절사 
            allow_out = fix(allow_in / ldc_bas_amt) * ldc_bas_amt
        case "2"    ' 절상 
            strNum = (allow_in / ldc_bas_amt)
            if  instr(1, Cstr(strNum), Parent.gComNumDec) = 0 then
                allow_out = fix(strNum) * ldc_bas_amt
            else
                allow_out = (fix(strNum) + 1)  * ldc_bas_amt
            end if
        case ELSE   ' 사사오입 
            allow_out = Round(allow_in / ldc_bas_amt, 0)  * ldc_bas_amt
    end select

    F_round = allow_out 

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

	If lgSpreadFlg = 1 Then
        If Frm1.vspdData.MaxRows < 1 Then
           Exit Function
        End If
    
        With frm1.vspdData
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData	
				ggoSpread.CopyRow
                SetSpreadColor .ActiveRow, .ActiveRow
    
				.ReDraw = True
    		    .Focus
			End If
		End With
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 

	If lgSpreadFlg = 1 Then
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.EditUndo  
        call initdata()
    End If
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
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
    	If lgSpreadFlg = 1 Then
             .vspdData.ReDraw = False
             .vspdData.focus()
             ggoSpread.Source = .vspdData
             ggoSpread.InsertRow .vspdData.ActiveRow, imRow
             SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
             .vspdData.ReDraw = True
		End If
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
    
    If lgSpreadFlg = 1 Then
        If Frm1.vspdData.MaxRows < 1 then
           Exit function
	    End if	
    
        With Frm1.vspdData 
        	.focus()
        	ggoSpread.Source = frm1.vspdData
        	lDelRows = ggoSpread.DeleteRow
        End With
	End if	
    
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
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

    Call MakeKeyStream("X")
	topleftOK = false
	frm1.txtPrevNext.value = "P"
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
	
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
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

    Call MakeKeyStream("X")
	topleftOK = false
	frm1.txtPrevNext.value = "N"
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
	
    FncNext = True                                                               '☜: Processing is OK
	
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, True)
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

    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")      
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
	End Select 
    
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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

	If LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgSpreadFlg="       & lgSpreadFlg
    strVal = strVal     & "&topleftOK="       & topleftOK                   '☜: Query Key
	if lgSpreadFlg = "1" then
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
	else
		strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
	end if
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & frm1.txtPrevNext.value
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
		
	If LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	With frm1
		.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
        ggoSpread.Source = frm1.vspdData
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData.Col = C_SUB_CD  	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SUB_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData.Col = C_SUB_CD  	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SUB_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                    .vspdData.Col = C_SUB_CD   : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
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
		
	If LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
	Call SetToolbar("1100111111101111")												'⊙: Set ToolBar
    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
    lgSpreadFlg = 1
    lgBlnFlgChgValue = False
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
    ggoSpread.Source = Frm1.vspdData1
    Frm1.vspdData1.MaxRows = 0
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

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp
'=======================================================================================================
Function OpenCode(iWhere, strCode, Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case 1
	        arrParam(0) = "공제코드 팝업"			        <%' 팝업 명칭 %>
	    	arrParam(1) = "HDA010T"							    <%' TABLE 명칭 %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "pay_cd = " & FilterVar("*", "''", "S") & "  AND code_type = " & FilterVar("2", "''", "S") & " "			    	<%' Where Condition%>
	    	arrParam(5) = "공제코드" 			        <%' TextBox 명칭 %>
	
	    	arrField(0) = "allow_cd"							<%' Field명(0)%>
	    	arrField(1) = "allow_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "공제코드"	   		    	<%' Header명(0)%>
	    	arrHeader(1) = "공제코드명"	    		     		<%' Header명(1)%>
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
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
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 1
		    	.vspdData.Col = C_SUB_CD_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_SUB_CD
		    	.vspdData.text = arrRet(0) 
				.vspdData.action = 0		    	
        End Select

	End With

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere, strData)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case 0
	        arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	        arrParam(1) = ""			' Name Cindition
	    Case 1
	        arrParam(0) = frm1.txtEmp_no.value
	        arrParam(1) = ""
	    Case 2
            arrParam(0) = strData
            arrParam(1) = ""
	End Select
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus	
		Exit Function
	Else
		Call SetEmpName(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet, iWhere)
	With frm1
	    if iWhere = 2 then
            ggoSpread.Source = Frm1.vspdData
	        With .VspdData
    		    .Row = .ActiveRow
                .Col = C_VALUE_EMP_NO
                .Text = arrRet(0)
                .Col = C_VALUE_NAME
                .Text = arrRet(1)
	        End With
	    else
		    .txtEmp_no.value = arrRet(0)
		    .txtName.value = arrRet(1)
		    Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		    .txtEmp_no.focus
		    lgBlnFlgChgValue = False
        end if
	End With
End Sub

'========================================================================================================

Sub txtbonus_bas_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtbonus_bas_amt_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtbonus_rate_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtadd_rate_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtminus1_rate_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtminus2_rate_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtsplendor_rate_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtsub_tot_amt_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtReal_prov_amt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			    Case C_SUB_CD_POP
			    	.Col = C_SUB_CD
			    	.Row = Row
			    	Call OpenCode(1, .value, Row)
			End Select
		End If
    
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
    Dim iDx

	Frm1.vspdData.Row = Row
    Select Case Col
        Case C_SUB_CD_NM
            Frm1.vspdData.Row = Row
            Frm1.vspdData.Col = C_SUB_CD_NM
            iDx = Frm1.vspdData.Text
            Frm1.vspdData.Row = Row
   	        Frm1.vspdData.Col = C_SUB_CD
            Frm1.vspdData.Text = iDx
    End Select    
  	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_Change(ByVal Col, ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_SUB_CD
            Frm1.vspdData.Col = C_SUB_CD
            iDx = CommonQueryRs(" allow_nm "," HDA010T ", " allow_cd= " & FilterVar(Frm1.vspdData.text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            if  iDx = true then
			    Frm1.vspdData.col = C_SUB_CD_NM
			    Frm1.vspdData.Text = Replace(lgF0, Chr(11), "")
            else
			    Frm1.vspdData.col = C_SUB_CD_NM
			    Frm1.vspdData.Text = ""
                Call DisplayMsgBox("970000", "x","공제코드","x")
                Frm1.vspdData.col = C_SUB_CD
                Frm1.vspdData.Action = 0 ' go to 
                vspdData_Change = true
            end if
    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

	lgSpreadFlg = 1

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP2C"   

    Set gActiveSpdSheet = frm1.vspdData1

	lgSpreadFlg = 2
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
    
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If

End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
     End If

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
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

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	topleftOK = true	
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then
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
' Name : txtbonus_yymm_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtbonus_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtbonus_yymm_dt.Action = 7 
        frm1.txtbonus_yymm_dt.focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtBonus_yymm_dt_KeyDown()
'   Event Desc : txtbonus_yymm_dt
'==========================================================================================
Sub txtBonus_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_OnChange()

    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_OnChange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여조회및조정</font></td>
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
			    	    	    	<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    	    	<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName 1, ''"></TD>
			    	            	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    	    	<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14></TD>
			            	    </TR>
			            	    <TR>
			            	    	<TD CLASS="TD5" NOWRAP>정산년월</TD>
			            	    	<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtbonus_yymm_dt name=txtbonus_yymm_dt CLASS=FPDTYYYYMM title=FPDATETIME tag="12X1" ALT="정산년월"></OBJECT>');</SCRIPT></TD>
			            	    	<TD CLASS="TD5" NOWRAP>지급구분</TD>
			            	    	<TD CLASS="TD6"><SELECT NAME="txtBonus_type" CLASS ="cbonormal" tag="12" ALT="지급구분"></SELECT></TD>
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
						        <TD COLSPAN=4>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>관련기본사항</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>부서코드</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24></TD>
							            	<TD CLASS=TD5 NOWRAP>직종코드</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOcpt_type" ALT="직종" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>급호봉</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd1" TYPE=TEXT SIZE="13" TAG=24XXXU ALT="급호">
							            	                     <INPUT NAME="txtPay_grd2" TYPE=TEXT SIZE="5" TAG=24XXXU ALT="호봉">호봉</TD>
							            	<TD CLASS=TD5 NOWRAP>급여구분</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_cd" ALT="급여구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24></TD>
							            </TR>	
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>세액구분/입퇴사구분</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTax_cd" ALT="세액구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24>&nbsp;/
							            	<INPUT NAME="txtExcept_type" ALT="입퇴사구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24></TD>
							            	<TD CLASS=TD5 NOWRAP>근속개월</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtDuty_mm name=txtDuty_mm CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X6Z" ALT="근속개월"></OBJECT>');</SCRIPT>
							            	</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>지급일</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProv_dt NAME="txtProv_dt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="설정일"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>배우자/부양자</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpouse" ALT="배우자" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24>&nbsp;/
							            	 <INPUT NAME="txtSupp_cnt" ALT="부양자" TYPE="Text" MAXLENGTH=20 SiZE=20 tag=24></TD>
							            </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
						    <TR>
						        <TD COLSPAN=4>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>상여율</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>상여기준금</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtBonus_bas name=txtBonus_bas CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="상여기준금"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>상여기본지급율</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtBonus_rate name=txtBonus_rate CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X7Z" ALT="상여지급율"></OBJECT>');</SCRIPT>&nbsp;%</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>가산율</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtAdd_rate name=txtAdd_rate CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X7Z" ALT="가산율"></OBJECT>');</SCRIPT>&nbsp;%</TD>
							            	<TD CLASS=TD5 NOWRAP>-----------------</TD>
						            	    <TD CLASS=TD6 NOWRAP>--------------------</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>일할계산차감율</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMinus2_rate name=txtMinus2_rate CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X7Z" ALT="일할계산차감율"></OBJECT>');</SCRIPT>&nbsp;%</TD>
							            	<TD CLASS=TD5 NOWRAP>총상여지급율</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtProv_rate name=txtProv_rate CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X7Z" ALT="상여지급율"></OBJECT>');</SCRIPT>&nbsp;%</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>근태차감</TD>
							            	<TD CLASS=TD6 NOWRAP>
							            			<table><tr>	<td>차감율</td>
							            						<td><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMinus1_rate name=txtMinus1_rate CLASS=FPDS65 title=FPDOUBLESINGLE tag="24X7Z" ALT="근태차감율"></OBJECT>');</SCRIPT></td>
							            						<td>&nbsp;%&nbsp;&nbsp;차감금액</td>
							            						<td><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMinus_amt name=txtMinus_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="근태차감금액"></OBJECT>');</SCRIPT>
							            						</td></tr></table>
							            	</TD>
							            	<TD CLASS=TD5 NOWRAP>생산장려율</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSplendor_rate name=txtSplendor_rate CLASS=FPDS90 title=FPDOUBLESINGLE tag="24X8Z" ALT="생산장려율"></OBJECT>');</SCRIPT>&nbsp;%</TD>

							            </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
						    <TR>
						        <TD COLSPAN=4>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>지급액</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>상여금</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtBonus name=txtBonus CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="상여금"></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5 NOWRAP>생산장려금</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSplendor_amt name=txtSplendor_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="생산장려금"></OBJECT>');</SCRIPT>&nbsp;</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>상여총액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtProv_tot_amt name=txtProv_tot_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="상여총액"></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5 NOWRAP>공제총액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_tot_amt name=txtSub_tot_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="공제총액"></OBJECT>');</SCRIPT>&nbsp;</TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>실지급액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtReal_prov_amt name=txtReal_prov_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="실지급액"></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5 NOWRAP></TD>
							            	<TD CLASS=TD6 NOWRAP></TD>
							            </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="58%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT="100%" WIDTH=* COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:CookiePage 1">급여조회</a>
				</TD>
				<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
	    </TD>   
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtsave_fund"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtincome_tax"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtres_tax"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtmed_insur"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtanut"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtemp_insur"   TAG="24">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

