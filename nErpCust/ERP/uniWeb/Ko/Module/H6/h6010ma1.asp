<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 급여조회및조정 
*  3. Program ID           : H6010ma1
*  4. Program Name         : H6010ma1
*  5. Program Desc         : 급여조회및조정 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/29
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Hwang Jeong-won
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
Const BIZ_PGM_ID = "H6010mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID  = "h6012ma1"                                '☜:cookie page 연결되있음..!(6월9일삭제금지..TGS)
Const C_SHEETMAXROWS    = 7	                                          '☜: Visble row

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
Dim gSpreadFlg
Dim topleftOK
Dim lgStrPrevKey1
Dim lsInternal_cd
Dim lgSpreadChange
Dim lgSpreadChange1

Dim C_SUB_CD
Dim C_SUB_CD_POP
Dim C_SUB_CD_NM
Dim C_SUB_AMT

Dim C_ALLOW_CD
Dim C_ALLOW_CD_POP
Dim C_ALLOW_CD_NM
Dim C_TAX_TYPE
Dim C_ALLOW_AMT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_ALLOW_CD = 1
        C_ALLOW_CD_POP = 2
        C_ALLOW_CD_NM = 3
        C_TAX_TYPE = 4
        C_ALLOW_AMT = 5
    End If

    If pvSpdNo = "B" Then
        C_SUB_CD = 1
        C_SUB_CD_POP = 2
        C_SUB_CD_NM = 3
        C_SUB_AMT = 4
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
    lgStrPrevKey1      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gSpreadFlg		  = 1
	lsInternal_cd     = ""
	
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
	lgSpreadChange    = False
	lgSpreadChange1   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
    Frm1.txtPayYymm.focus()

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtPayYymm.Year = strYear 		 '년월일 default value setting
	frm1.txtPayYymm.Month = strMonth 
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
Function CookiePage(Byval Kubun)                            '☜: 6/9   지우지 말아주세요....TGS
    On Error Resume Next
    Const CookieSplit = 4877	
    
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then                                       '☜: h6012ma1.asp 의 쿠기값을 받고 있음.
		strTemp = ReadCookie("EMP_NO")                      '         절대수정금지 요망........!    
		If strTemp = "" then Exit Function
		
        frm1.txtPayYymm.text = ReadCookie("PAY_YYMM_DT")
		frm1.txtEmpNo.value = strTemp
    	frm1.txtType.value = ReadCookie("PROV_TYPE_HIDDEN")
		MainQuery()              
		WriteCookie "PAY_YYMM_DT" , ""
	    WriteCookie "EMP_NO"      , ""
        WriteCookie "PROV_TYPE_HIDDEN"   , ""
        
	ElseIf Kubun = 1 Then 
        WriteCookie "PAY_YYMM_DT" , frm1.txtPayYymm.text
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
	
    lgKeyStream  = Frm1.txtEmpNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream  = lgKeyStream & Frm1.txtPayYymm.Year & Right("0" & Frm1.txtPayYymm.Month,2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtType.Value & Parent.gColSep
    lgKeyStream  = lgKeyStream & lgUsrIntCd & Parent.gColSep
   
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    	
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0040' AND (MINOR_CD <= '1' OR MINOR_CD >= 'A' ) AND MINOR_CD NOT IN ('$','Q' ,'C' ,'B','Z') ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtType, iCodeArr, iNameArr, Chr(11))

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
            ggoSpread.Source = Frm1.vspdData
            ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_ALLOW_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                             ' ☜:☜: Hide maxcols
           .ColHidden = True
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk
           
           ggoSpread.SSSetEdit    C_ALLOW_CD     , "수당코드", 10,,,3,2
           ggoSpread.SSSetButton  C_ALLOW_CD_POP
           ggoSpread.SSSetEdit    C_ALLOW_CD_NM  , "수당",  12,,,15
	       ggoSpread.SSSetEdit    C_TAX_TYPE     , "Tax_type",  2
           ggoSpread.SSSetFloat   C_ALLOW_AMT    , "수당액",    18, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

           Call ggoSpread.MakePairsColumn(C_ALLOW_CD,C_ALLOW_CD_POP)    'sbk

           Call ggoSpread.SSSetColHidden(C_TAX_TYPE,C_TAX_TYPE,True)

	       .ReDraw = true
	
           Call SetSpreadLock("A")
   
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With frm1.vspdData1
            ggoSpread.Source = Frm1.vspdData1

            ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_SUB_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("B") 'sbk

           ggoSpread.SSSetEdit    C_SUB_CD     , "공제코드", 10,,,3,2       
           ggoSpread.SSSetButton  C_SUB_CD_POP
	       ggoSpread.SSSetEdit    C_SUB_CD_NM  , "공제",  12,,,15
           ggoSpread.SSSetFloat   C_SUB_AMT    , "공제금액",    18, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

           Call ggoSpread.MakePairsColumn(C_SUB_CD,C_SUB_CD_POP)    'sbk

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
            ggoSpread.Source = frm1.vspdData
            .vspdData.ReDraw = False
            ggoSpread.SpreadLock      C_ALLOW_CD, -1, C_ALLOW_CD, -1
            ggoSpread.SpreadLock      C_ALLOW_CD_POP, -1, C_ALLOW_CD_POP, -1
            ggoSpread.SpreadLock      C_ALLOW_CD_NM, -1, C_ALLOW_CD_NM, -1
    		ggoSpread.SSSetRequired   C_ALLOW_AMT, -1, -1
    		ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
            .vspdData.ReDraw = True
        End If

        If pvSpdNo = "B" Then
            ggoSpread.Source = frm1.vspdData1
            .vspdData1.ReDraw = False
            ggoSpread.SpreadLock      C_SUB_CD, -1, C_SUB_CD, -1
            ggoSpread.SpreadLock      C_SUB_CD_POP, -1, C_SUB_CD_POP, -1
            ggoSpread.SpreadLock      C_SUB_CD_NM, -1, C_SUB_CD_NM, -1
    		ggoSpread.SSSetRequired   C_SUB_AMT, -1, -1
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
		If gSpreadFlg = 1 Then
			ggoSpread.Source = frm1.vspdData
			.vspdData.ReDraw = False
			ggoSpread.SSSetRequired   C_ALLOW_CD, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected  C_ALLOW_CD_NM, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired   C_ALLOW_AMT, pvStartRow, pvEndRow
			.vspdData.ReDraw = True
		Else
			ggoSpread.Source = frm1.vspdData1
			.vspdData1.ReDraw = False                
			ggoSpread.SSSetRequired   C_SUB_CD, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected  C_SUB_CD_NM, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired   C_SUB_AMT, pvStartRow, pvEndRow
			.vspdData1.ReDraw = True
		End If
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

            C_ALLOW_CD = iCurColumnPos(1)
            C_ALLOW_CD_POP = iCurColumnPos(2)
            C_ALLOW_CD_NM = iCurColumnPos(3)
            C_TAX_TYPE = iCurColumnPos(4)
            C_ALLOW_AMT = iCurColumnPos(5)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_SUB_CD = iCurColumnPos(1)
            C_SUB_CD_POP = iCurColumnPos(2)
            C_SUB_CD_NM = iCurColumnPos(3)
            C_SUB_AMT = iCurColumnPos(4)
            
    End Select    
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	
	Call ggoOper.LockField(Document, "N")		
	
	Call AppendNumberPlace("6", "1", "0")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtPayYymm, Parent.gDateFormat, 2)

    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100100111011111")												'⊙: Set ToolBar
   
    Call InitComboBox
	Call CookiePage(0)                                                             '☜: Check Cookie
    
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
    Dim iDx
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	Else
		ggoSpread.Source = Frm1.vspdData1
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    End If    
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmpNo_Onchange()  then
       Exit Function
    End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
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
	Call SetToolbar("1100111111011111")							                 '⊙: Set ToolBar
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
    Dim dblSub_tot_amt, dblPay_tot_amt, dblNon_tax1, dblNon_tax2
    Dim dblNon_tax3, dblNon_tax4, dblNon_tax_amt
    Dim lRow
    Dim dblBonus_amt
    Dim lFlag
    Dim close_Dt, input_Dt
    Dim strTax_cd  
        
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
        
    If lgBlnFlgChgValue = False Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			ggoSpread.Source = frm1.vspdData1
			If ggoSpread.SSCheckChange = False Then
				IntRetCD = DisplayMsgBox("900001","X","X","X")                   '⊙: No data changed!!
				Exit Function
			End If
		End If        
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	With Frm1
		ggoSpread.Source = frm1.vspdData
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .vspdData.Text           
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_ALLOW_CD_NM
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
						Call DisplayMsgBox("800145","X","X","X")
						.vspdData.Col = C_ALLOW_CD
						.vspdData.Action = 0
						Set gActiveElement = document.activeElement
						Exit Function
					End if
			End Select
		next
 
 		ggoSpread.Source = frm1.vspdData1
		For lRow = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRow
			.vspdData1.Col = 0

			Select Case .vspdData1.Text           
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData1.Col = C_SUB_CD_NM
					If IsNull(Trim(.vspdData1.Text)) OR Trim(.vspdData1.Text) = "" Then
						Call DisplayMsgBox("800176","X","X","X")
						.vspdData1.Col = C_SUB_CD
						.vspdData1.Action = 0
						Set gActiveElement = document.activeElement
						Exit Function
					End if
			End Select
		next
    End with    

    IntRetCD = CommonQueryRs(" emp_no "," hdf020t ", " prov_type = 'N' and emp_no=" & FilterVar(frm1.txtEmpNo.value , "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
        Call DisplayMsgBox("970029","X","임금지급대상자인지","X")
        Exit Function
    End if
    
    lFlag = DisplayMsgBox("800439",Parent.VB_YES_NO,"X","X")
    If lFlag = 6 Then
		frm1.txtTaxFlag.value = "Y"
	Else
		frm1.txtTaxFlag.value = "N"
	End If		    

	Dim strType
	strType = frm1.txtType.value
	if strType = "1" Then
		strType = "!"
	Elseif strType = "B" then
		strType = "xxxx"
	Elseif strType = "P" then
		strType = "xxxx"
	Elseif strType = "Z" then
		strType = "@"
	End if
	
    close_Dt = UniConvYYYYMMDDToDate(Parent.gDateFormat, frm1.txtPayYymm.Year, frm1.txtPayYymm.Month, "01")
    close_Dt = UniConvDateToYYYYMM(close_Dt, Parent.gDateFormat, Parent.gServerDateType)

    strReturn_value = "Y"
    strSQL = " org_cd = '1'  AND pay_gubun ='Z'  AND PAY_TYPE =  " & FilterVar(strType , "''", "S")
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
        Call DisplayMsgBox("800294","X","X","X")
        exit function
    end if

'   자동기표 처리 
    strTran_flag = "N"

    strSQL = " pay_yymm= " & FilterVar(replace(close_Dt, Parent.gServerDateType, ""), "''", "S")
    strSQL = strSQL & " AND emp_no =  " & FilterVar(frm1.txtEmpNo.value , "''", "S")
    strSQL = strSQL & " AND prov_type =  " & FilterVar(frm1.txtType.value , "''", "S")
    
    IntRetCD = CommonQueryRs(" IsNull(tran_flag,'N'), IsNull(tax_cd,'1') "," hdf070t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If  IntRetCd = true then
        strTran_flag = Trim(Replace(lgF0, Chr(11), ""))
        strTax_cd = Trim(Replace(lgF1, Chr(11), ""))
    End if

    If  strTran_flag = "Y" then
        Call DisplayMsgBox("800408","X","X","X")'이미 자동기표가 처리되었습니다. 회계자동기표처리를 취소한 후 작업하시기 바랍니다.
        Exit Function
    End If

'*** 공제총액 합계 
    dblSub_tot_amt = 0
    dblPay_tot_amt = 0
    dblNon_tax1 = 0
    dblNon_tax2 = 0
    dblNon_tax3 = 0
    dblNon_tax4 = 0
    dblNon_tax_amt = 0
    
		dblSub_tot_amt = 0
		With Frm1
			ggoSpread.Source = frm1.vspdData1
			For lRow = 1 To .vspdData1.MaxRows
				.vspdData1.Row = lRow
				.vspdData1.Col = 0
	            
				if  .vspdData1.Text = ggoSpread.DeleteFlag then
					.vspdData1.Col = C_SUB_AMT
	                
				else
					.vspdData1.Col = C_SUB_AMT
					dblSub_tot_amt = dblSub_tot_amt + UNICDbl(.vspdData1.text)
				end if
			next
	        
			.txtSub_tot_amt.text = UNIFormatNumber(dblSub_tot_amt, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
			
        End with

    Call MakeKeyStream("X")
	Call DisableToolBar(Parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
	
	If gSpreadFlg = 1 Then
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
	Else
        If Frm1.vspdData1.MaxRows < 1 Then
           Exit Function
        End If
    
        With frm1.vspdData1
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData1
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
	If gSpreadFlg = 1 Then
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo
	Else
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.EditUndo
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
		If gSpreadFlg = 1 Then
			.vspdData.ReDraw = False
			.vspdData.focus
			ggoSpread.Source = .vspdData
            ggoSpread.InsertRow .vspdData.ActiveRow, imRow
            SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
			.vspdData.ReDraw = True
		Else
			.vspdData1.ReDraw = False
			.vspdData1.focus
			ggoSpread.Source = .vspdData1
            ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
            SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
			.vspdData1.ReDraw = True
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
    
    If gSpreadFlg = 1 Then
		If Frm1.vspdData.MaxRows < 1 then
			Exit function
		End if	
		
		With Frm1.vspdData 
    		.focus
    		ggoSpread.Source = frm1.vspdData 
    		lDelRows = ggoSpread.DeleteRow
		End With
	Else
		If Frm1.vspdData1.MaxRows < 1 then
			Exit function
		End if	
		
		With Frm1.vspdData1
    		.focus
    		ggoSpread.Source = frm1.vspdData1
    		lDelRows = ggoSpread.DeleteRow
		End With
	End If
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
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

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgSpreadFlg="       & gSpreadFlg
	if gSpreadFlg = "1" then
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
		
	If   LayerShowHide(1) = False Then
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
                    .vspdData.Col = C_ALLOW_CD  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_ALLOW_AMT : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData.Col = C_ALLOW_CD  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_ALLOW_AMT : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                    .vspdData.Col = C_ALLOW_CD  : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
		Next
	
	    .txtMaxRows.value     = lGrpCnt-1	
	    .txtSpread.value      = strDel & strVal
	   
	    ggoSpread.Source = frm1.vspdData1

		strVal = ""
		strDel = ""
        lGrpCnt = 1
		For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData1.Col = C_SUB_CD  	: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_SUB_AMT   : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData1.Col = C_SUB_CD  	: strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_SUB_AMT   : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                    .vspdData1.Col = C_SUB_CD    : strDel = strDel & Trim(.vspdData1.Text) & Parent.gRowSep									
                    lGrpCnt = lGrpCnt + 1
           End Select
		Next
	
	   .txtMaxRows1.value     = lGrpCnt-1	
	   .txtSpread1.value      = strDel & strVal
	
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
		
	If   LayerShowHide(1) = False Then
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
	Call SetToolbar("1100111111011111")												'⊙: Set ToolBar
	Call ggoOper.LockField(Document, "Q")
	Set gActiveElement = document.ActiveElement   
	lgBlnFlgChgValue = False
	frm1.vspdData.focus
	lgIntFlgMode = Parent.OPMD_UMODE

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

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = frm1.txtEmpNo.value
	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd        			' Internal_cd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmpNo.focus
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
		.txtEmpNo.value = arrRet(0)
		.txtEmpNm.value = arrRet(1)
		.txtEmpNo.focus
	End With
End Sub

'========================================================================================================
' Name : OpenAllow()
' Desc : Allow Popup
'========================================================================================================
Function OpenAllow(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = "수당팝업"				' 팝업 명칭 
		arrParam(2) = frm1.vspdData.Text			' Code Condition
		arrParam(4) = " code_type = '1'"			' Where Condition
		arrParam(5) = "수당"					' 조건필드의 라벨 명칭 
	Else
		arrParam(0) = "공제팝업"				' 팝업 명칭 
		arrParam(2) = frm1.vspdData1.Text			' Code Condition
		arrParam(4) = " code_type = '2'"			' Where Condition
		arrParam(5) = "공제"					' 조건필드의 라벨 명칭 
	End If
	
	arrParam(1) = "hda010t"							' TABLE 명칭 
	arrParam(3) = ""								' Name Cindition
		
    arrField(0) = "allow_cd"						' Field명(0)
    arrField(1) = "allow_nm"						' Field명(1)
    arrField(2) = "tax_type"						' Field명(2)
        
    arrHeader(0) = "코드"						' Header명(0)
    arrHeader(1) = "코드명"						' Header명(1)
    arrHeader(2) = "세액구분"		       ' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 			
			frm1.vspdData.Col = C_ALLOW_CD
			frm1.vspdData.action =0
		Else
			frm1.vspdData1.Col = C_SUB_CD
			frm1.vspdData1.action =0
		End If
		Exit Function
	Else
		Call SetAllow(arrRet, iWhere)
	End If	
			
End Function

'========================================================================================================
'	Name : SetAllow()
'	Description : Allow Popup에서 Return되는 값 setting
'========================================================================================================
Function SetAllow(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 			
			.vspdData.Col = C_ALLOW_CD_NM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_TAX_TYPE
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_ALLOW_CD
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		Else
			.vspdData1.Col = C_SUB_CD_NM
			.vspdData1.Text = arrRet(1)
			.vspdData1.Col = C_SUB_CD
			.vspdData1.Text = arrRet(0)
			.vspdData1.action =0
		End If
	End With
End Function

'========================================================================================================
'   Event Name : txtEmpNo_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmpNo_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtEmpNm.value = ""

    If  frm1.txtEmpNo.value = "" Then
		frm1.txtEmpNo.value = ""
        frm1.txtEmpNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmpNo.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmpNo.focus
            Set gActiveElement = document.ActiveElement
            txtEmpNo_Onchange = true
        Else
            frm1.txtEmpNm.value = strName
        End if 
    End if
    
End Function

'========================================================================================================

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
' Function Name : vspdData_Click
' Function Desc : gSpreadFlg Setting
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

	gSpreadFlg = 1

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
' Function Name : vspdData1_Click
' Function Desc : gSpreadFlg Setting
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SP2C"   

    Set gActiveSpdSheet = frm1.vspdData1

	gSpreadFlg = 2
   
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

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
     End If
End Sub    

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================%>
Function vspdData_Change(ByVal Col , ByVal Row )
       
   Dim IntRetCD

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

   Select Case Col
         Case  C_ALLOW_CD
           	IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd='*' And code_type='1' And allow_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           	If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                Call DisplayMsgBox("800145","X","X","X")             '☜ : 등록되지 않은 코드입니다.
    	    	frm1.vspdData.Col = C_ALLOW_CD_NM
           		frm1.vspdData.Text=""
           		vspdData_Change = true
            Else
    	    	frm1.vspdData.Col = C_ALLOW_CD_NM
            	frm1.vspdData.Text=Trim(Replace(lgF1,Chr(11),""))
           	End If
   End Select    

   If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
   End If
   
   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row
   lgSpreadChange = True
	
End Function

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Private Sub vspdData1_Change(ByVal Col , ByVal Row )
   Dim IntRetCD

   Frm1.vspdData1.Row = Row
   Frm1.vspdData1.Col = Col

   Select Case Col
         Case  C_SUB_CD
           	IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd='*' And code_type='2' And allow_cd =  " & FilterVar(frm1.vspdData1.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           	If IntRetCD=False And Trim(frm1.vspdData1.Text)<>"" Then
                Call DisplayMsgBox("800176","X","X","X")             '☜ : 등록되지 않은 코드입니다.
    	    	frm1.vspdData1.Col = C_SUB_CD_NM
           		frm1.vspdData1.Text=""
            Else
    	    	frm1.vspdData1.Col = C_SUB_CD_NM
            	frm1.vspdData1.Text=Trim(Replace(lgF1,Chr(11),""))
           	End If
   End Select    

   If Frm1.vspdData1.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData1.text) < CDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
   End If
   
   ggoSpread.Source = frm1.vspdData1
   ggoSpread.UpdateRow Row
   lgSpreadChange1 = True

End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================%>
Private Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_ALLOW_CD_POP Then
		    .Row = Row
		    .Col = C_ALLOW_CD

		    Call OpenAllow(0)
    End If
    
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================%>
Private Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData1 
	
    ggoSpread.Source = frm1.vspdData1
   
    If Row > 0 And Col = C_SUB_CD_POP Then
		    .Row = Row
		    .Col = C_SUB_CD

		    Call OpenAllow(1)        
    End If
    
    End With
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
'=======================================
'   Event Name : txtRevoke_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtPayYymm_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtPayYymm.Action = 7
        frm1.txtPayYymm.focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtPayYymm_KeyDown()
'   Event Desc : 조회조건부의 txtPayYymm_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtPayYymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여조회및조정</font></td>
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
			            	    	<TD CLASS="TD5" NOWRAP>정산년월</TD>
			            	    	<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtPayYymm name=txtPayYymm CLASS=FPDTYYYYMM title=FPDATETIME tag="12X1" ALT="정산년월"></OBJECT>');</SCRIPT></TD>
			            	    	<TD CLASS="TD5" NOWRAP>지급구분</TD>
			            	    	<TD CLASS="TD6"><SELECT NAME="txtType" CLASS ="cbonormal" tag="12" ALT="지급구분"></SELECT></TD>
			            	    </TR>
			    	            <TR>
			    	            	<TD CLASS="TD5">사원</TD>
									<TD CLASS="TD6">
									<INPUT TYPE=TEXT NAME="txtEmpNo" SIZE=10 MAXLENGTH=13 tag="12XXXU"  ALT="사원"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
									<INPUT TYPE=TEXT NAME="txtEmpNm" tag="14X"></TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
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
                                    
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>부서</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
							            	<TD CLASS=TD5 NOWRAP>직종코드</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOcpt_type" ALT="직종" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>급호봉</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd1" TYPE=TEXT SIZE="13" TAG="24XXXU" ALT="급호">
							            	                     <INPUT NAME="txtPay_grd2" TYPE=TEXT SIZE="5"  TAG="24XXXU" ALT="호봉">호봉</TD>
							            	<TD CLASS=TD5 NOWRAP>급여구분</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_cd" ALT="급여구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
							            </TR>	
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>연장비과세적용구분</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTax_cd" ALT="연장비과세적용구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
							            	<TD CLASS=TD5 NOWRAP>입퇴사구분</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtExcept_type" ALT="입퇴사구분" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="24XXXU"></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>지급일</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtProv_dt NAME="txtProv_dt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="지급일"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5>&nbsp;</TD>
							            	<TD CLASS=TD6>&nbsp;</TD>
							            </TR>
							            <TR>
											<TD CLASS=TD5 NOWRAP>일급</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtDDPay name=txtDDPay CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="일급"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>통상일급</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtComDDPay name=txtComDDPay CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="통상일급"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>배우자</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpouse" ALT="배우자" TYPE="Text" MAXLENGTH=20 SiZE=13 tag="24XXXU"></TD>
							            	<TD CLASS=TD5 NOWRAP>부양자</TD>
							            	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSupp_cnt" ALT="부양자" TYPE="Text" MAXLENGTH=20 SiZE=13 tag="24XXXU"></TD>
							            </TR>
							        </TABLE>
							        
							    </TD>
							</TR>
						    <TR>
						        <TD COLSPAN=4>
                                    
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>월차발생</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMm_holy_crt name=txtMm_holy_crt CLASS=FPDS40 title=FPDOUBLESINGLE tag="24X6" ALT="월차발생"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>월차지급</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMm_holy_prov name=txtMm_holy_prov CLASS=FPDS40 title=FPDOUBLESINGLE tag="24X6" ALT="월차지급"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>월차사용</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMm_holy_use name=txtMm_holy_use CLASS=FPDS40 title=FPDOUBLESINGLE tag="24X6" ALT="월차사용"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>월차적치</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMm_accum name=txtMm_accum CLASS=FPDS40 title=FPDOUBLESINGLE tag="24X6" ALT="월차적치"></OBJECT>');</SCRIPT></TD>
							            </TR>
							        </TABLE>
							        
							    </TD>
							</TR>
						    <TR>
						        <TD COLSPAN=4>
                                    
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>연장비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax1 name=txtNon_tax1 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="연장비과세"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>비과세총액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax_sum name=txtNon_tax_sum CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="비과세총액"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>식대비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax2 name=txtNon_tax2 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="식대비과세"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>과세분금액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtTax_amt name=txtTax_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="과세분금액"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>기타비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax3 name=txtNon_tax3 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="기타비과세"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>급여총액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtProv_tot_amt name=txtProv_tot_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="급여총액"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>기자비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax4 name=txtNon_tax4 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="기자비과세"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>공제총액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_tot_amt name=txtSub_tot_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="공제총액"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>국외근로비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax5 name=txtNon_tax5 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="국외근로비과세"></OBJECT>');</SCRIPT></TD>
							            	<TD CLASS=TD5 NOWRAP>실지급액</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtReal_prov_amt name=txtReal_prov_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="실지급액"></OBJECT>');</SCRIPT></TD>
							            </TR>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>연구비과세</TD>
							            	<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax6 name=txtNon_tax6 CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="연구비과세"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>국민,건강,고용</TD>
							              <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEtc_non_tax name=txtEtc_non_tax  CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2" ALT="국민,건강,고용"></OBJECT>');</SCRIPT></TD>							            	
								<TD CLASS=TD5 NOWRAP></TD>
							            	<TD CLASS=TD6 NOWRAP></TD>
							            </TR>							            
							        </TABLE>
							        
							    </TD>
							</TR>
							<TR>
								<TD HEIGHT="35%" WIDTH="50%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT="35%" WIDTH="50%" COLSPAN=2>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtsave_fund"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtincome_tax"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtres_tax"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtmed_insur"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtanut"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtemp_insur"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtTaxFlag"     TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows1" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

