<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h3001ma1
*  4. Program Name         : 퇴직금조회및조정 
*  5. Program Desc         : 퇴직금 조회,등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/16
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "ha105mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "ha105bb1.asp"
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const C_SHEETMAXROWS1   = 21                                           '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgStrComDateType		                                            'Company Date Type을 저장(년월 Mask에 사용함.)
Dim lsInternal_cd
Dim lgStrPrevKey1
Dim topleftOK

Dim C_BONUS_YYMM                                               'Column Dimant for Spread Sheet 
Dim C_BONUS_TYPE 															
Dim C_BONUS_TYPE_NM 															
Dim C_BONUS_TYPE_NM_POP 
Dim C_BONUS_AMT   

Dim C_PAY_YYMM                                               'Column constant for Spread Sheet 
Dim C_ALLOW_CD														
Dim C_ALLOW_NM 													
Dim C_ALLOW_NM_POP 														
Dim C_ALLOW_AMT 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_PAY_YYMM = 1
        C_ALLOW_CD = 2
        C_ALLOW_NM = 3
        C_ALLOW_NM_POP = 4
        C_ALLOW_AMT = 5
    End If

    If pvSpdNo = "B" Then   
        C_BONUS_YYMM = 1
        C_BONUS_TYPE = 2
        C_BONUS_TYPE_NM = 3
        C_BONUS_TYPE_NM_POP = 4
        C_BONUS_AMT = 5
    End If
    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgStrPrevKey       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey1       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey1 = ""                                     '⊙: initializes Previous Key Index
    lgSortKey          = 1                                      '⊙: initializes sort direction
    lsInternal_cd      = ""
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    Dim strYear,strMonth,strDay
	frm1.txtRetire_yymm.focus
	Call ggoOper.FormatDate(frm1.txtRetire_yymm, parent.gDateFormat, 2)
	
	Call ExtractDateFrom("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtRetire_yymm.Year	= strYear
	frm1.txtRetire_yymm.Month	= strMonth
	frm1.txtRetire_yymm.Day		= strDay
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
Function CookiePage(ByVal flgs)
End Function
	
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)    
       lgKeyStream       = UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtRetire_yymm.Year,Right("0" & frm1.txtRetire_yymm.Month,2),frm1.txtRetire_yymm.Day) & parent.gColSep
       lgKeyStream       = lgKeyStream & Trim(frm1.txtEmp_no.value) & parent.gColSep
       lgKeyStream       = lgKeyStream & UniConvDateAToB(Trim(frm1.txtRetire_dt.Text),parent.gDateFormat, parent.gServerDateFormat) & parent.gColSep
       lgKeyStream       = lgKeyStream & Trim(frm1.txtCalcu_logic.value) & parent.gColSep
       lgKeyStream       = lgKeyStream & Trim(frm1.txtPay_logic.value) & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
     
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0116", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '회사근무시간일 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtCalcu_logic,iCodeArr, iNameArr,Chr(11))

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0117", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) '퇴직평균급여산정방식 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtPay_logic,iCodeArr, iNameArr,Chr(11))
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
    Dim iSum    	
	With frm1

		For intRow = 1 To .vspdData.MaxRows			
			.vspdData.Row = intRow
			.vspdData.Col = C_ALLOW_NM
			If  Trim(.vspdData.Text) = "총합계" Or Trim(.vspdData.Value) = "월소계" Then
                 ggoSpread.Source = .vspdData
                .vspdData.col = C_PAY_YYMM
                .vspddata.text = ""
                 ggoSpread.SSSetProtected   C_ALLOW_AMT , intRow, intRow    
			
				.vspdData.Col = C_ALLOW_NM                             
                if Trim(.vspdData.Text) = "월소계"  then
					Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Sub_Total)                
                else
                    Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Total)
                end if
			End If
		Next
        intRow = 1
       
	    For intRow = 1 To .vspdData1.MaxRows			
	    	.vspdData1.Row = intRow
	    	.vspdData1.Col = C_BONUS_TYPE_NM
			If  Trim(.vspdData1.Text) = "총합계" Then
                 ggoSpread.Source = .vspdData1
                .vspdData1.col = C_BONUS_YYMM
                .vspddata1.text = ""
                 ggoSpread.SSSetProtected   C_BONUS_AMT , intRow, intRow                
                Call SetSpreadBackColor(frm1.vspdData1,intRow,-1,intRow,-1,parent.C_RGB_Total)
			End If
	    Next	
    End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    Dim strMaskYM	

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	
	
    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 
    	
	    With Frm1.vspdData
            ggoSpread.Source = Frm1.vspdData
            ggoSpread.Spreadinit "V20021123",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols = C_ALLOW_AMT + 1                                                   '☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              '☜:☜: Hide maxcols
           .ColHidden = True        

           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk

                ggoSpread.SSSetMask	     C_PAY_YYMM,    	"급여년월",10,2, strMaskYM
                ggoSpread.SSSetEdit      C_ALLOW_CD,        "",20
                ggoSpread.SSSetEdit      C_ALLOW_NM,        "수당코드" ,18,,,20,2
                ggoSpread.SSSetButton    C_ALLOW_NM_POP
                ggoSpread.SSSetFloat     C_ALLOW_AMT,       "금액", 22,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

           Call ggoSpread.MakePairsColumn(C_ALLOW_NM,C_ALLOW_NM_POP)    'sbk

           Call ggoSpread.SSSetColHidden(C_ALLOW_CD,C_ALLOW_CD,True)
	
	       .ReDraw = true
	
           lgActiveSpd = "M"

           Call SetSpreadLock("A")
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With Frm1.vspdData1
            ggoSpread.Source = Frm1.vspdData1
            ggoSpread.Spreadinit "V20021123",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols = C_BONUS_AMT + 1                                                   ' ☜:☜: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:   

           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("B") 

                ggoSpread.SSSetMask	     C_BONUS_YYMM,    	"상여년월",10,2, strMaskYM
                ggoSpread.SSSetEdit      C_BONUS_TYPE,      "" ,5
                ggoSpread.SSSetEdit      C_BONUS_TYPE_NM,   "상여구분" ,18,,,50,2
                ggoSpread.SSSetButton    C_BONUS_TYPE_NM_POP
                ggoSpread.SSSetFloat     C_BONUS_AMT,       "상여금", 22,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

           Call ggoSpread.MakePairsColumn(C_BONUS_TYPE_NM,C_BONUS_TYPE_NM_POP)    'sbk

           Call ggoSpread.SSSetColHidden(C_BONUS_TYPE,C_BONUS_TYPE,True)
	
	       .ReDraw = true
	
           lgActiveSpd = "S"
	
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
             ggoSpread.Source = Frm1.vspdData
             .vspdData.ReDraw = False
             ggoSpread.SpreadLock     C_PAY_YYMM , -1, C_PAY_YYMM, -1
             ggoSpread.SpreadLock     C_ALLOW_CD , -1, C_ALLOW_CD, -1
             ggoSpread.SpreadLock     C_ALLOW_NM , -1, C_ALLOW_NM, -1
             ggoSpread.SpreadLock     C_ALLOW_NM_POP , -1, C_ALLOW_NM_POP, -1
             ggoSpread.SSSetRequired  C_ALLOW_AMT, -1, -1
             ggoSpread.SSSetProtected  .vspdData.MaxCols   , -1, -1
             .vspdData.ReDraw = True
        End If

        If pvSpdNo = "B" Then
             ggoSpread.Source = Frm1.vspdData1
             .vspdData1.ReDraw = False
             ggoSpread.SpreadLock     C_BONUS_YYMM , -1, C_BONUS_YYMM, -1
             ggoSpread.SpreadLock     C_BONUS_TYPE_NM , -1, C_BONUS_TYPE_NM, -1
             ggoSpread.SpreadLock     C_BONUS_TYPE_NM_POP , -1, C_BONUS_TYPE_NM_POP, -1
             ggoSpread.SSSetRequired  C_BONUS_AMT, -1, -1
             ggoSpread.SSSetProtected  .vspdData1.MaxCols   , -1, -1
             .vspdData1.ReDraw = True
        End If
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

   If Trim(lgActiveSpd) = "" Then
      lgActiveSpd = "M"
   End If

   Select Case UCase(Trim(lgActiveSpd))
       Case  "M"
                With Frm1
                        ggoSpread.Source = .vspdData
                       .vspdData.ReDraw  = False
                        ggoSpread.SSSetRequired    C_PAY_YYMM , pvStartRow, pvEndRow
                        ggoSpread.SSSetRequired    C_ALLOW_NM , pvStartRow, pvEndRow
                        ggoSpread.SSSetRequired    C_ALLOW_AMT , pvStartRow, pvEndRow
                       .vspdData.ReDraw = True
                End With
       Case  "S"
                With Frm1
                        ggoSpread.Source = .vspdData1
                       .vspdData1.ReDraw = False
                        ggoSpread.SSSetRequired    C_BONUS_YYMM , pvStartRow, pvEndRow
                        ggoSpread.SSSetRequired    C_BONUS_TYPE_NM , pvStartRow, pvEndRow
                        ggoSpread.SSSetRequired    C_BONUS_AMT , pvStartRow, pvEndRow
                       .vspdData1.ReDraw = True
                End With
   End Select
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

            C_PAY_YYMM = iCurColumnPos(1)
            C_ALLOW_CD = iCurColumnPos(2)
            C_ALLOW_NM = iCurColumnPos(3)
            C_ALLOW_NM_POP = iCurColumnPos(4)
            C_ALLOW_AMT = iCurColumnPos(5)    
            
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_BONUS_YYMM = iCurColumnPos(1)
            C_BONUS_TYPE = iCurColumnPos(2)
            C_BONUS_TYPE_NM = iCurColumnPos(3)
            C_BONUS_TYPE_NM_POP = iCurColumnPos(4)
            C_BONUS_AMT = iCurColumnPos(5)
            
    End Select    
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 ' go to 
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
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtRetire_yymm, parent.gDateFormat, 2)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	
	Call AppendNumberPlace("6","3","0")
	Call ggoOper.FormatNumber(frm1.txtTax_rate, "99.99", "0", False, 2)	                '세율 
            
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100000100101111")												'⊙: Set ToolBar
    
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
        If ggoSpread.SSCheckChange = False Then
               ggoSpread.Source = Frm1.vspdData1
            If ggoSpread.SSCheckChange = True Then
	        	IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
	        	If IntRetCD = vbNo Then
	        		Exit Function
	        	End If
            End If
        Else    
	        	IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
	        	If IntRetCD = vbNo Then
	        		Exit Function
	        	End If
        End If    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If
    
    lgCurrentSpd = "M"
    Call MakeKeyStream("X")

    If DbQuery = False Then
		Exit Function
	End If															'☜: Query db data
       
    FncQuery = True																'☜: Processing is OK
   
End Function

'========================================================================================================
' Name : FncQuery1
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery1()
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strWhere
    FncQuery1 = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    strWhere = " emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " And retire_dt = "
    strWhere = strWhere & " (SELECT MAX(retire_dt) FROM hga040t WHERE Emp_no =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & ")"
    IntRetCD = CommonQueryRs(" MAX(retire_dt) "," hga040t ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If IntRetCD=False Or CStr(Trim(Replace(lgF0,Chr(11),""))) = "" Then
    IntRetCD = CommonQueryRs(" retire_dt "," haa010t ", " emp_no =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False Or CStr(Trim(Replace(lgF0,Chr(11),""))) = "" Then
            Call DisplayMsgbox("800399","X","X","X")	                        '정규사원입니다 . 퇴직기초자료를 등록 한후 계산 하십시요.
            Exit Function
		End If
    End If
    If Trim(frm1.txtCalcu_logic.value) = "" Then
        Call DisplayMsgbox("970021","X",frm1.txtCalcu_logic.alt,"X")	                    '계산공식은 입력필수항목입니다.
        frm1.txtCalcu_logic.focus
        Exit Function
    End If

    If Trim(frm1.txtPay_logic.value) = "" Then
        Call DisplayMsgbox("970021","X",frm1.txtPay_logic.alt,"X")	            '평균급여산정은 입력필수항목입니다.
        frm1.txtPay_logic.focus
        Exit Function
    End If
    
    Call MakeKeyStream("X")
    Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery1 = false Then
		Call RestoreToolBar()
		Exit Function
	End If
           
    FncQuery1 = True																'☜: Processing is OK
   
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim strWhere
    Dim iRow   
    Dim TempDate
	Dim strYear,strMonth,strDay,ChkDate : ChkDate = False

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    lgCurrentSpd = "M"
    Select Case UCase(Trim(lgActiveSpd))
      Case  "M"
        ggoSpread.Source = Frm1.vspdData

        If ggoSpread.SSCheckChange = False Then
               ggoSpread.Source = Frm1.vspdData1
            If ggoSpread.SSCheckChange = False Then
               IntRetCD = DisplayMsgbox("900001","X","X","X")                           '⊙: No data changed!!
                Exit Function
            End If
        End If
    
        If Not chkField(Document, "2") Then
           Exit Function
        End If
	
	    ggoSpread.Source = Frm1.vspdData
        If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
           Exit Function
        End If
      
      Case  "S"
        ggoSpread.Source = Frm1.vspdData1

        If ggoSpread.SSCheckChange = False Then
               ggoSpread.Source = Frm1.vspdData
            If ggoSpread.SSCheckChange = False Then
               IntRetCD = DisplayMsgbox("900001","X","X","X")                           '⊙: No data changed!!
                Exit Function
            End If
        End If
    
        If Not chkField(Document, "2") Then
           Exit Function
        End If
	
	    ggoSpread.Source = Frm1.vspdData1
        If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
           Exit Function
        End If
   End Select

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    strWhere = " emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " And retire_dt = "
    strWhere = strWhere & " (SELECT MAX(retire_dt) FROM hga040t WHERE Emp_no =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & ")"
    IntRetCD = CommonQueryRs(" MAX(retire_dt) "," hga040t ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If IntRetCD=False Or CStr(Trim(Replace(lgF0,Chr(11),""))) = "" Then
    
		IntRetCD = CommonQueryRs(" retire_dt "," haa010t ", " emp_no =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False Or CStr(Trim(Replace(lgF0,Chr(11),""))) = "" Then
            Call DisplayMsgbox("800399","X","X","X")	                        '정규사원입니다 . 퇴직기초자료를 등록 한후 계산 하십시요.
            Exit Function
        Else 
            Call ExtractDateFrom(Trim(Replace(lgF0,Chr(11),"")),parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
			frm1.txtRetire_yymm.Year	= strYear
			frm1.txtRetire_yymm.Month	= strMonth
			frm1.txtRetire_yymm.Day		= strDay        
		End If
    Else 
		
		Call ExtractDateFrom(Trim(Replace(lgF0,Chr(11),"")),parent.gAPDateFormat,parent.gAPDateSeperator,strYear,strMonth,strDay)
		
		frm1.txtRetire_yymm.Year	= strYear
		frm1.txtRetire_yymm.Month	= strMonth
		frm1.txtRetire_yymm.Day		= strDay        
		
    End If
    
    With Frm1.vspdData
		
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
                    .Col = C_ALLOW_AMT
                    If UNICDbl(.text) = 0  Then
                        Call DisplayMsgbox("800379","X","X","X")	            '수당금액은 입력항목입니다.
                        .Action = 0 ' go to 
                        Exit Function
                    End If
   	                .Col = C_PAY_YYMM
				    If .Text <> "" Then
						TempDate = lgConvDateAndFormatDate(.Text,parent.gComDateType,strYear,strMonth,strDay)
						ChkDate = CheckDateFormat(Trim(TempDate),parent.gDateFormat)				    
				    	If ChkDate = False And Not Trim(.Text) = parent.gComDateType And IsDate(strYear & parent.gServerDateType & strMonth  & parent.gServerDateType &  strDay) = False Then
				    		Call DisplayMsgbox("140318","X","X","X")	         '년월을 올바로 입력하세요.
				    		.Text = ""
							.Action = 0 ' go to 
							Set gActiveElement = document.activeElement
							Exit Function
				    	End If
				    	If Trim(.Text) = parent.gComDateType  Then
  							.Text = ""
							.Action = 0 ' go to 
							Set gActiveElement = document.activeElement
							Exit Function
						End If					
						ChkDate = False
				    End If
           End Select

        Next
    End With
    
    With Frm1.vspdData1
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
                    .Col = C_BONUS_AMT
                    If UNICDbl(.text) = 0  Then
                        Call DisplayMsgbox("970021","X","상여금","X")	    '상여금은 입력항목입니다.
                        .Action = 0 ' go to 
                        Exit Function
                    End If
   	                .Col = C_BONUS_YYMM
				    If .Text <> "" Then							
						TempDate = lgConvDateAndFormatDate(.Text,parent.gComDateType,strYear,strMonth,strDay)    
						ChkDate = CheckDateFormat(Trim(TempDate),parent.gDateFormat)				    				    	
				    	If ChkDate = False And Not Trim(.Text) = parent.gComDateType And IsDate(strYear & parent.gServerDateType & strMonth  & parent.gServerDateType &  strDay) = False Then
				    		Call DisplayMsgbox("140318","X","X","X")	         '년월을 올바로 입력하세요.
				    		.Text = ""
							.Action = 0 ' go to 
							Set gActiveElement = document.activeElement
							Exit Function				    	
						End If
						If Trim(.Text) = parent.gComDateType  Then
  							.Text = ""
							.Action = 0 ' go to 
							Set gActiveElement = document.activeElement
							Exit Function
						End If
						ChkDate = False
					End If
           End Select

        Next
    End With

    If DbSave = False Then
		Exit Function
	End If			                                                         '☜: Save db data
  
    FncSave = True                                                                   '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
       Case  "M"
                 If Frm1.vspdData.MaxRows < 1 Then
                    Exit Function
                 End If
    
                 With Frm1
	    
		              If .vspdData.ActiveRow > 0 Then
			              .vspdData.ReDraw = False
		
                          ggoSpread.Source = .vspdData	
                	     .vspdData.Row = .vspdData.ActiveRow
                	     .vspdData.Col = C_PAY_YYMM
                	     If (.vspdData.Text = "") Then
                   	        Exit function
                	     Else
                              ggoSpread.CopyRow
                	     End If

                         SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

                        .vspdData.Row  = .vspdData.ActiveRow
                        .vspdData.Col  = C_PAY_YYMM
                        .vspdData.Text = ""
                        .vspdData.Col  = C_ALLOW_CD
                        .vspdData.Text = ""
                        .vspdData.Col  = C_ALLOW_NM
                        .vspdData.Text = ""
    
                        .vspdData.ReDraw = True
                        .vspdData.focus
		              End If
	              End With
        Case  "S"
                  If Frm1.vspdData1.MaxRows < 1 Then
                     Exit Function
                  End If
    
	              With Frm1
	    
		              If .vspdData1.ActiveRow > 0 Then
			             .vspdData1.ReDraw = False
		
                          ggoSpread.Source = .vspdData1	
                	     .vspdData1.Row = .vspdData1.ActiveRow
                	     .vspdData1.Col = C_BONUS_YYMM
                	     If (.vspdData1.Text = "") Then
                   	        Exit function
                	     Else
                              ggoSpread.CopyRow
                	     End If
                         SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow

                        .vspdData1.Row  = .vspdData1.ActiveRow
                        .vspdData1.Col  = C_BONUS_YYMM
                        .vspdData1.Text = ""
                        .vspdData1.Col  = C_BONUS_TYPE
                        .vspdData1.Text = ""
                        .vspdData1.Col  = C_BONUS_TYPE_NM
                        .vspdData1.Text = ""
    
                        .vspdData1.ReDraw = True
                        .vspdData1.focus
		              End If
	              End With
	              
    End Select
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
                  ggoSpread.Source = Frm1.vspdData	
        Case  "S"
                  ggoSpread.Source = Frm1.vspdData1	
    End Select     
    ggoSpread.EditUndo  

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

    Dim imRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If

    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
                  With Frm1
                         .vspdData.ReDraw = False
                         .vspdData.Focus
                          ggoSpread.Source = .vspdData
                        '-----합계된 Row이면 삽입을 Disable시킴 
                	     .vspdData.Row = .vspdData.ActiveRow
                	     .vspdData.Col = C_PAY_YYMM
                	     If (.vspdData.Text = "") Then
                    	     .vspdData.Col = 0
                    	     If (.vspdData.Text = ggoSpread.InsertFlag) Then
                                ggoSpread.InsertRow .vspdData.ActiveRow, imRow
                        	 Else
                       	        Exit function
                       	     End If
                	     Else
                            ggoSpread.InsertRow .vspdData.ActiveRow, imRow
                	     End If
                        '--------------------------------------                	     
                         SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
                         .vspdData.ReDraw = True
                  End With
        Case  Else
                  With Frm1
                         .vspdData1.ReDraw = False
                         .vspdData1.Focus
                          ggoSpread.Source = .vspdData1
                        '-----합계된 Row이면 삽입을 Disable시킴 
                	     .vspdData1.Row = .vspdData1.ActiveRow
                	     .vspdData1.Col = C_BONUS_YYMM
                	     If (.vspdData1.Text = "") Then
                    	     .vspdData.Col = 0
                    	     If (.vspdData.Text = ggoSpread.InsertFlag) Then
                                ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
                        	 Else
                       	        Exit function
                       	     End If
                	     Else
                             ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
                	     End If
                        '--------------------------------------                	     
                         SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
                         .vspdData1.ReDraw = True
                  End With
    End Select 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
	dim i
   Select Case UCase(Trim(lgActiveSpd))
       Case  "M"
                
                If Frm1.vspdData.MaxRows < 1 then
                   Exit function
                End if	
                With Frm1.vspdData 
                	.focus
                	ggoSpread.Source = frm1.vspdData
                	for i = .SelBlockRow to .SelBlockRow2
                		.Row = i
                		.Col = C_ALLOW_NM
                		If (.Text = "월소계" Or .Text = "총합계") Then
                			
                		Else
                			lDelRows = ggoSpread.DeleteRow(i,i)
                		End If
                	Next
                End With
       Case  "S"
                If Frm1.vspdData1.MaxRows < 1 then
                   Exit function
                End if	
                With Frm1.vspdData1 
                	.focus
                	ggoSpread.Source = frm1.vspdData1 
                	for i = .SelBlockRow to .SelBlockRow2
                		.Row = i
                		.Col = C_BONUS_TYPE_NM
                		If (.Text = "월소계" Or .Text = "총합계") Then
                			
                		Else
                			lDelRows = ggoSpread.DeleteRow(i,i)
                		End If
                	Next
                End With

   End Select

   Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
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
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

	ggoSpread.Source = Frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
   
	If lgCurrentSpd = "M" Then
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
    Else   
       strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows         '☜: Max fetched data
    End If   
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG
End Function

'========================================================================================================
' Name : DbQuery1
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery1()
    Dim strVal
    Err.Clear                                                                        '☜: Clear err status

    DbQuery1 = False                                                                 '☜: Processing is NG
    
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

    strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                    '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery1 = True                                                                   '☜: Processing is NG
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
	Dim strYear,strMonth,strDay
	
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

  	With Frm1
		.txtMode.value      = parent.UID_M0002                                          '☜: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    	If lgCurrentSpd = "M" Then
            For lRow = 1 To .vspdData.MaxRows
    
                .vspdData.Row = lRow
                .vspdData.Col = 0
             
                Select Case .vspdData.Text
 
                    Case ggoSpread.InsertFlag                                      '☜: Insert
                                                         strVal = strVal & "C" & parent.gColSep
                                                         strVal = strVal & lRow & parent.gColSep
                                                         strVal = strVal & Trim(lgCurrentSpd) & parent.gColSep
                                                         strVal = strVal & Trim(.txtRetire_dt.Text) & parent.gColSep                                                         
                                                         strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData.Col = C_PAY_YYMM    : Call lgConvDateAndFormatDate(.vspdData.Text,parent.gComDateType,strYear,strMonth,strDay)
													     strVal = strVal & Trim(strYear & strMonth) & parent.gColSep
                         .vspdData.Col = C_ALLOW_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                         .vspdData.Col = C_ALLOW_AMT   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
                         lGrpCnt = lGrpCnt + 1
                    Case ggoSpread.UpdateFlag                                      '☜: Update
                                                         strVal = strVal & "U" & parent.gColSep
                                                         strVal = strVal & lRow & parent.gColSep
                                                         strVal = strVal & Trim(lgCurrentSpd) & parent.gColSep
                                                         strVal = strVal & Trim(.txtRetire_dt.Text) & parent.gColSep                                                         
                                                         strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData.Col = C_PAY_YYMM    : Call lgConvDateAndFormatDate(.vspdData.Text,parent.gComDateType,strYear,strMonth,strDay)
													     strVal = strVal & Trim(strYear & strMonth) & parent.gColSep
                         .vspdData.Col = C_ALLOW_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                         .vspdData.Col = C_ALLOW_AMT   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                         lGrpCnt = lGrpCnt + 1
                    Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                         strDel = strDel & "D" & parent.gColSep
                                                         strDel = strDel & lRow & parent.gColSep
                                                         strDel = strDel & Trim(lgCurrentSpd) & parent.gColSep
                                                         strDel = strDel & Trim(.txtRetire_dt.Text) & parent.gColSep
                                                         strDel = strDel & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData.Col = C_PAY_YYMM    : Call lgConvDateAndFormatDate(.vspdData.Text,parent.gComDateType,strYear,strMonth,strDay)														 
														 strDel = strDel & Trim(strYear & strMonth) & parent.gColSep
                         .vspdData.Col = C_ALLOW_CD    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                         lGrpCnt = lGrpCnt + 1
                End Select
            Next
        Else
            For lRow = 1 To .vspdData1.MaxRows
    
                .vspdData1.Row = lRow
                .vspdData1.Col = 0
             
                Select Case .vspdData1.Text
 
                    Case ggoSpread.InsertFlag                                      '☜: Insert
                                                            strVal = strVal & "C" & parent.gColSep
                                                            strVal = strVal & lRow & parent.gColSep
                                                            strVal = strVal & Trim(lgCurrentSpd) & parent.gColSep
                                                            strVal = strVal & Trim(.txtRetire_dt.Text) & parent.gColSep
                                                            strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData1.Col = C_BONUS_YYMM    : Call lgConvDateAndFormatDate(.vspdData1.Text,parent.gComDateType,strYear,strMonth,strDay)
															strVal = strVal & Trim(strYear & strMonth) & parent.gColSep
                         .vspdData1.Col = C_BONUS_TYPE    : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                         .vspdData1.Col = C_BONUS_AMT     : strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep   
                         lGrpCnt = lGrpCnt + 1
                    Case ggoSpread.UpdateFlag                                      '☜: Update
                                                            strVal = strVal & "U" & parent.gColSep
                                                            strVal = strVal & lRow & parent.gColSep
                                                            strVal = strVal & Trim(lgCurrentSpd) & parent.gColSep
                                                            strVal = strVal & Trim(.txtRetire_dt.Text) & parent.gColSep
                                                            strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData1.Col = C_BONUS_YYMM    : Call lgConvDateAndFormatDate(.vspdData1.Text,parent.gComDateType,strYear,strMonth,strDay)
															strVal = strVal & Trim(strYear & strMonth) & parent.gColSep
                         .vspdData1.Col = C_BONUS_TYPE    : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                         .vspdData1.Col = C_BONUS_AMT     : strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep   
                         lGrpCnt = lGrpCnt + 1
                    Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                            strDel = strDel & "D" & parent.gColSep
                                                            strDel = strDel & lRow & parent.gColSep
                                                            strDel = strDel & Trim(lgCurrentSpd) & parent.gColSep
                                                            strDel = strDel & Trim(.txtRetire_dt.Text) & parent.gColSep
                                                            strDel = strDel & Trim(.txtEmp_no.value) & parent.gColSep
                         .vspdData1.Col = C_BONUS_YYMM    : Call lgConvDateAndFormatDate(.vspdData1.Text,parent.gComDateType,strYear,strMonth,strDay)
															strDel = strDel & Trim(strYear & strMonth) & parent.gColSep															
                         .vspdData1.Col = C_BONUS_TYPE    : strDel = strDel & Trim(.vspdData1.Text) & parent.gRowSep
                         lGrpCnt = lGrpCnt + 1
                End Select
            Next
        End If
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
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '☜: Processing is OK
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    if lgCurrentSpd = "S" then
	 	If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows <= 0 And Not frm1.Sflag.value Then
	 		Call DisplayMsgbox("900014","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
	 	End If
    End If
	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
	Call ggoOper.LockField(Document, "Q")
   call InitData()	
    frm1.vspdData.focus

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables

 	If lgCurrentSpd = "M" Then
	   lgCurrentSpd       = "S"
       DbSave()
    Else
        ggoSpread.Source = Frm1.vspdData
        Frm1.vspdData.MaxRows = 0
        lgCurrentSpd = "M"
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

	   Call MakeKeyStream(1)
      		Call DisableToolBar(parent.TBC_QUERY)
      		If DBQuery = False Then
      			Call RestoreToolBar ()
      			Exit Function
      		End If
	End If

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function
'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function ExeReflectOk()				                '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 
	IntRetCD =DisplayMsgbox("800154","X","X","X")    '퇴직금계산이 완료되었습니다.
    Call MainQuery
End Function
'======================================================================================================
' Function Name : ExeReflectNo
' Function Desc : ExeReflect가 실패했을 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function ExeReflectNo()				                '☆: 실행된 자료가 없습니다 
	Dim IntRetCD 
    Call DisplayMsgbox("800161","X","X","X")
End Function

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
	    Case 0
	        arrParam(0) = "수당코드 팝업"			        ' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		        ' TABLE 명칭 
	        arrParam(2) = ""                		            ' Code Condition
	        arrParam(3) = strCode						        ' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  "  ' Where Condition
	        arrParam(5) = "수당코드"			            ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					        ' Field명(0)
            arrField(1) = "ALLOW_NM"				            ' Field명(1)
    
            arrHeader(0) = "수당코드"				        ' Header명(0)
            arrHeader(1) = "수당코드명"
	    Case 1
	        arrParam(0) = "상여구분코드 팝업"			    ' 팝업 명칭 
	    	arrParam(1) = "b_minor"	    						' TABLE 명칭 
	    	arrParam(2) = ""                  		        	' Code Condition
	        arrParam(3) = strCode						        ' Name Cindition
	    	arrParam(4) = " major_cd=" & FilterVar("h0040", "''", "S") & " "                  ' Where Condition
	    	arrParam(5) = "상여코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"					    	' Field명(0)
	    	arrField(1) = "minor_nm"    				    	' Field명(1)
    
	    	arrHeader(0) = "상여코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "상여코드명"	   		            ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    Select Case iWhere
	        Case 0
	        	frm1.vspdData.Col = C_ALLOW_NM
	        	frm1.vspdData.action =0
	        Case 1
	        	frm1.vspdData1.Col = C_BONUS_TYPE_NM
	        	frm1.vspdData1.action =0
        End Select
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
 	    If iWhere = 0 Then
           	ggoSpread.Source = frm1.vspdData
            ggoSpread.UpdateRow Row
        Else
           	ggoSpread.Source = frm1.vspdData1
            ggoSpread.UpdateRow Row
        End If
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)
	    With frm1
	    	Select Case iWhere
	    	    Case 0
	    	        .vspdData.Col = C_ALLOW_CD
	    	    	.vspdData.text = arrRet(0) 
	    	    	.vspdData.Col = C_ALLOW_NM
	    	    	.vspdData.text = arrRet(1)   
	    	    	.vspdData.action =0
	    	    Case 1
	    	        .vspdData1.Col = C_BONUS_TYPE
	    	    	.vspdData1.text = arrRet(0) 
	    	    	.vspdData1.Col = C_BONUS_TYPE_NM
	    	    	.vspdData1.text = arrRet(1)   
	    	    	.vspdData1.action =0
            End Select
	    End With

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

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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

		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field

		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
    lgActiveSpd      = "M"
End Sub

'========================================================================================================
' Function Name : vspdData_Click
' Function Desc : gSpreadFlg Setting
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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

Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
     End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	    With Frm1.vspdData
	    	ggoSpread.Source = Frm1.vspdData
	    	If Row > 0 Then
	    		Select Case Col
	    		       Case C_ALLOW_NM_POP
	    		        	.Col = Col - 1
	    		        	.Row = Row
	    		        	Call OpenCode(.Text,0,Row)
	    		End Select
	    	End If
    
	    End With
End Sub
'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	    With Frm1.vspdData1
	    	ggoSpread.Source = Frm1.vspdData1
	    	If Row > 0 Then
	    		Select Case Col
	    		       Case C_BONUS_TYPE_NM_POP
	    		        	.Col = Col - 1
	    		        	.Row = Row
	    		        	Call OpenCode(.Text,1,Row)
	    		End Select
	    	End If
    
	    End With
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

    Select Case Col
	    Case C_ALLOW_NM
            IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd = " & FilterVar("*", "''", "S") & "  and code_type = " & FilterVar("1", "''", "S") & "  And allow_nm =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD=False And Trim(frm1.vspdData.Text)<>""  Then
                Call DisplayMsgbox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
                frm1.vspdData.focus
                
            ElseIf CountStrings(lgF0, Chr(11) ) > 1 Then    ' 같은명일 경우 pop up
                Call OpenCode(frm1.vspdData.Text, 0 , Row)
            Else
    	        frm1.vspdData.Col = C_ALLOW_CD
                frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
            End If
    End Select    
             
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
          Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change( Col ,  Row)

    Dim iDx
    Dim IntRetCD
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col

    Select Case Col
	    Case C_BONUS_TYPE_NM
            IntRetCD = CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0040", "''", "S") & " And minor_nm =  " & FilterVar(frm1.vspdData1.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD=False And Trim(frm1.vspdData1.Text)<>""  Then
                Call DisplayMsgbox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
                frm1.vspdData.focus
                
            ElseIf CountStrings(lgF0, Chr(11) ) > 1 Then    ' 같은명일 경우 pop up
                Call OpenCode(frm1.vspdData1.Text, 1 , Row)
            Else
    	        frm1.vspdData1.Col = C_BONUS_TYPE
                frm1.vspdData1.Text=Trim(Replace(lgF0,Chr(11),""))
            End If
    End Select    
             
    If Frm1.vspdData1.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData1.text) < CDbl(Frm1.vspdData1.TypeFloatMin) Then
          Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
Dim strYear,strMonth,strDay,TempDate,ChkDate : ChkDate = False

     With frm1.vspdData
		If Col <> NewCol And NewCol > 0 Then
			If Col = C_PAY_YYMM Then
				.Row = Row
				.Col = Col
				If .Text <> "" Then									
					TempDate = lgConvDateAndFormatDate(.Text,parent.gComDateType,strYear,strMonth,strDay)					
					ChkDate = CheckDateFormat(Trim(TempDate),parent.gDateFormat)				    					
					If ChkDate = False And IsDate(strYear & parent.gServerDateType & strMonth  & parent.gServerDateType &  strDay) = False Then
						Call DisplayMsgbox("140318","X","X","X")	'년월을 올바로 입력하세요.
						.Text = ""
						.Action = 0 ' go to 
						Set gActiveElement = document.activeElement
					End If
				End If
			End If
		End If
    End With
End Sub
'========================================================================================================
'   Event Name : vspdData1_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData1_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
Dim strYear,strMonth,strDay,TempDate,ChkDate : ChkDate = False

     With frm1.vspdData1
		If Col <> NewCol And NewCol > 0 Then
			If Col = C_BONUS_YYMM Then
				.Row = Row
				.Col = Col
				If .Text <> "" Then					
					TempDate = lgConvDateAndFormatDate(.Text,parent.gComDateType,strYear,strMonth,strDay)					
					ChkDate = CheckDateFormat(Trim(TempDate),parent.gDateFormat)				    
					If ChkDate = False And IsDate(strYear & parent.gServerDateType & strMonth  & parent.gServerDateType &  strDay) = False Then
						Call DisplayMsgbox("140318","X","X","X")	'년월을 올바로 입력하세요.
						.Text = ""
						.Action = 0 ' go to 
						Set gActiveElement = document.activeElement
					End If
				End If
			End If
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
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			topleftOK = true
			lgCurrentSpd = "M"
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			lgCurrentSpd = "S"
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S"
End Sub

'=======================================================================================================
'   Event Name : txtRetire_yymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtRetire_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRetire_yymm.Action = 7
        frm1.txtRetire_yymm.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtRetire_yymm_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtRetire_yymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'==========================================================================================
'   Event Name : btnCb_retire_calcu_OnClick()
'   Event Desc : 퇴직금재계산 
'==========================================================================================
Sub btnCb_retire_calcu_OnClick()
    Call FncQuery1()
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
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            Call ggoOper.ClearField(Document, "2")

            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
    
End Function

'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split(parent.gDateFormat,parent.gComDateType)
	
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
End Function

'========================================================================================================
' Function Name : lgConvDateAndFormatDate()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function lgConvDateAndFormatDate(Byval strDate,strDateType,strYear,strMonth,strDay)

	Dim i,ArrType,ArrDate,ArrgType,strTempDate,strgDate,TempYear,TempMonth
	strYear = "" : strMonth = "" : strDay = "" : strgDate = "" : TempYear = "" : TempMonth = ""

	If Trim(strDate) = "" Then
		lgConvDateAndFormatDate=""
		Exit Function
	End If
	
    ArrType = Split(parent.gDateFormatYYYYMM,strDateType)
    ArrDate = Split(Trim(strDate),strDateType)
    If IsArray(ArrType) And IsArray(ArrDate) Then
		For i=0 To Ubound(ArrType)
			If Instr(UCase(ArrType(i)),"Y") Then
				TempYear = ArrDate(i)
				If Len(Trim(ArrType(i))) <= 2 Then
					strYear = ConvertYYToYYYY(TempYear)
				Else
					strYear = TempYear
				End If
			ElseIf Instr(UCase(ArrType(i)),"M") Then
				TempMonth = ArrDate(i)
				If Len(Trim(ArrType(i))) >= 3 Then
					strMonth = ConvertMMMToMM(TempMonth)
				Else
					strMonth = TempMonth
				End If
			End If
		Next
		strDay = "01"	
	End If
		    
    ArrgType = Split(parent.gDateFormat,strDateType)
    If IsArray(ArrgType) Then
		ReDim strTempDate(Ubound(ArrgType))
		For i=0 To Ubound(ArrgType)
			If Instr(UCase(ArrgType(i)),"Y") Then
				strTempDate(i) = TempYear
			ElseIf Instr(UCase(ArrgType(i)),"M") Then
				strTempDate(i) = TempMonth
			ElseIf Instr(UCase(ArrgType(i)),"D") Then
				strTempDate(i) = strDay
			End If
			If i < Ubound(ArrgType) Then
				strTempDate(i) = strTempDate(i) & strDateType
			End If
		Next
	End If
	
	For i = 0 To Ubound(strTempDate)
		strgDate = strgDate & strTempDate(i)
	Next
	
	lgConvDateAndFormatDate = strgDate

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>퇴직금조회및조정</font></td>
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
								<TD CLASS=TD5  NOWRAP>퇴직년월</TD>
			                    <TD CLASS=TD6  NOWRAP><script language =javascript src='./js/ha105ma1_fpDateTime1_txtRetire_yymm.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
			    	    		                       <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>계산공식</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtCalcu_logic" ALT="계산공식" STYLE="WIDTH: 250px" TAG="1XN"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>평균급여산정방법</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtPay_logic" ALT="평균급여산정방법" STYLE="WIDTH: 150px" TAG="1XN"><OPTION VALUE=""></OPTION></SELECT></TD>
			            	</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%> >
							<TR HEIGHT=*>
								<TD WIDTH="50%" HEIGHT=* VALIGN=TOP>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            		    <TR>
									        <TD HEIGHT="100%"><script language =javascript src='./js/ha105ma1_vaSpread_vspdData.js'></script></TD>
									    </TR>    
					            	</TABLE>
								</TD>
								<TD WIDTH=*     HEIGHT=* VALIGN=TOP>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            		    <TR>
									        <TD HEIGHT="100%"><script language =javascript src='./js/ha105ma1_vaSpread1_vspdData1.js'></script></TD>
									    </TR>    
					            	</TABLE>
								</TD>
							</TR>
            			    <TR HEIGHT="40%">
            					<TD WIDTH=* HEIGHT=* VALIGN=TOP>
	            					<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>입사일</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDateTime3_txtEntr_dt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>퇴직금</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle1_txtRetire_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>퇴사일</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDateTime4_txtRetire_dt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>단체보험</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle2_txtCorp_insur_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>총근속개월수</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle3_txtTot_duty_mm.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>명예수당</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle4_txtHonor_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>급여평균</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle5_txtPay_avr_amt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>기타수당</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle6_txtEtc_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>상여평균</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle7_txtBonus_avr_amt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>갑근세</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle8_txtIncome_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP>연월차평균</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle9_txtYear_avr_amt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>주민세</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle10_txtRes_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>평균임금</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle11_txtAvr_wages_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>퇴직전환</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle12_txtRetire_anu_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP></TD>
								            <TD CLASS="TD6" NOWRAP></TD>
              						        <TD CLASS="TD5" NOWRAP>총지급액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle13_txtTot_prov_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
								            <TD CLASS="TD5" NOWRAP></TD>
								            <TD CLASS="TD6" NOWRAP></TD>
              						        <TD CLASS="TD5" NOWRAP>실지급액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle14_txtReal_prov_amt.js'></script></TD> 
	                   					</TR>
						            </TABLE>
					            </TD>
            					<TD WIDTH=* HEIGHT=* VALIGN=TOP>
	            					<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>소득공제</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle15_txtIncome_sub_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>근속년수</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle16_txtDuty_cnt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>특별공제</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle17_txtSpecial_sub_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>산출세액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle18_txtCalc_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>퇴직소득금액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle19_txtIncome_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>연평균산출세액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle20_txtAvr_calc_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>공제부족액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle21_txtSub_short_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>세액공제</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle22_txtTax_short_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>과세표준</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle23_txtTax_std_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>결정갑근세</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle24_txtDeci_income_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>연평균과세표준</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle25_txtAvr_tax_std_amt.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>결정주민세</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle26_txtDeci_res_tax_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP>세율</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle27_txtTax_rate.js'></script></TD> 
              						        <TD CLASS="TD5" NOWRAP>기타공제</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha105ma1_fpDoubleSingle28_txtEtc_sub_amt.js'></script></TD> 
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6" NOWRAP></TD> 
              						        <TD CLASS="TD5" NOWRAP><INPUT TYPE=HIDDEN NAME="Sflag" TAG="14"></TD>
	                   						<TD CLASS="TD6" NOWRAP></TD>  
	                   					</TR>
						            	<TR>
              						        <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6" NOWRAP></TD> 
              						        <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6" NOWRAP></TD> 
	                   					</TR>
						            </TABLE>
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
	        <TABLE <%=LR_SPACE_TYUPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnCb_retire_calcu" CLASS="CLSMBTN" Flag=1>퇴직금재계산</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

