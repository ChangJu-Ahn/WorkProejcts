<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 시스템마감입력 
*  3. Program ID           : H1019ma1
*  4. Program Name         : H1019ma1
*  5. Program Desc         : 기준정보관리/시스템마감입력 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/28
*  8. Modified date(Last)  : 2003/06/10
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H1019mb1.asp"						           '☆: Biz Logic ASP Name
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
Dim lgStrComDateType

Dim C_PAY_TYPE
Dim C_PAY_TYPE_NM
Dim C_CLOSE_TYPE
Dim C_CLOSE_TYPE_NM
Dim C_CLOSE_DT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_PAY_TYPE			= 1
	 C_PAY_TYPE_NM		= 2
	 C_CLOSE_TYPE		= 3
	 C_CLOSE_TYPE_NM	= 4
	 C_CLOSE_DT			= 5
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
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
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
    lgKeyStream  = "1" & parent.gColSep
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr   
    
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0104", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtclose_type1, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0104", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtclose_type2, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0104", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtclose_type3, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0104", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtclose_type4, iCodeArr, iNameArr, Chr(11))
   
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox2()
    Dim iCodeArr 
    Dim iNameArr   
    
     ggoSpread.Source = frm1.vspdData
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0104", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CLOSE_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CLOSE_TYPE_NM

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " and (minor_cd >= " & FilterVar("2", "''", "S") & " and minor_cd <= " & FilterVar("9", "''", "S") & ") ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAY_TYPE
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAY_TYPE_NM
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
			.Col = C_PAY_TYPE
			intIndex = .Value
			.col = C_PAY_TYPE_NM
			.Value = intindex					
			.Row = intRow
			.Col = C_CLOSE_TYPE
			intIndex = .Value
			.col = C_CLOSE_TYPE_NM
			.Value = intindex
		Next	
	End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

    Dim strMaskYM	

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_CLOSE_DT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
        
         Call  GetSpreadColumnPos("A")

         ggoSpread.SSSetCombo  C_PAY_TYPE,        "상여구분",        10
         ggoSpread.SSSetCombo  C_PAY_TYPE_NM,     "상여구분",        43
         ggoSpread.SSSetCombo  C_CLOSE_TYPE,      "상여마감구분",    10
         ggoSpread.SSSetCombo  C_CLOSE_TYPE_NM,   "상여마감구분",    43
         ggoSpread.SSSetMask   C_CLOSE_DT,        "상여마감년,월",   30, 2, strMaskYM
         
         Call ggoSpread.SSSetColHidden(C_PAY_TYPE,		C_PAY_TYPE,		True)
         Call ggoSpread.SSSetColHidden(C_CLOSE_TYPE,   C_CLOSE_TYPE,	True)

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
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SpreadLock    C_PAY_TYPE		, -1	, C_PAY_TYPE
         ggoSpread.SpreadLock    C_PAY_TYPE_NM	, -1	, C_PAY_TYPE_NM
         ggoSpread.SpreadLock    C_CLOSE_TYPE	, -1	, C_CLOSE_TYPE
         ggoSpread.SpreadLock    C_CLOSE_TYPE_NM, -1	, C_CLOSE_TYPE_NM
         ggoSpread.SSSetRequired C_CLOSE_DT		, -1	, C_CLOSE_DT
         ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SSSetProtected	C_PAY_TYPE		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_PAY_TYPE_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_CLOSE_TYPE	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_CLOSE_TYPE_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_CLOSE_DT		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
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
            
            C_PAY_TYPE			= iCurColumnPos(1)
			C_PAY_TYPE_NM		= iCurColumnPos(2)
			C_CLOSE_TYPE		= iCurColumnPos(3)
			C_CLOSE_TYPE_NM		= iCurColumnPos(4)
			C_CLOSE_DT			= iCurColumnPos(5)           	 
           	  
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call  AppendNumberPlace("7", "7", "3")
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call  ggoOper.FormatDate(frm1.txtclose_dt1,  parent.gDateFormat, 2)
    Call  ggoOper.FormatDate(frm1.txtclose_dt2,  parent.gDateFormat, 2)
    Call  ggoOper.FormatDate(frm1.txtclose_dt3,  parent.gDateFormat, 3)

    Call InitSpreadSheet                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables
	Call SetToolbar("1100110100010111")												'⊙: Set ToolBar
 
    Call InitComboBox
    Call InitComboBox2
	Call CookiePage (0)                                                             '☜: Check Cookie
	call MainQuery()

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

    Call MakeKeyStream("X")
    
    If DbQuery = False Then
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
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
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
    
    
    If DbDelete = False Then
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
    Dim strCloseDt
    Dim tmpDT
	Dim strYear, strMonth, strDay
	
	
    FncSave = False                                                              '☜: Processing is NG   
    
    Err.Clear                                                                    '☜: Clear err status
    tmpDT = Replace( parent.gDateFormatYYYYMM,"YYYY","1900")
    tmpDT = Replace(tmpDT            ,"YY"  ,"00")
    tmpDT = Replace(tmpDT            ,"MM"  ,"01")
        
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag
   	                .vspdData.Col = C_CLOSE_DT   	                
   	                
   	                If  CompareDateByFormat(tmpDT, .vspdData.Text, tmpDT, "입력일", 970023,  parent.gDateFormatYYYYMM,  parent.gComDateType, True) = False Then   	                
				       Exit Function
   	                End if   	                
   	                
   	                call  ExtractDateFrom(.vspdData.Text,  parent.gDateFormatYYYYMM, parent.gComDateType,strYear, strMonth, strDay)   	                
   	                if (Cint(strMonth) >12 or Cint(strMonth) < 1) then
   						IntRetCD =  DisplayMsgBox("800401","X","X","X")
   						Exit Function
   	                End if   	        
				    
            End Select            
        Next
	End With

     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

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

	With Frm1
           .vspdData.Col  = C_PAY_TYPE
           .vspdData.Row  = .vspdData.ActiveRow
           .vspdData.Text = ""

           .vspdData.Col  = C_PAY_TYPE_NM
           .vspdData.Row  = .vspdData.ActiveRow
           .vspdData.Text = ""
           
           .vspdData.Col  = C_CLOSE_DT
           .vspdData.Row  = .vspdData.ActiveRow
           .vspdData.Text = ""

			.vspdData.ReDraw = True
			.vspdData.focus		
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
    Call Initdata()
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
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1            
        
       .vspdData.ReDraw = True
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
    	.focus
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

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
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
		
	Call LayerShowHide(1)
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
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
 
               Case  ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & "1" & parent.gColSep   ' 사업장 
                    .vspdData.Col = C_PAY_TYPE 	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CLOSE_TYPE: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CLOSE_DT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & "1" & parent.gColSep   ' 사업장 
                    .vspdData.Col = C_PAY_TYPE 	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CLOSE_TYPE: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CLOSE_DT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                 strDel = strDel & "D" & parent.gColSep
                                                 strDel = strDel & lRow & parent.gColSep
                                                 strDel = strDel & "1" & parent.gColSep   ' 사업장 
                    .vspdData.Col = C_PAY_TYPE : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep									
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
		
	Call LayerShowHide(1)
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    IF  Frm1.vspdData.MaxRows > 0 then
	    Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    else
        Call SetToolbar("1100110100111111")
    end if

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
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
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Function

'========================================================================================================
' Name : txtClose_dt1_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtClose_dt1_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")  
        frm1.txtClose_dt1.Action = 7 		
        frm1.txtClose_dt1.Focus        
    End If
End Sub

'========================================================================================================
' Name : txtClose_dt2_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtClose_dt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtClose_dt2.Action = 7 
        frm1.txtClose_dt2.Focus
    End If
End Sub

'========================================================================================================
' Name : txtClose_dt3_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtClose_dt3_DblClick(Button)
    If Button = 1 Then
        frm1.txtClose_dt3.Action = 7 
        frm1.txtClose_dt3.Focus        
    End If
End Sub

'========================================================================================================
' Name : txtClose_dt4_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtClose_dt4_DblClick(Button)
    If Button = 1 Then
        frm1.txtClose_dt4.Action = 7 
        frm1.txtClose_dt4.Focus        
    End If
End Sub

'========================================================================================================
Sub txtClose_type1_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_type2_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_type3_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_type4_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_dt1_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_dt2_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_dt3_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtClose_dt4_Change()
    lgBlnFlgChgValue = True
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
        Case C_PAY_TYPE_NM
            Frm1.vspdData.Row = Row
            Frm1.vspdData.Col = C_PAY_TYPE_NM
            iDx = Frm1.vspdData.value
            Frm1.vspdData.Row = Row
   	        Frm1.vspdData.Col = C_PAY_TYPE
            Frm1.vspdData.value = iDx
        Case C_CLOSE_TYPE_NM
            Frm1.vspdData.Row = Row
            Frm1.vspdData.Col = C_CLOSE_TYPE_NM
            iDx = Frm1.vspdData.value
            Frm1.vspdData.Row = Row
   	        Frm1.vspdData.Col = C_CLOSE_TYPE
            Frm1.vspdData.value = iDx
    End Select    
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
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

    Dim iDx
    Dim tmpDT
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1.vspdData
		If Col = C_CLOSE_DT Then
			.Row = Row
			.Col = Col

            tmpDT = Replace( parent.gDateFormatYYYYMM,"YYYY","1900")
            tmpDT = Replace(tmpDT            ,"YY"  ,"00")
            tmpDT = Replace(tmpDT            ,"MM"  ,"01")

	        Call  CompareDateByFormat(tmpDT, .Text, tmpDT, "입력일", 970023,  parent.gDateFormatYYYYMM,  parent.gComDateType, True)	        
		End If
    End With

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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
    Call InitComboBox2
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split( parent.gDateFormat, parent.gComDateType)
	
	If  parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType =  parent.gComDateType
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>시스템마감입력</font></td>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
						        <TD>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>급여시스템 마감입력</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
              						        <TD CLASS="TD5" NOWRAP>급여마감구분</TD>
	                   						<TD CLASS="TD6"><SELECT NAME="txtClose_type1" ALT="급여마감구분" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
              						        <TD CLASS="TD5" NOWRAP>급여마감월</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h1019ma1_fpDateTime2_txtClose_dt1.js'></script></TD>
											<TD CLASS="TD6"><INPUT TYPE=text NAME="txtFlgMode1" STYLE="WIDTH: 0px"    TAG="24"></TD>
							            </TR>
                                        <TR>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD6"></TD>
                                        </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
						    <TR>
						        <TD>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>연월차시스템 마감입력</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS=TD5 NOWRAP>연월차마감구분</TD>
							            	<TD CLASS="TD6"><SELECT NAME="txtClose_type2" ALT="연월차마감구분" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
							            	<TD CLASS=TD5 NOWRAP>연월차마감월</TD>
							            	<TD CLASS="TD6"><script language =javascript src='./js/h1019ma1_fpDateTime2_txtClose_dt2.js'></script></TD>
											<TD CLASS="TD6"><INPUT TYPE=HIDDEN NAME="txtFlgMode2" STYLE="WIDTH: 0px"    TAG="24"></TD>
							            </TR>
                                        <TR>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD6"></TD>
                                        </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
						    <TR>
						        <TD>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>연말정산 마감입력</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
              						        <TD CLASS="TD5" NOWRAP>연말정산마감구분</TD>
	                   						<TD CLASS="TD6"><SELECT NAME="txtClose_type3" ALT="연말정산마감구분" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
              						        <TD CLASS="TD5" NOWRAP>연말정산마감년도</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h1019ma1_fpDateTime2_txtClose_dt3.js'></script></TD>
											<TD CLASS="TD6"><INPUT TYPE=HIDDEN NAME="txtFlgMode3" STYLE="WIDTH: 0px"    TAG="24"></TD>
							            </TR>
                                        <TR>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD6"></TD>
                                        </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
						    <TR>
						        <TD>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>근태시스템 마감입력</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							            <TR>
							            	<TD CLASS="TD5" NOWRAP>근태마감구분</TD>
							            	<TD CLASS="TD6"><SELECT NAME="txtClose_type4" ALT="근태마감구분" STYLE="WIDTH: 150px" TAG="22"></SELECT></TD>
							            	<TD CLASS="TD5" NOWRAP>근태마감일</TD>
							            	<TD CLASS="TD6"><script language =javascript src='./js/h1019ma1_fpDateTime2_txtClose_dt4.js'></script></TD>
											<TD CLASS="TD6"><INPUT TYPE=HIDDEN NAME="txtFlgMode4" STYLE="WIDTH: 0px"    TAG="24"></TD>
							            </TR>
                                        <TR>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD5"></TD>
                                            <TD CLASS="TD6"></TD>
                                            <TD CLASS="TD6"></TD>
                                        </TR>
							        </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
				            <TR>
				            	<TD WIDTH=100% HEIGHT=100% valign=top>
				            		<TABLE <%=LR_SPACE_TYPE_20%>>
				            		<TR>
				            			<TD HEIGHT="100%">
				            				<script language =javascript src='./js/h1019ma1_vaSpread1_vspdData.js'></script>
				            			</TD>
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
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
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

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</D
