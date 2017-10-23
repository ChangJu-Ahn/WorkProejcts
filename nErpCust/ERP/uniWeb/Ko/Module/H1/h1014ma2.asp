<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 연차분할계산표 
*  3. Program ID           : H1014ma2
*  4. Program Name         : H1014ma2
*  5. Program Desc         : 기준정보관리/연차분할계산표 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30
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

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H1014mb2.asp"                                      'Biz Logic ASP 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_STRT_YY
Dim C_STRT_YY_NM
Dim C_STRT_MM
Dim C_STRT_DD
Dim C_BAR	
Dim C_END_YY
Dim C_END_YY_NM
Dim C_END_MM
Dim C_END_DD
Dim C_YEAR_CNT10
Dim C_YEAR_CNT8
Dim C_YEAR_RETR

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_STRT_YY		= 1	
	 C_STRT_YY_NM	= 2
	 C_STRT_MM		= 3
	 C_STRT_DD		= 4
	 C_BAR			= 5
	 C_END_YY		= 6
	 C_END_YY_NM	= 7
	 C_END_MM		= 8
	 C_END_DD		= 9
	 C_YEAR_CNT10	= 10
	 C_YEAR_CNT8	= 11
	 C_YEAR_RETR	= 12	  
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
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
Sub MakeKeyStream(pRow)
    lgKeyStream  = ""
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0098", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_STRT_YY
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_STRT_YY_NM

     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_END_YY
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_END_YY_NM

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
			.Col = C_STRT_YY
			intIndex = .Value
			.col = C_STRT_YY_NM
			.Value = intindex

			.Row = intRow
			.Col = C_END_YY
			intIndex = .Value
			.col = C_END_YY_NM
			.Value = intindex
		Next	

	End With
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
        .MaxCols = C_YEAR_RETR + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	
       Call  AppendNumberPlace("6","2","0")
       Call  GetSpreadColumnPos("A")

         ggoSpread.SSSetCombo    C_STRT_YY,     "입사년",  17
		 ggoSpread.SSSetCombo    C_STRT_YY_NM,  "입사년",  17
		 ggoSpread.SSSetMask     C_STRT_MM,     "월",      07,2,"99"
		 ggoSpread.SSSetMask     C_STRT_DD,     "일",      07,2,"99"
		 ggoSpread.SSSetEdit     C_BAR,         "",            3,2
         ggoSpread.SSSetCombo    C_END_YY,      "입사년",  17		
		 ggoSpread.SSSetCombo    C_END_YY_NM,   "입사년",  17
		 ggoSpread.SSSetMask     C_END_MM,      "월",      07,2,"99"
		 ggoSpread.SSSetMask     C_END_DD,      "일",      07,2,"99"
		 ggoSpread.SSSetMask     C_YEAR_CNT10,  "만근연차",16,2,"99"
		 ggoSpread.SSSetMask     C_YEAR_CNT8,   "8할연차", 16,2,"99"
         ggoSpread.SSSetMask     C_YEAR_RETR,   "퇴사연차",16,2,"99"
         
         Call ggoSpread.SSSetColHidden(C_STRT_YY, C_STRT_YY, True)
         Call ggoSpread.SSSetColHidden(C_END_YY , C_END_YY,  True)
	
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

     ggoSpread.SpreadLock		C_STRT_YY	 , -1, C_STRT_YY
     ggoSpread.SpreadLock		C_STRT_YY_NM , -1, C_STRT_YY_NM
     ggoSpread.SpreadLock		C_STRT_MM	 , -1, C_STRT_MM
     ggoSpread.SpreadLock		C_STRT_DD	 , -1, C_STRT_DD
     ggoSpread.SpreadLock		C_BAR		 , -1, C_BAR
     ggoSpread.SSSetRequired    C_END_YY	 , -1, C_END_YY
     ggoSpread.SSSetRequired     C_END_YY_NM  , -1, C_END_YY_NM
     ggoSpread.SSSetRequired     C_END_MM	 , -1, C_END_MM
     ggoSpread.SSSetRequired     C_END_DD	 , -1, C_END_DD
     ggoSpread.SSSetRequired     C_YEAR_CNT10 , -1, C_YEAR_CNT10
     ggoSpread.SSSetRequired     C_YEAR_CNT8  , -1, C_YEAR_CNT8
     ggoSpread.SSSetRequired     C_YEAR_RETR  , -1, C_YEAR_RETR
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
    
    .vspdData.ReDraw = False
    
         ggoSpread.SSSetProtected   C_STRT_YY	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_STRT_YY_NM , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_STRT_MM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_STRT_DD	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected   C_BAR		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_END_YY		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_END_YY_NM  , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_END_DD		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_END_MM		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_YEAR_CNT10 , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_YEAR_CNT8  , pvStartRow, pvEndRow
         ggoSpread.SSSetRequired    C_YEAR_RETR  , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	.vspdData.MaxCols, pvStartRow, pvEndRow
        
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
            
            C_STRT_YY		= iCurColumnPos(1)	
			C_STRT_YY_NM	= iCurColumnPos(2)
			C_STRT_MM		= iCurColumnPos(3)
			C_STRT_DD		= iCurColumnPos(4)
			C_BAR			= iCurColumnPos(5)
			C_END_YY		= iCurColumnPos(6)
			C_END_YY_NM		= iCurColumnPos(7)
			C_END_MM		= iCurColumnPos(8)
			C_END_DD		= iCurColumnPos(9)
			C_YEAR_CNT10	= iCurColumnPos(10)
			C_YEAR_CNT8		= iCurColumnPos(11)
			C_YEAR_RETR		= iCurColumnPos(12)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call InitComboBox
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End if
       
    FncQuery = True                                                              '☜: Processing is OK
    
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
    Dim intStrt
    Dim intEnd
    Dim lRow

	Dim strStrt, strEnd
	Dim intstrt_yy
	Dim intend_yy
	Dim intstrt_mm
	Dim intstrt_dd
	Dim intend_mm
	Dim intend_dd

	Dim intYear_cnt10
	Dim intYear_cnt8
	Dim intYear_retr
        
	Dim strWhere
	
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
            if  .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
                .vspdData.Col = C_STRT_YY
                intstrt_yy = .vspdData.Text

                .vspdData.Col = C_STRT_MM
                intstrt_mm = .vspdData.Text
                
                if  IsZeroNumeric(.vspdData.Text) = false OR .vspdData.Text > "12" then
                    call  DisplayMsgBox("970027","x","입사월","x")
                    .vspdData.Row = lRow
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if
                
                .vspdData.Col = C_STRT_DD
                intstrt_dd = .vspdData.Text
                if  IsZeroNumeric(.vspdData.Text) = false OR .vspdData.Text > "31" then
                    call  DisplayMsgBox("970027","x","입사일","x")
                    .vspdData.Row = lRow
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if

                .vspdData.Col = C_END_YY
                intend_yy = .vspdData.Text
                .vspdData.Col = C_END_MM
                intend_mm = .vspdData.Text
                if  IsZeroNumeric(.vspdData.Text) = false OR .vspdData.Text > "12" then
                    call  DisplayMsgBox("970027","x","입사월","x")
                    .vspdData.Row = lRow
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if

                .vspdData.Col = C_END_DD
                intend_dd = .vspdData.Text
                if  IsZeroNumeric(.vspdData.Text) = false OR .vspdData.Text > "31" then
                    call  DisplayMsgBox("970027","x","입사일","x")
                    .vspdData.Row = lRow
                    .vspdData.Action = 0 ' go to 
                    exit function
                    
                end if

                if  intend_yy <> "" then
                    if  Cint(intend_yy) < Cint(intstrt_yy) then
                        call  DisplayMsgBox("970027","x","입사년","x")
                        .vspdData.Col = C_END_YY
                        .vspdData.Row = lRow
                        .vspdData.Action = 0 ' go to 
                        exit function
                    elseif Cint(intend_yy) = Cint(intstrt_yy) then
                        if  Cint(intend_mm) < Cint(intstrt_mm) then
                            call  DisplayMsgBox("970027","x","월","x")
                            .vspdData.Col = C_END_MM
                            .vspdData.Row = lRow
                            .vspdData.Action = 0 ' go to 
                            exit function
                        else
							if  (Cint(intend_mm) = Cint(intstrt_mm) AND Cint(intend_dd) <= Cint(intstrt_dd)) then
                                call  DisplayMsgBox("970027","x","일","x")
                                .vspdData.Col = C_END_DD
                                .vspdData.Row = lRow
                                .vspdData.Action = 0 ' go to 
                                exit function
                            end if
                        end if
                    end if
                end if
				
				
                .vspdData.Col = C_YEAR_CNT10
                intYear_cnt10 = .vspdData.Text
				if isZeroNumeric(intYear_cnt10) = false then
					call  DisplayMsgBox("970027","x","만근연차","x")
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if
				
                .vspdData.Col = C_YEAR_CNT8
                intYear_cnt8 = .vspdData.Text
                if isZeroNumeric(intYear_cnt8) = false then
					call  DisplayMsgBox("970027","x","8할연차","x")
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if

                .vspdData.Col = C_YEAR_RETR
                intYear_retr = .vspdData.Text
                if isZeroNumeric(intYear_retr) = false then
					call  DisplayMsgBox("970027","x","퇴사연차","x")
                    .vspdData.Action = 0 ' go to 
                    exit function
                end if

 '               if Cint(intYear_cnt10) = 0 and Cint(intYear_cnt8) = 0 and Cint(intYear_retr) = 0 then
'                    call  DisplayMsgBox("970027","x","만근년차/8할연차/퇴사연차","x")
 '                   .vspdData.Col = C_YEAR_CNT10
 '                   .vspdData.Row = lRow
 '                   .vspdData.Action = 0 ' go to 
  '                  exit function
   '             end if

            end if
       Next
	

	End With

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
           .Col  = C_STRT_YY
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_STRT_YY_NM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_STRT_MM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_STRT_DD
           .Row  = .ActiveRow
           .Text = ""

           .Col  = C_END_YY
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_END_YY_NM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_END_MM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_END_DD
           .Row  = .ActiveRow
           .Text = ""
    End With

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo 
    call initdata() 
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow
	Dim iRow
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
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1  
			.vspdData.Row = iRow
			.vspdData.Col = C_YEAR_CNT10
			.vspdData.Text = 0
			.vspdData.Col = C_YEAR_CNT8
			.vspdData.Text = 0
			.vspdData.Col = C_YEAR_RETR
			.vspdData.Text = 0       
       Next       
              
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function


'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
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
	
    DbSave = False                                                          
    
    If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case  ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                   .vspdData.Col = C_STRT_YY    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_STRT_MM	: strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_STRT_DD	: strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_END_YY  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_END_MM     : strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_END_DD     : strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_YEAR_CNT10 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_YEAR_CNT8  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_YEAR_RETR	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                  strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                   .vspdData.Col = C_STRT_YY    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_STRT_MM	: strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_STRT_DD	: strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_END_YY  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_END_MM     : strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_END_DD     : strVal = strVal & Right("0" & Trim(.vspdData.Text), 2) & parent.gColSep
                   .vspdData.Col = C_YEAR_CNT10 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_YEAR_CNT8  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_YEAR_RETR	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                  strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                   .vspdData.Col = C_STRT_YY    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_STRT_MM	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_STRT_DD	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
 
    DbDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    DbDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("110011110011111")									
    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData								'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    
    Call InitVariables															'⊙: Initializes local global variables
	call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'========================================================================================================
'	Name : OpenDept()
'	Description : 부서POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept.value			<%' 조건부에서 누른 경우 Code Condition%>
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			<%' Grid에서 누른 경우 Code Condition%>
	End If
	arrParam(1) = ""								<%' Name Cindition%>
	arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.action = 0		
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'========================================================================================================
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'========================================================================================================
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtDept.value = arrRet(0)
			.txtDeptNm.value = arrRet(1)
			.txtDept.focus
		Else 'spread
			.vspdData.Col = C_DEPT_CD_NM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(0)
		    .vspdData.action = 0						
		End If
	End With

End Function


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		 ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_STRT_YY_NM
            iDx = frm1.vspdData.value
            Frm1.vspdData.Col = C_STRT_YY
            frm1.vspdData.value =iDx 
         Case  C_END_YY_NM
            iDx = frm1.vspdData.value
            Frm1.vspdData.Col = C_END_YY
            frm1.vspdData.value =iDx 
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
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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

function IsZeroNumeric(pNo)
	DIM strNO
	strNo = CSTR(right("0" & Trim(pNO), 2))
		if isNumeric(left(strNo,1)) = true and isNumeric(right(strNo,1)) = true then
			IsZeroNumeric = true
		else
			IsZeroNumeric = false
	end if
end function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연차분할계산표</font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h1014ma2_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

