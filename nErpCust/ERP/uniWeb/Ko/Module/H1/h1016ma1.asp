<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single Sample
*  3. Program ID           : H1016ma1
*  4. Program Name         : H1016ma1
*  5. Program Desc         : 기준정보관리/기초금액계산식등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/08
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H1016mb1.asp"                                      'Biz Logic ASP 
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

Dim gSelframeFlg			   ' 현재 TAB의 위치를 나타내는 Flag
Dim gCounts
Dim isFirst   '첫화면이 열리는지 여부 
Dim lgStrPrevKey1

Dim C_BAS_AMT_TYPE 
Dim C_BAS_AMT_TYPE_POP
Dim C_BAS_AMT_TYPE_NM
Dim C_DILIG_CD
Dim C_DILIG_CD_POP
Dim C_DILIG_CD_NM
Dim C_TEXT1
Dim C_TEXT11
Dim C_COMPUTE4_N
Dim C_TEXT12
Dim C_TEXT13
Dim C_TEXT14
Dim C_TEXT2
Dim C_COMPUTE1_N
Dim C_TEXT3
Dim C_COMPUTE2_N
Dim C_TEXT4
Dim C_COMPUTE3_N
Dim C_TEXT5
Dim C_STD_AMT

Dim C_ALLOW_CD_NM
Dim C_ALLOW_CD

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
         C_BAS_AMT_TYPE		= 1
		 C_BAS_AMT_TYPE_POP = 2
		 C_BAS_AMT_TYPE_NM	= 3
		 C_DILIG_CD			= 4
		 C_DILIG_CD_POP		= 5
		 C_DILIG_CD_NM		= 6
		 C_TEXT1			= 7
		 C_TEXT11			= 8
		 C_COMPUTE4_N		= 9
		 C_TEXT12			= 10
		 C_TEXT13			= 11
		 C_TEXT14			= 12
		 C_TEXT2			= 13
		 C_COMPUTE1_N		= 14
		 C_TEXT3			= 15
		 C_COMPUTE2_N		= 16
		 C_TEXT4			= 17
		 C_COMPUTE3_N		= 18
		 C_TEXT5			= 19
		 C_STD_AMT			= 20
        
    ElseIf pvSpdNo = "B" Then
         C_ALLOW_CD_NM		= 1
		 C_ALLOW_CD			= 2
    End If
    
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
    lgStrPrevKey1	  = ""                                      '⊙: initializes Previous Key Index
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
    If lgCurrentSpd = "M" Then
       lgKeyStream = Frm1.txtocpt_type.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
       lgKeyStream = lgKeyStream & Frm1.txtbas_amt_type.Value & parent.gColSep
    Else
    	frm1.vspdData.Row = pRow
		frm1.vspdData.Col = C_BAS_AMT_TYPE
        lgKeyStream = frm1.vspdData.Text & parent.gColSep     'You Must append one character( parent.gColSep)
        lgKeyStream = lgKeyStream & Frm1.txtocpt_type.Value & parent.gColSep
    End If 
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

	' 급여구분 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtocpt_type, iCodeArr, iNameArr, Chr(11))    
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox2()
    Dim iCodeArr 
    Dim iNameArr
    ' 수당코드 
    Call  CommonQueryRs(" ALLOW_CD,ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ALLOW_CD_NM
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ALLOW_CD
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	     ggoSpread.Source = frm1.vspdData1
	    With frm1.vspdData1
	    	For intRow = 1 To .MaxRows			
	    		.Row = intRow

	    		.Col = C_ALLOW_CD         ' 수당코드 
	    		intIndex = .value
	    		.col = C_ALLOW_CD_NM
	    		.value = intindex

	    	Next	
	    End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	If pvSpdNo = "" OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")	
		With frm1.vspdData
 
			    ggoSpread.Source = frm1.vspdData	
			    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

			    .ReDraw = false

			    .MaxCols = C_STD_AMT + 1                                                <%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			    .ColHidden = True

			    .MaxRows = 0
			    ggoSpread.ClearSpreadData
	
		Call AppendNumberPlace("6","6","2")
		Call GetSpreadColumnPos("A")     

		 ggoSpread.SSSetEdit    C_BAS_AMT_TYPE,      "기초금액코드",11,,,5
		 ggoSpread.SSSetButton  C_BAS_AMT_TYPE_POP
		 ggoSpread.SSSetEdit    C_BAS_AMT_TYPE_NM,   "기초금액",12,,,20,2
		 ggoSpread.SSSetEdit    C_DILIG_CD,          "관련근태코드",11
		 ggoSpread.SSSetButton  C_DILIG_CD_POP
		 ggoSpread.SSSetEdit    C_DILIG_CD_NM,       "관련근태",12

		 ggoSpread.SSSetEdit	C_TEXT1,             "", 3, 2
		 ggoSpread.SSSetEdit	C_TEXT11,            "", 2, 2
		 ggoSpread.SSSetFloat   C_COMPUTE4_N,        "기본급",  8,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		 ggoSpread.SSSetEdit	C_TEXT12,            "", 2, 2
		 ggoSpread.SSSetEdit	C_TEXT13,            "", 10, 2
		 ggoSpread.SSSetEdit	C_TEXT14,            "", 2, 2
		 ggoSpread.SSSetEdit	C_TEXT2,             "", 2, 2
		 ggoSpread.SSSetFloat   C_COMPUTE1_N,        "",        8,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		 ggoSpread.SSSetEdit	C_TEXT3,             "", 2, 2
		 ggoSpread.SSSetFloat   C_COMPUTE2_N,        "",        8,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		 ggoSpread.SSSetEdit	C_TEXT4,             "", 2, 2
		 ggoSpread.SSSetFloat   C_COMPUTE3_N,        "",        8,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		 ggoSpread.SSSetEdit	C_TEXT5,             "", 3, 2
		 ggoSpread.SSSetFloat   C_STD_AMT,           "정액",10, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		 
		 Call ggoSpread.MakePairsColumn(C_BAS_AMT_TYPE,  C_BAS_AMT_TYPE_POP)
		 Call ggoSpread.MakePairsColumn(C_DILIG_CD	 ,  C_DILIG_CD_POP)
		                           
		.ReDraw = true

		Call SetSpreadLock 
    
		End With
    
	End if
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then		

		Call initSpreadPosVariables("B")	
		With frm1.vspdData1
 
		    ggoSpread.Source = frm1.vspdData1	
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

		    .ReDraw = false
		    .MaxCols = C_ALLOW_CD + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True
		    
		    .MaxRows = 0
	
		 Call AppendNumberPlace("6","2","0")
		 Call GetSpreadColumnPos("B")

		 ggoSpread.SSSetCombo    C_ALLOW_CD_NM,   "수당코드명", 15
		 ggoSpread.SSSetCombo    C_ALLOW_CD,      "수당코드",    5
		 
		 Call ggoSpread.SSSetColHidden(C_ALLOW_CD,  C_ALLOW_CD, True)

		.ReDraw = true
	
		Call SetSpreadLock1 
    
		End With
    End if
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData

        .ReDraw = False
        	 ggoSpread.SpreadLock C_BAS_AMT_TYPE		, -1, C_BAS_AMT_TYPE
             ggoSpread.SpreadLock C_BAS_AMT_TYPE_POP	, -1, C_BAS_AMT_TYPE_POP
             ggoSpread.SpreadLock C_BAS_AMT_TYPE_NM		, -1, C_BAS_AMT_TYPE_NM           
             ggoSpread.SpreadLock C_DILIG_CD_NM			, -1, C_DILIG_CD_NM
             ggoSpread.SpreadLock C_TEXT1				, -1, C_TEXT1
             ggoSpread.SpreadLock C_TEXT11				, -1, C_TEXT11
             ggoSpread.SpreadLock C_TEXT12				, -1, C_TEXT12
             ggoSpread.SpreadLock C_TEXT13				, -1, C_TEXT13
             ggoSpread.SpreadLock C_TEXT14				, -1, C_TEXT14
             ggoSpread.SpreadLock C_TEXT2				, -1, C_TEXT2
             ggoSpread.SpreadLock C_TEXT3				, -1, C_TEXT3
             ggoSpread.SpreadLock C_TEXT4				, -1, C_TEXT4
             ggoSpread.SpreadLock C_TEXT5				, -1, C_TEXT5
             ggoSpread.SSSetProtected					.MaxCols,-1,-1
        .ReDraw = True

    End With

End Sub

Sub SetSpreadLock1()

	With frm1.vspdData1
	
        ggoSpread.Source = frm1.vspdData1

        .ReDraw = False
        	ggoSpread.SpreadLock C_ALLOW_CD			, -1, C_ALLOW_CD
            ggoSpread.SpreadLock C_ALLOW_CD_NM		, -1, C_ALLOW_CD_NM
            ggoSpread.SSSetProtected				.MaxCols,-1,-1
        .ReDraw = True

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
    
        .ReDraw = False
             ggoSpread.SSSetRequired	C_BAS_AMT_TYPE		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_BAS_AMT_TYPE_NM	, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_DILIG_CD_NM		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT1				, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT11			, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT12			, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT13			, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT14			, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT2				, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT3				, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT4				, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_TEXT5				, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected						.MaxCols,-1,-1
        .ReDraw = True
    
    End With

End Sub

Sub SetSpreadColor1(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData1
	
        ggoSpread.Source = frm1.vspdData1
    
        .ReDraw = False
             ggoSpread.SSSetRequired		C_ALLOW_CD		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired		C_ALLOW_CD_NM	, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected						.MaxCols, pvStartRow, pvEndRow
        .ReDraw = True
    
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
			
			C_BAS_AMT_TYPE		= iCurColumnPos(1)
			C_BAS_AMT_TYPE_POP	= iCurColumnPos(2)
			C_BAS_AMT_TYPE_NM	= iCurColumnPos(3)
			C_DILIG_CD			= iCurColumnPos(4)
			C_DILIG_CD_POP		= iCurColumnPos(5)
			C_DILIG_CD_NM		= iCurColumnPos(6)
			C_TEXT1				= iCurColumnPos(7)
			C_TEXT11			= iCurColumnPos(8)
			C_COMPUTE4_N		= iCurColumnPos(9)
			C_TEXT12			= iCurColumnPos(10)
			C_TEXT13			= iCurColumnPos(11)
			C_TEXT14			= iCurColumnPos(12)
			C_TEXT2				= iCurColumnPos(13)
			C_COMPUTE1_N		= iCurColumnPos(14)
			C_TEXT3				= iCurColumnPos(15)
			C_COMPUTE2_N		= iCurColumnPos(16)
			C_TEXT4				= iCurColumnPos(17)
			C_COMPUTE3_N		= iCurColumnPos(18)
			C_TEXT5				= iCurColumnPos(19)
			C_STD_AMT			= iCurColumnPos(20)            
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ALLOW_CD_NM		= iCurColumnPos(1)
			C_ALLOW_CD			= iCurColumnPos(2)
            
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
    
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox
    Call InitComboBox2
    Call SetToolbar("1100110100111111")										        '버튼 툴바 제어 
    gCounts = 0
    isFirst = true
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
    Dim ChgOK
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
     
    ChgOK = false
     
    ggoSpread.Source = Frm1.vspdData1

    If  ggoSpread.SSCheckChange = True Then
		ChgOK = True
    End If

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		ChgOK = True    
    End If
	
	If  ChgOK Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")		'☜: Data is changed.  Do you want to display it? 
			
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If    
    Call  ggoOper.ClearField(Document, "2")
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	call txtBas_amt_type_Onchange()

    Call InitVariables                                                           '⊙: Initializes local global variables
    lgCurrentSpd = "M"
   Call MakeKeyStream("X")  

    gCounts = 0  
    isFirst = true

    lgCurrentSpd = "M"  ' Master
    
	Call  DisableToolBar( parent.TBC_QUERY)

	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

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
    dim lRow
    DIM strCD, strNm
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear    
	
	frm1.ChgSave1.value = "F"
	frm1.ChgSave2.value = "F"

    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		frm1.ChgSave1.value = "T"
    End If

    ggoSpread.Source = Frm1.vspdData1
    If  ggoSpread.SSCheckChange = True Then
		frm1.ChgSave2.value = "T"
    End If
	
	If frm1.ChgSave1.value = "F" and frm1.ChgSave2.value="F" Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
		Exit Function
	End If    
    
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
	End If

    ggoSpread.Source = frm1.vspdData1
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
	End If

    ggoSpread.Source = frm1.vspdData
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				.vspdData.Col = C_BAS_AMT_TYPE_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","기초금액코드","X")
					.vspddata.Action = 0
					
       	            exit function
				 end if 
				
                .vspdData.Col = C_DILIG_CD
                strCD = .vspddata.text
                
                .vspdData.Col = C_DILIG_CD_NM
                strNm = .vspddata.text
                
                if (Trim(strCD) <> "" AND Trim(strNM) = "") then
                    Call  DisplayMsgBox("970000","X","근태코드","X")
					.vspddata.Action = 0
					exit function
                end if
                
            end if
        next

    end with
	lgCurrentSpd = "M"    
    Call  DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    FncSave = True                                                              '☜: Processing is OK
    
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
    
	        With Frm1.vspdData
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData	
	        		ggoSpread.CopyRow
                    SetSpreadColor .ActiveRow, .ActiveRow

                    .Col  = C_BAS_AMT_TYPE
					.Row  = .ActiveRow
					.Text = ""

					.Col  = C_BAS_AMT_TYPE_NM
					.Row  = .ActiveRow
					.Text = ""
    
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With
			ggoSpread.Source = Frm1.vspdData1	
			ggoSpread.ClearSpreadData	        
        Case  Else

            If Frm1.vspdData1.MaxRows < 1 Then
                Exit Function
            End If
    
	        With Frm1.vspdData1
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData1
	        		ggoSpread.CopyRow
                    SetSpreadColor1 .ActiveRow, .ActiveRow

                    .Col  = C_ALLOW_CD
					.Row  = .ActiveRow
					.Text = ""

					.Col  = C_ALLOW_CD_NM
					.Row  = .ActiveRow
					.Text = ""
    
	        		.ReDraw = True
	        		.focus
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
    if lgCurrentSpd = "M" then
         ggoSpread.Source = Frm1.vspdData	
         ggoSpread.EditUndo  
    else
         ggoSpread.Source = Frm1.vspdData1
         ggoSpread.EditUndo  
    end if
    Call initdata()
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

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
      
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.ClearSpreadData
			ggoSpread.Source = frm1.vspdData
                  With Frm1
                         .vspdData.ReDraw = False
                         .vspdData.Focus
                          ggoSpread.Source = .vspdData
                          ggoSpread.InsertRow .vspdData.ActiveRow, imRow
                          SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
                          
                          For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1 
							.vspdData.Row = iRow
							.vspdData.Col = C_TEXT1
							.vspdData.value = "(1)"
							.vspdData.Col = C_TEXT11
							.vspdData.value = "("

							.vspdData.Col = C_TEXT12
							.vspdData.value = "+"
							.vspdData.Col = C_TEXT13
							.vspdData.value = "수당합계"
							.vspdData.Col = C_TEXT14
							.vspdData.value = ")"

							.vspdData.Col = C_TEXT2
							.vspdData.value = "*"
							.vspdData.Col = C_TEXT3
							.vspdData.value = "/"
							.vspdData.Col = C_TEXT4
							.vspdData.value = "*"
							.vspdData.Col = C_TEXT5
							.vspdData.value = "(2)"
                          Next
                          
                         .vspdData.ReDraw = True
                  End With
        Case  Else
                  With Frm1
                         .vspdData1.ReDraw = False
                         .vspdData1.Focus
                          ggoSpread.Source = .vspdData1
                          ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
                          SetSpreadColor1 .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
                         .vspdData1.ReDraw = True
                  End With
    End Select 

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

    if  lgCurrentSpd = "M" then
        If Frm1.vspdData.MaxRows < 1 then
           Exit function
	    End if	
        With Frm1.vspdData 
        	.focus
        	 ggoSpread.Source = frm1.vspdData 
        	lDelRows =  ggoSpread.DeleteRow
        End With
    ELSE
        If Frm1.vspdData1.MaxRows < 1 then
           Exit function
	    End if	
        With Frm1.vspdData1 
        	.focus
        	 ggoSpread.Source = frm1.vspdData1 
        	lDelRows =  ggoSpread.DeleteRow
        End With
    END IF
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
			Call InitComboBox2     		
	End Select     
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
        strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With


	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    DbQuery = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================

Function DbQuery2() 

    DbQuery2 = False
    
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData1.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey1                 '☜: Next key tag
    End With


	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery2 = True
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

    if frm1.ChgSave1.value = "T"  then
         ggoSpread.Source = frm1.vspdData 
	    With Frm1
           For lRow = 1 To .vspdData.MaxRows
               .vspdData.Row = lRow
               .vspdData.Col = 0
               Select Case .vspdData.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Create
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                                                          strVal = strVal & .txtOcpt_type.value & parent.gColSep
                        .vspdData.Col = C_BAS_AMT_TYPE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE4_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE1_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE2_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE3_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_STD_AMT       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                                                          strVal = strVal & .txtOcpt_type.value & parent.gColSep
                        .vspdData.Col = C_BAS_AMT_TYPE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE4_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE1_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE2_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_COMPUTE3_N    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_STD_AMT       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & lgCurrentSpd & parent.gColSep
                                                          strDel = strDel & .txtOcpt_type.value & parent.gColSep
                        .vspdData.Col = C_BAS_AMT_TYPE  : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
	       .txtMaxRows.value     = lGrpCnt-1	
	       .txtSpread.value      = strDel & strVal
	    End With
    elseif  frm1.ChgSave2.value = "T" then
         ggoSpread.Source = frm1.vspdData1
	    With Frm1
           For lRow = 1 To .vspdData1.MaxRows
               .vspdData1.Row = lRow
               .vspdData1.Col = 0
               Select Case .vspdData1.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Create
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lgCurrentSpd & parent.gColSep
                                                          strVal = strVal & .txtOcpt_type.value & parent.gColSep
															
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_BAS_AMT_TYPE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData1.Col = C_ALLOW_CD 	: strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                   
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & lgCurrentSpd & parent.gColSep
                                                          strDel = strDel & .txtOcpt_type.value & parent.gColSep
                                                          
                        frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_BAS_AMT_TYPE  : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData1.Col = C_ALLOW_CD 	: strDel = strDel & Trim(.vspdData1.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
	       .txtMaxRows.value     = lGrpCnt-1	
	       .txtSpread.value      = strDel & strVal
	    End With
    end if	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
	
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()

    Call SetToolbar("1100111100111111")

	if lgStrPrevKey1 <> "" and isFirst = false then
		exit function
	end if

	if lgStrPrevKey1 <> "" or isFirst = true then
		isFirst = false		' 첫화면이 열리고나서 오른쪽 그리드 세팅하기 위해 
		Call DisableToolBar(parent.TBC_QUERY)
		call vspdData_click(1,frm1.vspdData.activerow)
	end if		
	frm1.vspdData.focus
End Function

Function DbQueryOk2()
    lgIntFlgMode =  parent.OPMD_UMODE    

	Call InitData()
    Call SetToolbar("1100111100111111")
	
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
'	frm1.vspdData1.focus    
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")
   
    Call InitVariables															'⊙: Initializes local global variables
    lgCurrentSpd = "M"
    Call MakeKeyStream("X")    
   
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
    
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
' Name : OpenCd()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCd(iwhere,row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True

    arrParam(0) = "기초금액코드 팝업"	' 팝업 명칭 
	arrParam(1) = "HDA010T"				 	' TABLE 명칭 
	arrParam(3) = ""	                    ' Name Cindition
    arrParam(4) = "pay_cd=" & FilterVar("*", "''", "S") & " "                        ' Where Condition

    arrField(0) = "allow_CD"				' Field명(0)
    arrField(1) = "allow_NM"				' Field명(1)

	arrParam(5) = "기초금액코드"

    arrHeader(0) = "기초금액코드"		' Header명(0)
    arrHeader(1) = "기초금액명"			' Header명(1)

    Select Case iwhere
        case 0
	        arrParam(2) = frm1.txtBas_amt_type.value	' Code Condition
        case 1
        	frm1.vspdData.Row = row
            frm1.vspdData.Col = C_BAS_AMT_TYPE     
        
            arrParam(2) = frm1.vspdData.Text	                    ' Code Condition
        case 2
        	frm1.vspdData.Row = row
		    frm1.vspdData.Col = C_DILIG_CD
        
            arrParam(0) = "근태코드 팝업"	    ' 팝업 명칭 
        	arrParam(1) = "HCA010T"				 	' TABLE 명칭 
            arrParam(2) = frm1.vspdData.Text        ' Code Condition
        	arrParam(3) = ""	                    ' Name Cindition
        	arrParam(4) = ""                        ' Where Condition
            arrField(0) = "DILIG_CD"				' Field명(0)
            arrField(1) = "DILIG_NM"				' Field명(1)

	        arrParam(5) = "근태코드"

            arrHeader(0) = "근태코드"		    ' Header명(0)
            arrHeader(1) = "근태코드명"			' Header명(1)

    end select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
        Select Case iwhere
            case 0
		        Frm1.txtBas_amt_type.focus
		    case 1
		        Frm1.vspdData.Col = C_BAS_AMT_TYPE
				Frm1.vspdData.action =0
		    case 2				
		        Frm1.vspdData.Col = C_DILIG_CD
				Frm1.vspdData.action =0
		end select	
		Exit Function
	Else
		Call SubSetCd(arrRet, iwhere,row)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCd()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCd(arrRet, iwhere,row)
	With Frm1
        Select Case iwhere
            case 0
		        .txtBas_amt_type.value = arrRet(0)
		        .txtBas_amt_type_nm.value = arrRet(1)
		        .txtBas_amt_type.focus
		    case 1
		        .vspdData.Col = C_BAS_AMT_TYPE_NM
		        .vspdData.Text = arrRet(1)
		        .vspdData.Col = C_BAS_AMT_TYPE
		        .vspdData.Text = arrRet(0)
				.vspdData.action =0
		         ggoSpread.Source = frm1.vspdData
                 ggoSpread.UpdateRow Row
		    case 2
		        .vspdData.Col = C_DILIG_CD_NM
		        .vspdData.Text = arrRet(1)
		        .vspdData.Col = C_DILIG_CD
		        .vspdData.Text = arrRet(0)
				.vspdData.action =0
		         ggoSpread.Source = frm1.vspdData
                 ggoSpread.UpdateRow Row
		end select
	End With


End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_BAS_AMT_TYPE
            If Trim(Frm1.vspdData.Text) = "" Then
  	            Frm1.vspdData.Col = C_BAS_AMT_TYPE_NM
                Frm1.vspdData.Text = ""
            Else            
                IntRetCd =  CommonQueryRs(" allow_nm "," HDA010T "," pay_cd=" & FilterVar("*", "''", "S") & "  and allow_cd =  " & FilterVar(Frm1.vspdData.Text, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                If IntRetCd = false then
			        Call  DisplayMsgBox("970000","X","기초금액코드","X")
  	                Frm1.vspdData.Col = C_BAS_AMT_TYPE_NM
                    Frm1.vspdData.Text = ""
                Else
		            Frm1.vspdData.Col = C_BAS_AMT_TYPE_NM
		            Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
                End if 
            End if 
         Case  C_DILIG_CD
            If Trim(Frm1.vspdData.Text) = "" Then
  	            Frm1.vspdData.Col = C_DILIG_CD_NM
                Frm1.vspdData.Text = ""
            Else            
                IntRetCd =  CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(Frm1.vspdData.Text, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                If IntRetCd = false then
                    Call  DisplayMsgBox("970000","X","근태코드","X")
  	                Frm1.vspdData.Col = C_DILIG_CD
                
  	                Frm1.vspdData.Col = C_DILIG_CD_NM
                    Frm1.vspdData.Text = ""
                Else
		            Frm1.vspdData.Col = C_DILIG_CD_NM
		            Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
                End if 
            End if 
    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

    lgCurrentSpd = "M"

End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
    Dim iDx

   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

    Select Case Col
         Case  C_ALLOW_CD_NM
            iDx = Frm1.vspdData1.value
            Frm1.vspdData1.Col = C_ALLOW_CD
            Frm1.vspdData1.value = iDx
         Case Else
    End Select    
             
   	If Frm1.vspdData1.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData1.text) <  UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData1
     ggoSpread.UpdateRow Row

    lgCurrentSpd = "S"

End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim flagTxt
    Call SetPopupMenuItemInf("1101111111")    
	
	 gMouseClickStatus = "SPC"
     Set gActiveSpdSheet = frm1.vspdData     
	 ggoSpread.Source = frm1.vspdData
	With Frm1
		.vspdData.Row = Row
		.vspdData.Col = 0
		flagTxt = .vspdData.Text
		If flagTxt =  ggoSpread.InsertFlag or flagTxt =  ggoSpread.UpdateFlag or flagTxt =  ggoSpread.DeleteFlag Then
			Exit Sub
		End If
	End With	

    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

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
    ggoSpread.Source       = Frm1.vspdData1
	Frm1.vspdData1.MaxRows = 0

	lgCurrentSpd = "S"
	lgStrPrevKey1 = ""
    
	Call MakeKeyStream(Row)
	
	Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery2 = false Then
		Call  RestoreToolBar()
		Exit Sub
	End If

	lgCurrentSpd = "M"
    Set gActiveSpdSheet = frm1.vspdData
    
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101011111")

     gMouseClickStatus = "SP1C"

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
    lgCurrentSpd = "S"
    Set gActiveSpdSheet = frm1.vspdData1    
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
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
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

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub 
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
     End If

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
			    Case C_BAS_AMT_TYPE_POP
					.Col = Col - 1
			    	.Row = Row
					Call OpenCd(1, row)
			    Case C_DILIG_CD_POP
					.Col = Col - 1
			    	.Row = Row
					Call OpenCd(2, row)
			    End Select
		End If
    
	End With
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 5
    
    If  gMouseClickStatus = "SPCRP" Then
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  :	iRet =  DisplayMsgBox("900030", "X", iColumnLimit, "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars =  parent.SS_SCROLLBAR_NONE
    
        ggoSpread.Source = Frm1.vspdData
    
        ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars =  parent.SS_SCROLLBAR_BOTH
    End If   

    If  gMouseClickStatus = "SP1CRP" Then
       ACol = Frm1.vspdData1.ActiveCol
       ARow = Frm1.vspdData1.ActiveRow

       If ACol > iColumnLimit Then
          Frm1.vspdData1.Col = iColumnLimit : Frm1.vspdData1.Row = 0  :	iRet =  DisplayMsgBox("900030", "X", Trim(frm1.vspdData1.Text), "X")
          Exit Function  
       End If   
    
       Frm1.vspdData1.ScrollBars =  parent.SS_SCROLLBAR_NONE
    
        ggoSpread.Source = Frm1.vspdData1
    
        ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData1.Col = ACol
       Frm1.vspdData1.Row = ARow
    
       Frm1.vspdData1.Action = 0    
    
       Frm1.vspdData1.ScrollBars =  parent.SS_SCROLLBAR_BOTH
    End If   
 End Function

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
    Dim iRet

	If Not (Row <> NewRow And NewRow > 0) Then    
	   Exit Sub
	End If
     ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
	lgActiveSpd      = "M"
	lgCurrentSpd	="M"    
End Sub
'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S"
	lgCurrentSpd	="S"        
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    lgCurrentSpd	=  "M"
    call MakeKeyStream("X")
    if frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	
    	If lgStrPrevKey <> "" Then                         
      	   
      	   Call  DisableToolBar( parent.TBC_QUERY)
      	   
      	   If DbQuery = false Then
				Call  RestoreToolBar()
				Exit Sub
		   End If
    	End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    lgCurrentSpd	="S"                     
    If frm1.vspdData1.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData1,NewTop) Then	           
    	
    	If lgStrPrevKey1 <> "" Then   
    		
      	   Call  DisableToolBar( parent.TBC_QUERY)
      	   Call MakeKeyStream(frm1.vspdData.activeRow)
      	   If DbQuery2 = false Then
				Call  RestoreToolBar()
				Exit Sub
		   End If
    	End If
    End if
  
End Sub
'========================================================================================================
'   Event Name : txtBas_amt_type_change
'   Event Desc :
'========================================================================================================
Sub txtBas_amt_type_Onchange()
    Dim IntRetCd


    If frm1.txtBas_amt_type.value = "" Then
		frm1.txtBas_amt_type_nm.value = ""
        
    Else
        IntRetCd =  CommonQueryRs(" allow_NM "," HDA010T "," pay_cd=" & FilterVar("*", "''", "S") & "  and allow_CD =  " & FilterVar(frm1.txtBas_amt_type.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			frm1.txtBas_amt_type_nm.value = ""
        Else
			frm1.txtBas_amt_type_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if
    
    gCounts = 0  
End Sub

'========================================================================================================
'   Event Name : txtOcpt_type_Onchange
'   Event Desc :
'========================================================================================================
Function txtOcpt_type_Onchange()
    gCounts = 0  
End Function
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기초금액계산식등록</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% colspan=2></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100% colspan=2>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
							    <TD CLASS=TD6 NOWRAP>
                                    <SELECT NAME="txtOcpt_type" ALT="급여구분" CLASS=cboNormal TAG="12"></SELECT>
                                </TD>
								<TD CLASS="TD5" NOWRAP>기초금액코드</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBas_amt_type" MAXLENGTH=3 SIZE=10 ALT ="기초금액코드" tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCd 0,'X'">
						                               <INPUT NAME="txtBas_amt_type_nm" MAXLENGTH=20 SIZE=20 ALT ="기초금액코드명" tag=14XXXU></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% colspan=2  valign=top></TD>
				</TR>
				<TR>
					<TD WIDTH=78% HEIGHT=100% valign=top>
					<!-- org 82%-->
						<script language =javascript src='./js/h1016ma1_vaSpread_vspdData.js'></script>
					</TD>
					<TD WIDTH=22% HEIGHT=100% valign=top>
					<!-- org 18%-->
						<script language =javascript src='./js/h1016ma1_vaSpread1_vspdData1.js'></script>
                    </TD>
                </TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave1" tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave2" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

