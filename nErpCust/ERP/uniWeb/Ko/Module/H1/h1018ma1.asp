<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : H1018ma1
*  4. Program Name         : H1018mb1
*  5. Program Desc         : 단수처리기준등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : TGS 최용철 
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID =  "h1018mb1.asp"                                  'Biz Logic ASP 
Const TAB1 = 1
Const TAB2 = 2

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
Dim lgStrPrevKey1

Dim C_ALLOW_CD
Dim C_ALLOW_POP
Dim C_ALLOW_NM
Dim C_BAS_AMT
Dim C_BELOW 
Dim C_BELOW_NM
Dim C_PROC_BAS
Dim C_PROC_BAS_NM

Dim C_ATTEND_TYPE
Dim C_ATTEND_TYPE_NM
Dim C_DECI_PLACE
Dim C_PROC_BAS1
Dim C_PROC_BAS_NM1
Dim C_FORMAT

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
    
		 C_ALLOW_CD    = 1
		 C_ALLOW_POP   = 2
		 C_ALLOW_NM    = 3															<%'Spread Sheet의 Column별 상수 %>
		 C_BAS_AMT     = 4
		 C_BELOW       = 5
		 C_BELOW_NM    = 6
		 C_PROC_BAS    = 7
		 C_PROC_BAS_NM = 8		 

    ElseIf pvSpdNo = "B" Then
    
		 C_ATTEND_TYPE     = 1
		 C_ATTEND_TYPE_NM  = 2
		 C_DECI_PLACE      = 3
		 C_PROC_BAS1       = 4
		 C_PROC_BAS_NM1    = 5
		 C_FORMAT          = 6
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
	lgStrPrevKey1	  = ""
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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream  =  Frm1.txtAllow_cd.Value &  parent.gColSep                     'You Must append one character( parent.gColSep)
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox(ByVal pvSpdNo)
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx 
    
   ' TAB1
     If pvSpdNo = "" OR pvSpdNo = "A" Then   
		 ggoSpread.Source = Frm1.vspdData

		'미만/이하(이상,이하,미만,초과)
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'h0051' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		iCodeArr = lgF0
		iNameArr = lgF1
		   
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_BELOW
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_BELOW_NM
    
		'처리(절사버림,절사올림,사사오입)
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'h0052' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PROC_BAS
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PROC_BAS_NM
	End If

  ' TAB2  근태관련        
	If pvSpdNo = "" OR pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData1     

		'근태구분(근무일수/근무시간)
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'h0124' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		iCodeArr = lgF0
		iNameArr = lgF1
		   
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ATTEND_TYPE
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ATTEND_TYPE_NM

    
		'처리(절사버림,절사올림,사사오입)
		Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'h0052' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
		iCodeArr = lgF0
		iNameArr = lgF1
		 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PROC_BAS1
		 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PROC_BAS_NM1

	End if	
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	If gSelframeFlg = TAB1 Then
       ggoSpread.Source = Frm1.vspdData
	
		With frm1.vspdData
			For intRow = 1 To .MaxRows			
				.Row = intRow
				.Col = C_BELOW
				intIndex = .value
				.col = C_BELOW_NM
				.value = intindex	
				
				.Row = intRow
				.Col = C_PROC_BAS
				intIndex = .value
				.col = C_PROC_BAS_NM
				.value = intindex					
			Next	
		End With
	Else
	 ggoSpread.Source = Frm1.vspdData1	
		 With frm1.vspdData1
			For intRow = 1 To .MaxRows			
				.Row = intRow
				.Col = C_ATTEND_TYPE
				intIndex = .value
				.col = C_ATTEND_TYPE_NM
				.value = intindex	

	
				.Row = intRow
				.Col = C_PROC_BAS1
				intIndex = .value
				.col = C_PROC_BAS_NM1
				.value = intindex					
			Next	
		End With
	End if
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
           ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread    
           .ReDraw = false
           .MaxCols = C_PROC_BAS_NM + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
           .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
           .ColHidden = True
           
           .MaxRows = 0
           ggoSpread.ClearSpreadData

		   Call GetSpreadColumnPos("A") 

		   ggoSpread.SSSetEdit    C_ALLOW_CD    , "수당/공제코드", 12 ,,,3,2
		   ggoSpread.SSSetButton  C_ALLOW_POP    
		   ggoSpread.SSSetEdit    C_ALLOW_NM    , "수당/공제코드명", 30,,,20,2
		   ggoSpread.SSSetFloat   C_BAS_AMT     , "기준액" , 22 , parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetCombo   C_BELOW       , "미만/이하CD", 10 , 0 
		   ggoSpread.SSSetCombo   C_BELOW_NM    , "미만/이하", 15 , 0 
		   ggoSpread.SSSetCombo   C_PROC_BAS    , "처리CD", 10 , 0
		   ggoSpread.SSSetCombo   C_PROC_BAS_NM , "처리", 20 , 0

					 
		   Call ggoSpread.MakePairsColumn(C_ALLOW_CD	,  C_ALLOW_POP)
		   
		   
		   Call ggoSpread.SSSetColHidden(C_BELOW		,  C_BELOW		, True)
		   Call ggoSpread.SSSetColHidden(C_PROC_BAS	,  C_PROC_BAS	, True)

		  .ReDraw = true

		Call SetSpreadLock("A") 
    
		End With
    
    End if
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then		

		Call initSpreadPosVariables("B")	
		With frm1.vspdData1
 
		    ggoSpread.Source = frm1.vspdData1	
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

		    .ReDraw = false
		    .MaxCols = C_FORMAT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True
		    .MaxRows = 0
		
		    Call AppendNumberPlace("6","1","0")		   
		    Call GetSpreadColumnPos("B")
		    
		    ggoSpread.SSSetCombo   C_ATTEND_TYPE       , "코드", 4 , 0 
		    ggoSpread.SSSetCombo   C_ATTEND_TYPE_NM    , "구분", 30 , 0 
		    ggoSpread.SSSetFloat   C_DECI_PLACE     , "소수점자리수" , 20 , "6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0","4" 
		    ggoSpread.SSSetCombo   C_PROC_BAS1    , "처리CD", 10 , 0
		    ggoSpread.SSSetCombo   C_PROC_BAS_NM1 , "처리", 20 , 0
		    ggoSpread.SSSetEdit    C_FORMAT    , "포맷", 20,0
		    
		    Call ggoSpread.SSSetColHidden(C_ATTEND_TYPE	,  C_ATTEND_TYPE, True)
		    Call ggoSpread.SSSetColHidden(C_PROC_BAS1	,  C_PROC_BAS1	, True)
		     
		   .ReDraw = true
	
		Call SetSpreadLock("B") 
    
		End With
	End if
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then

        ggoSpread.Source = Frm1.vspdData

        With frm1.vspdData
        	.ReDraw = False

        	ggoSpread.SpreadLock     C_ALLOW_NM		, -1, C_ALLOW_NM  
			ggoSpread.SpreadLock     C_ALLOW_CD		, -1, C_ALLOW_CD
			ggoSpread.SpreadLock     C_ALLOW_POP	, -1, C_ALLOW_POP
			ggoSpread.SSSetRequired	 C_BAS_AMT		, -1, -1
			ggoSpread.SSSetRequired	 C_BELOW_NM		, -1, -1
			ggoSpread.SSSetRequired	 C_PROC_BAS_NM  , -1, -1
			ggoSpread.SSSetProtected .MaxCols		, -1, -1       	

        	.ReDraw = True
        End With
        
    ElseIf pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData1

        With frm1.vspdData1
        	.ReDraw = False

        	ggoSpread.SpreadLock     C_ATTEND_TYPE_NM , -1, -1
			ggoSpread.SSSetRequired	 C_DECI_PLACE	  , -1, -1
			ggoSpread.SSSetRequired	 C_PROC_BAS_NM1   , -1, -1
			ggoSpread.SpreadLock     C_FORMAT		  , -1, -1        	
			ggoSpread.SSSetProtected .MaxCols		   ,-1 ,-1

        	.ReDraw = True
        End With
    End If
                
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	If gSelframeFlg = TAB1 Then
		With frm1
    
		.vspdData.ReDraw = False

	   ggoSpread.SSSetRequired    C_ALLOW_CD		, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_ALLOW_NM		, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_BAS_AMT			, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_BELOW_NM		, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_PROC_BAS_NM		, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected	.vspdData.MaxCols	, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
    
		End With
    Else    
		With frm1
    
		.vspdData1.ReDraw = False

		ggoSpread.SSSetRequired    C_ATTEND_TYPE_NM , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_DECI_PLACE	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_PROC_BAS_NM1 , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected   C_FORMAT		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	.vspdData1.MaxCols, pvStartRow, pvEndRow
		.vspdData1.ReDraw = True
    
		End With
	End If

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
        if gSelframeFlg = TAB1 then
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
		else       
			For iDx = 1 To  frm1.vspdData1.MaxCols - 1
			    Frm1.vspdData1.Col = iDx
			    Frm1.vspdData1.Row = iRow
			    If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <>  parent.UC_PROTECTED Then
			       Frm1.vspdData1.Col = iDx
			       Frm1.vspdData1.Row = iRow
			       Frm1.vspdData1.Action = 0 ' go to 
			       Exit For
			    End If
			    
			Next
        end if  
    End If 
    Call  RestoreToolBar()  
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
            
            C_ALLOW_CD    = iCurColumnPos(1)
			C_ALLOW_POP   = iCurColumnPos(2)
			C_ALLOW_NM    = iCurColumnPos(3)															<%'Spread Sheet의 Column별 상수 %>
			C_BAS_AMT     = iCurColumnPos(4)
			C_BELOW       = iCurColumnPos(5)
			C_BELOW_NM    = iCurColumnPos(6)
			C_PROC_BAS    = iCurColumnPos(7)
			C_PROC_BAS_NM = iCurColumnPos(8)            
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)			
			
			C_ATTEND_TYPE     = iCurColumnPos(1)
			C_ATTEND_TYPE_NM  = iCurColumnPos(2)
			C_DECI_PLACE      = iCurColumnPos(3)
			C_PROC_BAS1       = iCurColumnPos(4)
			C_PROC_BAS_NM1    = iCurColumnPos(5)
			C_FORMAT          = iCurColumnPos(6)			
   
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
    
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox("")
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    
    frm1.txtAllow_cd.Focus
    
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    

	Call ClickTab1    

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

    If gSelframeFlg = TAB1 Then
        ggoSpread.Source = Frm1.vspdData
       If  ggoSpread.SSCheckChange = True Then
		   IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		   If IntRetCD = vbNo Then
			   Exit Function
		   End If
       End If
    Else
        ggoSpread.Source = Frm1.vspdData1
       If  ggoSpread.SSCheckChange = True Then
		   IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		   If IntRetCD = vbNo Then
			   Exit Function
		   End If
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

  	txtAllow_cd_Onchange()                                             '☜: enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	If DBQuery = False Then
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
    Dim BAS_AMT                                                '저장시 그리드에입력된 기준액 타당성 체크 
   	Dim lRow
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    if gSelframeFlg = TAB1 then
         ggoSpread.Source = frm1.vspdData
        If  ggoSpread.SSCheckChange = False Then
            IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If

		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If
    else
         ggoSpread.Source = frm1.vspdData1
        If  ggoSpread.SSCheckChange = False Then
            IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
            Exit Function
        End If

		If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		   Exit Function
		End If		
    end if

    If gSelframeFlg = TAB1 Then
         ggoSpread.Source = frm1.vspdData

        With Frm1.vspdData
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = 0
            Select Case .Text
                Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag

   	                .Col = C_ALLOW_NM
   	                
					if trim(.Text) = "" then
						IntRetCD = DisplayMsgBox("970000","X","수당코드","X")
						.focus
					    exit Function
					end if 
   	                
   	                .Col = C_BAS_AMT
                    BAS_AMT =  UNICDbl(.text)
                   	
                   	If BAS_AMT<= 0 then
	                    Call  DisplayMsgBox("800172","X","X","X")	'기준액은 0 보다 커야 합니다.
	                   .Row = lRow
  	                   .Col = C_BAS_AMT
                       .focus
                       Set gActiveElement = document.activeElement
                       Exit Function
                    End if 
           End Select
        Next
	    End With
    End if

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
	If gSelframeFlg = TAB1 Then
        lgCurrentSpd = "M"

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
                
               .Col = C_ALLOW_NM
               .Text = ""                                   
               
               .Col = C_ALLOW_CD
			   .Text = ""                                   
				
				.ReDraw = True
    		    .Focus
			End If
		End With
	Else
	    lgCurrentSpd = "S"

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
                
               .Col = C_ATTEND_TYPE_NM
               .Text = ""
               .Col = C_ATTEND_TYPE
               .Text = ""
    
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
    If gSelframeFlg = TAB1 Then
        ggoSpread.Source = Frm1.vspdData	
        ggoSpread.EditUndo
    Else
        ggoSpread.Source = Frm1.vspdData1
        ggoSpread.EditUndo
    End If
    
    Call  initData()  
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
	
	If gSelframeFlg = TAB1 Then
	   lgCurrentSpd = "M"

		With frm1
	    .vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
        End With

    Else
	   lgCurrentSpd = "S"

		With frm1
	    .vspdData1.ReDraw = False
		.vspdData1.focus
		ggoSpread.Source = .vspdData1
        ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
        SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
		.vspdData1.ReDraw = True
        End With
	End If

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
    
    If gSelframeFlg = TAB1 Then

       If Frm1.vspdData.MaxRows < 1 then
          Exit function
	   End if	
       With Frm1.vspdData 
    	   .focus
    	    ggoSpread.Source = frm1.vspdData 
    	   lDelRows =  ggoSpread.DeleteRow
       End With
    Else

       If Frm1.vspdData1.MaxRows < 1 then
          Exit function
	   End if	
       With Frm1.vspdData1
    	   .focus
    	    ggoSpread.Source = frm1.vspdData1
    	   lDelRows =  ggoSpread.DeleteRow
       End With
    End If
       
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

    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
            Call InitComboBox("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
            Call InitComboBox("B")
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

    If gSelframeFlg = TAB1 Then
	
        ggoSpread.Source = frm1.vspdData	
       If  ggoSpread.SSCheckChange = True Then
	       IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		   If IntRetCD = vbNo Then
			  Exit Function
		   End If
       End If

    Else
        ggoSpread.Source = frm1.vspdData1
       If  ggoSpread.SSCheckChange = True Then
	       IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		   If IntRetCD = vbNo Then
			  Exit Function
		   End If
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

	Call LayerShowHide(1)
	
	Dim strVal

    If gSelframeFlg = Tab1 Then    

       With Frm1
	       strVal = BIZ_PGM_ID & "?txtMode="            &  parent.UID_M0001						         
           strVal = strVal     & "&lgCurrentSpd="       & gSelframeFlg                      '☜: Next key tag
		   strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
           strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
           strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey                 '☜: Next key tag
       End With
    Else

       With Frm1
	       strVal = BIZ_PGM_ID & "?txtMode="            &  parent.UID_M0001						         
           strVal = strVal     & "&lgCurrentSpd="       & gSelframeFlg                      '☜: Next key tag
		   strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
           strVal = strVal     & "&txtMaxRows="         & .vspdData1.MaxRows
           strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey1
       End With    
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    If gSelframeFlg = TAB1 Then

	   With Frm1
    
           For lRow = 1 To .vspdData.MaxRows
        
               .vspdData.Row = lRow
               .vspdData.Col = 0
            
               Select Case .vspdData.Text
    
                   Case  ggoSpread.InsertFlag                                      '☜: Update추가 
                                                      strVal = strVal & "C"  &  parent.gColSep
                                                      strVal = strVal & lRow &  parent.gColSep
                                                      strVal = strVal & gSelframeFlg &  parent.gColSep
                        .vspdData.Col = C_ALLOW_CD	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                        .vspdData.Col = C_BAS_AMT	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                        .vspdData.Col = C_BELOW	    : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                        .vspdData.Col = C_PROC_BAS  : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                         lGrpCnt = lGrpCnt + 1
               
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U"  &  parent.gColSep
                                                      strVal = strVal & lRow &  parent.gColSep
                                                      strVal = strVal & gSelframeFlg &  parent.gColSep
                       .vspdData.Col = C_ALLOW_CD	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                       .vspdData.Col = C_BAS_AMT	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                       .vspdData.Col = C_BELOW  	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                       .vspdData.Col = C_PROC_BAS   : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                 
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                      strDel = strDel & "D"  &  parent.gColSep
                                                      strDel = strDel & lRow &  parent.gColSep
                                                      strDel = strDel & gSelframeFlg &  parent.gColSep
                       .vspdData.Col = C_ALLOW_CD   : strDel = strDel & Trim(.vspdData.Text) &  parent.gRowSep	'삭제시 key만								
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
    
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
    	   .txtMaxRows.value     = lGrpCnt-1	
    	   .txtSpread.value      = strDel & strVal
    	   .lgCurrentSpd.value   = TAB1
    
	   End With
	   
	Else

	   With Frm1
           For lRow = 1 To .vspdData1.MaxRows
               .vspdData1.Row = lRow
               .vspdData1.Col = 0
            
               Select Case .vspdData1.Text
    
                   Case  ggoSpread.InsertFlag                                      '☜: Update추가 
                                                      strVal = strVal & "C"  &  parent.gColSep
                                                      strVal = strVal & lRow &  parent.gColSep
                                                      strVal = strVal & gSelframeFlg &  parent.gColSep
                        .vspdData1.Col = C_ATTEND_TYPE	: strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep
                        .vspdData1.Col = C_DECI_PLACE	: strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep
                        .vspdData1.Col = C_PROC_BAS1: strVal = strVal & Trim(.vspdData1.Text) &  parent.gRowSep
                         lGrpCnt = lGrpCnt + 1
               
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U"  &  parent.gColSep
                                                      strVal = strVal & lRow &  parent.gColSep
                                                      strVal = strVal & gSelframeFlg &  parent.gColSep
                       .vspdData1.Col = C_ATTEND_TYPE	: strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep
                       .vspdData1.Col = C_DECI_PLACE	: strVal = strVal & Trim(.vspdData1.Text) &  parent.gColSep
                       .vspdData1.Col = C_PROC_BAS1 : strVal = strVal & Trim(.vspdData1.Text) &  parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                 
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                      strDel = strDel & "D"  &  parent.gColSep
                                                      strDel = strDel & lRow &  parent.gColSep
                                                      strDel = strDel & gSelframeFlg &  parent.gColSep
                       .vspdData1.Col = C_ATTEND_TYPE   : strDel = strDel & Trim(.vspdData1.Text) &  parent.gRowSep	'삭제시 key만								
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
    
           .txtMode.value        =  parent.UID_M0002
           .txtUpdtUserId.value  =  parent.gUsrID
           .txtInsrtUserId.value =  parent.gUsrID
    	   .txtMaxRows.value     = lGrpCnt-1	
    	   .txtSpread.value      = strDel & strVal
    	   .lgCurrentSpd.value   = TAB2
    
	   End With
    End If

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
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("110011110011111")									
	If gSelframeFlg = Tab1 Then    
		frm1.vspdData.focus
	else
		frm1.vspdData1.focus
	end if	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables

	Call  DisableToolBar( parent.TBC_QUERY)					'Query 버튼을 disable시킴 

	If FncQuery = False Then  '이화면에서만 fncQuery호출로 바꿈 lsm
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
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	    Case "1"

	        arrParam(0) = "수당/공제코드팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtAllow_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value		' Name Cindition
	        arrParam(4) = " PAY_CD = '*'"               ' Where Condition

	        arrParam(5) = "수당/공제코드"			    ' TextBox 명칭 
	
            arrField(0) = "allow_cd"					' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당/공제코드"				' Header명(0)
            arrHeader(1) = "수당/공제코드명"			    ' Header명(1)

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	If arrRet(0) = "" Then
		frm1.txtAllow_cd.focus	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)  
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtAllow_cd.value = arrRet(0)
		        .txtAllow_nm.value = arrRet(1)		
				.txtAllow_cd.focus
        End Select
	End With
End Sub
'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_ALLOW_POP
	        arrParam(0) = "수당/공제코드팝업"			    ' 팝업 명칭 
	    	arrParam(1) = "HDA010T"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                   			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = " PAY_CD = '*'"                       ' Where Condition

	    	arrParam(5) = "수당/공제코드" 			        ' TextBox 명칭 
	
	    	arrField(0) = "allow_cd"							' Field명(0)
	    	arrField(1) = "allow_nm"    						' Field명(1)
    
	    	arrHeader(0) = "수당/공제코드"	   		    	' Header명(0)
	    	arrHeader(1) = "수당/공제코드명"	    		' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If gSelframeFlg = TAB1 Then
		   ggoSpread.Source = Frm1.vspdData
			  Select Case iWhere
			    Case C_ALLOW_POP
			        frm1.vspdData.Col = C_ALLOW_CD
			    	frm1.vspdData.action =0
		      End Select
		Else
		   ggoSpread.Source = Frm1.vspdData1
			  Select Case iWhere
			    Case C_ALLOW_POP
			        frm1.vspdData1.Col = C_ALLOW_CD
			    	frm1.vspdData1.action =0		    	
		      End Select
		End If
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)

         ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

    If gSelframeFlg = TAB1 Then
       ggoSpread.Source = Frm1.vspdData

       With frm1.vspdData
		  Select Case iWhere
		    Case C_ALLOW_POP
		        .Col = C_ALLOW_NM
		    	.text = arrRet(1)   
		        .Col = C_ALLOW_CD
		    	.text = arrRet(0) 
		    	.action =0
          End Select
       End With

    Else
       ggoSpread.Source = Frm1.vspdData1

       With frm1.vspdData1
		  Select Case iWhere
		    Case C_ALLOW_POP
		        .Col = C_ALLOW_NM
		    	.text = arrRet(1)   
		        .Col = C_ALLOW_CD
		    	.text = arrRet(0) 
		    	.action =0		    	
          End Select
       End With
    End If

End Function

'========================================================================================================
'   Event Name : vspdData_Change       
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col        	
    	    Case C_ALLOW_CD
                        
            IntRetCD=  CommonQueryRs(" allow_nm "," HDA010T "," allow_CD = '" & frm1.vspdData.Text & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
            If IntRetCD=False  Then
                frm1.vspdData.Col = C_ALLOW_NM
                frm1.vspdData.Text=""
                Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
                exit sub
            Else
    	        frm1.vspdData.Col = C_ALLOW_NM
                frm1.vspdData.Text= Trim(Replace(lgF0,Chr(11),""))
            End If
    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD
       
   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

    Select Case Col
         Case  C_ATTEND_TYPE_NM 
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_ATTEND_TYPE
                Frm1.vspdData1.value = iDx
       
         Case  C_PROC_BAS_NM1     
                iDx = Frm1.vspdData1.value
   	            Frm1.vspdData1.Col = C_PROC_BAS1
                Frm1.vspdData1.value = iDx
                
         Case  C_DECI_PLACE
				iDx = Frm1.vspdData1.value
				frm1.vspddata1.Col = C_FORMAT
				frm1.vspddata1.value = get_format(idx)
         
         Case Else
    End Select    
             
   	If Frm1.vspdData1.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData1.text) <  UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData1
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_BELOW_NM
                .Col = Col
                intIndex = .Value        '  COMBO의 VALUE값 
				.Col = C_BELOW           '  CODE값란으로 이동 
				.Value = intIndex        '  CODE란의 값은 COMBO의 VALUE값이된다.
				
		    Case C_PROC_BAS_NM
                .Col = Col
                intIndex = .Value   
				.Col = C_PROC_BAS  
				.Value = intIndex 
			  
		End Select
	End With

   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData1
		
		.Row = Row
    
        Select Case Col
            Case C_ATTEND_TYPE_NM
                .Col = Col
                intIndex = .Value        '  COMBO의 VALUE값 
				.Col = C_ATTEND_TYPE     '  CODE값란으로 이동 
				.Value = intIndex        '  CODE란의 값은 COMBO의 VALUE값이된다.
				
		    Case C_PROC_BAS_NM1
                .Col = Col
                intIndex = .Value   
				.Col = C_PROC_BAS1  
				.Value = intIndex 
			  
		End Select
	End With

   	 ggoSpread.Source = frm1.vspdData1
     ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
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
    
End Sub

'======================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
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
    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
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

Sub vspdData1_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData1.MaxRows = 0 then
		exit sub
	end if
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

Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
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

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
    
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_ALLOW_POP
            Call OpenCode("", C_ALLOW_POP, Row)
    End Select    
End Sub

'========================================================================================================
'   Event Name : txtallow_cd_OnChange()             
'   Event Desc : allow_cd  field 입력값 check
'========================================================================================================
Function txtallow_cd_OnChange()
    Dim iDx
    Dim IntRetCd   
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
		txtAllow_cd_Onchange = True
    Else
        IntRetCd =  CommonQueryRs(" allow_nm "," HDA010T "," allow_cd = '" & frm1.txtallow_cd.value & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false  Then
        
            'Call  DisplayMsgBox("800145","X","X","X")            '수당정보에 등록되지않은 코드입니다.
            frm1.txtallow_nm.value=""             
    		txtAllow_cd_Onchange = false 
        ELSE   '수당코드 
        
            frm1.txtallow_nm.value=Trim(Replace(lgF0,Chr(11),""))
    		txtAllow_cd_Onchange = True
        END IF
    END IF  
End Function 

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


Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData1,NewTop) Then	           
    	If lgStrPrevKey1 <> "" Then                         
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If
			Call  DisableToolBar( parent.TBC_QUERY)					'Query 버튼을 disable시킴 
			If DBQuery = False Then
				Call  RestoreToolBar()
				Exit Sub
			End If
    	End If
    End if
End Sub

Function ClickTab1()
	Dim IntRetCD

	If gSelframeFlg = TAB1 Then Exit Function
	
	 ggoSpread.Source = frm1.vspdData1
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)                                               <%'첫번째 Tab%>

	 ggoSpread.Source = frm1.vspdData
	gSelframeFlg = TAB1

End Function

Function ClickTab2()
	Dim IntRetCD

	If gSelframeFlg = TAB2 Then Exit Function
	
	 ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call changeTabs(TAB2)

	 ggoSpread.Source = frm1.vspdData1

	gSelframeFlg = TAB2

End Function


function get_Format( no )

	DIM retFormat
	
	retFormat = ""
	select case no
		case 0
			retFormat = "###" 
		case 1
			retFormat = "###" &  parent.gComNumDec & "0"
		case 2
			retFormat = "###" &  parent.gComNumDec & "00"
		case 3
			retFormat = "###" &  parent.gComNumDec & "000"
		case 4
			retFormat = "###" &  parent.gComNumDec & "0000"
	end select
	
	get_format = retFormat
	
end function

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>단수처리기준등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>근태단수처리기준등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>

					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    			
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
							    <TD CLASS="TD5" NOWRAP>수당/공제코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT ID=txtAllow_cd NAME="txtAllow_cd" MAXLENGTH="3" SIZE="10" ALT ="수당코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
								                       <INPUT ID=txtAllow_nm NAME="txtAllow_nm" SIZE="20"  ALT ="수당코드명" tag="14XXXU"></TD>
								<TD CLASS="TDT" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>                     
							</TR>
						</TABLE>
						</FIELDSET>
						
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%" WIDTH=100%>
									<script language =javascript src='./js/h1018ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" WIDTH=100%>
									<script language =javascript src='./js/h1018ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no  noresize framespacing=0></IFRAME></TD>
	</TR>
	
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd"   TAG="24">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
