<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 
*  2. Function Name        : Multi Sample
*  3. Program ID           : h1017ma1
*  4. Program Name         : h1017mb1
*  5. Program Desc         : 기본급산출기준식등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/02
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : chcho
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
Const BIZ_PGM_ID =  "h1017mb1.asp"                                  'Biz Logic ASP 
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
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)

Dim C_PAY_CD
Dim C_PAY_NM
Dim C_BASIC_AMT
Dim C_BASIC_AMT_NM
Dim C_DIVIDE      
Dim C_DIVIDE_BY 
Dim C_DIVIDE_BY_NM
Dim C_PAY_BAS_MM  
Dim C_PAY_BAS_MM_NM
Dim C_PAY_BAS_DD  
Dim C_PAY_PROV_MM 
Dim C_PAY_PROV_MM_NM
Dim C_PAY_PROV_DD   
Dim C_DILIG_MM      
Dim C_DILIG_MM_NM   
Dim C_DILIG_DD 

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_PAY_CD         = 1  'HIDDEN CODE
	 C_PAY_NM         = 2															<%'Spread Sheet의 Column별 상수 %>
	 C_BASIC_AMT      = 3	'HIDDEN CODE
	 C_BASIC_AMT_NM   = 4													
	 C_DIVIDE         = 5  '/
	 C_DIVIDE_BY      = 6  'HIDDEN CODE
	 C_DIVIDE_BY_NM   = 7
	 C_PAY_BAS_MM     = 8  'HIDDEN CODE
	 C_PAY_BAS_MM_NM  = 9 															<%'Spread Sheet의 Column별 상수 %>
	 C_PAY_BAS_DD     = 10														
	 C_PAY_PROV_MM    = 11 'HIDDEN CODE 
	 C_PAY_PROV_MM_NM = 12
	 C_PAY_PROV_DD    = 13
	 C_DILIG_MM       = 14 'HIDDEN CODE
	 C_DILIG_MM_NM    = 15
	 C_DILIG_DD       = 16	  
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
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
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
    lgKeyStream       =  Frm1.cboPay_cd.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
    Dim    iCodeArr
    Dim    iNameArr
	'급여구분  ComboBox   Condition Area..!
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.cboPay_cd,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서 
End Sub

Sub InitComboBox2()
    
    Dim    iCodeArr
    Dim    iNameArr
  	
	'급여구분  ComboBox   Condition Area..!
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1    
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAY_CD
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAY_NM
    
    '기준급 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0101", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_BASIC_AMT
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_BASIC_AMT_NM
    
    '분할기준 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0102", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DIVIDE_BY
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DIVIDE_BY_NM
    '급여기준월 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0103", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAY_BAS_MM
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAY_BAS_MM_NM
    
    '급여지급월 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0103", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAY_PROV_MM
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAY_PROV_MM_NM
    
    '근태기준월 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0103", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DILIG_MM
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DILIG_MM_NM
End Sub
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	 ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow

			.Col = C_PAY_CD             '급여구분 
			intIndex = .value
			.col = C_PAY_NM
			.value = intIndex
						
			.Col = C_BASIC_AMT          '기준급 
			intIndex = .value
			.col = C_BASIC_AMT_NM
			.value = intIndex
			
			.Col = C_DIVIDE_BY          '분할기준 
			intIndex = .value
			.col = C_DIVIDE_BY_NM
			.value = intIndex
			
			.Col = C_PAY_BAS_MM         '급여기준월 
			intIndex = .value
			.col = C_PAY_BAS_MM_NM
			.value = intIndex
			
			.Col = C_PAY_PROV_MM        '급여지급월 
			intIndex = .value
			.col = C_PAY_PROV_MM_NM
			.value = intIndex
			
			.Col = C_DILIG_MM           '근태기준월 
			intIndex = .value
			.col = C_DILIG_MM_NM
			.value = intIndex
		Next	
	End With
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
            Case C_PAY_NM
                .Col = Col
                intIndex = .value    '  COMBO의 VALUE값 
				.Col = C_PAY_CD      '  CODE값란으로 이동 
				.value = intIndex    '  CODE란의 값은 COMBO의 VALUE값이된다.
		    Case C_BASIC_AMT_NM
                .Col = Col
                intIndex = .value   
				.Col = C_BASIC_AMT  
				.value = intIndex 
			Case C_DIVIDE_BY_NM
                .Col = Col
                intIndex = .value   
				.Col = C_DIVIDE_BY    
				.value = intIndex 
			Case C_PAY_BAS_MM_NM
                .Col = Col
                intIndex = .value   
				.Col = C_PAY_BAS_MM    
				.value = intIndex 
			Case C_PAY_PROV_MM_NM
                .Col = Col
                intIndex = .value   
				.Col = C_PAY_PROV_MM    
				.value = intIndex 
			Case C_DILIG_MM_NM
                .Col = Col
                intIndex = .value   
				.Col = C_DILIG_MM
				.value = intIndex 				  
		End Select
	End With

   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim strMaskYM
	strMaskYM = "99"
	
	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_DILIG_DD + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData

       Call  AppendNumberPlace("6","2","0")     
       Call  GetSpreadColumnPos("A")

         ggoSpread.SSSetCOMBO C_PAY_CD        , "급여구분CO", 16, 0 
         ggoSpread.SSSetCOMBO C_PAY_NM        , "급여구분", 15, 0 
         ggoSpread.SSSetCOMBO C_BASIC_AMT     , "기준급", 16 , 0 
         ggoSpread.SSSetCOMBO C_BASIC_AMT_NM  , "기준급", 16 , 0   
         ggoSpread.SSSetEdit  C_DIVIDE        , "" ,2,2
         ggoSpread.SSSetCOMBO C_DIVIDE_BY     , "분할기준CD", 10 , 0
         ggoSpread.SSSetCOMBO C_DIVIDE_BY_NM  , "분할기준", 12 , 0
         ggoSpread.SSSetCOMBO C_PAY_BAS_MM    , "급여기준월CD", 12 , 0
         ggoSpread.SSSetCOMBO C_PAY_BAS_MM_NM , "급여기준월", 12 , 0        
         ggoSpread.SSSetMask  C_PAY_BAS_DD    , "급여기준일",12,2, strMaskYM        
         ggoSpread.SSSetCOMBO C_PAY_PROV_MM   , "급여지급월CD", 12 , 0
         ggoSpread.SSSetCOMBO C_PAY_PROV_MM_NM, "급여지급월", 12 , 0        
         ggoSpread.SSSetMask  C_PAY_PROV_DD   , "급여지급일",12,2, strMaskYM        
         ggoSpread.SSSetCOMBO C_DILIG_MM      , "근태기준월CD" , 12 , 0
         ggoSpread.SSSetCOMBO C_DILIG_MM_NM   , "근태기준월" , 12 , 0        
         ggoSpread.SSSetMask  C_DILIG_DD      , "근태기준일",12,2, strMaskYM        
    	 
    	 Call ggoSpread.SSSetColHidden(C_PAY_CD		,  C_PAY_CD		, True)
    	 Call ggoSpread.SSSetColHidden(C_BASIC_AMT	,  C_BASIC_AMT	, True)
    	 Call ggoSpread.SSSetColHidden(C_DIVIDE_BY	,  C_DIVIDE_BY	, True)
    	 Call ggoSpread.SSSetColHidden(C_PAY_BAS_MM	,  C_PAY_BAS_MM	, True)
    	 Call ggoSpread.SSSetColHidden(C_PAY_PROV_MM	,  C_PAY_PROV_MM, True)
    	 Call ggoSpread.SSSetColHidden(C_DILIG_MM	,  C_DILIG_MM	, True)
    	 
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
       
        ggoSpread.SpreadLock     C_PAY_NM		, -1, C_PAY_NM 
        ggoSpread.SSSetRequired	C_BASIC_AMT_NM  , -1, -1
        ggoSpread.SpreadLock     C_DIVIDE		, -1, C_DIVIDE  
        ggoSpread.SSSetRequired	C_DIVIDE_BY_NM	, -1, -1       
        ggoSpread.SSSetRequired	C_PAY_BAS_MM_NM	, -1, -1
        ggoSpread.SSSetRequired	C_PAY_BAS_DD	, -1, -1
        ggoSpread.SSSetRequired	C_PAY_PROV_MM_NM, -1, -1       
        ggoSpread.SSSetRequired	C_PAY_PROV_DD	, -1, -1
        ggoSpread.SSSetRequired	C_DILIG_MM_NM	, -1, -1
        ggoSpread.SSSetRequired	C_DILIG_DD		, -1, -1
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
     ggoSpread.SSSetRequired		C_PAY_NM			, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_BASIC_AMT_NM		, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected		C_DIVIDE			, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_DIVIDE_BY_NM		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_PAY_BAS_MM_NM		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_PAY_BAS_DD		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_PAY_PROV_MM_NM	, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_PAY_PROV_DD		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_DILIG_MM_NM		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_DILIG_DD			, pvStartRow, pvEndRow
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
            
            C_PAY_CD         = iCurColumnPos(1) 
			C_PAY_NM         = iCurColumnPos(2)															<%'Spread Sheet의 Column별 상수 %>
			C_BASIC_AMT      = iCurColumnPos(3)	
			C_BASIC_AMT_NM   = iCurColumnPos(4)													
			C_DIVIDE         = iCurColumnPos(5) 
			C_DIVIDE_BY      = iCurColumnPos(6) 
			C_DIVIDE_BY_NM   = iCurColumnPos(7)
			C_PAY_BAS_MM     = iCurColumnPos(8) 
			C_PAY_BAS_MM_NM  = iCurColumnPos(9) 															<%'Spread Sheet의 Column별 상수 %>
			C_PAY_BAS_DD     = iCurColumnPos(10)														
			C_PAY_PROV_MM    = iCurColumnPos(11)
			C_PAY_PROV_MM_NM = iCurColumnPos(12)
			C_PAY_PROV_DD    = iCurColumnPos(13)
			C_DILIG_MM       = iCurColumnPos(14)
			C_DILIG_MM_NM    = iCurColumnPos(15)
			C_DILIG_DD       = iCurColumnPos(16)	              
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
'    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetDefaultVal
    Call InitComboBox
    Call InitComboBox2
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
    Call SetDefaultVal
    Call MakeKeyStream("X")
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
              
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
 	Dim intRow
	Dim intIndex 
	Dim len_count
   
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
    
    Dim PAY_BAS_DD  , PAY_PROV_DD  ,DILIG_DD ,lRow
    
    With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag

   	                .vspdData.Col = C_PAY_BAS_DD
   	                
   	                For len_count = 1 to Len(.vspdData.Text)
					     
						'msgbox Mid(.vspdData.Text,i,1) & "의 아스키 값 : " & asc(Mid(.vspdData.Text,i,1)) 
						If (asc(Mid(.vspdData.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData.Action = 0
							Exit Function						
						End If
					Next	
   	                             
			        If CInt(.vspdData.Text) < 10 Then
			            .vspdData.Text = "0" & Cstr(Cint(.vspdData.Text))
			        End If
			        If CInt(.vspdData.Text) > 31 Then
			            Call  DisplayMsgBox("800094","X","X","X")
	                        .vspdData.Text = "00" & .vspdData.Text
  	                        .vspdData.Action=0
                            Set gActiveElement = document.activeElement
                            Exit Function
			            PAY_BAS_DD = .vspdData.Text
			        End If
                    
   	                .vspdData.Col = C_PAY_PROV_DD
   	                
   	                For len_count = 1 to Len(.vspdData.Text)
					     
						If (asc(Mid(.vspdData.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData.Action = 0
							Exit Function						
						End If
					Next	
   	                             
			        If CInt(.vspdData.Text) < 10 Then
			            .vspdData.Text = "0" & Cstr(Cint(.vspdData.Text))
			        End If
			        If CInt(.vspdData.Text) > 31 Then
			            Call  DisplayMsgBox("800094","X","X","X")
	                        .vspdData.Text = "00" & .vspdData.Text
  	                        .vspdData.Action=0
                            Set gActiveElement = document.activeElement
                            Exit Function
			        End If
			        
			       .vspdData.Col = C_DILIG_DD
			       
			       For len_count = 1 to Len(.vspdData.Text)
					     
						If (asc(Mid(.vspdData.Text,len_count,1)) < 48) OR (asc(Mid(.vspdData.Text,len_count,1)) > 58) Then
							call  DisplayMsgBox("126404", "x","x","x")
							.vspdData.Action = 0
							Exit Function						
						End If
					Next	
			       
			        If CInt(.vspdData.Text) < 10 Then
			            .vspdData.Text = "0" & Cstr(Cint(.vspdData.Text))
			        End If
			        If CInt(.vspdData.Text) > 31 Then
			            Call  DisplayMsgBox("800094","X","X","X")
	                        .vspdData.Text = "00" & .vspdData.Text
  	                        .vspdData.Action=0
                            Set gActiveElement = document.activeElement
                            Exit Function
			        End If
                   
            End Select
        Next
	End With
   
	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSAVE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
            
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy_bak()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
            .Col = C_PAY_NM
            .Text = ""
        
                                   
            .ReDraw = True
            .Col = C_PAY_CD
		    .Focus
		    .Action = 0 ' go to 
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

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
           .Col  = C_PAY_NM
           .Row  = .ActiveRow
           .Text = ""
           
           .Col  = C_PAY_CD
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
    Call InitComboBox2
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

	if LayerShowHide(1) = false then
	exit Function
	end if
	
	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
		strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
    
  		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If
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
    
	if LayerShowHide(1) = false then
	exit Function
	end if

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
                                                
                    .vspdData.Col = C_PAY_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BASIC_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   
                    .vspdData.Col = C_DIVIDE_BY	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_BAS_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_BAS_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_PROV_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_PROV_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_MM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_DD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                 
                    .vspdData.Col = C_PAY_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BASIC_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   
                    .vspdData.Col = C_DIVIDE_BY     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_BAS_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_BAS_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_PROV_MM	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_PROV_DD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_MM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_DD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                               
                    .vspdData.Col = C_PAY_CD	    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
	If DbDELETE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
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
    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call InitVariables															'⊙: Initializes local global variables
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function


'========================================================================================================
'	Name : OpenMajor()
'	Description : Major PopUp
'========================================================================================================
Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Major코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_MAJOR"				 		' TABLE 명칭 
	arrParam(2) = frm1.txtMajorCd.value			' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "Major코드"			
	
    arrField(0) = "major_cd"					' Field명(0)
    arrField(1) = "major_nm"				    ' Field명(1)
    
    arrHeader(0) = "Major코드"		        ' Header명(0)
    arrHeader(1) = "Major코드명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

'========================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetMajor(Byval arrRet)
	With frm1
		.txtMajorCd.value = arrRet(0)
		.txtMajorNm.value = arrRet(1)		
	End With
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

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기본급기준식등록</font></td>
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
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> width=100%></TD>
			   </TR>
				<TR>
					<TD HEIGHT=20 width=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>급여구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPay_cd" ALT="급여구분" STYLE="WIDTH: 100px" TAG="1XN"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TDT" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
			    </TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%" WIDTH=100%>
									<script language =javascript src='./js/h1017ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%> ><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD> 
	</TR>

</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
