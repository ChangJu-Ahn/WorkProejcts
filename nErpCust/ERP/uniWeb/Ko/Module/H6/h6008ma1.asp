<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          	: Basis Architect
'*  2. Function Name        	: 급여관리 
'*  3. Program ID           	: h6008ma1.asp
'*  4. Program Name         	: h6008ma1.asp
'*  5. Program Desc         	: 급여계산수당/공제등록 
'*  6. Modified date(First) 	: 2001/05/17
'*  7. Modified date(Last)  	: 2003/06/13
'*  8. Modifier (First)     	: Song Bong-kyu
'*  9. Modifier (Last)      	: Lee SiNa
'* 10. Comment             	:
'* 11. Common Coding Guide   : this mark(☜) means that "Do not change"
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->

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
Const BIZ_PGM_ID      = "h6008mb1.asp"						           '☆: Biz Logic ASP Name

Const TAB1 = 1										                   'Tab의 위치 
Const TAB2 = 2
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
Dim lgStrPrevKey1

Dim C_PayCd
Dim C_PayNm
Dim C_AllowNameCd
Dim C_AllowNameNm
Dim C_AllowCdCd
Dim C_AllowCdPop
Dim C_AllowCdNm
Dim C_AllowSeq
Dim C_CalcYn1

Dim C_PayCd2
Dim C_PayNm2
Dim C_AllowNameCd2
Dim C_AllowNameNm2
Dim C_AllowCdCd2
Dim C_AllowCdPop2
Dim C_AllowCdNm2
Dim C_AllowSeq2
Dim C_LendBaseCd2
Dim C_LendBaseNm2
Dim C_CalcYn2

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_PayCd = 1															'Spread Sheet의 Column별 상수 %>
        C_PayNm = 2															
        C_AllowNameCd = 3															
        C_AllowNameNm = 4															
        C_AllowCdCd = 5															
        C_AllowCdPop =  6
        C_AllowCdNm = 7															
        C_AllowSeq = 8														
        C_CalcYn1 = 9	

    ElseIf pvSpdNo = "B" Then
        C_PayCd2 = 1
        C_PayNm2 = 2
        C_AllowNameCd2 = 3
        C_AllowNameNm2 = 4
        C_AllowCdCd2 = 5
        C_AllowCdPop2 =  6
        C_AllowCdNm2 = 7
        C_AllowSeq2 = 8
        C_LendBaseCd2 = 9
        C_LendBaseNm2 = 10
        C_CalcYn2 = 11
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

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
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
    lgKeyStream       = Frm1.cboPayCd.Value & Parent.gColSep           'You Must append one character(Parent.gColSep)
End Sub        

'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitComboBox(ByVal pvSpdNo)
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    If pvSpdNo = "" OR pvSpdNo = "A" Then
   
        ggoSpread.Source = Frm1.vspdData

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," h_pay_cd "," MINOR_CD LIKE " & FilterVar("%", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PayCd
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PayNm

        Call SetCombo2(frm1.cboPayCd, iCodeArr, iNameArr,Chr(11))

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0066", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AllowNameCd
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_AllowNameNm
    
        ggoSpread.SetCombo "Y"               & vbtab & "N"              , C_CalcYn1
    End If

' TAB2        
	If pvSpdNo = "" OR pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData2

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," h_pay_cd "," MINOR_CD LIKE " & FilterVar("%", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PayCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PayNm2

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0067", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AllowNameCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_AllowNameNm2
    
        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0099", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_LendBaseCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_LendBaseNm2

        ggoSpread.SetCombo "Y"               & vbtab & "N"              , C_CalcYn2
    End If
   
End Sub

'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitSpreadComboBox(ByVal pvSpdNo)
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    If pvSpdNo = "" OR pvSpdNo = "A" Then
   
        ggoSpread.Source = Frm1.vspdData

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," h_pay_cd "," MINOR_CD LIKE " & FilterVar("%", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PayCd
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PayNm

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0066", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AllowNameCd
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_AllowNameNm
    
        ggoSpread.SetCombo "Y"               & vbtab & "N"              , C_CalcYn1
    End If

' TAB2        
	If pvSpdNo = "" OR pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData2

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," h_pay_cd "," MINOR_CD LIKE " & FilterVar("%", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PayCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PayNm2

        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0067", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AllowNameCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_AllowNameNm2
    
        Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0099", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
        iCodeArr = lgF0
        iNameArr = lgF1

        ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_LendBaseCd2
        ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_LendBaseNm2

        ggoSpread.SetCombo "Y"               & vbtab & "N"              , C_CalcYn2
    End If
   
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
			    .Col = C_PayCd
			    intIndex = .value
			    .col = C_PayNm
			    .value = intindex	
			
                .Col = C_AllowNameCd
                intIndex = .value
                .col = C_AllowNameNm
                .value = intindex					
            Next	
	   End With
    Else	   
       ggoSpread.Source = Frm1.vspdData2
	   With frm1.vspdData2
            For intRow = 1 To .MaxRows			
			    .Row = intRow
			    .Col = C_PayCd2
			    intIndex = .value
			    .col = C_PayNm2
			    .value = intindex	
			
                .Col = C_AllowNameCd2
                intIndex = .value
                .col = C_AllowNameNm2
                .value = intindex					

                .Col = C_LendBaseCd2
                intIndex = .value
                .col = C_LendBaseNm2
                .value = intindex					
            Next	
	   End With
    End If
	
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

    	With frm1.vspdData
            ggoSpread.Source = Frm1.vspdData

            ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    'sbk

    	   .ReDraw = false

           .MaxCols = C_CalcYn1 + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

    	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0              
            ggoSpread.ClearSpreadData
            
            Call GetSpreadColumnPos("A") 'sbk

            ggoSpread.SSSetCombo C_PayCd        , "",2
            ggoSpread.SSSetCombo C_PayNm        , "급여구분", 25                           
            ggoSpread.SSSetCombo C_AllowNameCd  , "",2
            ggoSpread.SSSetCombo C_AllowNameNm  , "계산종류", 25                           
            ggoSpread.SSSetEdit  C_AllowCdCd    , "수당코드",15,,,10,2
            ggoSpread.SSSetButton C_AllowCdPop
            ggoSpread.SSSetEdit  C_AllowCdNm    , "수당코드명", 20,,,20,2
    		ggoSpread.SSSetFloat C_AllowSeq     , "계산순서", 14,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z","1","99"
            ggoSpread.SSSetCombo C_CalcYn1      , "계산여부", 10

            Call ggoSpread.MakePairsColumn(C_AllowCdCd,C_AllowCdPop)    'sbk

            Call ggoSpread.SSSetColHidden(C_PayCd,C_PayCd,True)
            Call ggoSpread.SSSetColHidden(C_AllowNameCd,C_AllowNameCd,True)

    		.ReDraw = true

            Call SetSpreadLock("A") 
    	
        End With
    
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With frm1.vspdData2
            ggoSpread.Source = Frm1.vspdData2

            ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols = C_CalcYn2 + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True                                                            ' ☜:☜:

           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 'sbk

            ggoSpread.SSSetCombo C_PayCd2        , "",2
            ggoSpread.SSSetCombo C_PayNm2        , "급여구분", 20
            ggoSpread.SSSetCombo C_AllowNameCd2  , "",2
            ggoSpread.SSSetCombo C_AllowNameNm2  , "계산종류", 20                           
            ggoSpread.SSSetEdit  C_AllowCdCd2    , "공제코드", 15,,,10,2
	    	ggoSpread.SSSetButton C_AllowCdPop2
            ggoSpread.SSSetEdit  C_AllowCdNm2    , "공제코드명", 20,,,20,2
	    	ggoSpread.SSSetFloat C_AllowSeq2     , "계산순서", 14,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z","1","99"
            ggoSpread.SSSetCombo C_LendBaseCd2   , "",1
            ggoSpread.SSSetCombo C_LendBaseNm2   , "대부상환기준", 15
            ggoSpread.SSSetCombo C_CalcYn2      , "계산여부", 10

            Call ggoSpread.MakePairsColumn(C_AllowCdCd2,C_AllowCdPop2)    'sbk

            Call ggoSpread.SSSetColHidden(C_PayCd2,C_PayCd2,True)
            Call ggoSpread.SSSetColHidden(C_AllowNameCd2,C_AllowNameCd2,True)
            Call ggoSpread.SSSetColHidden(C_LendBaseCd2,C_LendBaseCd2,True)

	    	.ReDraw = true	

            Call SetSpreadLock("B")

        End With
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then

        ggoSpread.Source = Frm1.vspdData

        With frm1.vspdData
        	.ReDraw = False

        	ggoSpread.SpreadLock    C_PayNm, -1, C_PayNm, -1
        	ggoSpread.SpreadLock    C_AllowNameCd, -1, C_AllowNameCd, -1
        	ggoSpread.SpreadLock    C_AllowNameNm, -1, C_AllowNameNm, -1
        	ggoSpread.SpreadLock    C_AllowCdCd, -1, C_AllowCdCd, -1
        	ggoSpread.SpreadLock    C_AllowCdPop, -1, C_AllowCdPop, -1
        	ggoSpread.SpreadLock    C_AllowCdNm, -1, C_AllowCdNm, -1
        	ggoSpread.SSSetRequired C_AllowSeq, -1, -1
        	ggoSpread.SSSetRequired C_CalcYn1, -1, -1
        	ggoSpread.SSSetProtected   .MaxCols   , -1, -1

        	.ReDraw = True
        End With
        
    ElseIf pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData2

        With frm1.vspdData2
        	.ReDraw = False

        	ggoSpread.SpreadLock    C_PayNm2, -1, C_PayNm2, -1
        	ggoSpread.SpreadLock    C_AllowNameCd2, -1, C_AllowNameCd2, -1
        	ggoSpread.SpreadLock    C_AllowNameNm2, -1, C_AllowNameNm2, -1
        	ggoSpread.SpreadLock    C_AllowCdCd2, -1, C_AllowCdCd2, -1
        	ggoSpread.SpreadLock    C_AllowCdPop2, -1, C_AllowCdPop2, -1
        	ggoSpread.SpreadLock    C_AllowCdNm2, -1, C_AllowCdNm2, -1
        	ggoSpread.SSSetRequired C_AllowSeq2, -1, -1
        	ggoSpread.SSSetRequired C_LendBaseCd2, -1, -1
        	ggoSpread.SSSetRequired C_LendBaseNm2, -1, -1
        	ggoSpread.SSSetRequired C_CalcYn2, -1, -1
        	ggoSpread.SSSetProtected   .MaxCols   , -1, -1

        	.ReDraw = True
        End With
    End If
                
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	If gSelframeFlg = TAB1 Then
		With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired	   C_PayNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	   C_AllowNameNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	   C_AllowCdCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_AllowCdNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired     C_AllowSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired     C_CalcYn1, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
    
		End With
    Else    
		With frm1
    
		.vspdData2.ReDraw = False

		ggoSpread.SSSetRequired	   C_PayNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	   C_AllowNameNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	   C_AllowCdCd2, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_AllowCdNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired     C_AllowSeq2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired     C_LendBaseNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired     C_CalcYn2, pvStartRow, pvEndRow

		.vspdData2.ReDraw = True
    
		End With
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
            C_PayCd = iCurColumnPos(1)		
            C_PayNm = iCurColumnPos(2)											
            C_AllowNameCd = iCurColumnPos(3)
            C_AllowNameNm = iCurColumnPos(4)
            C_AllowCdCd = iCurColumnPos(5)
            C_AllowCdPop =  iCurColumnPos(6)
            C_AllowCdNm = iCurColumnPos(7)
            C_AllowSeq = iCurColumnPos(8)
            C_CalcYn1 = iCurColumnPos(9)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_PayCd2 = iCurColumnPos(1)
            C_PayNm2 = iCurColumnPos(2)
            C_AllowNameCd2 = iCurColumnPos(3)
            C_AllowNameNm2 = iCurColumnPos(4)
            C_AllowCdCd2 = iCurColumnPos(5)
            C_AllowCdPop2 =  iCurColumnPos(6)
            C_AllowCdNm2 = iCurColumnPos(7)
            C_AllowSeq2 = iCurColumnPos(8)
            C_LendBaseCd2 = iCurColumnPos(9)
            C_LendBaseNm2 = iCurColumnPos(10)
            C_CalcYn2 = iCurColumnPos(11)
    End Select    
End Sub

'======================================================================================================
'	기능: ClickTab1()
'	설명: Tab Click시 필요한 기능을 수행한다.
'         Header Tab처리 부분 (Header Tab이 있는 경우만 사용)
'=======================================================================================================
Function ClickTab1()
	Dim IntRetCD
	
	If gSelframeFlg = TAB1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)                                               <%'첫번째 Tab%>
	gSelframeFlg = TAB1
    lgCurrentSpd = "M"    

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
	
End Function

'======================================================================================================
'	기능: ClickTab2()
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab2()	
	Dim IntRetCD
	
	If gSelframeFlg = TAB2 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
		
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
    lgCurrentSpd = "S"    

    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData2
	
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화 
'======================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	
    Call AppendNumberPlace("6","2","0")
            
    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    
    lgCurrentSpd = "M"
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    

    Call InitComboBox("")
    Call ClickTab1
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    Err.Clear                                                               <%'Protect system from crashing%>

    If gSelframeFlg = TAB1 Then
	    ggoSpread.Source = frm1.vspdData
    Else
    	ggoSpread.Source = frm1.vspdData2
	End If            

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
   		If IntRetCD = vbNo Then
  			Exit Function
   		End If
	End If
    
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
        															
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call MakeKeyStream("X")
    Call DisableToolBar(Parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
       
    FncQuery = True                                                              '☜: Processing is OK
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'======================================================================================================
Function FncSave() 
    
    Dim lRow
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
		
		If Not ggoSpread.SSDefaultCheck Then  		'Not chkField(Document, "2") Or
			Call changeTabs(TAB1)
			Exit Function
		End If

        With Frm1
            For lRow = 1 To .vspdData.MaxRows
                .vspdData.Row = lRow
                .vspdData.Col = 0
                Select Case .vspdData.Text
               
                    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                        .vspdData.Col = C_AllowCdNm
                        If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                            Call DisplayMsgBox("800092","X","X","X")
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if
                End Select
            Next
        End With
		
	Else
		ggoSpread.Source = frm1.vspdData2
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
		
		If Not ggoSpread.SSDefaultCheck Then  		'Not chkField(Document, "2") Or
			Call changeTabs(TAB2)
			Exit Function
		End If

        With Frm1
            For lRow = 1 To .vspdData2.MaxRows
                .vspdData2.Row = lRow
                .vspdData2.Col = 0
                Select Case .vspdData2.Text
               
                    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                        .vspdData2.Col = C_AllowCdNm2
                        If IsNull(Trim(.vspdData2.Text)) OR Trim(.vspdData2.Text) = "" Then
                            Call DisplayMsgBox("800092","X","X","X")
                            .vspdData2.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if
                End Select
            Next
        End With
	End If
          
    Call MakeKeyStream("X")

    Call DisableToolBar(Parent.TBC_SAVE)
	IF DBSAVE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True                                                          
    
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'======================================================================================================
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
    
				.ReDraw = True
    		    .Focus
			End If
		End With
	Else
	    lgCurrentSpd = "S"

        If Frm1.vspdData2.MaxRows < 1 Then
           Exit Function
        End If

        With frm1.vspdData2
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData2	
				ggoSpread.CopyRow
                SetSpreadColor .ActiveRow, .ActiveRow
    
				.ReDraw = True
    		    .Focus
			End If
		End With
	End If

    Set gActiveElement = document.ActiveElement   
	
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'======================================================================================================
Function FncCancel() 

	If gSelframeFlg = TAB1 Then
	    ggoSpread.Source = frm1.vspdData	
	Else
	    ggoSpread.Source = frm1.vspdData2	
	End If
	ggoSpread.EditUndo                                                  '☜: Protect system from crashing

	Call InitData

End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'======================================================================================================
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
	    .vspdData2.ReDraw = False
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
        ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
        SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
		.vspdData2.ReDraw = True
        End With
	End If

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	   
    Set gActiveElement = document.ActiveElement   
	
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    If gSelframeFlg = TAB1 Then
	   lgCurrentSpd = "M"
	    With frm1.vspdData 
			.focus
    		ggoSpread.Source = frm1.vspdData 
    		lDelRows = ggoSpread.DeleteRow
		End With
	Else
	   lgCurrentSpd = "S"
		With frm1.vspdData2
			.focus
    		ggoSpread.Source = frm1.vspdData2 
    		lDelRows = ggoSpread.DeleteRow
		End With
	End If    

    Set gActiveElement = document.ActiveElement   
	
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
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
            Call InitSpreadComboBox("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
            Call InitSpreadComboBox("B")
	End Select 

	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False

    If gSelframeFlg = TAB1 Then		
	    ggoSpread.Source = frm1.vspdData	
	Else
		ggoSpread.source = frm1.vspdData2
	End If

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
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
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    If gSelframeFlg = Tab1 Then
	   lgCurrentSpd = "M"
       strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '☜: Next key tag
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
    Else   
	   lgCurrentSpd = "S"
       strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '☜: Next key tag
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey1             '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData2.MaxRows         '☜: Max fetched data
    End If   
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
    Call SetToolbar("1100111100111111")										        '버튼 툴바 제어 
    If gSelframeFlg = TAB1 Then
    	Frm1.vspdData.focus
    Else
    	Frm1.vspdData2.focus
	End If            
End Function
'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'======================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
		
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) = False then
    		Exit Function 
    	End if

  	With Frm1
		.txtMode.value      = Parent.UID_M0002                                            '☜: Delete
		.txtKeyStream.value = lgKeyStream
		If gSelframeFlg = TAB1 Then
           .lgCurrentSpd.value = "M"
		Else
           .lgCurrentSpd.value = "S"
		End If
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

	If gSelframeFlg = TAB1 Then
		ggoSpread.Source = .vspdData

		For lRow = 1 To .vspdData.MaxRows
    
	        .vspdData.Row = lRow
		    .vspdData.Col = 0

			Select Case .vspdData.Text

			Case ggoSpread.InsertFlag									    '☜: 신규 
                                                 strVal = strVal & "C" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_PayCd        : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_AllowNameCd  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_AllowCdCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '4
                .vspdData.Col = C_AllowSeq     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '5
                .vspdData.Col = C_CalcYn1      

				Select Case Trim(.vspdData.Text)
				Case "Y"
					strVal = strVal & "Y" & Parent.gRowSep			'5
				Case "N"								
					strVal = strVal & "N" & Parent.gRowSep			'5
				End Select

                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag
                                                 strVal = strVal & "U" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_PayCd        : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_AllowNameCd  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_AllowCdCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '4
                .vspdData.Col = C_AllowSeq     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '5
                .vspdData.Col = C_CalcYn1      

				Select Case Trim(.vspdData.Text)
				Case "Y"
					strVal = strVal & "Y" & Parent.gRowSep			'5
				Case "N"								
					strVal = strVal & "N" & Parent.gRowSep			'5
				End Select

                lGrpCnt = lGrpCnt + 1
                    
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>
                                                 strDel = strDel & "D" & Parent.gColSep                   '0
                                                 strDel = strDel & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_PayCd        : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_AllowNameCd  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_AllowCdCd    : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep   '4
                
                lGrpCnt = lGrpCnt + 1
                
			End Select
		Next
	Else  
	    
		ggoSpread.Source = .vspdData2 

		For lRow = 1 To .vspdData2.MaxRows
    
	        .vspdData2.Row = lRow
		    .vspdData2.Col = 0
        
			Select Case .vspdData2.Text

			Case ggoSpread.InsertFlag									    <%'☜: 신규 %>
                                                 strVal = strVal & "C" & Parent.gColSep                    '0
                                                 strVal = strVal & lRow & Parent.gColSep                    '1
                .vspdData2.Col = C_PayCd2       : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '2
                .vspdData2.Col = C_AllowNameCd2 : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '3
                .vspdData2.Col = C_AllowCdCd2   : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '4
                .vspdData2.Col = C_AllowSeq2    : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '5
                .vspdData2.Col = C_LendBaseCd2  : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '6
                .vspdData2.Col = C_CalcYn2   

				Select Case Trim(.vspdData2.Text)
				Case "Y"
					strVal = strVal & "Y" & Parent.gRowSep			'7
				Case "N"								
					strVal = strVal & "N" & Parent.gRowSep			'7
				End Select

                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag
                                                 strVal = strVal & "U" & Parent.gColSep                    '0
                                                 strVal = strVal & lRow & Parent.gColSep                    '1
                .vspdData2.Col = C_PayCd2       : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '2
                .vspdData2.Col = C_AllowNameCd2 : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '3
                .vspdData2.Col = C_AllowCdCd2   : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '4
                .vspdData2.Col = C_AllowSeq2    : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '5
                .vspdData2.Col = C_LendBaseCd2  : strVal = strVal & Trim(.vspdData2.Text) & Parent.gColSep   '6
                .vspdData2.Col = C_CalcYn2     

				Select Case Trim(.vspdData2.Text)
				Case "Y"
					strVal = strVal & "Y" & Parent.gRowSep			'7
				Case "N"								
					strVal = strVal & "N" & Parent.gRowSep			'7
				End Select

                lGrpCnt = lGrpCnt + 1
                    
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>
                                                 strDel = strDel & "D" & Parent.gColSep                    '0
                                                 strDel = strDel & lRow & Parent.gColSep                    '1
                .vspdData2.Col = C_PayCd2       : strDel = strDel & Trim(.vspdData2.Text) & Parent.gColSep   '2
                .vspdData2.Col = C_AllowNameCd2 : strDel = strDel & Trim(.vspdData2.Text) & Parent.gColSep   '3
                .vspdData2.Col = C_AllowCdCd2   : strDel = strDel & Trim(.vspdData2.Text) & Parent.gRowSep   '4
                
                lGrpCnt = lGrpCnt + 1
                
			End Select
		Next
		
	End If
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 %>
	
	End With

    DbSave  = True                                                               '☜: Processing is NG
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'======================================================================================================
Function DbSaveOk()													        ' 저장 성공후 실행 로직 
	Call InitVariables

    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = Frm1.vspdData2
	Frm1.vspdData2.MaxRows = 0
    ggoSpread.ClearSpreadData

    lgCurrentSpd = "M"

    DBQuery()
End Function

'========================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If gSelframeFlg = TAB1 Then
       ggoSpread.Source = Frm1.vspdData
	   With frm1.vspdData
			Select Case iWhere
				Case C_AllowCdPop
					arrParam(0) = "수당코드 팝업"			' 팝업 명칭 
					arrParam(1) = "HDA010T"				 		' TABLE 명칭 
					.Col = C_AllowCdCd
					arrParam(2) = .value                        ' Code Condition
					.Col = C_AllowCdNm
					arrParam(3) = ""'.value				        ' Name Cindition
					arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "   "				' Where Condition
					arrParam(5) = "수당코드"			    ' TextBox 명칭 
	
					arrField(0) = "allow_cd"					' Field명(0)
					arrField(1) = "allow_nm"				    ' Field명(1)
    
					arrHeader(0) = "수당코드"				' Header명(0)
					arrHeader(1) = "수당코드명"			    ' Header명(1)
			End Select
		End With
	Else
       ggoSpread.Source = Frm1.vspdData2
	   With frm1.vspdData2
			Select Case iWhere
				Case C_AllowCdPop2
					arrParam(0) = "공제코드 팝업"			' 팝업 명칭 
					arrParam(1) = "HDA010T"				 		' TABLE 명칭 
					.Col = C_AllowCdCd2
					arrParam(2) = .value                         ' Code Condition
					.Col = C_AllowCdNm2
					arrParam(3) = ""'.value				        ' Name Cindition
					arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & "  "				' Where Condition
					arrParam(5) = "공제코드"			    ' TextBox 명칭 
	
					arrField(0) = "allow_cd"					' Field명(0)
					arrField(1) = "allow_nm"				    ' Field명(1)
    
					arrHeader(0) = "공제코드"				' Header명(0)
					arrHeader(1) = "공제코드명"			    ' Header명(1)
			End Select
		End With
	End If
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If gSelframeFlg = TAB1 Then
		   ggoSpread.Source = Frm1.vspdData
			Select Case iWhere
			    Case C_AllowCdPop
			        frm1.vspdData.Col = C_AllowCdCd
					frm1.vspdData.action =0		    		
			End Select
		Else
		   ggoSpread.Source = Frm1.vspdData2
			Select Case iWhere
			    Case C_AllowCdPop2
			        frm1.vspdData2.Col = C_AllowCdCd2
					frm1.vspdData2.action =0		    				    		
			End Select
		End If
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
        ggoSpread.UpdateRow Row
	End If	

End Function
'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

    If gSelframeFlg = TAB1 Then
       ggoSpread.Source = Frm1.vspdData
       
	   With frm1.vspdData
			Select Case iWhere
			    Case C_AllowCdPop
		    		.Col = C_AllowCdNm
		    		.text = arrRet(1)
			        .Col = C_AllowCdCd
					.text = arrRet(0) 
					.action =0		    		
			End Select
		End With
	Else
       ggoSpread.Source = Frm1.vspdData2
       
	   With frm1.vspdData2
			Select Case iWhere
			    Case C_AllowCdPop2
		    		.Col = C_AllowCdNm2
		    		.text = arrRet(1)
			        .Col = C_AllowCdCd2
					.text = arrRet(0) 
					.action =0		    				    		
			End Select
		End With
	End If

End Function

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData
        ggoSpread.Source = Frm1.vspdData

		If Row > 0 Then
			Select Case Col
                Case C_AllowCdPop
				    .Col = Col
				    .Row = Row
    				Call OpenCode("", C_AllowCdPop, Row)
			End Select
		End If
	End With
			
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData2
        ggoSpread.Source = Frm1.vspdData2

		If Row > 0 Then
			Select Case Col
                Case C_AllowCdPop2
				    .Col = Col
				    .Row = Row
    				Call OpenCode("", C_AllowCdPop2, Row)
			End Select
		End If
	End With
			
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
       
    Dim iDx, IntRetCD
    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_PayNm
                 iDx = Frm1.vspdData.value
                 Frm1.vspdData.Col   = C_PayCd
                 Frm1.vspdData.value = iDx
         Case  C_AllowNameNm
                 iDx = Frm1.vspdData.value
                 Frm1.vspdData.Col   = C_AllowNameCd
                 Frm1.vspdData.value = iDx
	     Case C_AllowCdCd
           	IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd=" & FilterVar("*", "''", "S") & "  And code_type=" & FilterVar("1", "''", "S") & "  And allow_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           	If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                Call DisplayMsgBox("800145","X","X","X")             '☜ : 등록되지 않은 코드입니다.
    	    	frm1.vspdData.Col = C_AllowCdNm
           		frm1.vspdData.Text=""
            Else
    	    	frm1.vspdData.Col = C_AllowCdNm
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

End Sub

'======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'======================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )
       
    Dim iDx, IntRetCD
    
   	Frm1.vspdData2.Row = Row
   	Frm1.vspdData2.Col = Col

    Select Case Col
         Case  C_PayNm2
                 iDx = Frm1.vspdData2.value
                 Frm1.vspdData2.Col   = C_PayCd2
                 Frm1.vspdData2.value = iDx
         Case  C_AllowNameNm2
                 iDx = Frm1.vspdData2.value
                 Frm1.vspdData2.Col   = C_AllowNameCd2
                 Frm1.vspdData2.value = iDx
 	     Case C_AllowCdCd2
           	IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd=" & FilterVar("*", "''", "S") & "  And code_type=" & FilterVar("2", "''", "S") & " And allow_cd =  " & FilterVar(frm1.vspdData2.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           	If IntRetCD=False And Trim(frm1.vspdData2.Text)<>"" Then
                Call DisplayMsgBox("800176","X","X","X")             '☜ : 등록되지 않은 코드입니다.
    	    	frm1.vspdData2.Col = C_AllowCdNm2
           		frm1.vspdData2.Text=""
            Else
    	    	frm1.vspdData2.Col = C_AllowCdNm2
            	frm1.vspdData2.Text=Trim(Replace(lgF1,Chr(11),""))
           	End If
         Case  C_LendBaseNm2
                 iDx = Frm1.vspdData2.value
                 Frm1.vspdData2.Col   = C_LendBaseCd2
                 Frm1.vspdData2.value = iDx
         Case Else
    End Select    

   	If Frm1.vspdData2.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData2.text) < CDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData2
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
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
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
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
   	ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

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
Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
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
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
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

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
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
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
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
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'======================================================================================================
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
'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
   
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData2_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData2
		
		.Row = Row
   
	End With

   	ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'	Name : cboPayCd_OnChange()
'	Description : combobox 변화시 
'========================================================================================================
Sub cboPayCd_OnChange()
    lgBlnFlgChgValue = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %> >&nbsp;<% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>급여계산수당</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>공제등록</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5>급여구분</TD>
								<TD CLASS=TD656>
								    <SELECT NAME="cboPayCd" STYLE="Width:100px;" tag="12N" ALT="급여구분"></SELECT>
								</TD>
							</TR>
							</TABLE>				
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=2 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<!-- 첫번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=2 CELLPADDING=2 CLASS="BasicTB">
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/h6008ma1_vaSpread_vspdData.js'></script>
							</TD>
						</TR>
						</TABLE>
						</DIV>	

						<!-- 두번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
						<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=2 CELLPADDING=2 CLASS="BasicTB">
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/h6008ma1_vaSpread1_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="h6008mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME> 
<!--		<TD HEIGHT=120><IFRAME NAME="MyBizASP" SRC="h6008mb1.asp" WIDTH=100% HEIGHT=100% FRAMEBORDER=1 SCROLLING=YES noresize framespacing=0></IFRAME> -->
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">


<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<!--
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
-->
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

