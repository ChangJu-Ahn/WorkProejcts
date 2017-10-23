<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Human Resource
'*  2. Function Name        : 급/상여소급분관리 
'*  3. Program ID           : h8004ma1.asp
'*  4. Program Name         : h8004ma1.asp
'*  5. Program Desc         : 소급급/상여월별조회 
'*  6. Modified date(First) : 2001/05/28
'*  7. Modified date(Last)  : 2003/06/13
'*  8. Modifier (First)     : Song Bong-kyu
'*  9. Modifier (Last)      : Lee SiNa
'* 10. Comment              :
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
Const BIZ_PGM_ID = "h8004mb1.asp"                                      'Biz Logic ASP 
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
Dim lsInternal_cd

Dim C_AllowNm 
Dim C_Amount  												
Dim C_AllowNm2
Dim C_Amount2 												
Dim C_AllowNm3
Dim C_Amount3 												
Dim C_AllowNm4
Dim C_Amount4 												
Dim C_AllowNm5
Dim C_Amount5 												
Dim C_AllowNm6
Dim C_Amount6 												
Dim C_AllowNm1S 
Dim C_Amount1S  												
Dim C_AllowNm2S
Dim C_Amount2S 												
Dim C_AllowNm3S
Dim C_Amount3S 												
Dim C_AllowNm4S
Dim C_Amount4S 												
Dim C_AllowNm5S
Dim C_Amount5S 												
Dim C_AllowNm6S
Dim C_Amount6S 												
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_AllowNm   = 1
        C_Amount    = 2
        C_AllowNm1S   = 1
        C_Amount1S    = 2

    ElseIf pvSpdNo = "B" Then
        C_AllowNm2   = 1
        C_Amount2    = 2
        C_AllowNm2S   = 1
        C_Amount2S    = 2
    
    ElseIf pvSpdNo = "C" Then
        C_AllowNm3   = 1
        C_Amount3    = 2
        C_AllowNm3S   = 1
        C_Amount3S    = 2
    
    ElseIf pvSpdNo = "D" Then
        C_AllowNm4   = 1
        C_Amount4    = 2
        C_AllowNm4S   = 1
        C_Amount4S    = 2
    
    ElseIf pvSpdNo = "E" Then
        C_AllowNm5   = 1
        C_Amount5    = 2
        C_AllowNm5S   = 1
        C_Amount5S    = 2
    
    ElseIf pvSpdNo = "F" Then
        C_AllowNm6   = 1
        C_Amount6    = 2
        C_AllowNm6S   = 1
        C_Amount6S    = 2
        
    End If
    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
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
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtYymm.focus()			'년월 default value setting
	
	frm1.txtYymm.Year = strYear 		 '년월일 default value setting
	frm1.txtYymm.Month = strMonth
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
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
   Dim strYear
   Dim strMonth
   Dim strYymm

    strYear = frm1.txtYymm.year
    strMonth = frm1.txtYymm.month
    
    If len(strMonth) = 1 then
		strMonth = "0" & strMonth
	End if

	strYymm = strYear & strMonth

	lgKeyStream       = strYymm & Parent.gColSep                'You Must append one character(Parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtPayCd.Value & Parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & Parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData(ByVal pvSpdNo)
	Dim intRow
	Dim intIndex 
	    
    If pvSpdNo = "" OR Left(pvSpdNo,1) = "A" Then
        ggoSpread.Source = frm1.vspdData1S
        ggoSpread.UpdateRow 1 
	    frm1.vspdData1S.Col = 0
        frm1.vspdData1S.Text = "합계"
    
	    frm1.vspdData1S.Col = C_Amount1S
	    frm1.vspdData1S.Row = 1
	    
	    frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData1,C_Amount,1,frm1.vspdData1.MaxRows,False,-1,-1,"V")

        Call SetSpreadLock("A")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "B" Then	        
        ggoSpread.Source = frm1.vspdData2S
        ggoSpread.UpdateRow 1 
        
	    frm1.vspdData2S.Col = 0
        frm1.vspdData2S.Text = "합계"	
	    frm1.vspdData2S.Col = C_Amount2S
	    frm1.vspdData2S.Row = 1	
	    frm1.vspdData2S.Text = FncSumSheet(frm1.vspdData2,C_Amount2,1,frm1.vspdData2.MaxRows,False,-1,-1,"V")	        

        Call SetSpreadLock("B")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "C" Then	        
        ggoSpread.Source = frm1.vspdData3S
        ggoSpread.UpdateRow 1 

	    frm1.vspdData3S.Col = 0
        frm1.vspdData3S.Text = "합계"
	    frm1.vspdData3S.Col = C_Amount3S
	    frm1.vspdData3S.Row = 1
	    frm1.vspdData3S.Text = FncSumSheet(frm1.vspdData3,C_Amount3,1,frm1.vspdData3.MaxRows,False,-1,-1,"V")

        Call SetSpreadLock("C")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "D" Then	        
        ggoSpread.Source = frm1.vspdData4S
        ggoSpread.UpdateRow 1 

	    frm1.vspdData4S.Col = 0
        frm1.vspdData4S.Text = "합계"
	    frm1.vspdData4S.Col = C_Amount4S
	    frm1.vspdData4S.Row = 1
	    frm1.vspdData4S.Text = FncSumSheet(frm1.vspdData4,C_Amount4,1,frm1.vspdData4.MaxRows,False,-1,-1,"V")

        Call SetSpreadLock("D")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "E" Then	        
        ggoSpread.Source = frm1.vspdData5S
        ggoSpread.UpdateRow 1 
        
	    frm1.vspdData5S.Col = 0
        frm1.vspdData5S.Text = "합계"
	    frm1.vspdData5S.Col = C_Amount5S
	    frm1.vspdData5S.Row = 1
	    frm1.vspdData5S.Text = FncSumSheet(frm1.vspdData5,C_Amount5,1,frm1.vspdData5.MaxRows,False,-1,-1,"V")

        Call SetSpreadLock("E")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "F" Then	        
        ggoSpread.Source = frm1.vspdData6S
        ggoSpread.UpdateRow 1 
        
	    frm1.vspdData6S.Col = 0
        frm1.vspdData6S.Text = "합계"
	    frm1.vspdData6S.Col = C_Amount6S
	    frm1.vspdData6S.Row = 1
	    frm1.vspdData6S.Text = FncSumSheet(frm1.vspdData6,C_Amount6,1,frm1.vspdData6.MaxRows,False,-1,-1,"V")

        Call SetSpreadLock("F")
    End If
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "A" Then

        Call initSpreadPosVariables("A")   'sbk 

        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
    	    With frm1.vspdData1
                ggoSpread.Source = frm1.vspdData1
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call GetSpreadColumnPos("A") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm    , "원지급수당", 16
                ggoSpread.SSSetFloat C_Amount     , "수당액", 16, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then
    	    With frm1.vspdData1S
                ggoSpread.Source = Frm1.vspdData1S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount1S + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData

               .ScrollBars   = 0
               .DisplayColHeaders = False

                Call GetSpreadColumnPos("B") 'sbk

               .Col = 0 
               .Row = 1
               .Text = "합계"

                ggoSpread.SSSetEdit  C_AllowNm1S      ,"", 16
                ggoSpread.SSSetFloat C_Amount1S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With            
        End If

        Call SetSpreadLock("A")    
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "B" Then
        Call initSpreadPosVariables("B")
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
      	    With frm1.vspdData2
                ggoSpread.Source = Frm1.vspdData2
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false
    	       
               .MaxCols = C_Amount2 + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call GetSpreadColumnPos("C") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm2    , "인상분수당", 16
                ggoSpread.SSSetFloat C_Amount2     , "수당액", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then        
    	    With frm1.vspdData2S
                ggoSpread.Source = Frm1.vspdData2S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount2S  + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData
    	
               .ScrollBars   = 0
               .DisplayColHeaders = False

               .Col = 0 
               .Row = 1
               .Text = "합계"

                Call GetSpreadColumnPos("D") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm2S      ,"", 16
                ggoSpread.SSSetFloat C_Amount2S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With
        End If

        Call SetSpreadLock("B")    
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "C" Then
        Call initSpreadPosVariables("C")
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
    	    With frm1.vspdData3
                ggoSpread.Source = Frm1.vspdData3
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount3 + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>

    	       .Col = .MaxCols                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call GetSpreadColumnPos("E") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm3    , "소급분수당", 16
                ggoSpread.SSSetFloat C_Amount3     , "수당액", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then                
    	    With frm1.vspdData3S
                ggoSpread.Source = Frm1.vspdData3S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount3S + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData

               .ScrollBars   = 0
               .DisplayColHeaders = False

                Call GetSpreadColumnPos("F") 'sbk

               .Col = 0 
               .Row = 1
               .Text = "합계"

                ggoSpread.SSSetEdit  C_AllowNm3S      ,"", 16
                ggoSpread.SSSetFloat C_Amount3S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With
        End If

        Call SetSpreadLock("C")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "D" Then
        Call initSpreadPosVariables("D")
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
    	    With frm1.vspdData4
                ggoSpread.Source = Frm1.vspdData4
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount4 + 1										<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData
    	
                Call GetSpreadColumnPos("G") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm4    , "원지급공제", 16
                ggoSpread.SSSetFloat C_Amount4     , "공제액", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then                
    	    With frm1.vspdData4S
                ggoSpread.Source = Frm1.vspdData4S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false
    	
               .MaxCols = C_Amount4S  + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData

               .ScrollBars   = 0
               .DisplayColHeaders = False
    	
                Call GetSpreadColumnPos("H") 'sbk

               .Col = 0 
               .Row = 1
               .Text = "합계"

                ggoSpread.SSSetEdit  C_AllowNm4S      ,"", 16
                ggoSpread.SSSetFloat C_Amount4S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With
        End If

        Call SetSpreadLock("D")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "E" Then
        Call initSpreadPosVariables("E")

        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
    	    With frm1.vspdData5
                ggoSpread.Source = Frm1.vspdData5
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount5 + 1										<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData
    	        	
                Call GetSpreadColumnPos("I") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm5    , "인상분공제", 16
                ggoSpread.SSSetFloat C_Amount5     , "공제액", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then
    	    With frm1.vspdData5S
                ggoSpread.Source = Frm1.vspdData5S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount5S + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData
    	
               .ScrollBars   = 0
               .DisplayColHeaders = False
    	        	
                Call GetSpreadColumnPos("J") 'sbk

               .Col = 0 
               .Row = 1
               .Text = "합계"

                ggoSpread.SSSetEdit  C_AllowNm5S      ,"", 16
                ggoSpread.SSSetFloat C_Amount5S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With
        End If

        Call SetSpreadLock("E")
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "F" Then
        Call initSpreadPosVariables("F")

        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "1" Then
    	    With frm1.vspdData6
                ggoSpread.Source = Frm1.vspdData6
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount6 + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>

    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True                                                            ' ☜:☜:

               .MaxRows = 0
                ggoSpread.ClearSpreadData
    	        	
                Call GetSpreadColumnPos("K") 'sbk

                ggoSpread.SSSetEdit  C_AllowNm6    , "소급분공제", 16
                ggoSpread.SSSetFloat C_Amount6     , "공제액", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	    	.ReDraw = true	
            End With
        End If
        
        If pvSpdNo = "" OR Mid(pvSpdNo,2,1) = "S" Then
    	    With frm1.vspdData6S
                ggoSpread.Source = Frm1.vspdData6S
                ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_Amount6S + 1										'☜: 최대 Columns의 항상 1개 증가시킴 %>	   
    	       .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
               .ColHidden = True           

               .MaxRows = 0
                ggoSpread.ClearSpreadData

               .ScrollBars   = 0
               .DisplayColHeaders = False
    	        	
                Call GetSpreadColumnPos("L") 'sbk

               .Col = 0 
               .Row = 1
               .Text = "합계"

                ggoSpread.SSSetEdit  C_AllowNm6S      ,"", 16
                ggoSpread.SSSetFloat C_Amount6S       ,"", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

    	       .ReDraw = true	
            End With
        End If

        Call SetSpreadLock("F")
    End If
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "A" Then
        ggoSpread.Source = Frm1.vspdData1   
        ggoSpread.SpreadLockWithOddEvenRowColor()
        
        ggoSpread.Source = Frm1.vspdData1S
        With frm1.vspdData1S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm1S , -1, C_AllowNm1S, -1
             ggoSpread.SpreadLock      C_Amount1S , -1, C_Amount1S, -1
             ggoSpread.SSSetProtected   .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "B" Then
        ggoSpread.Source = Frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
        
        ggoSpread.Source = Frm1.vspdData2S
        With frm1.vspdData2S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm2S , -1, C_AllowNm2S, -1
             ggoSpread.SpreadLock      C_Amount2S , -1, C_Amount2S, -1
             ggoSpread.SSSetProtected  .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "C" Then    
        ggoSpread.Source = Frm1.vspdData3
        ggoSpread.SpreadLockWithOddEvenRowColor()

        ggoSpread.Source = Frm1.vspdData3S
        With frm1.vspdData3S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm3S , -1, C_AllowNm3S, -1
             ggoSpread.SpreadLock      C_Amount3S , -1, C_Amount3S, -1
             ggoSpread.SSSetProtected  .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "D" Then
        ggoSpread.Source = Frm1.vspdData4
        ggoSpread.SpreadLockWithOddEvenRowColor()
      
        ggoSpread.Source = Frm1.vspdData4S
        With frm1.vspdData4S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm4S , -1, C_AllowNm4S, -1
             ggoSpread.SpreadLock      C_Amount4S , -1, C_Amount4S, -1
             ggoSpread.SSSetProtected  .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "E" Then    
        ggoSpread.Source = Frm1.vspdData5
        ggoSpread.SpreadLockWithOddEvenRowColor()
      
        ggoSpread.Source = Frm1.vspdData5S
        With frm1.vspdData5S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm5S , -1, C_AllowNm5S, -1
             ggoSpread.SpreadLock      C_Amount5S , -1, C_Amount5S, -1
             ggoSpread.SSSetProtected  .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

    If pvSpdNo = "" OR Left(pvSpdNo,1) = "F" Then    
        ggoSpread.Source = Frm1.vspdData6
        ggoSpread.SpreadLockWithOddEvenRowColor()

        ggoSpread.Source = Frm1.vspdData6S
        With frm1.vspdData6S
	    	.ReDraw = False
             ggoSpread.SpreadLock      C_AllowNm6S , -1, C_AllowNm6S, -1
             ggoSpread.SpreadLock      C_Amount6S , -1, C_Amount6S, -1
             ggoSpread.SSSetProtected  .MaxCols   , -1, -1
	    	.ReDraw = True
        End With
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
         ggoSpread.Source = Frm1.vspdData1
        .vspdData1.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount  , pvStartRow, pvEndRow
        .vspdData1.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData1S
        .vspdData1S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm1S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount1S  , pvStartRow, pvEndRow
        .vspdData1S.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData2
        .vspdData2.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm2 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount2  , pvStartRow, pvEndRow
        .vspdData2.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData2S
        .vspdData2S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm2S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount2S  , pvStartRow, pvEndRow
        .vspdData2S.ReDraw = True

         ggoSpread.Source = Frm1.vspdData3
        .vspdData3.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm3 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount3  , pvStartRow, pvEndRow
        .vspdData3.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData3S
        .vspdData3S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm3S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount3S  , pvStartRow, pvEndRow
        .vspdData3S.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData4
        .vspdData4.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm4 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount4  , pvStartRow, pvEndRow
        .vspdData4.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData4S
        .vspdData4S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm4S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount4S  , pvStartRow, pvEndRow
        .vspdData4S.ReDraw = True

         ggoSpread.Source = Frm1.vspdData5
        .vspdData5.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm5 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount5  , pvStartRow, pvEndRow
        .vspdData5.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData5S
        .vspdData5S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm5S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount5S  , pvStartRow, pvEndRow
        .vspdData5S.ReDraw = True

         ggoSpread.Source = Frm1.vspdData6
        .vspdData6.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm6 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount6  , pvStartRow, pvEndRow
        .vspdData6.ReDraw = True
    
         ggoSpread.Source = Frm1.vspdData6S
        .vspdData6S.ReDraw = False
         ggoSpread.SSSetProtected    C_AllowNm6S , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected    C_Amount6S  , pvStartRow, pvEndRow
        .vspdData6S.ReDraw = True

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
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_AllowNm   = iCurColumnPos(1)
            C_Amount    = iCurColumnPos(2)															
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_AllowNm1S   = iCurColumnPos(1)
            C_Amount1S    = iCurColumnPos(2)															
    
       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm2  = iCurColumnPos(1)
            C_Amount2   = iCurColumnPos(2)

       Case "D"
            ggoSpread.Source = frm1.vspdData2S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm2S  = iCurColumnPos(1)
            C_Amount2S   = iCurColumnPos(2)	
    
       Case "E"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm3  = iCurColumnPos(1)
            C_Amount3   = iCurColumnPos(2)	
     
       Case "F"
            ggoSpread.Source = frm1.vspdData3S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm3S  = iCurColumnPos(1)
            C_Amount3S   = iCurColumnPos(2)	
   
       Case "G"
            ggoSpread.Source = frm1.vspdData4
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm4  = iCurColumnPos(1)
            C_Amount4   = iCurColumnPos(2)	
     
       Case "H"
            ggoSpread.Source = frm1.vspdData4S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm4S  = iCurColumnPos(1)
            C_Amount4S   = iCurColumnPos(2)	
  
       Case "I"
            ggoSpread.Source = frm1.vspdData5
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm5  = iCurColumnPos(1)
            C_Amount5   = iCurColumnPos(2)	
    
       Case "J"
            ggoSpread.Source = frm1.vspdData5S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm5S  = iCurColumnPos(1)
            C_Amount5S   = iCurColumnPos(2)	
  
       Case "K"
            ggoSpread.Source = frm1.vspdData6
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm6  = iCurColumnPos(1)
            C_Amount6   = iCurColumnPos(2)	
  
       Case "L"
            ggoSpread.Source = frm1.vspdData6S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_AllowNm6S  = iCurColumnPos(1)
            C_Amount6S   = iCurColumnPos(2)															
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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData1.Col = iDx
              Frm1.vspdData1.Row = iRow
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
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
 
    Call ggoOper.FormatDate(frm1.txtYymm, Parent.gDateFormat, 2)
           
    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if

	If txtPayCd_Onchange() Then         
        Exit Function
    End if

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
	Call DisableToolBar(Parent.TBC_QUERY)
	
	If DbQuery = False Then
		Call RestoreToolBar()
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
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
       
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
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
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
            
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	gActiveSpdSheet.id = "vaSpread1" Then
		ggoSpread.Source = frm1.vspdData1S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	gActiveSpdSheet.id = "vaSpread2" Then
		ggoSpread.Source = frm1.vspdData2S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	gActiveSpdSheet.id = "vaSpread3" Then
		ggoSpread.Source = frm1.vspdData3S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	gActiveSpdSheet.id = "vaSpread4" Then
		ggoSpread.Source = frm1.vspdData4S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	gActiveSpdSheet.id = "vaSpread5" Then
		ggoSpread.Source = frm1.vspdData5S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	gActiveSpdSheet.id = "vaSpread6" Then
		ggoSpread.Source = frm1.vspdData6S 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    If IsEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
    End If
	    
    Select Case gActiveSpdSheet.id
		Case "vaSpread1"
			Call InitSpreadSheet("A1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData1S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("AS")
		    ggoSpread.Source = frm1.vspdData1S 
	        Call ggoSpread.ReOrderingSpreadData()

			ggoSpread.Source = frm1.vspdData1S
			ggoSpread.InsertRow 
        	Call InitData("A")
		Case "vaSpread2"
			Call InitSpreadSheet("B1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData2S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("BS")
		    ggoSpread.Source = frm1.vspdData2S 
	        Call ggoSpread.ReOrderingSpreadData()

			ggoSpread.Source = frm1.vspdData2S
			ggoSpread.InsertRow  
	       	Call InitData("B")
		Case "vaSpread3"
			Call InitSpreadSheet("C1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData3S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("CS")
		    ggoSpread.Source = frm1.vspdData3S 
	        Call ggoSpread.ReOrderingSpreadData()

			ggoSpread.Source = frm1.vspdData3S
			ggoSpread.InsertRow
        	Call InitData("C")
		Case "vaSpread4"
			Call InitSpreadSheet("D1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData4S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("DS")
		    ggoSpread.Source = frm1.vspdData4S 
	        Call ggoSpread.ReOrderingSpreadData()

			ggoSpread.Source = frm1.vspdData4S
			ggoSpread.InsertRow
        	Call InitData("D")
		Case "vaSpread5"
			Call InitSpreadSheet("E1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData5S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("ES")
		    ggoSpread.Source = frm1.vspdData5S 
	        Call ggoSpread.ReOrderingSpreadData()
	        
			ggoSpread.Source = frm1.vspdData5S
			ggoSpread.InsertRow 
        	Call InitData("E")
		Case "vaSpread6"
			Call InitSpreadSheet("F1")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData6S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("FS")
		    ggoSpread.Source = frm1.vspdData6S 
	        Call ggoSpread.ReOrderingSpreadData()
			ggoSpread.Source = frm1.vspdData6S
			ggoSpread.InsertRow
        	Call InitData("F")
	End Select 

End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
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
    
    If LayerShowHide(1) = False Then
	    Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
	lgCurrentSpd = "1"
    
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows          '☜: Max fetched data
    
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic

    DbQuery = True                                                                   '☜: Processing is NG
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call InsertSum()

    Call InitData("")
	Call SetToolbar("1100000000001111")										'⊙: Set ToolBar

	Call DbQueryTotal                                                           ' 총액(single)에 뿌려주는 Sub
	frm1.vspdData1.focus
End Function

Function InsertSum()	
	ggoSpread.Source = frm1.vspdData1S
	ggoSpread.InsertRow   
	 
	ggoSpread.Source = frm1.vspdData2S
	ggoSpread.InsertRow  
	  
	ggoSpread.Source = frm1.vspdData3S
	ggoSpread.InsertRow
	
	ggoSpread.Source = frm1.vspdData4S
	ggoSpread.InsertRow
	
	ggoSpread.Source = frm1.vspdData5S
	ggoSpread.InsertRow   
	
	ggoSpread.Source = frm1.vspdData6S
	ggoSpread.InsertRow 
End Function
  
'========================================================================================================
' Function Name : DbQueryNo
' Function Desc : Called by MB Area when query operation is not successful
'========================================================================================================
Function DbQueryNo()	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
	Call DisableToolBar(Parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'======================================================================================================
'	Name : DbQueryTotal()
'	Description : 총액 필드 계산 
'=======================================================================================================
Sub DbQueryTotal()
    Dim Allow_amt, Sub_amt
    
    frm1.vspdData1S.Col = C_Amount1S 
    frm1.vspdData1S.Row = 1
    Frm1.txtOrgAllowAmt.value = frm1.vspdData1S.Text

    frm1.vspdData2S.Col = C_Amount2S
    frm1.vspdData2S.Row = 1
    Frm1.txtIncAllowAmt.value = frm1.vspdData2S.Text

    frm1.vspdData4S.Col = C_Amount4S 
    frm1.vspdData4S.Row = 1
    Frm1.txtOrgSubAmt.value = frm1.vspdData4S.Text

    frm1.vspdData5S.Col = C_Amount5S 
    frm1.vspdData5S.Row = 1
    Frm1.txtIncSubAmt.value = frm1.vspdData5S.Text

    Allow_amt = UNICDbl(Frm1.txtIncAllowAmt.value) - UNICDbl(Frm1.txtOrgAllowAmt.value)
    Sub_amt   = UNICDbl(Frm1.txtIncSubAmt.value)   - UNICDbl(Frm1.txtOrgSubAmt.value)
 
    Frm1.txtAllowAmt.value = UNIFormatNumber(Allow_amt, ggAmtOfMoney.DecPoint,-2,0,parent.ggAmtOfMoney.RndPolicy,parent.ggAmtOfMoney.RndUnit)
    Frm1.txtSubAmt.value   = UNIFormatNumber(Sub_amt, ggAmtOfMoney.DecPoint,-2,0,parent.ggAmtOfMoney.RndPolicy,parent.ggAmtOfMoney.RndUnit)
    
End Sub

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
    arrParam(0) = "지급구분팝업"			' 팝업 명칭 
    arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
    arrParam(2) = frm1.txtPayCd.value		    ' Code Condition
    arrParam(3) = ""							' Name Cindition
    arrParam(4) = "MAJOR_CD = " & FilterVar("H0040", "''", "S") & ""			' Where Condition
    arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"					' Field명(0)
    arrField(1) = "MINOR_NM"				    ' Field명(1)
    
    arrHeader(0) = "지급구분"				' Header명(0)
    arrHeader(1) = "지급명"			        ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPayCd.focus	
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
        .txtPayCd.value = arrRet(0)
        .txtPayNm.value = arrRet(1)		
        .txtPayCd.focus
	End With
End Sub

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value  = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If

		Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		lgBlnFlgChgValue = False
	End With
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

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
	    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                              strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement

		    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
            txtEmp_no_Onchange = true                           
            Exit Function   
	    Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtPayCd_Onchange             
'   Event Desc :
'========================================================================================================
function txtPayCd_Onchange()
    Dim IntRetCd
    
    If frm1.txtPayCd.value = "" Then
		frm1.txtPayNm.value = ""
    Else
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtPayCd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call DisplayMsgBox("800054","X","X","X")	'등록되지 않은 코드입니다.
			 frm1.txtPayNm.value = ""
             frm1.txtPayCd.focus
            Set gActiveElement = document.ActiveElement   
            txtPayCd_Onchange = true    
        Else
			frm1.txtPayNm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End function

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

    gMouseClickStatus = "SPC"   

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

'======================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

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
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

    gMouseClickStatus = "SP2C"   

    Set gActiveSpdSheet = frm1.vspdData3
   
    If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData3
       
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
Sub vspdData4_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

    gMouseClickStatus = "SP3C"   

    Set gActiveSpdSheet = frm1.vspdData4
   
    If frm1.vspdData4.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData4
       
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
Sub vspdData5_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

    gMouseClickStatus = "SP4C"   

    Set gActiveSpdSheet = frm1.vspdData5
   
    If frm1.vspdData5.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData5
       
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
Sub vspdData6_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000001111")

    gMouseClickStatus = "SP5C"   

    Set gActiveSpdSheet = frm1.vspdData6
   
    If frm1.vspdData6.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData6
       
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

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData4S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData5S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData6S_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    Set gActiveSpdSheet = frm1.vspdData1S
   
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData3_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
    
End Sub    

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData4_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP3C" Then
       gMouseClickStatus = "SP3CR"
    End If
    
End Sub    

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData5_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP4C" Then
       gMouseClickStatus = "SP4CR"
    End If
    
End Sub    

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData6_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP5C" Then
       gMouseClickStatus = "SP5CR"
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData1.Col = pvCol1
    frm1.vspdData1S.ColWidth(pvCol1) = frm1.vspdData1.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData1S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData2.Col = pvCol1
    frm1.vspdData2S.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData2S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData3.Col = pvCol1
    frm1.vspdData3S.ColWidth(pvCol1) = frm1.vspdData3.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData3S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData4.Col = pvCol1
    frm1.vspdData4S.ColWidth(pvCol1) = frm1.vspdData4.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData4S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData5_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData5.Col = pvCol1
    frm1.vspdData5S.ColWidth(pvCol1) = frm1.vspdData5.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData5S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData6_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData6
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData6.Col = pvCol1
    frm1.vspdData6S.ColWidth(pvCol1) = frm1.vspdData6.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData6S
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

    ggoSpread.Source = frm1.vspdData1S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

    ggoSpread.Source = frm1.vspdData1
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
    Call GetSpreadColumnPos("C")

    ggoSpread.Source = frm1.vspdData2S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("E")

    ggoSpread.Source = frm1.vspdData3S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("F")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData3S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData3S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("E")

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("F")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData4_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("G")

    ggoSpread.Source = frm1.vspdData4S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("H")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData4S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData4S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("G")

    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("H")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData5_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("I")

    ggoSpread.Source = frm1.vspdData5S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("J")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData5S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData5S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("I")

    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("J")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData6_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData6
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("K")

    ggoSpread.Source = frm1.vspdData6S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("L")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData6S_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData6S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("K")

    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("L")
End Sub

'=======================================================================================================
'   Event Name : txtYymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYymm.Action = 7
        frm1.txtYymm.focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtpay_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtYymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>소급급/상여월별조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>조회년월</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/h8004ma1_txtYymm_txtYymm.js'></script>
									</TD>
									<TD CLASS=TD5>지급구분</TD>
									<TD CLASS="TD6" NOWRAP>
									   <INPUT NAME="txtPayCd" MAXLENGTH=3 SIZE=10 ALT ="지급구분" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
						               <INPUT NAME="txtPayNm" MAXLENGTH=20 SIZE=20 ALT ="지급구분명" tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR HEIGHT=120>
					<TD WIDTH=33% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread1_vspdData1.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread2_vspdData2.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread3_vspdData3.js'></script>
							</TD>
						</TR>
						<TR HEIGHT=22>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread1S_vspdData1S.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread2S_vspdData2S.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread3S_vspdData3S.js'></script>
							</TD>
						</TR>
						<TR>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread4_vspdData4.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread5_vspdData5.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread6_vspdData6.js'></script>
							</TD>
						</TR>						
						<TR HEIGHT=22>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread4S_vspdData4S.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread5S_vspdData5S.js'></script>
							</TD>
							<TD WIDTH="33%" >
								<script language =javascript src='./js/h8004ma1_vaSpread6S_vspdData6S.js'></script>
							</TD>
						</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_50%>>
						<TR>  
							<TD CLASS=TDT NOWRAP>
							   수당차액<script language =javascript src='./js/h8004ma1_txtAllowAmt_txtAllowAmt.js'></script>
							   =인상분수당액<script language =javascript src='./js/h8004ma1_txtIncAllowAmt_txtIncAllowAmt.js'></script>
							   -원지급수당액<script language =javascript src='./js/h8004ma1_txtOrgAllowAmt_txtOrgAllowAmt.js'></script></TD>
						</TR>
						<TR>
							<TD CLASS=TDT NOWRAP>
							    공제차액<script language =javascript src='./js/h8004ma1_txtSubAmt_txtSubAmt.js'></script>
							    =인상분공제액<script language =javascript src='./js/h8004ma1_txtIncSubAmt_txtIncSubAmt.js'></script>
							    -원지급공제액<script language =javascript src='./js/h8004ma1_txtOrgSubAmt_txtOrgSubAmt.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="h8004mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
<!--		<TD HEIGHT=100><IFRAME NAME="MyBizASP" SRC="h8004mb1.asp" WIDTH=100% HEIGHT=100% FRAMEBORDER=1 SCROLLING=YES noresize framespacing=0></IFRAME> -->
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd"   TAG="24">
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

