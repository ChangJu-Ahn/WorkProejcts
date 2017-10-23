<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h4003ma1
*  4. Program Name         : h4003ma1
*  5. Program Desc         : 근태관리/출퇴근시간등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : mok young bin
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h4003mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_ID1 = "h4003mb2.asp"                                      'Biz Logic ASP  

Const C_SHEETMAXROWS    =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>

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

Dim C_EMP_NO 
Dim C_EMP_NO_POP
Dim C_NAME 
Dim C_DEPT_CD 
Dim C_DEPT_NM 
Dim C_STRT_DATE_DT 
Dim C_STRT_HH   
Dim C_STRT_MM 
Dim C_END_DATE_DT
Dim C_END_HH 
Dim C_END_MM
Dim C_WK_TYPE_CD
Dim C_WK_TYPE 
Dim C_HOLI_TYPE_CD 
Dim C_HOLI_TYPE 
Dim C_INTERNAL_CD 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================

Sub initSpreadPosVariables()  
	C_EMP_NO  =       1
	C_EMP_NO_POP =    2
	C_NAME  =         3
	C_DEPT_CD =       4
	C_DEPT_NM =       5
	C_STRT_DATE_DT =  6  
	C_STRT_HH =       7    
	C_STRT_MM =       8  
	C_END_DATE_DT =   9 
	C_END_HH  =       10 
	C_END_MM =        11
	C_WK_TYPE_CD =    12 
	C_WK_TYPE =       13
	C_HOLI_TYPE_CD  = 14 
	C_HOLI_TYPE  =    15 
	C_INTERNAL_CD =   16
 
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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    	frm1.txtAttend_dt.Text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>
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
   
    lgKeyStream       = Frm1.txtAttend_dt.Text & parent.gColSep                                           'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtName.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtStrt_hh.text & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtStrt_mm.text & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEnd_hh.text & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEnd_mm.text & parent.gColSep
    if  Frm1.txtDept_cd.Value = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
        lgKeyStream = lgKeyStream & Frm1.txtInternal_cd.Value & parent.gColSep
    end if
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
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false
       .MaxCols = C_INTERNAL_CD + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       .MaxRows = 0
		Call GetSpreadColumnPos("A")  
	
       Call AppendNumberPlace("6","2","0")
        
        ggoSpread.SSSetEdit     C_EMP_NO,       "사번",          15,,, 13,2
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetEdit     C_NAME,         "성명",          15,,, 30,2
        ggoSpread.SSSetEdit     C_DEPT_CD,      "부서",          7,,, 10,2
        ggoSpread.SSSetEdit     C_DEPT_NM,      "부서명",          20,,, 40,2
        ggoSpread.SSSetDate     C_STRT_DATE_DT, "출근일", 11,2, parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetFloat    C_STRT_HH,      "출근(시)" ,11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
        ggoSpread.SSSetFloat    C_STRT_MM,      "시각(분)" ,11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
        ggoSpread.SSSetDate     C_END_DATE_DT,  "퇴근일", 11,2, parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetFloat    C_END_HH,       "퇴근(시)", 11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
        ggoSpread.SSSetFloat    C_END_MM,       "시각(분)" ,11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
        ggoSpread.SSSetEdit     C_WK_TYPE_CD,   "코드",          5,,, 5,2
        ggoSpread.SSSetEdit     C_WK_TYPE,      "근무조",          10,,, 10,2
        ggoSpread.SSSetEdit     C_HOLI_TYPE_CD, "코드",          5,,, 5,2
        ggoSpread.SSSetEdit     C_HOLI_TYPE,    "휴일",          10,,, 10,2
        ggoSpread.SSSetEdit     C_INTERNAL_CD,  "내부코드",          11,,, 30,2
        
        Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)	   
        
		call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)
        call ggoSpread.SSSetColHidden(C_WK_TYPE_CD,C_WK_TYPE_CD,True)
		call ggoSpread.SSSetColHidden(C_HOLI_TYPE_CD,C_HOLI_TYPE_CD,True)
		call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
        
	   .ReDraw = true
       Call SetSpreadLock 
    
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
            
			C_EMP_NO  =       iCurColumnPos(1)
			C_EMP_NO_POP =    iCurColumnPos(2)
			C_NAME  =         iCurColumnPos(3)
			C_DEPT_CD =       iCurColumnPos(4)
			C_DEPT_NM =       iCurColumnPos(5)
			C_STRT_DATE_DT =  iCurColumnPos(6) 
			C_STRT_HH =       iCurColumnPos(7)    
			C_STRT_MM =       iCurColumnPos(8) 
			C_END_DATE_DT =   iCurColumnPos(9) 
			C_END_HH  =       iCurColumnPos(10)
			C_END_MM =        iCurColumnPos(11)
			C_WK_TYPE_CD =    iCurColumnPos(12)
			C_WK_TYPE =       iCurColumnPos(13)
			C_HOLI_TYPE_CD  = iCurColumnPos(14) 
			C_HOLI_TYPE  =    iCurColumnPos(15)
			C_INTERNAL_CD =   iCurColumnPos(16)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SpreadLock      C_EMP_NO , -1, C_EMP_NO
      ggoSpread.SpreadLock      C_EMP_NO_POP , -1, C_EMP_NO_POP
      ggoSpread.SpreadLock      C_NAME , -1, C_NAME
      ggoSpread.SpreadLock      C_DEPT_CD , -1, C_DEPT_CD
      ggoSpread.SpreadLock      C_DEPT_NM , -1, C_DEPT_NM
      ggoSpread.SpreadLock      C_INTERNAL_CD , -1, C_INTERNAL_CD
      ggoSpread.SpreadLock      C_WK_TYPE_CD , -1, C_WK_TYPE_CD
      ggoSpread.SpreadLock      C_WK_TYPE , -1, C_WK_TYPE
      ggoSpread.SpreadLock      C_HOLI_TYPE_CD , -1, C_HOLI_TYPE_CD
      ggoSpread.SpreadLock      C_HOLI_TYPE , -1, C_HOLI_TYPE
      ggoSpread.SSSetRequired	C_STRT_DATE_DT, -1, -1
      ggoSpread.SSSetRequired	C_STRT_HH, -1, -1
      ggoSpread.SSSetRequired	C_STRT_MM, -1, -1
      ggoSpread.SSSetRequired	C_END_DATE_DT, -1, -1
      ggoSpread.SSSetRequired	C_END_HH  , -1, -1
      ggoSpread.SSSetRequired	C_END_MM  , -1, -1
	  ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    
      ggoSpread.SSSetProtected    C_NAME , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_EMP_NO , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DEPT_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DEPT_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_STRT_DATE_DT,pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_STRT_HH  , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_STRT_MM  , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_END_DATE_DT  , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_END_HH  , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_END_MM  , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_WK_TYPE_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_WK_TYPE , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_HOLI_TYPE_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_HOLI_TYPE ,pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_INTERNAL_CD , pvStartRow, pvEndRow
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
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
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
	Call AppendNumberPlace("6", "2", "0")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.FormatNumber(frm1.txtStrt_hh,23,0,false)
    Call ggoOper.FormatNumber(frm1.txtStrt_mm,59,0,false)
    
    Call ggoOper.FormatNumber(frm1.txtEnd_hh,23,0,false)
    Call ggoOper.FormatNumber(frm1.txtEnd_mm,59,0,false)
    
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    
    frm1.txtEmp_no.focus 
	Call CookiePage (0)
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	ggoSpread.ClearSpreadData  
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if
    
    If txtDept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
   
    Call SetSpreadLock                                   '자동입력때 풀어준 부분을 다시 조회할때 Lock시킴 
    
    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End if                                                  '☜: Query db data
       
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
    
   	Dim strStartDt
   	Dim strEndDt
   	Dim lRow

   	Dim strStartHh
   	Dim strEndHh
   	Dim strStartMm
   	Dim strEndMm
   	 
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	IntRetCD = CheckDate()  '날짜 및 출퇴근시간 체크 

	If IntRetCD = False Then
		Exit Function
    End if

	Call MakeKeyStream("X")
    
	If DbSave = False Then
	   Exit Function
	End If				                                                    '☜: Save db data
    
	FncSave = True                                                              '☜: Processing is OK

End Function

Function CheckDate()
   	Dim strStartDt
   	Dim strEndDt
   	Dim lRow

   	Dim strStartHh
   	Dim strEndHh
   	Dim strStartMm
   	Dim strEndMm

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_NAME
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
					    CheckDate = False
						Exit Function
					end if

   	                .vspdData.Col = C_STRT_DATE_DT
                    strStartDt = .vspdData.Text
                    
   	                .vspdData.Col = C_END_DATE_DT
                    strEndDt = .vspdData.Text
                    If .vspdData.Text = "" Then
                    Else
						If CompareDateByFormat(strStartDt,strEndDt,"출근일","퇴근일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
	                        .vspdData.Row = lRow
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
						    CheckDate = False
                            Exit Function
                        Else
                        End if 
                    End if  
            End Select
        Next
	End With

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_STRT_HH
                    strStartHh = .vspdData.Text
                    
   	                .vspdData.Col = C_END_HH
                    strEndHh = .vspdData.Text
                    
   	                .vspdData.Col = C_STRT_MM
                    strStartMm = .vspdData.Text
                    
   	                .vspdData.Col = C_END_MM
                    strEndMm = .vspdData.Text

                    '--  2005.08.11 -- 날짜 비교를 제대로 못함.
                    .vspdData.Col = C_STRT_DATE_DT
                    strStartDt = .vspdData.Text
                    
   	                .vspdData.Col = C_END_DATE_DT
                    strEndDt = .vspdData.Text
                    
                    If .vspdData.Text = "" Then
                    Else
                        If CInt(strStartHh) < 0 Then
	                        Call DisplayMsgBox("970021","X","출근시간","X")	'출근시간 입력필수항목입니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_STRT_HH
                            .vspdData.Action = 0 
                            Set gActiveElement = document.activeElement
						    CheckDate = False
                            Exit Function
                        End if 
                        If CInt(strEndHh) < 0 Then
	                        Call DisplayMsgBox("970021","X","퇴근시간","X")	'퇴근시간은 입력필수항목입니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_END_HH
                            .vspdData.Action = 0 
                            Set gActiveElement = document.activeElement
						    CheckDate = False
                            Exit Function
                        End if 
                        
                        If strStartDt = strEndDt Then
                            If CInt(strStartHh) > CInt(strEndHh) then
	                            Call DisplayMsgBox("800082","X","X","X")	'퇴근시각은 출근시각보다 커야 합니다.
	                            .vspdData.Row = lRow
  	                            .vspdData.Col = C_END_HH
                                .vspdData.Action = 0 
                                Set gActiveElement = document.activeElement
							    CheckDate = False
                                Exit Function
                            Elseif CInt(strStartHh) = CInt(strEndHh) then
                                If CInt(strStartMm) >= CInt(strEndMm) then
	                                Call DisplayMsgBox("800082","X","X","X")	'퇴근시각은 출근시각보다 커야 합니다.
	                                .vspdData.Row = lRow
  	                                .vspdData.Col = C_END_MM
                                    .vspdData.Action = 0 
                                    Set gActiveElement = document.activeElement
								    CheckDate = False
                                    Exit Function
                                End if
                            End if 
                        End if
                    End if  
            End Select
        Next
	End With

    CheckDate = True

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False           

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
           .Col  = C_NAME
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_EMP_NO
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DEPT_CD
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DEPT_NM
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
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow,iRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus	
	    ggoSpread.Source = .vspdData
	    
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    
	For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1        
    	.vspdData.Row = iRow	
        .vspdData.col = C_STRT_DATE_DT
        .vspdData.Text = frm1.txtAttend_dt.text
        .vspdData.col = C_END_DATE_DT
        .vspdData.Text = frm1.txtAttend_dt.text
        
        .vspdData.col = C_STRT_HH
        .vspdData.Text = frm1.txtStrt_hh.text
        .vspdData.col = C_STRT_MM
        .vspdData.Text = frm1.txtStrt_mm.text
        .vspdData.col = C_END_HH
        .vspdData.Text = frm1.txtEnd_hh.text
        .vspdData.col = C_END_MM
        .vspdData.Text = frm1.txtEnd_mm.text
    Next        	    
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
    	lDelRows = ggoSpread.DeleteRow
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

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
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel

    Dim iColSep, iRowSep
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
 	Dim iFormLimitByte						'102399byte
 	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
 	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
 	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size
    
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
 	
 	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
 	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
     
     '102399byte
     iFormLimitByte = parent.C_FORM_LIMIT_BYTE
     
     '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
 	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				
 
 	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
 	
 	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
    DbSave = False                                                          
    If  LayerShowHide(1) = False Then
		Exit Function
	End If

    lGrpCnt = 1

	With Frm1
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
					strVal = ""               
                                                    strVal = strVal & "C" & parent.gColSep 'array(0)
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & Trim(frm1.txtAttend_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_TYPE_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HOLI_TYPE_CD  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_DATE_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_HH	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_MM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_DATE_DT   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_HH	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_MM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INTERNAL_CD   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 
               Case ggoSpread.UpdateFlag                                      '☜: Update
					strVal = ""                              
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & Trim(frm1.txtAttend_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_DATE_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_HH	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STRT_MM           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_DATE_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_HH	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_END_MM            : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
					strDel = ""               

                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                                                    strDel = strDel & Trim(frm1.txtAttend_dt.Text) & parent.gColSep
                   .vspdData.Col = C_EMP_NO    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'삭제시 key만								
                    lGrpCnt = lGrpCnt + 1
           End Select
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If

			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   

			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)

			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
       Next
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
'	   .txtSpread.value      = strDel & strVal

	End With

    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
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
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100111100111111")									
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData  
	Call RemovedivTextArea    
    Call InitVariables															'⊙: Initializes local global variables
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
       Call RestoreToolBar()
       Exit Function
    End if
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	End If
    arrParam(1) = frm1.txtAttend_dt.Text
	arrParam(2) = lgUsrIntCd                              ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               .txtInternal_cd.value = arrRet(2)
               .txtDept_cd.focus
        End Select
	End With
End Function       		

'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
'	Dim arrParam(3)
	Dim arrParam(2)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Condition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
'	    arrParam(3) = frm1.txtAttend_dt.Text	    
	Else 'spread
	    frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	'    arrParam(3) = frm1.txtAttend_dt.Text		    
	End If
		
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
    Dim strAttendDt
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else 'spread
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_NM
			.vspdData.Text = arrRet(2)
 		
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			
			strAttendDt = UNIConvDate(frm1.txtAttend_dt.Text)			
            Call CommonQueryRs(" dbo.ufn_H_get_dept_cd( EMP_NO,"& strAttendDt & ")  DEPT_CD, dbo.ufn_H_get_internal_cd(EMP_NO,"& strAttendDt & ") internal_cd "," HAA010T "," EMP_NO =  " & FilterVar(arrRet(0) , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         	Frm1.vspdData.Col = C_DEPT_CD
		    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
		    	
         	Frm1.vspdData.Col = C_INTERNAL_CD
		    Frm1.vspdData.value = Trim(Replace(lgF1,Chr(11),""))
 
 			.vspdData.action =0 
		End If
	End With
End Sub

'======================================================================================================
'	Name : vspdData_ButtonClicked()
'	Description : 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    Dim strKeyStream              '근무조 , 휴일을 구하러간다.
    Dim strVal

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col

	Select Case Col
	    Case C_EMP_NO_POP
                Call OpenEmptName("1")
    End Select     

    frm1.vspdData.Col = C_STRT_DATE_DT
    strKeyStream		= frm1.vspdData.text & parent.gColSep 
    if frm1.vspdData.text = "" then		
		Exit sub 
	End If	  
	
    frm1.vspdData.Col = C_EMP_NO
    strKeyStream		= strKeyStream & frm1.vspdData.text  & parent.gColSep    
    if frm1.vspdData.text = "" then		
		Exit sub 
	End If
		
'	If  LayerShowHide(1) = False Then
'		Exit Sub
'	End If
		
	With Frm1
		strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001						         
	    strVal = strVal     & "&txtKeyStream="       & strKeyStream                       '☜: Query Key
	    strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
	    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
	End With
	Call RunMyBizASP(MyBizASP, strVal)			                                           '☜: Run Biz Logic

End Sub

'======================================================================================================
'	Name : autoInsert_ButtonClicked()
'	Description : 자동입력 
'=======================================================================================================
Sub autoInsert_ButtonClicked(Byval ButtonDown)

    Dim strKeyStream
    Dim strVal
    Dim IntRetCD 
    Dim strEmpNo
    Dim strInternalCd
    Dim strInternalCd2
    Dim strAttendDt

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")	'☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit sub
		End If
    End If

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Sub
    End If
    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Sub
    End if
    
    If txtDept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Sub
    End if

    If CInt(frm1.txtStrt_hh.Text) > CInt(frm1.txtEnd_hh.Text) then
        Call DisplayMsgBox("800082","X","X","X")	'퇴근시각은 출근시각보다 커야 합니다.
        frm1.txtEnd_hh.focus()
        Exit Sub
    Elseif CInt(frm1.txtStrt_hh.Text) = CInt(frm1.txtEnd_hh.Text) then
        If CInt(frm1.txtStrt_mm.Text) >= CInt(frm1.txtEnd_mm.Text) then
	        Call DisplayMsgBox("800082","X","X","X")	'퇴근시각은 출근시각보다 커야 합니다.
			frm1.txtEnd_mm.focus()
            Exit Sub
        End if
    End if 
	
    strEmpNo = frm1.txtEmp_no.value
    strInternalCd = frm1.txtInternal_cd.value
    strInternalCd2 = frm1.txtInternal_cd.value
    strAttendDt = UNIConvDate(frm1.txtAttend_dt.Text)
    
    If strEmpNo = "" then
        strEmpNo = "%"
    End if

    If strInternalCd = "" then
        strInternalCd = lgUsrIntCd
        Call CommonQueryRs(" COUNT(*) AS counts "," HAA010T a, HCA080T b ","  a.emp_no = b.emp_no AND a.emp_no LIKE  " & FilterVar(strEmpNo, "''", "S") & " AND a.internal_cd LIKE " & FilterVar(strInternalCd & "%", "''", "S") & " AND b.attend_dt   =  " & FilterVar(strAttendDt, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Else
        Call CommonQueryRs(" COUNT(*) AS counts "," HAA010T a, HCA080T b ","  a.emp_no = b.emp_no AND a.emp_no LIKE  " & FilterVar(strEmpNo, "''", "S") & " AND a.internal_cd =  " & FilterVar(strInternalCd, "''", "S") & " AND b.attend_dt   =   " & FilterVar(strAttendDt, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    End if

	If Trim(Replace(lgF0,Chr(11),"")) = 0 then
	Else
        Call DisplayMsgBox("800064","X","X","X")	'이미생성된 자료가 있습니다.
	    Exit Sub                                    '바로 return한다....자동입력을 멈춘다.
    End if

    If strInternalCd2 = "" then
        strInternalCd2 = lgUsrIntCd
        Call CommonQueryRs(" COUNT(*) AS counts "," HAA010T "," emp_no LIKE " & FilterVar(strEmpNo, "''", "S") & " AND internal_cd LIKE " & FilterVar(strInternalCd & "%", "''", "S")  & " AND entr_dt <= " & FilterVar(strAttendDt, "''", "S") & " AND (retire_dt IS NULL OR retire_dt >= " &  FilterVar(strAttendDt, "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Else
        Call CommonQueryRs(" COUNT(*) AS counts "," HAA010T "," emp_no LIKE " & FilterVar(strEmpNo, "''", "S") & " AND internal_cd =  " & FilterVar(strInternalCd2, "''", "S") & " AND entr_dt <=  " & FilterVar(strAttendDt, "''", "S") & " AND (retire_dt IS NULL OR retire_dt >=  " & FilterVar(strAttendDt, "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    End if

	If Trim(Replace(lgF0,Chr(11),"")) = 0 then
        Call DisplayMsgBox("800065","X","X","X")	'자동 입력할 사원이 없습니다.
	    Exit Sub                                    '바로 return한다....자동입력을 멈춘다.
	Else
    End if
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData  
	
	If  LayerShowHide(1) = False Then
		Exit Sub
	End If
	
	frm1.vspdData.MaxRows = 0
    strKeyStream       = frm1.txtAttend_dt.text & parent.gColSep 
    strKeyStream       = strKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
    if  Frm1.txtDept_cd.Value = "" then
        strKeyStream = strKeyStream & lgUsrIntCd & parent.gColSep
    else
        strKeyStream = strKeyStream & Frm1.txtInternal_cd.Value & parent.gColSep
    end if
	strKeyStream       = strKeyStream & Frm1.txtStrt_hh.text & parent.gColSep
	strKeyStream       = strKeyStream & Frm1.txtStrt_mm.text & parent.gColSep
	strKeyStream       = strKeyStream & Frm1.txtEnd_hh.text & parent.gColSep
	strKeyStream       = strKeyStream & Frm1.txtEnd_mm.text & parent.gColSep
	strKeyStream       = strKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
     
    With Frm1
    	strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0002                          'mb2 자동입력......						         
        strVal = strVal     & "&txtKeyStream="       & strKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
End Sub

Sub DBAutoQueryOk()
    Dim lRow
    
    With Frm1
        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
		For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
        Next

		ggoSpread.SpreadLock C_EMP_NO, -1,C_EMP_NO
		ggoSpread.SpreadLock C_EMP_NO_POP, -1,C_EMP_NO_POP

        .vspdData.ReDraw = TRUE
		ggoSpread.ClearSpreadData "T"            
    End With    
    Set gActiveElement = document.ActiveElement  
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strAttendDt ,strEmp_no
    
    strAttendDt = FilterVar(UNIConvDate(frm1.txtAttend_dt.Text) , "''", "S")    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_EMP_NO
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMP_NO
			strEmp_no = Frm1.vspdData.value
			
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_CD
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_NM
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_WK_TYPE_CD
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_WK_TYPE
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_HOLI_TYPE_CD
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_HOLI_TYPE
                Frm1.vspdData.value = ""
            Else
	            IntRetCd = FuncGetEmpInf2(iDx,lgUsrIntCd,strName,strDept_nm, strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)

	            if  IntRetCd < 0 then
	                if  IntRetCd = -1 then
                		Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                    else
                        Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                    end if
  	                Frm1.vspdData.Col = C_NAME
                    Frm1.vspdData.value = ""
                    
         		    Frm1.vspdData.Col = C_DEPT_CD
		       	    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_DEPT_NM
		       	    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_WK_TYPE_CD
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_WK_TYPE
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_HOLI_TYPE_CD
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_HOLI_TYPE
                    Frm1.vspdData.value = ""
                Else
		       	    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.value = strName
		       	    
                    Call CommonQueryRs(" dbo.ufn_H_get_dept_cd( EMP_NO,"& strAttendDt & ")  DEPT_CD, dbo.ufn_GetDeptName( dbo.ufn_H_get_dept_cd(EMP_NO,"& strAttendDt & "),"& strAttendDt &" ) DEPT_NM "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         		    Frm1.vspdData.Col = C_DEPT_CD
		       	    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
		       	                                  
  	                Frm1.vspdData.Col = C_DEPT_NM
		       	    Frm1.vspdData.value = Trim(Replace(lgF1,Chr(11),""))
                   
                    Dim strKeyStream              '근무조 , 휴일을 구하러간다 
                    Dim strVal
						
					frm1.vspdData.Col = C_STRT_DATE_DT
					strKeyStream		= frm1.vspdData.text & parent.gColSep 						
					frm1.vspdData.Col = C_EMP_NO
					strKeyStream		= strKeyStream & frm1.vspdData.text  & parent.gColSep		
						
                   	If  LayerShowHide(1) = False Then
						Exit Function
					End If

					With Frm1
						strVal = BIZ_PGM_ID1 & "?txtMode="      & parent.UID_M0001						         
					    strVal = strVal     & "&txtKeyStream=" & strKeyStream                       '☜: Query Key
					    strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
					    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
					End With
					Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
                End if 
            End if 
            
    End Select    
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
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
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If   
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If  
    frm1.vspdData.Row = Row   	

    
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
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
'			frm1.txtEmp_no.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtDept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtDept_cd.value = "" Then
		frm1.txtDept_nm.value = ""
		frm1.txtInternal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,UNIConvDate(frm1.txtAttend_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtDept_nm.value = ""
		    frm1.txtInternal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtDept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtDept_cd_Onchange = true
            Exit Function      
        else
           
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtInternal_cd.value = lsInternal_cd
        end if
    End if  
    
End Function


'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtAttend_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")     
        frm1.txtAttend_dt.Action = 7
        frm1.txtAttend_dt.focus
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
'=======================================================================================================
'   Event Name : txtAttend_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtAttend_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub
'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>출퇴근시간등록</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
			     		 <TD CLASS=TD5 NOWRAP>근태일</TD>       
				    	 <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h4003ma1_fpDateTime_txtAttend_dt.js'></script></TD>
						 <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
						 <TD CLASS=TD6 NOWRAP></TD>	
			           </TR>
		               <TR>		
				         <TD CLASS=TD5 NOWRAP>사원</TD>
				     	 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
				     	                      <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30 tag="14XXXU"></TD>
				    	 <TD CLASS=TD5 NOWRAP>부서코드</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                                  <INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
			                                  <INPUT NAME="txtInternal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
					   </TR>
		               <TR>		
					     <TD CLASS="TD5" NOWRAP>기본출근시간</TD>
						 <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h4003ma1_fpDoubleSingle2_txtStrt_hh.js'></script>&nbsp;:&nbsp;
						                        <script language =javascript src='./js/h4003ma1_fpDoubleSingle3_txtStrt_mm.js'></script></TD>
		                 <TD CLASS="TD5" NOWRAP>기본퇴근시간</TD>
						 <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h4003ma1_fpDoubleSingle4_txtEnd_hh.js'></script>&nbsp;:&nbsp;
						                        <script language =javascript src='./js/h4003ma1_fpDoubleSingle5_txtEnd_mm.js'></script></TD>
					   </TR>
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
	
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h4003ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
    <TR>
	    <TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: autoInsert_ButtonClicked('1')">자동입력</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
