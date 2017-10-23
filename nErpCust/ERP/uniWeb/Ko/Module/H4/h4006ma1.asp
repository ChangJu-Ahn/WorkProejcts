<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 
*  3. Program ID           	: h4006ma1
*  4. Program Name         	: h4006ma1
*  5. Program Desc         	: 근태관리/기간근태등록 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/
*  8. Modified date(Last)  	: 2003/06/11
*  9. Modifier (First)     	: mok young bin
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	:
======================================================================================================-->
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
Const BIZ_PGM_ID = "h4006mb1.asp"                                      'Biz Logic ASP 
Const C_SHEETMAXROWS =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>

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
Dim C_NAME_POP
Dim C_ROLL_PSTN 
Dim C_ROLL_PSTN_NM
Dim C_DILIG_CD 
Dim C_DILIG_NM
Dim C_DILIG_POP 
Dim C_DILIG_STRT_DT
Dim C_DILIG_END_DT 
Dim C_REMARK 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_EMP_NO  =       1
	C_EMP_NO_POP =    2
	C_NAME  =         3
	C_NAME_POP =      4
	C_ROLL_PSTN =     5
	C_ROLL_PSTN_NM =  6
	C_DILIG_CD =      7
	C_DILIG_NM =      8
	C_DILIG_POP =     9
	C_DILIG_STRT_DT = 10
	C_DILIG_END_DT =  11
	C_REMARK =        12

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
	
    Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtDilig_dt.focus 	
	frm1.txtDilig_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtDilig_dt.Month = strMonth 
	frm1.txtDilig_dt.Day = strDay
    
End Sub
	
'========================================================================================================
' Name : nfTB19029()	
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
   
    lgKeyStream       = Frm1.txtDilig_dt.Text & parent.gColSep                                           'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
    if  Frm1.txtDept_cd.Value = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
	    lgKeyStream       = lgKeyStream & Frm1.txtInternal_cd.value & parent.gColSep
    end if
	lgKeyStream       = lgKeyStream & Frm1.txtName.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtDilig_cd.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ROLL_PSTN
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ROLL_PSTN_NM         ''''''''DB에서 불러 gread에서 

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
		    ' Combo 일경우 
			.Row = intRow
			.Col = C_ROLL_PSTN
			intIndex = .value
			.col = C_ROLL_PSTN_NM
			.value = intindex
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
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    

	   .ReDraw = false
       .MaxCols = C_REMARK + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       .MaxRows = 0
		Call GetSpreadColumnPos("A") 		
		
        ggoSpread.SSSetEdit     C_NAME,          "성명",          17,,, 30,2
        ggoSpread.SSSetButton   C_NAME_POP
        ggoSpread.SSSetEdit     C_EMP_NO,        "사번",          15,,, 13,2
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetCombo    C_ROLL_PSTN,     "직위코드",      10
        ggoSpread.SSSetCombo    C_ROLL_PSTN_NM,  "직위",          15
        ggoSpread.SSSetEdit     C_DILIG_CD,      "근태코드",      10,,, 2,2
        ggoSpread.SSSetEdit     C_DILIG_NM,      "근태코드명",    16,,,20,2
        ggoSpread.SSSetButton   C_DILIG_POP
        ggoSpread.SSSetDate     C_DILIG_STRT_DT, "근태시작일",    15,2, parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetDate     C_DILIG_END_DT,  "근태종료일",    15,2, parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetEdit     C_REMARK,        "비고",          18,,, 20,2

        Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)
        
        call ggoSpread.SSSetColHidden(C_ROLL_PSTN,C_ROLL_PSTN,True)
        call ggoSpread.SSSetColHidden(C_DILIG_CD,C_DILIG_CD,True)
        call ggoSpread.SSSetColHidden(C_NAME_POP,C_NAME_POP,True)        
      
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
      ggoSpread.SpreadLock      C_NAME , -1, C_NAME
      ggoSpread.SpreadLock      C_NAME_POP , -1, C_NAME_POP
      ggoSpread.SpreadLock      C_EMP_NO , -1, C_EMP_NO
      ggoSpread.SpreadLock      C_EMP_NO_POP , -1, C_EMP_NO_POP
      ggoSpread.SpreadLock      C_ROLL_PSTN , -1, C_ROLL_PSTN
      ggoSpread.SpreadLock      C_ROLL_PSTN_NM , -1, C_ROLL_PSTN_NM
      ggoSpread.SpreadLock      C_DILIG_NM , -1, C_DILIG_NM
      ggoSpread.SpreadLock      C_DILIG_POP , -1, C_DILIG_POP
      ggoSpread.SpreadLock      C_DILIG_STRT_DT, -1, -1
      ggoSpread.SSSetRequired	C_DILIG_END_DT, -1, -1
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
      ggoSpread.SSSetProtected   C_NAME , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_EMP_NO , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_ROLL_PSTN , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_ROLL_PSTN_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_NM ,pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_STRT_DT , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_END_DT , pvStartRow, pvEndRow
      
    .vspdData.ReDraw = True
    
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
			C_NAME_POP =      iCurColumnPos(4)
			C_ROLL_PSTN =     iCurColumnPos(5)
			C_ROLL_PSTN_NM =  iCurColumnPos(6)
			C_DILIG_CD =      iCurColumnPos(7)
			C_DILIG_NM =      iCurColumnPos(8)
			C_DILIG_POP =     iCurColumnPos(9)
			C_DILIG_STRT_DT = iCurColumnPos(10)
			C_DILIG_END_DT =  iCurColumnPos(11)
			C_REMARK =        iCurColumnPos(12)
            
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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")	'⊙: Lock Field
    Call ggoOper.FormatDate(frm1.txtDilig_dt, parent.gDateFormat, 1)			
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 

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
    
    If txtDilig_cd_Onchange() Then        'enter key 로 조회시 근태코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
    Call DisableToolBar(parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
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
   	Dim strStrtDt
   	Dim strEndDt
   	Dim lRow
   	Dim strWhere
   	Dim strVDate
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

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_NAME

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
					.vspdData.Col = C_DILIG_CD

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
						Call DisplayMsgBox("800099", "X","X","x")
						Exit Function
					end if					
   	                .vspdData.Col = C_DILIG_STRT_DT
                    strStrtDt = .vspdData.Text
                    
   	                .vspdData.Col = C_DILIG_END_DT
                    strEndDt = .vspdData.Text
    
                    If .vspdData.Text = "" Then
                    Else
                        if not CompareDateByFormat(strStrtDt, strEndDt,"","", "970025", parent.gDateFormat, parent.gComDateType, false) then
	                        Call DisplayMsgBox("970025","X","근태시작일","근태종료일")	'근태시작일은 종료일보다 작아야합니다.
   							.vspdData.Col = C_DILIG_STRT_DT
							.vspdData.action =0
                            Exit Function
                        Else
                        End if 
                    End if 
                   
                    strVDate = UNIConvDate(strEndDt) '2003-12-11 lsn strStrtDt->strEndDt
					frm1.vspdData.col = C_EMP_NO
					strWhere = " EMP_NO=" & FilterVar(frm1.vspdData.Text, "''", "S") 
					strWhere = strWhere & " AND isnull(retire_dt," & FilterVar(strVDate, "''", "S") & ") < " & FilterVar(strVDate, "''", "S")

					intRetCD = CommonQueryRs(" Retire_dt " ," haa010t a",strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   

					IF intRetCD= true Then  '퇴직자의 경우 퇴직일 다음날 부터 입력 불가 
						frm1.vspdData.col = C_NAME
						Call DisplayMsgBox("800494","X",frm1.vspdData.text,UNIDateClientFormat(Trim(Replace(lgF0,Chr(11),""))))
						Set gActiveElement = document.activeElement
						exit function
					end if 
            End Select
        Next
	End With
	
    Dim strInput_emp_no
    Dim strClose_type
    Dim strClose_dt
    Dim strDilig_str_date
    Dim strDilig_end_date
    
    Dim counts
    Dim i
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag , ggoSpread.DeleteFlag , ggoSpread.UpdateFlag  

   	                .vspdData.Col = C_DILIG_STRT_DT
                    strDilig_str_date = .vspdData.Text
                                    
 	                .vspdData.Col = C_DILIG_END_DT
                    strDilig_end_date = .vspdData.Text
                                    
                    Call CommonQueryRs(" close_type, close_dt, emp_no, COUNT(close_dt) as counts "," hda270t ","  ORG_CD = " & FilterVar("1", "''", "S") & "  AND PAY_GUBUN = " & FilterVar("Z", "''", "S") & "  AND PAY_TYPE  = " & FilterVar("#", "''", "S") & "   GROUP BY emp_no,close_type,close_dt" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					If Trim(Replace(lgF3,Chr(11),""))="" Then
					Else
						counts = Trim(Replace(lgF3,Chr(11),""))
						For i = 1 to counts
						    strInput_emp_no = Trim(Replace(lgF2,Chr(11),""))
						    strClose_type = Trim(Replace(lgF0,Chr(11),""))
						    strClose_dt =(Trim(Replace(lgF1,Chr(11),"")))
	                
						    IF strClose_type = "1" THEN 
						    	strClose_dt = UNIDateAdd("d",-1,strClose_dt,parent.gAPDateFormat)
						    END IF 

							IF CompareDateByFormat(UNIConvDate(strDilig_str_date), strClose_dt,"","","",parent.gAPDateFormat,parent.gAPDateSeperator,false) or _
							   CompareDateByFormat(UNIConvDate(strDilig_end_date), strClose_dt,"","","",parent.gAPDateFormat,parent.gAPDateSeperator,false) then			    
						        Call DisplayMsgBox("800291","X","X","X")	     '근태 마감처리된 일 입니다 
'						        Call FncCancel()
						        Exit Function                                    '바로 return한다 
						    END IF 	 

						Next
					End if	                
            End Select
        Next
	End With

    Dim strEmpNo
    Dim strDilig_str_dt
    Dim strDilig_end_dt    
    '기간근태(hca050t)에서 기간에 (중복일자)속했는지를 check 한다. 만약 없으면 일일근태(hca060t)에 있는지도 check 한다.
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag ', ggoSpread.UpdateFlag
   	                .vspdData.Col = C_DILIG_STRT_DT
                    strDilig_str_dt = UNIConvDate(.vspdData.Text)
                                    
   	                .vspdData.Col = C_DILIG_END_DT
                    strDilig_end_dt = UNIConvDate(.vspdData.Text)
                                    
   	                .vspdData.Col = C_EMP_NO
                    strEmpNo = .vspdData.Text
   
                    Call CommonQueryRs(" isnull(count(emp_no),0) "," hca050t ","  emp_no =  " & FilterVar(strEmpNo, "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar(strDilig_str_dt, "''", "S") & " AND  " & FilterVar(strDilig_end_dt, "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar(strDilig_str_dt, "''", "S") & " AND  " & FilterVar(strDilig_end_dt, "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	                If Trim(Replace(lgF0,Chr(11),"")) = 0 then
                        Call CommonQueryRs(" isnull(count(emp_no),0) "," hca060t ","  emp_no =  " & FilterVar(strEmpNo, "''", "S") & " AND (dilig_dt between  " & FilterVar(strDilig_str_dt, "''", "S") & " AND  " & FilterVar(strDilig_end_dt, "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 		        	    
		        	    If Trim(Replace(lgF0,Chr(11),"")) > 0 then
                            Call DisplayMsgBox("800234","X","X","X")	'이 기간에 대해 이미 입력된 일일근태사항이 있습니다 
	                        Exit Function                                    '바로 return한다.
		        	    End if
	                Else
                        Call DisplayMsgBox("800067","X","X","X")	'이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
	                    Exit Function                                    '바로 return한다 
                    End if
            End Select
        Next
	End With

    Call MakeKeyStream("X")
    
    Call DisableToolBar(parent.TBC_SAVE)
	IF DBsave =  False Then
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
           .Col  = C_ROLL_PSTN
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_ROLL_PSTN_NM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DILIG_CD
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DILIG_NM
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_DILIG_STRT_DT
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
        .vspdData.col = C_DILIG_STRT_DT
        .vspdData.Text = frm1.txtDilig_dt.text
        .vspdData.col = C_DILIG_END_DT
        .vspdData.Text = frm1.txtDilig_dt.text
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
    Call InitComboBox
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

	If LayerShowHide(1) = False then
    		Exit Function 
    	End if
	
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
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
                                                      strVal = strVal & "C" & parent.gColSep 'array(0)
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strVal = strVal & Trim(frm1.txtDilig_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_END_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_NM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strVal = strVal & Trim(frm1.txtDilig_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_END_DT  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_NM	    : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'삭제시 key만								
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DisableToolBar(parent.TBC_DELETE)
	IF DBDELETE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
    
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

    Call InitVariables															'⊙: Initializes local global variables
	Call DisableToolBar(parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
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
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_DILIG_POP
	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""                 		    ' Code Condition
	        frm1.vspddata.col = C_DILIG_NM
	        arrParam(3) = frm1.vspddata.Text			' Name Cindition
	        arrParam(4) = " day_time = " & FilterVar("1", "''", "S") & "  "							' Where Condition
	        arrParam(5) = "근태코드"			    ' TextBox 명칭 
	
            arrField(0) = "dilig_cd"					' Field명(0)
            arrField(1) = "dilig_nm"				    ' Field명(1)
    
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_DILIG_NM
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_DILIG_POP
		        .vspdData.Col = C_DILIG_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_DILIG_NM
		    	.vspdData.text = arrRet(1)
				.vspdData.action =0
        End Select

	End With

End Function

'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	Else
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
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
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else 'spread
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_ROLL_PSTN_NM
			.vspdData.Text = arrRet(3)
			.vspdData.Col = C_ROLL_PSTN
            Call CommonQueryRs(" minor_cd "," b_minor "," major_cd=" & FilterVar("H0002", "''", "S") & " and minor_nm =  " & FilterVar(arrRet(3), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            .vspdData.Text = Replace(lgF0,Chr(11),"")
			.vspdData.Col = C_EMP_NO
			.vspdData.action =0
		End If
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
            Case C_ROLL_PSTN_NM
                .Col = Col
                intIndex = .Value
				.Col = C_ROLL_PSTN
				.Value = intIndex
            Case C_ROLL_PSTN
                .Col = Col
                intIndex = .Value
				.Col = C_ROLL_PSTN_NM
				.Value = intIndex
				
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
	Select Case iWhere
           
	    Case "3"
	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtDilig_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtDilig_nm.value    	' Name Cindition
	        arrParam(4) = " day_time = " & FilterVar("1", "''", "S") & "  "							' Where Condition
	        arrParam(5) = "근태코드"			    ' TextBox 명칭 
	
            arrField(0) = "dilig_cd"					' Field명(0)
            arrField(1) = "dilig_nm"				    ' Field명(1)
    
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
            
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtDilig_cd.focus
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
		    Case "3"
		        .txtDilig_cd.value = arrRet(0)
		        .txtDilig_nm.value = arrRet(1)
		        .txtDilig_cd.focus
        End Select
	End With
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
   	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

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
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_NAME_POP
                    Call OpenEmptName("1")
	    Case C_EMP_NO_POP	    
                    Call OpenEmptName("1")                    
	    Case C_DILIG_POP
                    Call OpenCode("", C_DILIG_POP, Row)
    End Select    
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

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_EMP_NO
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMP_NO
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_ROLL_PSTN
                Frm1.vspdData.Text = ""
  	            Frm1.vspdData.Col = C_ROLL_PSTN_NM
                Frm1.vspdData.Text = ""
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
  	                Frm1.vspdData.Col = C_ROLL_PSTN
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_ROLL_PSTN_NM
                    Frm1.vspdData.Text = ""
                    vspdData_Change = true
                Else
		       	    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.value = strName
		       	    
                    Call CommonQueryRs(" ROLL_PSTN "," HAA010T "," EMP_NO =  " & FilterVar(iDx , "''", "S") & " AND NAME =  " & FilterVar(strName , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		       	    Frm1.vspdData.Col = C_ROLL_PSTN
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		       	    Frm1.vspdData.Col = C_ROLL_PSTN_NM
		       	    Frm1.vspdData.Text = strRoll_pstn
                End if 
            End if 
         Case  C_DILIG_NM
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_DILIG_NM
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_DILIG_CD
                Frm1.vspdData.value = ""
            Else
                IntRetCd = CommonQueryRs(" DILIG_CD "," HCA010T "," day_time=" & FilterVar("1", "''", "S") & "  and DILIG_NM =  " & FilterVar(iDx , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If IntRetCd = false then
			        Call DisplayMsgBox("800099","X","X","X")	'해당근태는 근태코드정보에 존재하지 않습니다.
  	            Frm1.vspdData.Col = C_DILIG_CD
                Frm1.vspdData.value = ""			        
                Else
		       	    Frm1.vspdData.Col = C_DILIG_NM
                    iDx = Frm1.vspdData.value
                    Call CommonQueryRs(" DILIG_CD "," HCA010T "," day_time=" & FilterVar("1", "''", "S") & "  and  DILIG_NM =  " & FilterVar(iDx , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		       	    Frm1.vspdData.Col = C_DILIG_CD
		       	    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
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
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,UNIConvDate(frm1.txtDilig_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 부서코드는 부서마스타에 등록되지 않은코드입니다.
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

'========================================================================================================
'   Event Name : txtDilig_cd_change
'   Event Desc :
'========================================================================================================
Function txtDilig_cd_Onchange()
    Dim IntRetCd
    
    If frm1.txtDilig_cd.value = "" Then
		frm1.txtDilig_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" DILIG_NM "," HCA010T "," DAY_TIME = " & FilterVar("1", "''", "S") & "  and DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800099","X","X","X")	'근태코드정보에 등록되지 않은 코드입니다.
'			frm1.txtDilig_cd.value = ""
		    frm1.txtDilig_nm.value = ""
            frm1.txtDilig_cd.focus
            Set gActiveElement = document.ActiveElement 
            
            txtDilig_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtDilig_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End Function
'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtDilig_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDilig_dt.Action = 7
        frm1.txtDilig_dt.focus
    End If
End Sub

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
'   Event Name : txtDilig_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDilig_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기간근태등록</font></td>
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
				    	 <TD CLASS=TD6 ><script language =javascript src='./js/h4006ma1_txtDilig_dt_txtDilig_dt.js'></script></TD>
				    	 <TD CLASS=TD5 NOWRAP>부서코드</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                                  <INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
						                      <INPUT NAME="txtInternal_cd" ALT="내부코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
			           </TR>
		               <TR>		
				    	 <TD CLASS=TD5 NOWRAP>근태코드</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDilig_cd" ALT="근태코드" TYPE="Text" SiZE=3 MAXLENGTH=2  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('3')">
			                                  <INPUT NAME="txtDilig_nm" ALT="근태코드명" TYPE="Text" SiZE=15 MAXLENGTH=20  tag="14XXXU"></TD>
				         <TD CLASS=TD5 NOWRAP>사원</TD>
				     	 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
				     	                      <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
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
									<script language =javascript src='./js/h4006ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
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
