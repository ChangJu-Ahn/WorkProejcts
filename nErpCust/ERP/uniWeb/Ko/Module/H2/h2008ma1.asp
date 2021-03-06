<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 자격/면허등록 
*  3. Program ID           : H2007ma1
*  4. Program Name         : H2007ma1
*  5. Program Desc         : 인사기본자료관리/자격/면허등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/10
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H2008mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const CookieSplit = 1233
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

Dim C_LICN_KIND 
Dim C_LICN_KIND_NM 
Dim C_LICN_KIND_POP
Dim C_LICN_GRADE 
Dim C_PUBL_OFFICE 
Dim C_PUBL_OFFICE_NM
Dim C_PUBL_OFFICE_POP
Dim C_ACQ_DT 
Dim C_FNSH_DT
Dim C_LICN_UPDT_DT 
Dim C_LICN_NO

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_LICN_KIND = 1
	C_LICN_KIND_NM = 2
	C_LICN_KIND_POP = 3
	C_LICN_GRADE = 4
	C_PUBL_OFFICE = 5
	C_PUBL_OFFICE_NM = 6
	C_PUBL_OFFICE_POP = 7
	C_ACQ_DT = 8
	C_FNSH_DT = 9
	C_LICN_UPDT_DT = 10
	C_LICN_NO = 11
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
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
    lgKeyStream       = Frm1.txtEmp_no.Value & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtName.Value & parent.gColSep
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

		ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread 

		.ReDraw = false    
		.MaxCols = C_LICN_NO + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		.ColHidden = True

		.MaxRows = 0	
		ggoSpread.ClearSpreadData
		Call GetSpreadColumnPos("A")    

		ggoSpread.SSSetEdit     C_LICN_KIND,      "자격면허종류", 4,,, 4      ' hidden 자격면허코드 
		ggoSpread.SSSetEdit     C_LICN_KIND_NM,   "자격면허종류", 25,,, 20
		ggoSpread.SSSetButton   C_LICN_KIND_POP
		ggoSpread.SSSetEdit     C_LICN_GRADE,     "등급", 8,,, 4

		ggoSpread.SSSetEdit     C_PUBL_OFFICE,    "발행기관", 4,,,4
		ggoSpread.SSSetEdit     C_PUBL_OFFICE_NM, "발행기관", 22,,,20
		ggoSpread.SSSetButton   C_PUBL_OFFICE_POP

		ggoSpread.SSSetDate     C_ACQ_DT,         "취득일", 12,2, gDateFormat
		ggoSpread.SSSetDate     C_FNSH_DT,        "유효일", 12,2, gDateFormat
		ggoSpread.SSSetDate     C_LICN_UPDT_DT,   "차기갱신일", 12,2, gDateFormat

		ggoSpread.SSSetEdit     C_LICN_NO,        "자격증번호", 20,,, 20
    
		Call ggoSpread.MakePairsColumn(C_LICN_KIND_NM,C_LICN_KIND_POP)
		Call ggoSpread.MakePairsColumn(C_PUBL_OFFICE_NM,C_PUBL_OFFICE_POP)
   
		Call ggoSpread.SSSetColHidden(C_LICN_KIND,C_LICN_KIND,True)
		Call ggoSpread.SSSetColHidden(C_PUBL_OFFICE,C_PUBL_OFFICE,True)
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
        ggoSpread.SpreadLock    C_LICN_KIND,    -1, C_LICN_KIND
        ggoSpread.SpreadLock    C_LICN_KIND_NM, -1, C_LICN_KIND_NM
        ggoSpread.SpreadLock    C_LICN_KIND_POP, -1, C_LICN_KIND_POP
        ggoSpread.SpreadLock    C_PUBL_OFFICE,    -1, C_PUBL_OFFICE
        ggoSpread.SpreadLock    C_PUBL_OFFICE_NM, -1, C_PUBL_OFFICE_NM
        ggoSpread.SpreadLock    C_PUBL_OFFICE_POP, -1, C_PUBL_OFFICE_POP
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
        ggoSpread.SSSetRequired		C_LICN_KIND_NM, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_PUBL_OFFICE_NM, pvStartRow, pvEndRow
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
			C_LICN_KIND = iCurColumnPos(1)
			C_LICN_KIND_NM = iCurColumnPos(2)
			C_LICN_KIND_POP = iCurColumnPos(3)
			C_LICN_GRADE = iCurColumnPos(4)
			C_PUBL_OFFICE = iCurColumnPos(5)
			C_PUBL_OFFICE_NM = iCurColumnPos(6)
			C_PUBL_OFFICE_POP = iCurColumnPos(7)
			C_ACQ_DT = iCurColumnPos(8)
			C_FNSH_DT = iCurColumnPos(9)
			C_LICN_UPDT_DT = iCurColumnPos(10)
			C_LICN_NO = iCurColumnPos(11)      
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
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    
    frm1.txtemp_no.Focus
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    if  frm1.txtEmp_no.value = "" AND frm1.txtName.value <> "" then
        OpenEmpName(0)
        exit function
    else
        If  txtEmp_no_Onchange() then
            Exit Function
        End If
    end if

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
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
    Dim lRow 
    Dim strAcq_dt, strFnsh_dt, strLicn_updt_dt
    
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
    
    if  frm1.txtEmp_no.value = "" then
        Call chkField(Document, "1")
        Frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    end if

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
			.vspdData.Col = 0
            
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                    .vspdData.Col = C_ACQ_DT
                    strAcq_dt = UniConvDateToYYYYMMDD(.vspdData.text,gDateFormat,"")
                    .vspdData.Col = C_FNSH_DT
                    strFnsh_dt = UniConvDateToYYYYMMDD(.vspdData.text,gDateFormat,"")
                    .vspdData.Col = C_LICN_UPDT_DT
                    strLicn_updt_dt = UniConvDateToYYYYMMDD(.vspdData.text,gDateFormat,"")
				IF strAcq_dt <> "" AND strFnsh_dt <> "" THEN
                   IF strAcq_dt >= strFnsh_dt THEN
                        Call DisplayMsgBox("800036","x","x","x")
                        .vspdData.Col = C_FNSH_DT
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    END IF
				END IF

				IF strFnsh_dt <> "" AND strLicn_updt_dt <> "" THEN
                   IF strFnsh_dt >= strLicn_updt_dt THEN
                        Call DisplayMsgBox("800443","x","차기갱신일","유효일")
                        .vspdData.Col = C_LICN_UPDT_DT
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    END IF
                END IF
                
                IF strAcq_dt <> "" AND strLicn_updt_dt <> "" THEN
                   IF strAcq_dt >= strLicn_updt_dt THEN
                        Call DisplayMsgBox("800038","x","x","x")
                        .vspdData.Col = C_LICN_UPDT_DT
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    END IF
                END IF    	
                .vspdData.Col = C_LICN_KIND
              
                if Trim(.vspdData.text="") then
                     Call DisplayMsgBox("971012","x","자격면허종류","x")
	                .vspdData.Col=C_LICN_KIND_NM
                    .vspdData.Action = 0 ' go to 
                     Exit Function
				end if 
                .vspdData.Col = C_PUBL_OFFICE				
                if Trim(.vspdData.text="") then
                     Call DisplayMsgBox("971012","x","발행기관","x")
	                .vspdData.Col=C_PUBL_OFFICE_NM
                    .vspdData.Action = 0 ' go to 
                     Exit Function
				end if 				
            End Select
        Next
    End With

    Call MakeKeyStream("X")
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
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
    FncCopy = False                                                               '☜: Processing is NG

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
           .Col  = C_LICN_KIND
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_LICN_KIND_NM
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
    Call Initdata()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
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

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    Err.Clear                                                                        '☜: Clear err status

	If   LayerShowHide(1) = False Then
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

	Dim strRes_no

    DbSave = False                                                          
    
	If   LayerShowHide(1) = False Then
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
 
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                     strVal = strVal & "C" & parent.gColSep
                                                     strVal = strVal & lRow & parent.gColSep
                                                     strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_LICN_KIND	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_GRADE   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PUBL_OFFICE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ACQ_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FNSH_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_UPDT_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_NO      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U" & parent.gColSep
                                                     strVal = strVal & lRow & parent.gColSep
                                                     strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_LICN_KIND	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_GRADE   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PUBL_OFFICE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ACQ_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FNSH_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_UPDT_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_LICN_NO      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_LICN_KIND	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
    If DbDelete = False Then
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
	Call SetToolbar("110011110011111")									
	Frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
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
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	End If
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 		
		Set gActiveElement = document.ActiveElement
        call txtEmp_no_Onchange()
        .txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub


'========================================================================================================

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
	    Case 1  ' 자격면허코드 
	        arrParam(0) = "자격면허 팝업"				    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0030", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "자격면허코드" 				    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "자격면허코드"	       			<%' Header명(0)%>
	    	arrHeader(1) = "자격면허"	    				<%' Header명(1)%>
	    Case 2  ' 발행기관코드 
	        arrParam(0) = "발행기관 팝업"				    <%' 팝업 명칭 %>
	    	arrParam(1) = "B_minor"							    <%' TABLE 명칭 %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0031", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "발행기관코드" 				    <%' TextBox 명칭 %>
	
	    	arrField(0) = "minor_cd"							<%' Field명(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field명(1)%>
    
	    	arrHeader(0) = "발행기관코드"	       			<%' Header명(0)%>
	    	arrHeader(1) = "발행기관"	    				<%' Header명(1)%>

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case 1  ' 자격면허 
		        frm1.vspdData.Col = C_LICN_KIND_NM
		    	frm1.vspdData.action =0	    	
            Case 2
		        frm1.vspdData.Col = C_PUBL_OFFICE_NM
		    	frm1.vspdData.action =0
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

    lgBlnFlgChgValue = True

	With frm1

		Select Case iWhere
		    Case 1  ' 자격면허 
		        .vspdData.Col = C_LICN_KIND
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_LICN_KIND_NM
		    	.vspdData.text = arrRet(1)		    	
		    	.vspdData.action =0	    	
            Case 2
		        .vspdData.Col = C_PUBL_OFFICE
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_PUBL_OFFICE_NM
		    	.vspdData.text = arrRet(1)   
		    	.vspdData.action =0
        End Select

		lgBlnFlgChgValue = True

	End With

End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx,IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
      
	End If
    Select Case Col
         Case  C_LICN_KIND_NM
			IntRetCd = CommonQueryRs(" minor_cd "," B_minor "," major_cd=" & FilterVar("H0030", "''", "S") & " and minor_nm = " & FilterVar(Frm1.vspdData.text, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If  IntRetCd = false or Frm1.vspdData.text="" then			
				Frm1.vspdData.Col = C_LICN_KIND
				Frm1.vspdData.text =""
			else 	
				Frm1.vspdData.Col = C_LICN_KIND					
				Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
			end if
         Case  C_PUBL_OFFICE_NM	
			IntRetCd = CommonQueryRs(" minor_cd "," B_minor "," major_cd=" & FilterVar("H0031", "''", "S") & " and minor_nm = " & FilterVar(Frm1.vspdData.text, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCd = false  or Frm1.vspdData.text="" then			
			Frm1.vspdData.Col = C_PUBL_OFFICE
			Frm1.vspdData.text =""
			else 			
				Frm1.vspdData.Col = C_PUBL_OFFICE			
				Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
			end if
    End Select 
        
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
'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_LICN_KIND_POP     ' 자격면허 
            Call OpenCode(frm1.vspdData.Text, 1, Row)
	    Case C_PUBL_OFFICE_POP   ' 자격면허 
            Call OpenCode(frm1.vspdData.Text, 2, Row)
    End Select    
   	frm1.vspdData.Row = Row
    	
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
	frm1.txtDept_nm.value = ""
	frm1.txtRoll_pstn.value = ""
	frm1.txtEntr_dt.Text = ""
	frm1.txtPay_grd.value = ""
	Frm1.imgPhoto.src = ""

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtEmp_no.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
			strVal = "../../../CShared/image/default_picture.jpg"
    		Frm1.imgPhoto.src = strVal	    
	    
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			ggoSpread.Source = Frm1.vspdData    
			ggoSpread.ClearSpreadData             
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text = UNIDateClientFormat(strEntr_dt)
             'strEntr_dt는 Client Format(parent.gClientDateFormat) 그러므로 Client Format -->Company Format

			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if

    		Frm1.imgPhoto.src = strVal

        End if 
    End if
    
End Function 

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>자격/면허등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=10%>
                        <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=80 HEIGHT=90 HSPACE=10 VSPACE=0 BORDER=1>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" Maxlength=13 SiZE=13 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName(0)"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" Maxlength=30 SiZE=20 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_nm" ALT="부서명" TYPE="Text" SiZE=15 tag=14></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
							    <TD CLASS="TD6"><script language =javascript src='./js/h2008ma1_txtEntr_dt_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h2008ma1_vaSpread1_vspdData.js'></script>
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
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">인사마스타</a></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

