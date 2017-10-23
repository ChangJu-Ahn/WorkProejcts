<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : �������׻��׵�� 
*  3. Program ID           : H2003ma1
*  4. Program Name         : H2003ma1
*  5. Program Desc         : �λ�⺻�ڷ����/�������׵�� 
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
const	 CookieSplit = 1233
const	 C_SHEETMAXROWS    = 21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>
Const BIZ_PGM_ID = "H2003mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_JUMP_ID = "H2001ma1" 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim gUsrAuth    ' �ڷ���Ѱ��� 

Dim C_FAMILY_NM 
Dim C_REL_CD
Dim C_REL_CD_NM

Dim C_RES_NO 
Dim C_SCH_SHIP 
Dim C_SCH_SHIP_NM 

Dim C_OCCUP_NM
Dim C_COMP_NM 
Dim C_ROLL_PSTN
Dim C_SUPP_CD 
Dim C_SUPP_CD_NM

Dim C_RESIDE_TYPE 
Dim C_MDCL_INSUR 
Dim C_ALLOW_TYPE 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	 C_FAMILY_NM = 1
	 C_REL_CD = 2
	 C_REL_CD_NM = 3

	 C_RES_NO = 4
	 C_SCH_SHIP = 5
	 C_SCH_SHIP_NM = 6

	 C_OCCUP_NM = 7
	 C_COMP_NM = 8
	 C_ROLL_PSTN = 9
	 C_SUPP_CD = 10
	 C_SUPP_CD_NM= 11

	 C_RESIDE_TYPE = 12
	 C_MDCL_INSUR = 13
	 C_ALLOW_TYPE = 14

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
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
'	Description : Item Popup���� Return�Ǵ� �� setting
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
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0023", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    iCodeArr = lgF0
    iNameArr = lgF1
    
    ggoSpread.SetCombo replace(iCodeArr,Chr(11),vbTab), C_REL_CD
    ggoSpread.SetCombo replace(iNameArr,Chr(11),vbTab), C_REL_CD_NM

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    
    ggoSpread.SetCombo replace(iCodeArr,Chr(11),vbTab), C_SCH_SHIP
    ggoSpread.SetCombo replace(iNameArr,Chr(11),vbTab), C_SCH_SHIP_NM

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0024", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    
    ggoSpread.SetCombo replace(iCodeArr,Chr(11),vbTab), C_SUPP_CD
    ggoSpread.SetCombo replace(iNameArr,Chr(11),vbTab), C_SUPP_CD_NM
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
			.Col = C_REL_CD         ' �������� 
			intIndex = .value
			.col = C_REL_CD_NM
			.value = intindex					

			.Row = intRow
			.Col = C_SCH_SHIP       ' �з� 
			intIndex = .value
			.col = C_SCH_SHIP_NM
			.value = intindex					

			.Row = intRow
			.Col = C_SUPP_CD        ' �ξ� 
			intIndex = .value
			.col = C_SUPP_CD_NM
			.value = intindex					
		Next	
	End With

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'2007.04 res_no  ggoSpread.SSSetEdit     C_RES_NO,      "�ֹι�ȣ" ,14,,, 14,2 --> 20 ���� ���� 
'========================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM
	strMaskYM = "999999-9999999"
	Call initSpreadPosVariables()  
	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
	
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false    
        .MaxCols = C_ALLOW_TYPE + 1												<%'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															<%'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True	  
        .MaxRows = 0 
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData     

		Call GetSpreadColumnPos("A")  

        ggoSpread.SSSetEdit     C_FAMILY_NM,   "����", 10,,, 30    
        ggoSpread.SSSetCombo    C_REL_CD,      "����", 10
        ggoSpread.SSSetCombo    C_REL_CD_NM,   "����", 10
    

	    ggoSpread.SSSetEdit     C_RES_NO,      "�ֹι�ȣ" ,20,,, 20,2
        ggoSpread.SSSetCombo    C_SCH_SHIP,    "�����з�", 10
        ggoSpread.SSSetCombo    C_SCH_SHIP_NM, "�����з�", 10

        ggoSpread.SSSetEdit     C_OCCUP_NM,    "����", 10,,, 20,1
        ggoSpread.SSSetEdit     C_COMP_NM,     "�ٹ�ȸ��", 15,,, 30,1
        ggoSpread.SSSetEdit     C_ROLL_PSTN,   "����", 10,,, 10,1
        ggoSpread.SSSetCombo    C_SUPP_CD,     "�ξ�", 10
        ggoSpread.SSSetCombo    C_SUPP_CD_NM,  "�ξ�", 10

        ggoSpread.SSSetCheck    C_RESIDE_TYPE, "����", 08
        ggoSpread.SSSetCheck    C_MDCL_INSUR,  "�ǰ�����", 10
        ggoSpread.SSSetCheck    C_ALLOW_TYPE,  "����", 08
        
        Call ggoSpread.SSSetColHidden(C_REL_CD,C_REL_CD,True)	
        Call ggoSpread.SSSetColHidden(C_SCH_SHIP,C_SCH_SHIP,True)
        Call ggoSpread.SSSetColHidden(C_SUPP_CD,C_SUPP_CD,True)
        
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

			C_FAMILY_NM = iCurColumnPos(1)
			C_REL_CD = iCurColumnPos(2)
			C_REL_CD_NM = iCurColumnPos(3)

			C_RES_NO = iCurColumnPos(4)
			C_SCH_SHIP = iCurColumnPos(5)
			C_SCH_SHIP_NM = iCurColumnPos(6)

			C_OCCUP_NM = iCurColumnPos(7)
			C_COMP_NM = iCurColumnPos(8)
			C_ROLL_PSTN = iCurColumnPos(9)
			C_SUPP_CD = iCurColumnPos(10)
			C_SUPP_CD_NM= iCurColumnPos(11)

			C_RESIDE_TYPE = iCurColumnPos(12)
			C_MDCL_INSUR = iCurColumnPos(13)
			C_ALLOW_TYPE = iCurColumnPos(14)          
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_FAMILY_NM, -1, C_FAMILY_NM
        ggoSpread.SpreadLock    C_REL_CD, -1, C_REL_CD
        ggoSpread.SpreadLock    C_REL_CD_NM, -1, C_REL_CD_NM
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
        ggoSpread.SSSetRequired		C_FAMILY_NM, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_REL_CD, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_REL_CD_NM, pvStartRow, pvEndRow
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

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
    
    Call InitComboBox
    Call SetToolbar("1100110100101111")										        '��ư ���� ����    
    frm1.txtEmp_no.Focus
    Call InitComboBox
	Call CookiePage (0)                                                             '��: Check Cookie
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

    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = Frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	ggoSpread.ClearSpreadData    															
    If Not chkField(Document, "1") Then									         '��: This function check required field
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

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
' 2007.04 res_no KR �� �ƴѰ�� �ڸ���ó���κ� ����.
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
    Dim strAdmi_dt
    Dim strGrudt_dt
    Dim strNat_cd
    Dim cnt00   ' ���� 
    Dim cnt01   ' ���� 
    Dim cnt02   ' ���� 
    Dim cnt03   ' �� 
    Dim cnt04   ' �� 
    Dim cnt09   ' ����� 

    Dim res_no1, res_no2            ' �ֹι�ȣ 
    Dim intChk, intMod, intDef      ' �ֹι�ȣ 

    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    
    if  frm1.txtEmp_no.value = "" then
        Frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    end if

    call CommonQueryRs(" nat_cd "," HAA010T "," EMP_NO =  " & FilterVar(Frm1.txtemp_no.Value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strNat_cd = Replace(lgF0, Chr(11), "")  ' �ֹι�ȣ check�� ���ؼ� 

    cnt00 = 0
    cnt01 = 0
    cnt02 = 0
    cnt03 = 0
    cnt04 = 0
    cnt09 = 0

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow

            .vspdData.Col = 0
            if  .vspdData.Text <> ggoSpread.DeleteFlag then
                .vspdData.Col = C_REL_CD
                 select case .vspdData.text
                     case "00"   ' ���� 
                         cnt00 = cnt00 + 1
                     case "01"   ' ���� 
                         cnt01 = cnt01 + 1
                     case "02"   ' ���� 
                         cnt02 = cnt02 + 1
                     case "03"   ' �� 
                         cnt03 = cnt03 + 1
                     case "04"   ' �� 
                         cnt04 = cnt04 + 1
                     case "09"   ' ����� 
                         cnt09 = cnt09 + 1
                 end select
            end if
           .vspdData.Col = 0

           Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                    .vspdData.Col = C_RES_NO

                    if  Trim(Replace(.vspdData.Text,"-","")) <> "" then
                        ' �ֹι�ȣ Check **** Start
                        If  UCase(strNat_cd) = "KR" Then
                                      
							if instr(1,Trim(.vspdData.Text),"-") > 0 then
								.vspdData.Text = replace(Trim(.vspdData.Text),"-","")
							end if  								             								
							if len(.vspdData.Text)>13 then
								.vspdData.Text = mid(.vspdData.Text,1,13)
							end if
						    
                            res_no1 = Mid(.vspdData.Text, 1, 6)
                            res_no2 = Mid(.vspdData.Text, 7, 7)
                            
                            if  Len(res_no1 & res_no2) = 13  then
                                                    '��: If process fails
							    On Error Resume Next                                  
                                intChk = Cint(Mid(res_no1, 1, 1)) * 2 + Cint(Mid(res_no1, 2, 1)) * 3 + _
                                         Cint(Mid(res_no1, 3, 1)) * 4 + Cint(Mid(res_no1, 4, 1)) * 5 + _
                                         Cint(Mid(res_no1, 5, 1)) * 6 + Cint(Mid(res_no1, 6, 1)) * 7 + _
                                         Cint(Mid(res_no2, 1, 1)) * 8 + Cint(Mid(res_no2, 2, 1)) * 9 + _
                                         Cint(Mid(res_no2, 3, 1)) * 2 + Cint(Mid(res_no2, 4, 1)) * 3 + _
                                         Cint(Mid(res_no2, 5, 1)) * 4 + Cint(Mid(res_no2, 6, 1)) * 5
								if err.number <>0  then
									call DisplayMsgBox("970029","x","�ֹι�ȣ","x")
									Set gActiveElement = document.ActiveElement  									
                                    .vspdData.Action = 0
                                    exit function									
								end if
                                intMod = intChk Mod 11
                                intDef = 11 - intMod
                                If intDef = 10 Then
                                    intDef = 0
                                ElseIf intDef = 11 Then
                                    intDef = 1
                                End If

                                If Cstr(intDef) <> Mid(res_no2, 7, 1) Then
                                	IntRetCD = DisplayMsgBox("800345", parent.VB_YES_NO,"x","x")
                                    if  IntRetCD = VBNO then
                                        .vspdData.Row = lRow
                                        .vspdData.Col = C_RES_NO
                                        .vspdData.Action = 0 ' go to 
                                        exit function
                                    end if
                                End If

                            else
                            	IntRetCD = DisplayMsgBox("800345", parent.VB_YES_NO,"x","x")
                                if  IntRetCD = VBNO then
                                    .vspdData.Row = lRow
                                    .vspdData.Col = C_RES_NO
                                    .vspdData.Action = 0 ' go to 
                                    exit function
                                end if
                            end if
		
                            'frm1.txtRes_no.value = res_no1 & res_no2
						else
                                      						
							'if len(Trim(.vspdData.Text))>13 then
							'	.vspdData.Text = mid(Trim(.vspdData.Text),1,13)
							'end if                            
                        End If 'end if KR
                    end if 'end if trim 
           End Select
       Next

	End With
    if  cnt00 > 1 then  ' ���� 
        call DisplayMsgBox("970027", "X","������ ����(1��)","X")
        exit function
    end if
    if  cnt01 > 1 then  ' ���� 
        call DisplayMsgBox("970027", "X","������ ����(1��)","X")
        exit function
    end if
    if  cnt02 > 1 then  ' ���� 
        call DisplayMsgBox("970027", "X","������ ����(1��)","X")
        exit function
    end if
    if  cnt03 > 1 then  ' �� 
        call DisplayMsgBox("970027", "X","������ ��(1��)","X")
        exit function
    end if
    if  cnt04 > 1 then  ' �� 
        call DisplayMsgBox("970027", "X","������ ��(1��)","X")
        exit function
    end if
    if  cnt09 > 1 then  ' ����� 
        call DisplayMsgBox("970027", "X","������ �����(1��)","X")
        exit function
    end if

    Call MakeKeyStream("X")
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

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
    Call parent.FncExport(parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
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

    Err.Clear                                                                        '��: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
		

	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic

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
 
               Case ggoSpread.InsertFlag                                      '��: Insert
                                                    strVal = strVal & "C" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_FAMILY_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REL_CD	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RES_NO        ' �ֹι�ȣ '-' ���� 
                        strRes_no = Trim(.vspdData.Text)
                        strVal = strVal & strRes_no & parent.gColSep
                    .vspdData.Col = C_SCH_SHIP	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCCUP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COMP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUPP_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESIDE_TYPE   ' ���ű��� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gColSep
                        Else
                            strVal = strVal & "N" & parent.gColSep
                        End if
                    .vspdData.Col = C_MDCL_INSUR   ' �Ǻ����� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gColSep
                        Else
                            strVal = strVal & "N" & parent.gColSep
                        End if
                    .vspdData.Col = C_ALLOW_TYPE   ' ���籸�� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gRowSep
                        Else
                            strVal = strVal & "N" & parent.gRowSep
                        End if
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '��: Update
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_FAMILY_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REL_CD	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RES_NO        ' �ֹι�ȣ '-' ���� 
                        strRes_no = Trim(.vspdData.Text)
                        strVal = strVal & strRes_no & parent.gColSep
                    .vspdData.Col = C_SCH_SHIP	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCCUP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COMP_NM	  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUPP_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESIDE_TYPE   ' ���ű��� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gColSep
                        Else
                            strVal = strVal & "N" & parent.gColSep
                        End if
                    .vspdData.Col = C_MDCL_INSUR   ' �Ǻ����� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gColSep
                        Else
                            strVal = strVal & "N" & parent.gColSep
                        End if
                    .vspdData.Col = C_ALLOW_TYPE   ' ���籸�� 
                        If Trim(.vspdData.Text) = "1" then
                            strVal = strVal & "Y" & parent.gRowSep
                        Else
                            strVal = strVal & "N" & parent.gRowSep
                        End if
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '��: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & .txtEmp_no.value & parent.gColSep
                    .vspdData.Col = C_FAMILY_NM	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REL_CD	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

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
    
    FncDelete = False                                                      '��: Processing is NG

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
	
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If
    
    FncDelete = True                                                        '��: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
	Dim strVal

    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
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
    Call InitVariables															'��: Initializes local global variables
	Call MainQuery()
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
	    arrParam(0) = ""			' Code Condition
	    arrParam(1) = frm1.txtName.value			' Name Cindition
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)

		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 		
		Set gActiveElement = document.ActiveElement

        Call txtEmp_no_Onchange()

		lgBlnFlgChgValue = False
		.txtEmp_no.focus
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
	    Case 21  <%' �������б��ڵ� %>
	        arrParam(0) = "�������б� �˾�"			        <%' �˾� ��Ī %>
	    	arrParam(1) = "B_minor"							    <%' TABLE ��Ī %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0008", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "�������б��ڵ�" 			        <%' TextBox ��Ī %>
	
	    	arrField(0) = "minor_cd"							<%' Field��(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field��(1)%>
    
	    	arrHeader(0) = "�������б��ڵ�"	   		    	<%' Header��(0)%>
	    	arrHeader(1) = "�б���"	    		     		<%' Header��(1)%>
	    Case 22  <%' ����б��ڵ� %>
	        arrParam(0) = "����б� �˾�"				    <%' �˾� ��Ī %>
	    	arrParam(1) = "B_minor"							    <%' TABLE ��Ī %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0009", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "����б��ڵ�" 				    <%' TextBox ��Ī %>
	
	    	arrField(0) = "minor_cd"							<%' Field��(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field��(1)%>
    
	    	arrHeader(0) = "����б��ڵ�"	    	    	<%' Header��(0)%>
	    	arrHeader(1) = "�б���"	    		     		<%' Header��(1)%>
	    Case 23  <%' ���б��ڵ� %>
	        arrParam(0) = "���б� �˾�"				        <%' �˾� ��Ī %>
	    	arrParam(1) = "B_minor"							    <%' TABLE ��Ī %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0010", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "���б��ڵ�" 				        <%' TextBox ��Ī %>
	
	    	arrField(0) = "minor_cd"							<%' Field��(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field��(1)%>
    
	    	arrHeader(0) = "���б��ڵ�"	       		    	<%' Header��(0)%>
	    	arrHeader(1) = "�б���"	    		     		<%' Header��(1)%>
	    Case 31  <%' �������ڵ� %>
	        arrParam(0) = "������ �˾�"	   			    <%' �˾� ��Ī %>
	    	arrParam(1) = "B_minor"							    <%' TABLE ��Ī %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0011", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "�������ڵ�" 					    <%' TextBox ��Ī %>
	
	    	arrField(0) = "minor_cd"							<%' Field��(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field��(1)%>
    
	    	arrHeader(0) = "�������ڵ�"	       			<%' Header��(0)%>
	    	arrHeader(1) = "������"	        				<%' Header��(1)%>
	    Case 32  <%' ����/�������б������ڵ� %>
	        arrParam(0) = "���� �˾�"	    			    <%' �˾� ��Ī %>
	    	arrParam(1) = "B_minor"							    <%' TABLE ��Ī %>
	    	arrParam(2) = strCode                   			<%' Code Condition%>
	    	arrParam(3) = ""									<%' Name Cindition%>
	    	arrParam(4) = "major_cd=" & FilterVar("H0012", "''", "S") & ""			    	<%' Where Condition%>
	    	arrParam(5) = "�����ڵ�" 					    <%' TextBox ��Ī %>
	
	    	arrField(0) = "minor_cd"							<%' Field��(0)%>
	    	arrField(1) = "minor_nm"    						<%' Field��(1)%>
    
	    	arrHeader(0) = "�����ڵ�"		       			<%' Header��(0)%>
	    	arrHeader(1) = "������"	        				<%' Header��(1)%>
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 21, 22, 23     ' �б� 
		        .vspdData.Col = C_SCHOOL
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_SCHOOL_NM
		    	.vspdData.text = arrRet(1)   
		    Case 31, 32         ' ���� 
		        .vspdData.Col = C_MAJOR
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_MAJOR_NM
		    	.vspdData.text = arrRet(1)   
        End Select

	End With

End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
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
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
    End Select    
            
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_SCH_SHIP_NM
                .Col = Col
                intIndex = .Value
				.Col = C_SCH_SHIP
				.Value = intIndex
            Case C_SUPP_CD_NM
                .Col = Col
                intIndex = .Value
				.Col = C_SUPP_CD
				.Value = intIndex
            Case C_REL_CD_NM
                .Col = Col
                intIndex = .Value
				.Col = C_REL_CD
				.Value = intIndex
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : �÷���ư�� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
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
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
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
            frm1.txtEntr_dt.text = UNIDateClientFormat(strEntr_dt)  
            'strEntr_dt�� Client Format(parent.gClientDateFormat) �׷��Ƿ� Client Format -->Company Format

			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    		if   Replace(lgF0, Chr(11), "") > 0  then

				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '��: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '��: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '��: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '��: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '��: Key
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�������׵��</font></td>
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
                        <img src="../../../CShared/image/default_picture.jpg"  name="imgPhoto" WIDTH=80 HEIGHT=90 HSPACE=10 VSPACE=0 BORDER=1>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>���</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="���" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName(1)"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>����</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName" ALT="����" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>�μ���</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_nm" ALT="�μ���" TYPE="Text" MAXLENGTH=15 SiZE=15 tag=14></TD>
			            		<TD CLASS="TD5" NOWRAP>����</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="����" TYPE="Text" MAXLENGTH=15 SiZE=15 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>�Ի���</TD>
							    <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID="txtEntr_dt" NAME="txtEntr_dt" ALT="�Ի���" CLASS=FPDTYYYYMMDD TITLE=FPDATETIME tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
			            		<TD CLASS="TD5" NOWRAP>��ȣ</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="��ȣ" TYPE="Text" MAXLENGTH=15 SiZE=15 tag=14></TD>
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">�λ縶��Ÿ</a></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

