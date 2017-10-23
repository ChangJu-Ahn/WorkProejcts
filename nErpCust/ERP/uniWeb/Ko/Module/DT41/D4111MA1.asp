
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 전자세금계산서(SmartBill)
'*  2. Function Name        : 
'*  3. Program ID           : D4111ma1.asp
'*  4. Program Name         : 사용자관리
'*  5. Program Desc         :
'*  6. Modified date(First) : 2011/05/11
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>


<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "D4111mb1.asp"												<%'비지니스 로직 ASP명 %>

Dim C_user_id
Dim C_user_id_pop
Dim C_user_name
Dim C_user_reg_no
Dim C_user_reg_no_pop
Dim C_smartbill_id
Dim C_smartbill_pw1
Dim C_smartbill_pw
Dim C_Dept_Nm
Dim C_Tel_Num
Dim C_Email_id
 

Dim IsOpenPop          
Dim lgSortKey1
Dim lgSortKey2

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
    C_user_id			= 1      
    C_user_id_pop		= 2
    C_user_name		    = 3
    C_user_reg_no       = 4  
    C_user_reg_no_pop   = 5      
    C_smartbill_id      = 6    
    C_smartbill_pw1	    = 7
    C_smartbill_pw	    = 8
    C_Dept_Nm           = 9
    C_Tel_Num           = 10
    C_Email_id          = 11
    
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","BA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	 ggoSpread.Source = frm1.vspdData
	 ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
    .ReDraw = false
    .MaxCols = C_Email_id + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
    .Col = .MaxCols														'공통콘트롤 사용 Hidden Column
    .ColHidden = True
    .MaxRows = 0			
				
    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetEdit		C_user_id,		    "사용자ID", 10,,,30,2
	ggoSpread.SSSetButton	C_user_id_pop
	ggoSpread.SSSetEdit		C_user_name,	    "사용자명", 15,,,15,2
	ggoSpread.SSSetEdit		C_user_reg_no,	    "사업자등록번호", 15,,,10,2		
	ggoSpread.SSSetButton   C_user_reg_no_pop			
	ggoSpread.SSSetEdit		C_smartbill_id,	    "스마트빌ID", 15,,,12,1	
	ggoSpread.SSSetEdit		C_smartbill_pw1,	"스마트빌PW1", 15,,,35,1
	ggoSpread.SSSetEdit		C_smartbill_pw,  	"스마트빌PW", 15,,,35,1
	
	ggoSpread.SSSetEdit		C_Dept_Nm,  	    "부서명", 15,,,40,1
	ggoSpread.SSSetEdit		C_Tel_Num,  	    "전화번호", 13,,,20,1
	ggoSpread.SSSetEdit		C_Email_id,  	    "E-MAIL", 20,,,100,1
								
	call ggoSpread.MakePairsColumn(C_user_id,C_user_id_pop)

    Call ggoSpread.SSSetColHidden(C_smartbill_pw,	C_smartbill_pw,True) 

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired		C_smartbill_id,	-1, -1
    ggoSpread.SSSetRequired		C_smartbill_pw,	-1, -1
    ggoSpread.SSSetRequired		C_smartbill_pw1,	-1, -1
    ggoSpread.SSSetRequired		C_user_id,			-1, -1
        
    ggoSpread.SSSetProtected	C_user_reg_no,			-1, -1
    ggoSpread.SSSetProtected	C_user_reg_no_pop,			-1, -1
    ggoSpread.SSSetRequired		C_Dept_Nm,			-1, -1
    ggoSpread.SSSetRequired		C_Tel_Num,			-1, -1
    ggoSpread.SSSetRequired		C_Email_id,			-1, -1
            
    'ggoSpread.SSSetRequired		C_user_name,		-1, -1
    
    
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_smartbill_id,	pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_smartbill_pw1,	pvStartRow, pvEndRow    
    ggoSpread.SSSetRequired	C_smartbill_pw,	pvStartRow, pvEndRow
    
    ggoSpread.SSSetRequired	C_user_id,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_user_reg_no,			pvStartRow, pvEndRow
    
    ggoSpread.SSSetRequired	C_Dept_Nm,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_Tel_Num,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_Email_id,			pvStartRow, pvEndRow
    
    'ggoSpread.SSSetProtected	C_user_id_pop,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_user_name,		pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor1(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected C_user_id, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_user_id_pop, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_user_name, pvStartRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_user_id			= iCurColumnPos(1)      
            C_user_id_pop		= iCurColumnPos(2)
            C_user_name		    = iCurColumnPos(3)
            C_user_reg_no       = iCurColumnPos(4)             
            C_user_reg_no_pop   = iCurColumnPos(5)                               
            C_smartbill_id      = iCurColumnPos(6)    
            C_smartbill_pw1	    = iCurColumnPos(7)
            C_smartbill_pw	    = iCurColumnPos(8)
            C_Dept_Nm           = iCurColumnPos(9)
            C_Tel_Num           = iCurColumnPos(10)
            C_Email_id          = iCurColumnPos(11)
    End Select    
End Sub


Function OpenPopup(Byval strcode, Byval iWhere)
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	Dim strUserid

   If IsOpenPop = True Then Exit Function

   IsOpenPop = True


	Select Case iWhere
	    Case 1 
        
         frm1.vspddata.Col = C_user_id
         strUserid = Trim(frm1.vspddata.value) 
         
         if strUserid = "" then
            Call DisplayMsgBox("17A002","X","사용자ID(를)","X")	
            '17A002: %1을 입력하세요.               		            
            IsOpenPop = False
            Exit Function
         end if
                                                  
         arrParam(0) = "사업자등록번호팝업"
         arrParam(1) = "B_TAX_BIZ_AREA a (NOLOCK) INNER JOIN B_BIZ_PARTNER b (NOLOCK) on (b.BP_CD = a.TAX_BIZ_AREA_CD) " ' TABLE 명칭 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = "" ' Where Condition
         arrParam(5) = "사업자등록번호"    ' 조건필드의 라벨 명칭 
         

         arrField(0) = "replace(b.BP_RGST_NO,'-','')"     ' Field명(0)
         arrField(1) = "A.TAX_BIZ_AREA_NM"     ' Field명(1)

         arrHeader(0) = "사업자등록번호"    ' Header명(0)
         arrHeader(1) = "사업장명"     ' Header명(1)
	    		
'		Case 2

''			If frm1.txtCompCd.className = parent.UCN_PROTECTED Then
''               IsOpenPop = False
''               Exit Function
''            End If

'         arrParam(0) = "PU차수팝업"
'         arrParam(1) = "KMA_PU_CYCLE" ' TABLE 명칭 
'         arrParam(2) = strcode      ' Code Condition
'         arrParam(3) = ""       ' Name Cindition
'         arrParam(4) = "" ' Where Condition
'         arrParam(5) = "PU차수"    ' 조건필드의 라벨 명칭 
'         

'         arrField(0) = "PU_CYCLE"     ' Field명(0)
'         arrField(1) = "PU_CYCLE_NM"     ' Field명(1)

'         arrHeader(0) = "PU차수"    ' Header명(0)
'         arrHeader(1) = "PU차수명"     ' Header명(1)
            
		Case Else
		
	     Exit Function
   End Select
	 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False

   If arrRet(0) = "" Then
      Exit Function
   Else
		Call SetPopup(arrRet, iWhere)
   End If 
 
End Function



Function Open_User(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사용자"				        ' 팝업 명칭 
	arrParam(1) = "z_usr_mast_rec"				' TABLE 명칭
	arrParam(2) = strCode						' Code Condition
	arrParam(4) = ""							' Name Cindition
	arrParam(5) = "사용자"						' 조건필드의 라벨 명칭 

	arrField(0) = "usr_id"						' Field명(0)
	arrField(1) = "usr_nm"						' Field명(1)

	arrHeader(0) = "사용자"						' Header명(0)
	arrHeader(1) = "사용자명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUser(arrRet, iWhere)
	End If	

End Function

Function Open_User1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
   
    arrParam(0) = "사용자ID팝업"                                             ' 팝업 명칭 
    arrParam(1) = "XXSB_DTI_SM_USER (nolock)"                                  ' TABLE 명칭 
    arrParam(2) = Trim(frm1.txtuserId.value)                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "사용자ID"            
         
	arrField(0) = "FND_USER"					' Field명(0)
	arrField(1) = "FND_USER_NAME"				' Field명(1)
	arrField(2) = "FND_REGNO"					' Field명(2)

	arrHeader(0) = "사용자ID"					' Header명(0)
	arrHeader(1) = "사용자명"					' Header명(1)
	arrHeader(2) = "사업자등록번호"				' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=620px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtuserId.value = arrRet(0)
		frm1.txtuserNm.value = arrRet(1)
	End If	
	
End Function


'========================================================================================================
' Function Name : SetPopUp(Byval arrRet, Byval iWhere)
'========================================================================================================
Function SetPopup(Byval arrRet, Byval iWhere)
   With frm1
      Select Case iWhere
	     Case 1   ' 
           .vspdData.Col = C_user_reg_no
           .vspdData.Text = arrRet(0)
'           Call vspdData_Change(C_PuNo, .vspdData.Row)	     	     
	     Case 2   

      End Select
   End With
End Function


Function SetUser(Byval arrRet, Byval iWhere)
	With frm1 
		.vspdData.Col = C_user_id
		.vspdData.Text = arrRet(0)

		.vspdData.Col = C_user_name
		.vspdData.Text = arrRet(1)

		lgBlnFlgChgValue = True
	End With
End Function


Sub Form_Load()
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                           <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
          
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>   
 
    Call SetToolbar("1100110100001111")										<%'버튼 툴바 제어 %>

End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )

   Dim strUserid
   Dim strUserRegNo
   Dim strPW1
   Dim strPW


    Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        
        Case  C_user_id
            strUserid = Trim(Frm1.vspdData.value)
    
            If strUserid = "" Then
  	            Frm1.vspdData.Col = C_user_name
                Frm1.vspdData.value = ""
            Else					
				If CommonQueryRs(" a.USR_NM "," Z_USR_MAST_REC a (nolock)  "," a.USR_ID = '" & strUserid & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					Frm1.vspdData.Col = C_user_name
				    Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
				else
				    Call DisplayMsgBox("971001","X",strUserid,"X")	               		
				    '971001: %1 이(가) 존재하지 않습니다.
				    Frm1.vspdData.Col = C_user_id
				    Frm1.vspdData.text = ""
				    Frm1.vspdData.Col = C_user_name
				    Frm1.vspdData.text = ""				    
				END IF					
            End if    
            
        Case  C_user_reg_no
            strUserRegNo = Trim(Frm1.vspdData.value)
    
            If strUserRegNo = "" Then
  	            Frm1.vspdData.Col = C_user_name
                Frm1.vspdData.value = ""
            Else					
				If CommonQueryRs(" a.TAX_BIZ_AREA_CD, a.TAX_BIZ_AREA_NM, replace(b.BP_RGST_NO,'-','') "," B_TAX_BIZ_AREA a (NOLOCK) INNER JOIN B_BIZ_PARTNER b (NOLOCK) on (b.BP_CD = a.TAX_BIZ_AREA_CD) "," replace(b.BP_RGST_NO,'-','') = '" & strUserRegNo & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					Frm1.vspdData.Col = C_user_reg_no
				    Frm1.vspdData.text = Trim(Replace(lgF2,Chr(11),""))
				else
				    Call DisplayMsgBox("970000","X",strUserRegNo,"X")	               		
				    '970000:%1 이(가) 존재하지 않습니다.
				    Frm1.vspdData.Col = C_user_reg_no
				    Frm1.vspdData.text = ""
				END IF					
            End if    
            
         Case   C_smartbill_pw1
         
            strPW1 = Trim(Frm1.vspdData.value)
            
            if  strPW1 <> "" then
               Frm1.vspdData.Col = C_smartbill_pw
               frm1.vspdData.value = strPW1
            end if
                  
                                                      
    
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101011111") 

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

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
	Dim iColumnName

	If Row <= 0 Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

Sub vspdData_GotFocus()
	ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    'Call InitSpreadComboBox
    call SetSpreadColor1(-1, -1)
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
		    Select Case Col
			    Case C_user_id_pop
				   .Row = Row
		           .Col = C_user_id
		            Call Open_User(.Text, 1)
			    Case C_user_reg_no_pop
			        frm1.vspddata.Col = C_user_reg_no			    
				    Call OpenPopUp(.text, 1)        
		    End Select    
	    End If								
    End With                     
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
	If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
'    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
'    	If (lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "") Then <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
'      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
'			If DBQuery = False Then 
'			   Call RestoreToolBar()
'			   Exit Sub 
'			End If 
'    	End If
'    End if

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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

  '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    If frm1.txtUserId.value = "" then
        frm1.txtUserNm.value = ""
    End If
    
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function		  					<%'Query db data%>
       
    FncQuery = True
            
End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
  '-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
    

    
  '-----------------------
    'Save function call area
    '----------------------- 
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 

    frm1.vspdData.ReDraw = False

    if frm1.vspdData.maxrows < 1 then exit function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    'Key field clear
    frm1.vspdData.Col=C_user_id
    frm1.vspdData.Text=""

    frm1.vspdData.Col = C_user_name
    frm1.vspdData.Text=""

    frm1.vspdData.ReDraw = True

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow 
    Dim iIntCnt

    'On Error Resume Next                                                          '☜: If process fails
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
	    ggoSpread.InsertRow  .vspdData.ActiveRow, imRow	    	    
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
			.vspdData.Row = iIntCnt 			
			.vspddata.col = C_smartbill_pw1
			.vspddata.text = "**************"
												
			.vspdData.ReDraw = True
			
	    Next
	End With
 
'	With frm1
'	
'		.vspdData.focus
'		ggoSpread.Source = .vspdData
'		
'		.vspdData.ReDraw = False
'				ggoSpread.InsertRow ,imRow
'				ggoSpread.SSSetRequired		C_smartbill_id,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired		C_smartbill_pw,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired		C_user_id,			.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetProtected	C_user_name,		.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				
'				ggoSpread.SSSetRequired	    C_user_reg_no,		.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired	    C_smartbill_pw1,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired	    C_smartbill_pw,		.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1				
'				ggoSpread.SSSetRequired	    C_Dept_Nm,		    .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired	    C_Tel_Num,		    .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				ggoSpread.SSSetRequired	    C_Email_id,	    	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
'				

'		.vspdData.ReDraw = True
'    End With

    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtuserId=" & .hUserId.value 			'☆: 조회 조건 데이타  			 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtUserId=" & Trim(.txtuserId.value)			'☆: 조회 조건 데이타
		strVal = strVal & "&txtUserNm=" & Trim(.txtuserNm.value) 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

    End If        
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
       
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk(LngMaxRow)													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
   
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>

	call SetSpreadColor1(-1, -1)
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	
    DbSave = False                                                          
    
      If LayerShowHide(1) = False then
    	Exit Function 
    End if
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
    
  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

    For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag								'☜: 신규 
					
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'☜: C=Create
										
		            .vspdData.Col = C_user_id	    	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_user_reg_no		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_user_name 		'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            		            		            
		            .vspdData.Col = C_smartbill_id	    '5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_pw	    '6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Dept_Nm	        '7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Tel_Num	        '8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Email_id	        '9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag								'☜: 수정 
		
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update

		            .vspdData.Col = C_user_id		'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_user_reg_no		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_user_name 		'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            		            		            
		            .vspdData.Col = C_smartbill_id	    '5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_pw	    '6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Dept_Nm	        '7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Tel_Num	        '8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Email_id	        '9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep 
		         

		            lGrpCnt = lGrpCnt + 1
		            
		            
		        Case ggoSpread.DeleteFlag								'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update

		            .vspdData.Col = C_user_id		'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_user_reg_no		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
	.txtMaxRows.value = lGrpCnt - 1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function


Sub txtuserId_OnChange()
    lgBlnFlgChgValue = True

     'If frm1.fpdtWk_yymm.Text = "" Then
	'	Call DisplayMsgBox("800489","x",frm1.fpdtWk_yymm.alt,"x")
	'	frm1.fpdtWk_yymm.focus
	'	Exit Sub
	'End If  

   if Trim(frm1.txtuserId.value) <> "" then
    If CommonQueryRs(" FND_USER, FND_USER_NAME "," XXSB_DTI_SM_USER  (nolock) "," FND_USER = '" & Trim(frm1.txtuserId.value) & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
       frm1.txtuserId.value = Trim(Replace(lgF0,Chr(11),""))
       frm1.txtuserNm.value =  Trim(Replace(lgF1,Chr(11),""))
    else
	   frm1.txtuserId.value = ""
       frm1.txtuserNm.value =  ""
        Call DisplayMsgBox("970000","X",frm1.txtuserId.alt,"X")	               		
        '970000:%1 이(가) 존재하지 않습니다.              
	   Exit Sub 	  
    End if 
   end if 
	
	frm1.txtuserId.focus
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">사용자ID</TD>
									<TD CLASS="TD656" colspan =3>
										<INPUT TYPE=TEXT NAME="txtuserId" SIZE=10  MAXLENGTH=13 tag="11XXXU" ALT="사용자ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Open_User1()">
										<INPUT TYPE=TEXT NAME="txtuserNm" tag="14X">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

