'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY_ID  = "B82108mb1.asp"
Const BIZ_PGM_SAVE_ID = "B82108mb2.asp"

Const BIZ_PGM_JUMP_ID = "B82107ma1"
'========================================================================================================
Dim gSelframeFlg                                                '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg        
Dim IsOpenPop
Dim IsRouting
dim GradeVal        
'========================================================================================================
' Name : InitVariables()        
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode      = parent.OPMD_CMODE                       '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue  = False                                   '⊙: Indicates that no value changed
    lgIntGrpCount     = 0                                       '⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
    gblnWinEvent      = False
    lgBlnFlawChgFlg   = False
    IsRouting         = ""
End Sub

'========================================================================================================
' Name : SetDefaultVal()        
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.btn1.disabled = True
    frm1.btn2.disabled = True
    frm1.btn3.disabled = True
    frm1.btn4.disabled = True 
End Sub

'========================================================================================================
' Function Name : CookiePage
' Function Desc : 
'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877				
	Dim strTemp, arrVal

	If Kubun = 1 Then
	
	   WriteCookie CookieSplit , frm1.txtReqNo.value & parent.gRowSep 
	
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtarReqNo.value = arrVal(0)
		
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
		
		Call fncQuery()
					
		WriteCookie CookieSplit , ""
		
	End If

End Function

'========================================================================================================
' Function Name : JumpChgCheck1
' Function Desc : 
'========================================================================================================
Function JumpChgCheck1()
	Dim IntRetCD
	
	If frm1.txtReqNo.value = "" Then                                                'If there is no data.
           Call DisplayMsgBox("900002", "X", "X", "X")
           Exit Function
   	End If
   	If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
             Exit Function
        End If
   	
	If lgBlnFlgChgValue = True Then
	   IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
	   If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function
        
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Err.Clear
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = 'Y1007' ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtRgrade ,lgF0 ,lgF1 ,Chr(11))
    Call SetCombo2(frm1.txtTgrade ,lgF0 ,lgF1 ,Chr(11))
    Call SetCombo2(frm1.txtPgrade ,lgF0 ,lgF1 ,Chr(11))
    Call SetCombo2(frm1.txtQgrade ,lgF0 ,lgF1 ,Chr(11))
    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear
    Call LoadInfTB19029                                                   '☜: Load table , B_numeric_format
    Call AppendNumberPlace("6", "2", "0")
    Call AppendNumberPlace("7", "16", "4")        
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call FormatField        
    Call SetDefaultVal()
    Call SetToolbar("1110000000001111")
        
    Call InitVariables
    Call InitComboBox
    
    Call CookiePage(0)
   call frm1.txtarReqNo.focus()
        
End Sub

'=========================================
Sub FormatField()
    With frm1
        ' 날짜 OCX Foramt 설정 
        Call FormatDATEField(.txtTransDt)
        Call FormatDATEField(.txtEndDt)
        Call FormatDATEField(.txtReqDt)
    End With
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
    
    FncQuery = False                                                                                                                         '☜: Processing is NG
    Err.Clear                                                                   '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")             '☜: Data is changed.  Do you want to display it? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    If Not chkField(Document, "1") Then                                          '☜: This function check required field
       Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "2")                                      '☜: Clear Contents  Field
    Call InitVariables                                                           '⊙: Initializes local global variables
        
    Call DisableToolBar( parent.TBC_QUERY)
    
    If DbQuery = False Then
       Call  RestoreToolBar()
       Exit Function
    End If
              
    FncQuery = True                                                              '☜: Processing is OK

End Function
        
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False                                                                                                                                 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")            '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("1110000000001111")
    Call SetDefaultVal
    Call InitVariables                                                          '⊙: Initializes local global variables
            
    Set gActiveElement = document.ActiveElement  
    FncNew = True                                                                                                                                 '☜: Processing is OK

End Function
        
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False
    Err.Clear
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"x","x")                '☜: Do you want to delete? 
    If IntRetCD = vbNo Then
       Exit Function
    End If
            
    Call  DisableToolBar( parent.TBC_DELETE)
        
    If DbDelete = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncDelete = True
    
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
        
    FncSave = False    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                          '☜:There is no changed data. 
        Exit Function
    End If
   
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
     //권한 여부 체크 05.04.20
     dim configSet
    configSet = CommonQueryRs("case ITEM_R when 'Y' then 'R' else '' end+case ITEM_T when 'Y' then 'T' else '' end+case ITEM_P when 'Y'  then 'P' else '' end+case ITEM_Q when 'Y'  then 'Q' else '' end","B_CIS_ROUTING_USER","ITEM_ACCT = '" & TRIM(frm1.txtItemAcct.value) & "' AND ITEM_KIND = '" & TRIM(frm1.txtItemKind.value) & "' AND upper(USER_ID) = upper('" & parent.gUsrId & "')"  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
  
	IF INSTR(lgF0,frm1.txtUpdMode.value)= 0 then
		 call DisplayMsgBox("990016", "X","x","x")
		 call DbQueryok()          
		exit function
	
	end if       
    Call  DisableToolBar( parent.TBC_SAVE)
    
    If DbSave = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
            
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
    Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
       
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")            '☜: Data is changed.  Do you want to continue? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE                                                                                         '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: This function lock the suitable field
    Call SetToolbar("1110000000001111")
    Set gActiveElement = document.ActiveElement   
    frm1.txtStatus.Value = ""
    frm1.txtReqNo.Value = ""
    FncCopy = True                                                              '☜: Processing is OK
    lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
 Call parent.FncPrint()  
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
        Dim IntRetCD

        FncExit = False
        If lgBlnFlgChgValue = True Then
           IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")       '⊙: Data is changed.  Do you want to exit? 
           If IntRetCD = vbNo Then
              Exit Function
           End If
        End If

        FncExit = True
End Function

'========================================================================================================
' Name : FncPrev
' Desc : 
'========================================================================================================
Function FncPrev()

    Dim IntRetCD
    
    FncPrev = False 
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")                                    '☜: Will you destory previous data
       If IntRetCD = vbNo Then
         Exit Function
       End If
   End If
   
   If DbPrev = False Then
      Exit Function
   End If
    
   FncPrev = True 
End Function 

'========================================================================================================
' Name : FncNext
' Desc : 
'========================================================================================================
Function FncNext()

    Dim IntRetCD
    
    FncNext = False 
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")              '☜: Will you destory previous data
       If IntRetCD = vbNo Then
         Exit Function
       End If
   End If
   
   If DbNext = False Then
      Exit Function
   End If
   
   FncNext = True 
End Function 

'========================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function DbDelete() 
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status
              
    DbDelete = False                                                                      '☜: Processing is NG
              
    If LayerShowHide(1) = False Then
       Exit Function
    End If
    
    lgKeyStream = Trim(frm1.txtReqNo.value)  & parent.gColSep
    
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""                                    '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
       
    DbDelete = True                                                              '⊙: Processing is NG                                                              '⊙: Processing is NG
End Function       

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()                                   
       DbDeleteOk = false
       lgBlnFlgChgValue = False
       Call FncNew()
       DbDeleteOk = true
End Function

'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
     Dim strVal
     Err.Clear                                                                    '☜: Clear err status

     DbQuery = False                                                              '☜: Processing is NG

     If LayerShowHide(1) = False Then
       Exit Function
     End If
     
     If Trim(frm1.txtarReqNo.value) = "" Then
        Call DisplayMsgBox("971012", "X", frm1.txtarReqNo.Alt, "X")
        frm1.txtarReqNo.focus
        Exit Function
     End If
     
     strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	'☆: 조회 조건 데이타 
	
     Call RunMyBizASP(MyBizASP, strVal)	                                        '☜:  Run biz logic 
     
     DbQuery = True                                                             '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================================
Function DbPrev()
    Dim strVal
    
    DbPrev = False                                                                      '⊙: Processing is NG
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	       '☆: 조회 조건 데이타 
    strVal = strVal & "&PrevNextFlg=" & "P"
         
    Call RunMyBizASP(MyBizASP, strVal)                                                 '☜: 비지니스 ASP 를 가동 
    
    DbPrev = True
      	
End Function

'========================================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================================
Function DbNext()
    Dim IntRetCD
    
    DbNext = False                                                                      '⊙: Processing is NG
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	       '☆: 조회 조건 데이타 
    strVal = strVal & "&PrevNextFlg=" & "N"
                    
    Call RunMyBizASP(MyBizASP, strVal)                                                 '☜: 비지니스 ASP 를 가동 
    
    DbNext = True
    
End Function
'========================================================================================================
' Function Name : authChk
' Function Desc : 앞단계가 완료 되었는지 체크함.(LWS 2005.04.19)
' 
'========================================================================================================
function authChk(prev_chk)
	authChk=false
	
		gradeVal = replace(gradeVal,Chr(11),"")
		select case prev_chk
		case "1" 
			If mid(gradeVal,1,1) = "P" Then	authChk=true
			
		case "2"
		   If mid(gradeVal,2,1) <> "X" and trim(mid(gradeVal,2,1)) <> "" Then authChk=true
		case "3"
		   If mid(gradeVal,3,1) <> "X" and trim(mid(gradeVal,3,1)) <> "" Then authChk=true
		case "4"
		   If mid(gradeVal,4,1) <> "X" and trim(mid(gradeVal,4,1)) <> "" Then authChk=true
		         
	end select
end function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================================
Function DbQueryOk()
    Dim IntRetCD
                                                                        '☆: 조회 성공후 실행로직 
    DbQueryOk = false
	
    lgIntFlgMode = Parent.OPMD_UMODE					'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")												'⊙: This function lock the suitable field
        
    Call SetToolBar("11100000110011")
    
    frm1.txtReqReason.value = REPLACE(Trim(frm1.htxtReqReason.value) , chr(7), chr(13)&chr(10))
    
    Call SetDefaultVal()
    
    '사용자 ID별 로 권한을 체크하여 버튼을 활성화 한다.     
IntRetCD = CommonQueryRs("Isnull(R_GRADE,' ') +Isnull(T_GRADE,' ')+Isnull(P_GRADE,' ')+Isnull(q_GRADE,' ')","B_CIS_CHANGE_ITEM_NM_REQ","REQ_NO = '" & TRIM(frm1.txtarReqNo.value) & "' " ,GradeVal,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
   IntRetCD = CommonQueryRs("ITEM_R + ITEM_T + ITEM_P + ITEM_Q","B_CIS_ROUTING_USER","ITEM_ACCT = '" & TRIM(frm1.txtItemAcct.value) & "' AND ITEM_KIND = '" & TRIM(frm1.txtItemKind.value) & "' AND upper(USER_ID) = upper('" & parent.gUsrId & "')" ,IsRouting,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
  configSet = CommonQueryRs("case ITEM_R when 'Y' then '1' else '' end+case ITEM_T when 'Y' then '2' else '' end+case ITEM_P when 'Y'  then '3' else '' end+case ITEM_Q when 'Y'  then '4' else '' end","B_CIS_CONFIG","ITEM_ACCT = '" & TRIM(frm1.txtItemAcct.value) & "' AND ITEM_KIND = '" & TRIM(frm1.txtItemKind.value) & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    
 if mid(IsRouting,1,1)="Y" then
		frm1.btn1.disabled = False
 
    end if
    if mid(IsRouting,2,1)="Y" then
		if inStr(lgF0,"2") > 0 then
			if authChk( mid(lgF0,inStr(lgF0,"2")-1,1))=true then
				frm1.btn2.disabled = False
			end if
		end if
    end if
    
    if mid(IsRouting,3,1)="Y" then
    	if inStr(lgF0,"3") > 0 then
			if authChk( mid(lgF0,inStr(lgF0,"3")-1,1))=true then
				frm1.btn3.disabled = False
			end if
		end if
		
		
    end if
    if mid(IsRouting,4,1)="Y" then
		if inStr(lgF0,"4") > 0 then
		//msgbox lgF0
		//msgbox mid(lgF0,inStr(lgF0,"4")-1,1)
			if authChk( mid(lgF0,inStr(lgF0,"4")-1,1))=true then
				frm1.btn4.disabled = False
			end if
		end if
 
    end if

    
    DbQueryOk = true
	
End Function

'========================================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================================
Function DbSave()
       
       frm1.htxtReqReason.value = REPLACE(Trim(frm1.txtReqReason.value), chr(13)&chr(10) , chr(7))
       
       DbSave = False                                                                               '⊙: Processing is NG

       Call LayerShowHide(1)
       
       With frm1
            .txtFlgMode.value     = lgIntFlgMode
                        
            Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)                                                                    
            
       End With
              
       DbSave = True
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================================
Function DbSaveOk()
       DbSaveOk = false
       frm1.txtarReqNo.value = frm1.txtReqNo.value 
       Call InitVariables
       Call FncQuery()
       DbSaveOk = true
End Function
   
'========================================================================================================
' Name : _DblClick
' Desc : developer describe this line
'========================================================================================================

'========================================================================================================
' Name : _Change
' Desc : developer describe this line
'========================================================================================================

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)  
        
End Function

'======================================================================================================
'        Name : SubSetPopup()
'        Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetPopup(Byval arrRet, Byval arPopUp)
         
End Sub

'======================================================================================================
' 신규의뢰번호선택 PopUp
'======================================================================================================
Function OpenReqNo()
       Dim arrRet
       Dim arrParam(5), arrField(6)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Or UCase(frm1.txtArReqNo.className) = UCase(parent.UCN_PROTECTED) Then
          Exit Function
       End If   

       IsOpenPop = True
       
       arrParam(0) = Trim(frm1.txtArReqNo.value)
       
       arrField(0) = 1
       arrField(1) = 7
       arrField(2) = 8
                       
       iCalledAspName = AskPRAspName("B82107PA1")
        
       If Trim(iCalledAspName) = "" Then
          IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82107PA1", "X")
          IsOpenPop = False
          Exit Function
       End If
        Call SetFocusToDocument("M")
       frm1.txtArReqNo.focus
        arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
              "dialogWidth=760px; dialogHeight=410px; center: Yes; help: No; resizable: No; status: No;")

        IsOpenPop = False
                        
        If arrRet(0) <> "" Then
          frm1.txtArReqNo.Value = arrRet(0)
          Set gActiveElement = document.activeElement 
       End If
      
       
End Function

'======================================================================================================
' 접수검토 버튼 실행시...
'======================================================================================================
Function BtnR()

    Dim arrRet
    Dim arrParam(10), arrField(6)
    Dim iCalledAspName, IntRetCD 

    Err.Clear                                                                    '☜: Clear err status
        
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If IsOpenPop = True Then
       Exit Function
    End If   

    arrParam(0) = (Trim(frm1.htxtRDt.value))
    arrParam(1) = Trim(frm1.htxtRGrade.value)
    arrParam(2) = Trim(frm1.htxtRDesc.value)
    arrParam(3) = Trim(frm1.htxtRPerSon.value)
    arrParam(4) = Trim(frm1.htxtRPerSonNm.value)
    
    arrParam(5) = Trim(frm1.txtItemAcct.value)
    arrParam(6) = Trim(frm1.txtItemKind.value)
    
    If Mid(IsRouting,2,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtTGrade.value)
    ElseIf Mid(IsRouting,3,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtPGrade.value)
    ElseIf Mid(IsRouting,4,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtQGrade.value)
    Else
       arrParam(7) = ""   
    End If
    
    arrParam(8) = Trim(frm1.htxtStatus.value)
    
    If Mid(IsRouting,2,1) = "Y" And Trim(frm1.htxtTGrade.value) <> "" Then
       arrParam(9) = "X"
    ElseIf Mid(IsRouting,3,1) = "Y" And Trim(frm1.htxtPGrade.value) <> "" Then
       arrParam(9) = "X"
    ElseIf Mid(IsRouting,4,1) = "Y" And Trim(frm1.htxtQGrade.value) <> "" Then
       arrParam(9) = "X"
    End If
    
    arrParam(10)= "C" '변경의뢰 
    
    iCalledAspName = AskPRAspName("B82102PA1")
	
    If Trim(iCalledAspName) = "" Then
       IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82102PA1", "X")
       IsOpenPop = False
       Exit Function
    End If
    
    IsOpenPop = True
        
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=470px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    if UBound(arrRet) =0 then exit function
    If UBound(arrRet) > 0 Then
       If frm1.htxtRDt.value <> arrRet(0) Then
          frm1.htxtRDt.value = arrRet(0)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 1 Then
       If frm1.htxtRGrade.Value <> arrRet(1) Then
          frm1.htxtRGrade.Value = arrRet(1)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 2 Then
       If frm1.htxtRDesc.Value <> arrRet(2) Then
          frm1.htxtRDesc.Value = arrRet(2)
          lgBlnFlgChgValue = True
       End If 
    End If
    
    If UBound(arrRet) >= 3 Then
       If frm1.htxtRPerson.Value <> arrRet(3) Then
          frm1.htxtRPerson.Value = arrRet(3)
          lgBlnFlgChgValue = True
       End If

    End If  
  
    
    If UBound(arrRet) >= 4 Then
       If frm1.htxtRPersonNm.Value <> arrRet(4) Then
          frm1.htxtRPersonNm.Value = arrRet(4)
          lgBlnFlgChgValue = True
       End If
    End If
    
    '접수 
    frm1.txtUpdMode.value  = "R"
       
    Set gActiveElement = document.activeElement 
	
    Call SetFocusToDocument("M")
	
    '데이터가 변경되었으면 바로 저장한다.
    If lgBlnFlgChgValue = True Then 
       Call FncSave()
    End If
     
End Function


'======================================================================================================
' 기술검토 버튼 실행시...
'======================================================================================================
Function BtnT()
    Dim arrRet
    Dim arrParam(10), arrField(6)
    Dim iCalledAspName, IntRetCD 

    Err.Clear                                                                    '☜: Clear err status
        
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If IsOpenPop = True Then
       Exit Function
    End If   

    arrParam(0) = (Trim(frm1.htxtTDt.value))
    arrParam(1) = Trim(frm1.htxtTGrade.value)
    arrParam(2) = Trim(frm1.htxtTDesc.value)
    arrParam(3) = Trim(frm1.htxtTPerSon.value)
    arrParam(4) = Trim(frm1.htxtTPerSonNm.value)
    
    arrParam(5) = Trim(frm1.txtItemAcct.value)
    arrParam(6) = Trim(frm1.txtItemKind.value)
    
    If Mid(IsRouting,3,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtPGrade.value)
    ElseIf Mid(IsRouting,4,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtQGrade.value)
    Else
       arrParam(7) = ""   
    End If
    
    arrParam(8) = Trim(frm1.htxtStatus.value)
    
    If Mid(IsRouting,3,1) = "Y" And Trim(frm1.htxtPGrade.value) <> "" Then
       arrParam(9) = "X"
    ElseIf Mid(IsRouting,4,1) = "Y" And Trim(frm1.htxtQGrade.value) <> "" Then
       arrParam(9) = "X"
    End If
    
    arrParam(10)= "C" '변경의뢰 
    
    iCalledAspName = AskPRAspName("B82102PA2")
	
    If Trim(iCalledAspName) = "" Then
       IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82102PA2", "X")
       IsOpenPop = False
       Exit Function
    End If
    
    IsOpenPop = True
        
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=470px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    if UBound(arrRet) =0 then exit function
    If UBound(arrRet) > 0 Then
       If frm1.htxtTDt.value <> arrRet(0) Then
          frm1.htxtTDt.value = arrRet(0)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 1 Then
       If frm1.htxtTGrade.Value <> arrRet(1) Then
          frm1.htxtTGrade.Value = arrRet(1)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 2 Then
       If frm1.htxtTDesc.Value <> arrRet(2) Then
          frm1.htxtTDesc.Value = arrRet(2)
          lgBlnFlgChgValue = True
       End If 
    End If
    
    If UBound(arrRet) >= 3 Then
       If frm1.htxtTPerson.Value <> arrRet(3) Then
          frm1.htxtTPerson.Value = arrRet(3)
          lgBlnFlgChgValue = True
       End If
    End If  
    
    If UBound(arrRet) >= 4 Then
       If frm1.htxtTPersonNm.Value <> arrRet(4) Then
          frm1.htxtTPersonNm.Value = arrRet(4)
          lgBlnFlgChgValue = True
       End If
    End If
    
    '기술 
    frm1.txtUpdMode.value = "T"
       
    Set gActiveElement = document.activeElement 
	
    Call SetFocusToDocument("M")
    
    '데이터가 변경되었으면 바로 저장한다.
    If lgBlnFlgChgValue = True Then 
       Call FncSave()
    End If
     
End Function


'======================================================================================================
' 구매검토 버튼 실행시...
'======================================================================================================
Function BtnP()
    Dim arrRet
    Dim arrParam(10), arrField(6)
    Dim iCalledAspName, IntRetCD 

    Err.Clear                                                                    '☜: Clear err status
        
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If IsOpenPop = True Then
       Exit Function
    End If   

    arrParam(0) = (Trim(frm1.htxtPDt.value))
    arrParam(1) = Trim(frm1.htxtPGrade.value)
    arrParam(2) = Trim(frm1.htxtPDesc.value)
    arrParam(3) = Trim(frm1.htxtPPerSon.value)
    arrParam(4) = Trim(frm1.htxtPPerSonNm.value)
    
    arrParam(5) = Trim(frm1.txtItemAcct.value)
    arrParam(6) = Trim(frm1.txtItemKind.value)
    
    If Mid(IsRouting,4,1) = "Y" Then
       arrParam(7) = Trim(frm1.htxtQGrade.value)
    Else
       arrParam(7) = ""   
    End If
    
    arrParam(8) = Trim(frm1.htxtStatus.value)
    
    If Mid(IsRouting,4,1) = "Y" And Trim(frm1.htxtQGrade.value) <> "" Then
       arrParam(9) = "X"
    End If
    
    arrParam(10)= "C" '변경의뢰 
    
    iCalledAspName = AskPRAspName("B82102PA3")
	
    If Trim(iCalledAspName) = "" Then
       IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82102PA3", "X")
       IsOpenPop = False
       Exit Function
    End If
    
    IsOpenPop = True
        
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=470px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    if UBound(arrRet) =0 then exit function
    If UBound(arrRet) > 0 Then
       If frm1.htxtPDt.value <> arrRet(0) Then
          frm1.htxtPDt.value = arrRet(0)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 1 Then
       If frm1.htxtPGrade.Value <> arrRet(1) Then
          frm1.htxtPGrade.Value = arrRet(1)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 2 Then
       If frm1.htxtPDesc.Value <> arrRet(2) Then
          frm1.htxtPDesc.Value = arrRet(2)
          lgBlnFlgChgValue = True
       End If 
    End If
    
    If UBound(arrRet) >= 3 Then
       If frm1.htxtPPerson.Value <> arrRet(3) Then
          frm1.htxtPPerson.Value = arrRet(3)
          lgBlnFlgChgValue = True
       End If
    End If  
    
    If UBound(arrRet) >= 4 Then
       If frm1.htxtPPersonNm.Value <> arrRet(4) Then
          frm1.htxtPPersonNm.Value = arrRet(4)
          lgBlnFlgChgValue = True
       End If
    End If
    
    '구매 
    frm1.txtUpdMode.value =  "P"
       
    Set gActiveElement = document.activeElement 
	
    Call SetFocusToDocument("M")
    
    '데이터가 변경되었으면 바로 저장한다.
    If lgBlnFlgChgValue = True Then 
       Call FncSave()
    End If
         
End Function

'======================================================================================================
' 품질검토 버튼 실행시...
'======================================================================================================
Function BtnQ()
    Dim arrRet
    Dim arrParam(10), arrField(6)
    Dim iCalledAspName, IntRetCD 

    Err.Clear                                                                    '☜: Clear err status
        
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                   '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If IsOpenPop = True Then
       Exit Function
    End If   

    arrParam(0) = (Trim(frm1.htxtQDt.value))
    arrParam(1) = Trim(frm1.htxtQGrade.value)
    arrParam(2) = Trim(frm1.htxtQDesc.value)
    arrParam(3) = Trim(frm1.htxtQPerSon.value)
    arrParam(4) = Trim(frm1.htxtQPerSonNm.value)
    
    arrParam(5) = Trim(frm1.txtItemAcct.value)
    arrParam(6) = Trim(frm1.txtItemKind.value)
    
    arrParam(7) = ""
    arrParam(8) = Trim(frm1.htxtStatus.value)
    arrParam(9)= ""
    arrParam(10)= "C" '변경의뢰 
    
    iCalledAspName = AskPRAspName("B82102PA4")
	
    If Trim(iCalledAspName) = "" Then
       IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82102PA4", "X")
       IsOpenPop = False
       Exit Function
    End If
    
    IsOpenPop = True
        
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=470px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    if UBound(arrRet) =0 then exit function
    If UBound(arrRet) > 0 Then
       If frm1.htxtQDt.value <> arrRet(0) Then
          frm1.htxtQDt.value = arrRet(0)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 1 Then
       If frm1.htxtQGrade.Value <> arrRet(1) Then
          frm1.htxtQGrade.Value = arrRet(1)
          lgBlnFlgChgValue = True
       End If
    End If   
    
    If UBound(arrRet) >= 2 Then
       If frm1.htxtQDesc.Value <> arrRet(2) Then
          frm1.htxtQDesc.Value = arrRet(2)
          lgBlnFlgChgValue = True
       End If 
    End If
    
    If UBound(arrRet) >= 3 Then
       If frm1.htxtQPerson.Value <> arrRet(3) Then
          frm1.htxtQPerson.Value = arrRet(3)
          lgBlnFlgChgValue = True
       End If
    End If  
    
    If UBound(arrRet) >= 4 Then
       If frm1.htxtQPersonNm.Value <> arrRet(4) Then
          frm1.htxtQPersonNm.Value = arrRet(4)
          lgBlnFlgChgValue = True
       End If
    End If
    
    '품질 
    frm1.txtUpdMode.value =  "Q"
       
    Set gActiveElement = document.activeElement 
	
    Call SetFocusToDocument("M")
    
    '데이터가 변경되었으면 바로 저장한다.
    If lgBlnFlgChgValue = True Then 
       Call FncSave()
    End If
     
End Function