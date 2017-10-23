'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID  = "B82107mb1.asp"
Const BIZ_PGM_SAVE_ID = "B82107mb2.asp"
Const BIZ_PGM_DEL_ID  = "B82107mb3.asp"

'========================================================================================================
Dim gSelframeFlg                                                '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg        
Dim IsOpenPop
        
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
    
End Sub

'========================================================================================================
' Name : SetDefaultVal()        
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtReqDt.text  = BaseDt
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
              
              Call MainQuery()
                                   
              WriteCookie CookieSplit , ""
              
       End If

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
    Call SetToolbar("1110100000001111")
        
    Call InitVariables
    Call InitComboBox
    
    Call CookiePage(0)
     frm1.txtarReqNo.focus()
        
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
    
    Call SetToolbar("11101000000011")
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
       IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")           '☜: Data is changed.  Do you want to continue? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE                                                                                         '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                     '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                      '⊙: This function lock the suitable field
    Call SetToolbar("11101000000011")
    Set gActiveElement = document.ActiveElement   
    frm1.txtStatus.Value = ""
    frm1.txtReqNo.Value = ""
    FncCopy = True                                                            '☜: Processing is OK
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
              
    DbDelete = False                                                             '☜: Processing is NG
              
    If LayerShowHide(1) = False Then
       Exit Function
    End If
       
    strVal = BIZ_PGM_DEL_ID & "?txtReqNo=" & Trim(frm1.txtReqNo.value)		'☆: 조회 조건 데이타 
	
    Call RunMyBizASP(MyBizASP, strVal)
       
    DbDelete = True                                                              '⊙: Processing is NG
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
     
     strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	  '☆: 조회 조건 데이타 
	
     Call RunMyBizASP(MyBizASP, strVal)	     
     
     DbQuery = True                                                               '☜:  Run biz logic                                               '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================================
Function DbPrev()    
    Dim strVal

    DbPrev = False                                                                '⊙: Processing is NG
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	  '☆: 조회 조건 데이타 
    strVal = strVal & "&PrevNextFlg=" & "P"
          
    Call RunMyBizASP(MyBizASP, strVal)                                            '☜: 비지니스 ASP 를 가동 
    
    DbPrev = True
End Function

'========================================================================================================
' Function Name : DbNext
' Function Desc : This function is the previous data query and display
'========================================================================================================
Function DbNext()
    Dim IntRetCD    
    DbNext = False                                                               '⊙: Processing is NG
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtReqNo=" & Trim(frm1.txtarReqNo.value)	 '☆: 조회 조건 데이타 
    strVal = strVal & "&PrevNextFlg=" & "N"
                    
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜: 비지니스 ASP 를 가동 
    
    DbNext = True
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================================
Function DbQueryOk()
                                                                                               '☆: 조회 성공후 실행로직 
    DbQueryOk = false
       
    lgIntFlgMode = Parent.OPMD_UMODE                                                           '⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")                                                      '⊙: This function lock the suitable field         
    
    'A : 접수 E : 완료 S : 중단 T : 이관 
    'D : 반려 ( 변경은 반려가 없다. - 반려되면 바로 중단처리한다. )
    
    Call ggoOper.SetReqAttr(frm1.txtItemCd,  "Q")
        
    Select Case Trim(frm1.htxtStatus.Value)
           Case "R" , ""
           
                Call SetToolBar("11111000111011")
                                
           Case "A" , "E" , "S" , "T"
           
                If Trim(frm1.htxtStatus.Value) = "S" Then
                   '중단은 삭제할 수 있다. 
                   Call SetToolBar("11110000111011")
                Else   	
                   Call SetToolBar("11100000111011")
                End If 
                 
                '상태가 접수이상 이면 필수 입력및 코드정보에 관한 필드를 비활성화 한다.
                Call ggoOper.SetReqAttr(frm1.txtNewItemNm,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtNewItemNm2, "Q")
                Call ggoOper.SetReqAttr(frm1.txtNewSpec,    "Q")
                Call ggoOper.SetReqAttr(frm1.txtNewSpec2,   "Q")
                
                If Trim(frm1.htxtStatus.Value) <> "A"  Then
                   Call ggoOper.SetReqAttr(frm1.txtreq_user,  "Q") 
                   Call ggoOper.SetReqAttr(frm1.txtReqDt,  "Q") 
                   Call ggoOper.SetReqAttr(frm1.txtReqReason,  "Q")
                End If
               
           Case Else
           
    End Select
    
    frm1.txtReqReason.value = REPLACE(Trim(frm1.htxtReqReason.value) , chr(7), chr(13)&chr(10))
    
    DbQueryOk = true
       
End Function

'========================================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================================
Function DbSave() 
       
       if frm1.txtItemCd.value="" then
		 Call  DisplayMsgBox("129001","x","품목코드","x")
         exit function
  
       end if 
     
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
Sub txtReqDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtReqDt.Action = 7                                      ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P") 
       frm1.txtReqDt.Focus     
    End If
End Sub

'========================================================================================================
' Name : _Change
' Desc : developer describe this line
'========================================================================================================
Sub txtItemCd_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtreq_user_OnChange()
    lgBlnFlgChgValue = True
        Dim iDx
    Dim IntRetCd
 
    If frm1.txtreq_user.value = "" Then
        frm1.txtreq_user_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1006' and minor_cd="&filterVar(frm1.txtreq_user.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtreq_user_nm.value=""
        Else
            frm1.txtreq_user_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    
    
    
End Sub

Sub txtNewItemNm_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtNewItemNm2_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtNewSpec_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtNewSpec2_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtReqReason_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtRemark_OnChange()
    lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim arrParam(7), arrField(8), arrHeader(8)
        Dim sItemAcct , sItemKind, sItemLvl1, sItemLvl2, sItemLvl3

        If IsOpenPop = True  Then  
           Exit Function
        End If   

        IsOpenPop = True
        
        Select Case arPopUp
               Case 1 '의뢰자 
                    
                    If UCase(frm1.txtreq_user.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                                        
                    arrParam(0) = "의뢰자"       
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtreq_user.Value)
                    arrParam(4) = "MAJOR_CD = 'Y1006' "
                    arrParam(5) = "의뢰자"
       
                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"       
    
                    arrHeader(0) = "의뢰자"
                    arrHeader(1) = "의뢰자명"
                    frm1.txtreq_user.focus()
               Case Else
                    IsOpenPop = False
                    Exit Function
      End Select
        
      arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

      IsOpenPop = False
                
      If arrRet(0) = "" Then
         Exit Function
      Else
         Call SubSetPopup(arrRet,arPopUp)
      End If        
        
End Function

'======================================================================================================
'        Name : SubSetPopup()
'        Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetPopup(Byval arrRet, Byval arPopUp)
    
    With Frm1
        Select Case arPopUp 
               Case 1 '의뢰자 
                    .txtreq_user.value   = arrRet(0)
                    .txtreq_user_Nm.value = arrRet(1)           
               Case Else
                    Exit Sub
              End Select              
              
        End With
        
        lgBlnFlgChgValue = True
        
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
              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

        IsOpenPop = False
                        
        If arrRet(0) <> "" Then
          frm1.txtArReqNo.Value = arrRet(0)
          Set gActiveElement = document.activeElement 
       End If
      
       
End Function

'======================================================================================================
' 기준품목선택 PopUp
'======================================================================================================
Function OpenItemCd( )

       Dim arrRet
       Dim arrParam(5), arrField(40)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Then
          Exit Function
       End If   
       
       If lgBlnFlgChgValue = True Then
          If FncNew() = False Then
       	     Exit Function
       	  End If
       End If 

       IsOpenPop = True
       
       arrParam(0) = "CHANGE"
          
       iCalledAspName = AskPRAspName("B82101PA2")
       
       If Trim(iCalledAspName) = "" Then
          IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82101PA2", "X")
          IsOpenPop = False
          Exit Function
       End If
       
       arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
       
       IsOpenPop = False
        
       If arrRet(0) <> "" Then 
       	  If lgIntFlgMode = Parent.OPMD_UMODE Then
       	     If FncNew() = False Then
       	     	IsOpenPop = False
       	        Exit Function
             End If
       	  End If          
          frm1.txtItemCd.Value      = arrRet(0)
          frm1.txtItemNm.Value      = arrRet(1)
          frm1.txtNewItemNm.Value   = arrRet(1)
          frm1.txtSpec.Value        = arrRet(2)
          frm1.txtNewSpec.Value     = arrRet(2)
          frm1.txtItemAcct.Value    = arrRet(3)
          frm1.txtItemAcctNm.Value    = arrRet(4)
          frm1.txtItemKind.Value    = arrRet(5)
          frm1.txtItemKindNm.Value  = arrRet(6)
          frm1.txtItemLvl1.Value    = arrRet(7)
          frm1.txtItemLvl1NM.Value  = arrRet(8)
          frm1.txtItemLvl2.Value    = arrRet(9)
          frm1.txtItemLvl2Nm.Value  = arrRet(10)
          frm1.txtItemLvl3.Value    = arrRet(11)
          frm1.txtItemLvl3Nm.Value  = arrRet(12)
          frm1.txtSerialNo.Value    = arrRet(13)
          frm1.txtItemVer.Value     = arrRet(14)
          frm1.txtItemNm2.Value     = arrRet(16)
          frm1.txtNewItemNm2.Value  = arrRet(16)
       
          frm1.txtSpec2.Value       = arrRet(17)
          frm1.txtNewSpec2.Value    = arrRet(17)
          'frm1.txtItemUnit.Value    = arrRet(18)
          'frm1.cboPurType.Value     = arrRet(19)
          frm1.txtBasicItem.Value   = arrRet(21)
          frm1.txtBasicItemNm.Value = arrRet(22)
          'frm1.txtPurGroup.Value    = arrRet(23)
          'frm1.txtPurGroupNm.Value  = arrRet(24)
          'frm1.txtPurVendor.Value   = arrRet(25)
          'frm1.txtPurVendorNm.Value = arrRet(26)
          
          'If arrRet(27) = "Y"  Then
          '   frm1.rdoUnifyPurFlg1.Checked = True
          '   frm1.rdoUnifyPurFlg2.Checked = False
          'Else
          '   frm1.rdoUnifyPurFlg2.Checked = True
          '   frm1.rdoUnifyPurFlg1.Checked = False
          'End If   
          'frm1.txtNetWeight.Value= arrRet(28)
          'frm1.txtNetWeightUnit.Value= arrRet(29)
          'frm1.txtGrossWeight.Value= arrRet(30)
          'frm1.txtGrossWeightUnit.Value= arrRet(31)
          'frm1.txtCBM.Value= arrRet(32)
          'frm1.txtCBMInfo.Value= arrRet(33)
          'frm1.txtHSCd.Value= arrRet(34)
          'frm1.txtHSNm.Value= arrRet(35)
          frm1.htxtInternalCd.Value = arrRet(39)          
          
          Set gActiveElement = document.activeElement 
          
          lgBlnFlgChgValue = True
          
       End If       
       
       Call SetFocusToDocument("M")
       frm1.txtItemCd.focus
       
End Function

'======================================================================================================
' 표준규격 PopUp
'======================================================================================================
Function OpenCategory()
       Dim arrRet
       Dim arrParam(10), arrField(6)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Or UCase(frm1.txtNewSpec.className) = UCase(parent.UCN_PROTECTED) Then
           Exit Function
        End If   

       IsOpenPop = True
       
       arrParam(0) = Trim(frm1.txtItemAcct.value)
       arrParam(1) = Trim(frm1.txtItemKind.value)
       arrParam(2) = Trim(frm1.txtItemKindNm.value)
       arrParam(3) = Trim(frm1.txtItemLvl1.value)
       arrParam(4) = Trim(frm1.txtItemLvl1Nm.value)
       arrParam(5) = Trim(frm1.txtItemLvl2.value)
       arrParam(6) = Trim(frm1.txtItemLvl2Nm.value)
       arrParam(7) = Trim(frm1.txtItemLvl3.value)       
       arrParam(8) = Trim(frm1.txtItemLvl3Nm.value)
          
       iCalledAspName = AskPRAspName("B82101PA3")
       
       If Trim(iCalledAspName) = "" Then
          IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82101PA3", "X")
          IsOpenPop = False
          Exit Function
       End If
       
       If arrParam(0) = "" Then
       	  Call DisplayMsgBox("971012", "X", frm1.txtItemAcct.Alt, "X")
          frm1.txtItemAcct.focus
          IsOpenPop = False
          Exit Function
       End If 
       
       If arrParam(1) = "" Then
       	  Call DisplayMsgBox("971012", "X", frm1.txtItemKind.Alt, "X")
          frm1.txtItemKind.focus
          IsOpenPop = False
          Exit Function
       End If 
       
       If arrParam(3) = "" Then
       	  Call DisplayMsgBox("971012", "X", frm1.txtItemLvl1.Alt, "X")
          frm1.txtItemLvl1.focus
          IsOpenPop = False
          Exit Function
       End If 
       
       If arrParam(5) = "" Then
       	  Call DisplayMsgBox("971012", "X", frm1.txtItemLvl2.Alt, "X")
          frm1.txtItemLvl2.focus
          IsOpenPop = False
          Exit Function
       End If 
       
       If arrParam(7) = "" Then
       	  Call DisplayMsgBox("971012", "X", frm1.txtItemLvl3.Alt, "X")
          frm1.txtItemLvl3.focus
          IsOpenPop = False
          Exit Function
       End If  
       
       arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
              "dialogWidth=740px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
       
       IsOpenPop = False

       If arrRet(0) <> "" Then
          frm1.txtNewSpec.Value = arrRet(0)
          Set gActiveElement = document.activeElement 
          lgBlnFlgChgValue = True
       End If
       
       Call SetFocusToDocument("M")
       frm1.txtNewSpec.focus

End Function