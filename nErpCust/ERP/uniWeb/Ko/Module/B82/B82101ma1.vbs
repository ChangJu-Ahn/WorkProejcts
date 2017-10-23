'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID  = "B82101mb1.asp"
Const BIZ_PGM_SAVE_ID = "B82101mb2.asp"
Const BIZ_PGM_DEL_ID  = "B82101mb3.asp"		

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
    
    frm1.btnRun.disabled = True
    
    frm1.rdoDerive1.disabled = True
    frm1.rdoDerive2.disabled = True
    frm1.rdoDerive1.Checked  = False
    frm1.rdoDerive2.Checked  = True
    frm1.hrdoDerive.value    = "N" 
    
    frm1.rdoUnifyPurFlg1.Checked = False
    frm1.rdoUnifyPurFlg2.Checked = True    
    frm1.hrdoUnifyPurFlg.value   = "N"   
    
    Call ggoOper.SetReqAttr(frm1.cboItemVer, "Q")
    Call ggoOper.SetReqAttr(frm1.button1, "Q")
    
End Sub

'========================================================================================================
' Name : SetDefaultVal()        
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
    frm1.txtEndReqDt.text = BaseDt
    frm1.txtReqDt.text = BaseDt
    frm1.txtValidFromDt.text = BaseDt
	frm1.txtValidToDt.text =BaseDtTo
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
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = 'P1001' ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct ,lgF0 ,lgF1 ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = 'P1003' ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboPurType ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = 'Y1004' ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemVer ,lgF0  ,lgF1  ,Chr(11))
    
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
        Call FormatDATEField(.txtEndReqDt)
        Call FormatDATEField(.txtReqDt)
        Call FormatDATEField(.txtValidFromDt)
        Call FormatDATEField(.txtValidToDt)
        ' 숫자 OCX Foramt 설정 
        Call FormatDoubleSingleField(.txtNetWeight)
        Call FormatDoubleSingleField(.txtGrossWeight)
        Call FormatDoubleSingleField(.txtCBM)
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
       IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")                                 '☜: Data is changed.  Do you want to continue? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE                                                                                         '⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                                                    '⊙: This function lock the suitable field
    Call SetToolbar("11101000000011")
    Set gActiveElement = document.ActiveElement   
    frm1.txtStatus.Value = ""
    frm1.txtReqNo.Value = ""
    FncCopy = True                                                            '☜: Processing is OK
    frm1.btnRun.disabled = True
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
       
    strVal = BIZ_PGM_DEL_ID & "?txtReqNo=" & Trim(frm1.txtReqNo.value)			'☆: 조회 조건 데이타 
	
    Call RunMyBizASP(MyBizASP, strVal)
       
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
	
     Call RunMyBizASP(MyBizASP, strVal)	     
     
     DbQuery = True                                                    
     
End Function

'========================================================================================================
' Function Name : DbPrev
' Function Desc : This function is the previous data query and display
'========================================================================================================
Function DbPrev()
    
    DbPrev = False                                                                      '⊙: Processing is NG
    
    Dim strVal
    
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
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================================
Function DbQueryOk()
                                                                                               '☆: 조회 성공후 실행로직 
    DbQueryOk = false
       
    lgIntFlgMode = Parent.OPMD_UMODE                                                           '⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    'R : 의뢰 A : 접수 D : 반려 E : 완료 S : 중단 T : 이관 
    
    Select Case Trim(frm1.htxtStatus.Value)
           Case "R" , "" , "D"
                
                Call ggoOper.LockField(Document, "Q")                                                      '⊙: This function lock the suitable field
                
                If Trim(frm1.htxtStatus.Value) = "D" Then
                	
                   '상태가 반려이면 재의뢰 버튼을 활성화 한다. 
                   frm1.btnRun.disabled = False
                  
                   '품목코드에 영향을 미치는 컬럼은 비활성화 한다.
                   If frm1.htxtItemCd.Value <> "" Then
                      Call ggoOper.SetReqAttr(frm1.txtItemLvl1, "Q")
                      Call ggoOper.SetReqAttr(frm1.txtItemLvl2, "Q")
                      Call ggoOper.SetReqAttr(frm1.txtItemLvl3, "Q")
                      Call ggoOper.SetReqAttr(frm1.cboItemVer,  "Q")
                      Call ggoOper.SetReqAttr(frm1.cboItemAcct, "Q")
                      Call ggoOper.SetReqAttr(frm1.txtItemKind, "Q")
                      Call ggoOper.SetReqAttr(frm1.rdoDerive1,  "Q")
                      Call ggoOper.SetReqAttr(frm1.rdoDerive2,  "Q")
                      Call SetToolBar("11101000111011")
                   Else
                      Call SetToolBar("11111000111011")
                   End If 
                Else
                   '의뢰상태는 삭제/수정 가능하다.
                   Call SetToolBar("11111000111011")
                
                   '대,중,소 길이및 필수항목 속성을 변경한다.                
                   Call SetColumn(frm1.cboItemAcct.value , frm1.txtItemKind.value)
                End If      
                
           Case "A" , "E" , "S" , "T"
                
                Call SetToolBar("11100000111011")
                 
                '상태가 접수이상 이면 필수 입력및 코드정보에 관한 필드를 비활성화 한다.      
                Call ggoOper.SetReqAttr(frm1.txtItemLvl1, "Q")
                Call ggoOper.SetReqAttr(frm1.txtItemLvl2, "Q")
                Call ggoOper.SetReqAttr(frm1.txtItemLvl3, "Q")
                Call ggoOper.SetReqAttr(frm1.cboItemVer,  "Q")
                Call ggoOper.SetReqAttr(frm1.cboItemAcct, "Q")
                Call ggoOper.SetReqAttr(frm1.txtItemKind, "Q")
                Call ggoOper.SetReqAttr(frm1.rdoDerive1,  "Q")
                Call ggoOper.SetReqAttr(frm1.rdoDerive2,  "Q")                
                Call ggoOper.SetReqAttr(frm1.txtItemNm,  "Q")                
                Call ggoOper.SetReqAttr(frm1.txtSpec,  "Q")                
                Call ggoOper.SetReqAttr(frm1.txtItemUnit,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtPurVendor,  "Q")
                Call ggoOper.SetReqAttr(frm1.rdoUnifyPurFlg1,  "Q")
                Call ggoOper.SetReqAttr(frm1.rdoUnifyPurFlg2,  "Q")
                Call ggoOper.SetReqAttr(frm1.cboPurType,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtPurGroup,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtNetWeight,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtNetWeightUnit,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtGrossWeight,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtGrossWeightUnit,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtCBM,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtCBMInfo,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtHSCd,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtValidFromDt,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtValidToDt,  "Q")
                Call ggoOper.SetReqAttr(frm1.txtDocNo,  "Q")
                
                If Trim(frm1.htxtStatus.Value) <> "A"  Then
                   Call ggoOper.SetReqAttr(frm1.txtreq_user,  "Q") 
                   Call ggoOper.SetReqAttr(frm1.txtReqDt,  "Q") 
                   Call ggoOper.SetReqAttr(frm1.txtEndReqDt,  "Q")
                   Call ggoOper.SetReqAttr(frm1.txtReqReason,  "Q")
                   Call ggoOper.SetReqAttr(frm1.txtRemark,  "Q")
                   Call ggoOper.SetReqAttr(frm1.txtItemNm2,  "Q")
                   Call ggoOper.SetReqAttr(frm1.txtSpec2,  "Q")
                End If
               
           Case Else
           
    End Select
    
    frm1.txtReqReason.value = REPLACE(Trim(frm1.htxtReqReason.value) , chr(7), chr(13)&chr(10))
    frm1.txtRemark.value    = REPLACE(Trim(frm1.htxtRemark.value)    , chr(7), chr(13)&chr(10))
    
    DbQueryOk = true
       
End Function

'========================================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================================
Function DbSave() 
       
       frm1.htxtReqReason.value = REPLACE(Trim(frm1.txtReqReason.value), chr(13)&chr(10) , chr(7))
       frm1.htxtRemark.value    = REPLACE(Trim(frm1.txtRemark.value),    chr(13)&chr(10) , chr(7))
       
       DbSave = False                                                       '⊙: Processing is NG

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
       Call InitVariables
       Call FncQuery()
       DbSaveOk = true
End Function
   
'========================================================================================================
' Name : _DblClick
' Desc : developer describe this line
'========================================================================================================
Sub txtEndReqDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtEndReqDt.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P")
       frm1.txtEndReqDt.Focus
    End If
End Sub

Sub txtReqDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtReqDt.Action = 7                                      ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P") 
       frm1.txtReqDt.Focus     
    End If
End Sub

Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtValidFromDt.Action = 7                                ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P")  
       frm1.txtValidFromDt.Focus   
    End If
End Sub

Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtValidToDt.Action = 7                                  ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P")
       frm1.txtValidToDt.Focus
    End If
End Sub

'========================================================================================================
' Name : _Change
' Desc : developer describe this line
'========================================================================================================
Sub txtEndReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub cboItemAcct_OnChange()
    frm1.txtItemKind.value = ""
    frm1.txtItemKindNm.value = ""
    
    frm1.txtItemLvl1.value = ""
    frm1.txtItemLvl1Nm.value = ""
    
    frm1.txtItemLvl2.value = ""
    frm1.txtItemLvl2Nm.value = ""
    
    frm1.txtItemLvl3.value = ""
    frm1.txtItemLvl3Nm.value = ""
        
    //frm1.txtSerialNo.value = ""
        
    //frm1.txtBasicItem.value = ""
    //frm1.txtBasicItemNm.value = ""
    
    //frm1.cboItemVer.value   = ""
        
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtItemKind_OnChange()
    Dim IntRetCD
    Dim sItemAcct
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = filtervar(Trim(frm1.cboItemAcct.value),"''","S")
   
    If sItemAcct = "" Then
           Call DisplayMsgBox("971012", "X", "품목계정", "X")
           frm1.cboItemAcct.focus
           Exit Sub
    End If
                    
    If Trim(frm1.txtItemKind.value) = "" Then
       frm1.txtItemKindNm.Value = ""
    Else
       IntRetCD = CommonQueryRs("A.MINOR_NM","B_MINOR A, B_CIS_CONFIG B"," A.MAJOR_CD = 'Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = " & sItemAcct & " AND B.ITEM_KIND = " & filtervar(TRIM(frm1.txtItemKind.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false then
         frm1.txtItemKind.focus
       Else   
         frm1.txtItemKindNm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
   
   //call txtItemLvl1_OnChange()
   //call txtItemLvl2_OnChange()
   //call txtItemLvl3_OnChange()
     
    //frm1.txtSerialNo.value = ""
        
    //frm1.txtBasicItem.value = ""
    //frm1.txtBasicItemNm.value = ""
    
    //frm1.cboItemVer.value   = ""
    frm1.txtItemLvl1.value=""
    frm1.txtItemLvl1Nm.value=""
    frm1.txtItemLvl2.value=""
    frm1.txtItemLvl2Nm.value=""
    frm1.txtItemLvl3.value=""
    frm1.txtItemLvl3Nm.value=""
    
    '----------------------------------------------------------------
    ' 컬럼속성을 재설정한다.
    '----------------------------------------------------------------
    //Call SetColumn(sItemAcct , frm1.txtItemKind.value)
    '----------------------------------------------------------------
        
    lgBlnFlgChgValue = True
                    
End Sub

'========================================================================================================
Sub txtItemLvl1_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = filtervar(Trim(frm1.cboItemAcct.value),"''","S")
    sItemKind = filtervar(Trim(frm1.txtItemKind.value),"''","S")
    sItemLvl1 = filtervar(Trim(frm1.txtItemLvl1.value),"''","S")
    sItemLvl2 = filtervar(Trim(frm1.txtItemLvl2.value),"''","S")
    
    If Trim(frm1.txtItemLvl1.value) = "" Then
       frm1.txtItemLvl1Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT = " & sItemAcct & " AND ITEM_KIND = " & sItemKind & " AND ITEM_LVL = 'L1' AND CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl1.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
                //frm1.txtItemLvl1.value = ""
                frm1.txtItemLvl1Nm.value = ""
       Else   
         frm1.txtItemLvl1Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If

   
    frm1.txtItemLvl2.value=""
    frm1.txtItemLvl2Nm.value=""
    frm1.txtItemLvl3.value=""
    frm1.txtItemLvl3Nm.value=""
                     
   //call txtItemLvl2_OnChange()
  // call txtItemLvl3_OnChange()
   
    //frm1.txtSerialNo.value = ""
        
    //frm1.txtBasicItem.value = ""
    //frm1.txtBasicItemNm.value = ""
    
    //frm1.cboItemVer.value   = ""
    
    lgBlnFlgChgValue = True
End Sub

Sub txtItemLvl2_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind , sItemLvl1
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
   sItemAcct = filtervar(Trim(frm1.cboItemAcct.value),"''","S")
    sItemKind = filtervar(Trim(frm1.txtItemKind.value),"''","S")
    sItemLvl1 = filtervar(Trim(frm1.txtItemLvl1.value),"''","S")
    sItemLvl2 = filtervar(Trim(frm1.txtItemLvl2.value),"''","S")
   
    
    If Trim(frm1.txtItemLvl2.value) = "" Then
       frm1.txtItemLvl2Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT = " & sItemAcct & " AND ITEM_KIND = " & sItemKind & " AND PARENT_CLASS_CD = " & sItemLvl1 & " AND ITEM_LVL = 'L2' AND CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl2.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
         frm1.txtItemLvl2Nm.value = ""
       Else   
         frm1.txtItemLvl2Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
    
    frm1.txtItemLvl2.value=""
    frm1.txtItemLvl2Nm.value=""
    frm1.txtItemLvl3.value=""
    frm1.txtItemLvl3Nm.value=""
                     
                     
    //call txtItemLvl3_OnChange()
    //frm1.txtSerialNo.value = ""
        
   // frm1.txtBasicItem.value = ""
    //frm1.txtBasicItemNm.value = ""
    
    //frm1.cboItemVer.value   = ""
    lgBlnFlgChgValue = True
End Sub

Sub txtItemLvl3_OnChange()

    Dim IntRetCD
    Dim sItemAcct , sItemKind , sItemLvl1, sItemLvl2
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    sItemAcct = filtervar(Trim(frm1.cboItemAcct.value),"''","S")
    sItemKind = filtervar(Trim(frm1.txtItemKind.value),"''","S")
    sItemLvl1 = filtervar(Trim(frm1.txtItemLvl1.value),"''","S")
    sItemLvl2 = filtervar(Trim(frm1.txtItemLvl2.value),"''","S")
   
        
    If Trim(frm1.txtItemLvl3.value) = "" Then
       frm1.txtItemLvl3Nm.Value = ""
    Else
       IntRetCD = CommonQueryRs("CLASS_NAME","B_CIS_ITEM_CLASS","ITEM_ACCT = " & sItemAcct & " AND ITEM_KIND = " & sItemKind & " AND PARENT_CLASS_CD = " & sItemLvl2 & " AND ITEM_LVL = 'L3' AND CLASS_CD = " & filtervar(TRIM(frm1.txtItemLvl3.value),"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       If IntRetCd = false Then
         frm1.txtItemLvl3Nm.value = ""
       Else 
         frm1.txtItemLvl3Nm.value = Trim(Replace(lgF0,Chr(11),""))
       End If
    End If
    
    frm1.txtSerialNo.value = ""
        
    frm1.txtBasicItem.value = ""
    frm1.txtBasicItemNm.value = ""
    
    lgBlnFlgChgValue = True
End Sub

Sub txtSerialNo_OnChange()
    lgBlnFlgChgValue = True
End Sub



Sub rdoDerive1_OnChange()
    lgBlnFlgChgValue = True
    frm1.hrdoDerive.value = "Y"
End Sub

Sub rdoDerive2_OnChange()
    lgBlnFlgChgValue = True
    frm1.hrdoDerive.value = "N"
End Sub

Sub cboItemVer_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtItemNm_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtItemNm2_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtSpec_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtSpec2_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub cboPurType_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub rdoUnifyPurFlg1_OnChange()
    lgBlnFlgChgValue = True
    frm1.hrdoUnifyPurFlg.value = "Y"
End Sub

Sub rdoUnifyPurFlg2_OnChange()
    lgBlnFlgChgValue = True
    frm1.hrdoUnifyPurFlg.value = "N"
End Sub

Sub txtPurVendor_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtPurGroup_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtreq_user_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtNetWeight_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub txtNetWeight_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtGrossWeight_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCBM_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtHSCd_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtValidFromDt_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtValidToDt_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtDocNo_OnChange()
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
               Case 1 '품목구분 
                    If UCase(frm1.txtItemKind.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                    
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    If sItemAcct = "" Then
                           Call DisplayMsgBox("971012", "X", frm1.cboItemAcct.Alt , "X")
                           frm1.cboItemAcct.focus
                           IsOpenPop = False
                           Exit Function
                    End If
                    
                    arrParam(0) = frm1.txtItemKind.Alt
                    arrParam(1) = "B_MINOR A, B_CIS_CONFIG B"
                    arrParam(2) = Trim(frm1.txtItemKind.value)
                    arrParam(4) = "A.MAJOR_CD = 'Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = '" & sItemAcct & "'"
                    arrParam(5) = frm1.cboItemAcct.Alt

                    arrField(0) = "A.MINOR_CD"
                    arrField(1) = "A.MINOR_NM"
    
                    arrHeader(0) = frm1.txtItemKind.Alt
                    arrHeader(1) = frm1.txtItemKindNm.Alt
                    frm1.txtItemKind.focus
               Case 2 '대분류 
                    If UCase(frm1.txtItemLvl1.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                    
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    If sItemAcct = "" Then
                           Call DisplayMsgBox("971012", "X", frm1.cboItemAcct.Alt, "X")
                           frm1.cboItemAcct.focus
                           IsOpenPop = False
                           Exit Function
                    End If
                    
                    sItemKind = Trim(frm1.txtItemKind.value)
                    If sItemKind = "" Then
                           Call DisplayMsgBox("971012", "X", frm1.txtItemKind.Alt, "X")
                           frm1.txtItemKind.focus
                           IsOpenPop = False
                           Exit Function
                    End If
                    
                    arrParam(0) = frm1.txtItemLvl1.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl1.value)
                    arrParam(4) = "ITEM_ACCT = '" & sItemAcct & "' AND ITEM_KIND = '" & sItemKind & "' AND ITEM_LVL = 'L1' "
                    arrParam(5) = frm1.txtItemLvl1.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl1.Alt
                    arrHeader(1) = frm1.txtItemLvl1Nm.Alt
                     frm1.txtItemLvl1.focus
               Case 3 '중분류 
                    If UCase(frm1.txtItemLvl2.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                    
                    sItemLvl1 = Trim(frm1.txtItemLvl1.value)
                    If sItemLvl1 = "" Then
                           Call DisplayMsgBox("971012", "X", frm1.txtItemLvl1.Alt, "X")
                           frm1.txtItemLvl1.focus
                           IsOpenPop = False
                           Exit Function
                    End If
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    sItemKind = Trim(frm1.txtItemKind.value)
                    
                    arrParam(0) = frm1.txtItemLvl2.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl2.value)
                    arrParam(4) = "ITEM_ACCT = '" & sItemAcct & "' AND ITEM_KIND = '" & sItemKind & "' AND ITEM_LVL = 'L2' AND PARENT_CLASS_CD = '" & sItemLvl1 & "' "
                    arrParam(5) = frm1.txtItemLvl2.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl2.Alt
                    arrHeader(1) = frm1.txtItemLvl2Nm.Alt
                     frm1.txtItemLvl2.focus
               Case 4 '소분류 
                    If UCase(frm1.txtItemLvl3.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                    
                    sItemLvl2 = Trim(frm1.txtItemLvl2.value)
                    If sItemLvl2 = "" Then
                           Call DisplayMsgBox("971012", "X", frm1.txtItemLvl2.Alt, "X")
                           frm1.txtItemLvl2.focus
                           IsOpenPop = False
                           Exit Function
                    End If
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    sItemKind = Trim(frm1.txtItemKind.value)
                    
                    arrParam(0) = frm1.txtItemLvl3.Alt
                    arrParam(1) = "B_CIS_ITEM_CLASS"
                    arrParam(2) = Trim(frm1.txtItemLvl3.value)
                    arrParam(4) = "ITEM_ACCT = '" & sItemAcct & "' AND ITEM_KIND = '" & sItemKind & "' AND ITEM_LVL = 'L3' AND PARENT_CLASS_CD = '" & sItemLvl2 & "' "
                    arrParam(5) = frm1.txtItemLvl3.Alt

                    arrField(0) = "CLASS_CD"
                    arrField(1) = "CLASS_NAME"
    
                    arrHeader(0) = frm1.txtItemLvl3.Alt
                    arrHeader(1) = frm1.txtItemLvl3Nm.Alt
					frm1.txtItemLvl3.focus
               Case 9 '공급처 
                   If UCase(frm1.txtPurVendor.className) = UCase(parent.UCN_PROTECTED) Then 
                      IsOpenPop = False
                      Exit Function
                   End If
                   arrParam(0) = frm1.txtPurVendor.Alt                                 
                   arrParam(1) = "B_BIZ_PARTNER"
                   arrParam(2) = Trim(frm1.txtPurVendor.Value)
                   arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y' AND IN_OUT_FLAG = 'O'"       
                   arrParam(5) = frm1.txtPurVendor.Alt                                         
       
                   arrField(0) = "BP_CD"
                   arrField(1) = "BP_NM"
                   arrField(2) = "REPRE_NM"       
                   arrField(3) = "BP_RGST_NO"                            
    
                   arrHeader(0) = frm1.txtPurVendor.Alt                    
                   arrHeader(1) = frm1.txtPurVendorNm.Alt
                   arrHeader(2) = "대표자"
                   arrHeader(3) = "사업자등록번호"
				 frm1.txtPurVendor.focus
               Case 10 '구매그룹 
                   If UCase(frm1.txtPurGroup.className) = UCase(parent.UCN_PROTECTED) Then 
                      IsOpenPop = False
                      Exit Function
                   End If
                    
                   arrParam(0) = frm1.txtPurGroup.Alt
                   arrParam(1) = "B_Pur_Grp"
                   arrParam(2) = Trim(frm1.txtPurGroup.Value)
                   arrParam(4) = "USAGE_FLG='Y'"                     
                   arrParam(5) = frm1.txtPurGroup.Alt                 
       
                   arrField(0) = "PUR_GRP"       
                   arrField(1) = "PUR_GRP_NM"       
    
                   arrHeader(0) = frm1.txtPurGroup.Alt       
                   arrHeader(1) = frm1.txtPurGroupNm.Alt
					frm1.txtPurGroup.focus
               Case 12 '의뢰자 
                    
                   If UCase(frm1.txtreq_user.className) = UCase(parent.UCN_PROTECTED) Then 
                      IsOpenPop = False
                      Exit Function
                   End If
                                        
                   arrParam(0) = frm1.txtreq_user.Alt     
                   arrParam(1) = "B_MINOR"
                   arrParam(2) = Trim(frm1.txtreq_user.Value)
                   arrParam(4) = "MAJOR_CD = 'Y1006' "
                   arrParam(5) = frm1.txtreq_user.Alt
       
                   arrField(0) = "MINOR_CD"
                   arrField(1) = "MINOR_NM"       
    
                   arrHeader(0) = frm1.txtreq_user.Alt
                   arrHeader(1) = frm1.txtreq_user_Nm.Alt
                   frm1.txtreq_user.focus 
               Case 8, 13 , 14 '재고단위, Net중량단위 , 'Gross중량단위               
                    If arPopUp = 8 Then
                       If UCase(frm1.txtItemUnit.className) = UCase(parent.UCN_PROTECTED) Then 
                          IsOpenPop = False
                          Exit Function
                       End If       
                    ElseIf arPopUp = 13 Then       
                       If UCase(frm1.txtNetWeightUnit.className) = UCase(parent.UCN_PROTECTED) Then 
                          IsOpenPop = False
                          Exit Function
                       End If
                    Else
                       If UCase(frm1.txtGrossWeightUnit.className) = UCase(parent.UCN_PROTECTED) Then 
                          IsOpenPop = False
                          Exit Function
                       End If
                    End If 
                     
                   arrParam(0) = "단위팝업"       
                   arrParam(1) = "B_UNIT_OF_MEASURE"
                 
                   If arPopUp = 13 Then              
                      arrParam(2) = Trim(frm1.txtNetWeightUnit.Value)
                      frm1.txtNetWeightUnit.focus 
                   elseif  arPopUp = 14 Then   
					  arrParam(2) = Trim(frm1.txtGrossWeightUnit.Value)
                      frm1.txtGrossWeightUnit.focus 
                   Else
                  
                      arrParam(2) = Trim(frm1.txtItemUnit.Value)
                      frm1.txtItemUnit.focus 
                   End If
                   arrParam(3) = ""
                   arrParam(4) = "DIMENSION <> 'TM' "                     
                   arrParam(5) = "단위"
       
                   arrField(0) = "UNIT"       
                   arrField(1) = "UNIT_NM"       
    
                   arrHeader(0) = "단위"              
                   arrHeader(1) = "단위명"
               
               Case 15 'HS코드 
                   If UCase(frm1.txtHSCd.className) = UCase(parent.UCN_PROTECTED) Then 
                      IsOpenPop = False
                      Exit Function
                   End If
                    
                   arrParam(0) = frm1.txtHSCd.Alt   
                   arrParam(1) = "B_HS_CODE"                            
                   arrParam(2) = Trim(frm1.txtHSCd.Value)
                   arrParam(3) = ""
                   arrParam(4) = ""                     
                   arrParam(5) = frm1.txtHSCd.Alt 
       
                   arrField(0) = "HS_CD"       
                   arrField(1) = "HS_NM"
                   arrField(2) = "HS_SPEC"       
                   arrField(3) = "HS_UNIT"
                   
                   arrHeader(0) = frm1.txtHSCd.Alt               
                   arrHeader(1) = frm1.txtHSNm.Alt 
                   arrHeader(2) = "HS규격"
                   arrHeader(3) = "HS단위"
					frm1.txtHSCd.focus 
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

    lgBlnFlgChgValue = True
    
    With Frm1
        Select Case arPopUp
               Case 1 '품목구분 
                    .txtItemKind.value   = arrRet(0)
                    .txtItemKindNm.value = arrRet(1)
                  '----------------------------------------------------------------
                  ' 컬럼속성을 재설정한다.
                  '----------------------------------------------------------------
                    Call SetColumn(.cboItemAcct.value , arrRet(0))
                   
                     frm1.txtItemLvl1.value=""
                     frm1.txtItemLvl1Nm.value=""
                     frm1.txtItemLvl2.value=""
                     frm1.txtItemLvl2Nm.value=""
                     frm1.txtItemLvl3.value=""
                     frm1.txtItemLvl3Nm.value=""
                
                    
               Case 2 '대분류 
                    .txtItemLvl1.value   = arrRet(0)
                    .txtItemLvl1Nm.value = arrRet(1)
                    
                     frm1.txtItemLvl2.value=""
                     frm1.txtItemLvl2Nm.value=""
                     frm1.txtItemLvl3.value=""
                     frm1.txtItemLvl3Nm.value=""
                     
               Case 3 '중분류 
                    .txtItemLvl2.value   = arrRet(0)
                    .txtItemLvl2Nm.value = arrRet(1)
                    
                     frm1.txtItemLvl3.value=""
                     frm1.txtItemLvl3Nm.value=""
                    
               Case 4 '소분류 
                    .txtItemLvl3.value   = arrRet(0)
                    .txtItemLvl3Nm.value = arrRet(1)
                    
               Case 9 '공급처 
                    .txtPurVendor.value   = arrRet(0)
                    .txtPurVendorNm.value = arrRet(1)
    
               Case 10 '구매그룹 
                    .txtPurGroup.value   = arrRet(0)
                    .txtPurGroupNm.value = arrRet(1)
               
               Case 12 '의뢰자 
                    .txtreq_user.value   = arrRet(0)
                    .txtreq_user_Nm.value = arrRet(1)
                    
               Case 8, 13 , 14 '재고단위, Net중량단위 , 'Gross중량단위               
                    If arPopUp = 8 Then
                       .txtItemUnit.focus()
                       .txtItemUnit.value       = arrRet(0)
                    ElseIf arPopUp = 13 Then
					   .txtNetWeightUnit.focus()
                       .txtNetWeightUnit.value  = arrRet(0)
                       
                    Else
                       .txtGrossWeightUnit.focus()
                       .txtGrossWeightUnit.value = arrRet(0)
                      
                    End If 
               
               Case 15 'HS코드 
                    .txtHSCd.value = arrRet(0)
                    .txtHSNm.value = arrRet(1)
               
               Case Else
                    Exit Sub
              End Select              
              
        End With
End Sub

'======================================================================================================
'        Name : SetColumn()
'        Description : 품목계정과 품목구분에 따라서 컬럼의 길이와 속성 설정한다.
'=======================================================================================================
Sub SetColumn(Byval arItemAcct, Byval arItemKind)
    Dim IntRetCD
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
   
    IntRetCD = CommonQueryRs("A.ITEM_LVL1, A.ITEM_LVL2, A.ITEM_LVL3, A.ITEM_SEQNO, A.ITEM_LVL_D, A.ITEM_VER  ","B_CIS_CONFIG A"," A.ITEM_ACCT = '" & arItemAcct & "' AND A.ITEM_KIND = '" & arItemKind & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If IntRetCd = false then
       '관련컬럼 전부 Desable
       Call ggoOper.SetReqAttr(frm1.txtItemLvl1, "Q")
       Call ggoOper.SetReqAttr(frm1.txtItemLvl2, "Q")
       Call ggoOper.SetReqAttr(frm1.txtItemLvl3, "Q")
       Call ggoOper.SetReqAttr(frm1.cboItemVer,  "Q")
       frm1.rdoDerive1.disabled = False
       frm1.rdoDerive2.disabled = False
    Else
       '대분류 
       If CDbl(lgF0) > 0 Then 
          Call ggoOper.SetReqAttr(frm1.txtItemLvl1, "N")
          //frm1.txtItemLvl1.MAXLENGTH = CDbl(lgF0)
       Else
          Call ggoOper.SetReqAttr(frm1.txtItemLvl1, "Q")
          //frm1.txtItemLvl1.MAXLENGTH = 0
       End If
       '중분류 
       If CDbl(lgF1) > 0 Then 
          Call ggoOper.SetReqAttr(frm1.txtItemLvl2, "N")
          //frm1.txtItemLvl2.MAXLENGTH = CDbl(lgF1)
       Else
          Call ggoOper.SetReqAttr(frm1.txtItemLvl2, "Q")
          //frm1.txtItemLvl2.MAXLENGTH = 0
       End If
       '소분류 
       If CDbl(lgF2) > 0 Then 
          Call ggoOper.SetReqAttr(frm1.txtItemLvl3, "N")
          //frm1.txtItemLvl3.MAXLENGTH = CDbl(lgF2)
       Else
          Call ggoOper.SetReqAttr(frm1.txtItemLvl3, "Q")
          //frm1.txtItemLvl3.MAXLENGTH = 0
       End If   
       
       'Serial No : 코드생성할때 생성한다.  
           
       '파생번호  : 코드생성할때 생성한다.         
       If CDbl(lgF4) > 0 Then 
       	  If frm1.txtBasicItem.Value <> "" Then
       	     frm1.rdoDerive1.disabled = False
             frm1.rdoDerive2.disabled = False
       	  End If   
       Else
          frm1.rdoDerive1.disabled = True
          frm1.rdoDerive2.disabled = True
       End If
           
       '이슈부여  
       If CDbl(lgF5) > 0 Then 
       	  If frm1.txtBasicItem.Value <> "" Then
       	     Call ggoOper.SetReqAttr(frm1.cboItemVer, "D")
       	  End If   
       Else
          Call ggoOper.SetReqAttr(frm1.cboItemVer, "Q")
       End If
       
    End If
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
         frm1.txtArReqNo.focus()             
        iCalledAspName = AskPRAspName("B82101PA1")
        frm1.txtarReqNo.focus()
        arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

        IsOpenPop = False
         frm1.txtarReqNo.focus()                
        If arrRet(0) <> "" Then
          frm1.txtArReqNo.Value   = arrRet(0)
          frm1.txtArItemCd.Value  = arrRet(1)
          frm1.txtArItemNm.Value  = arrRet(2)
          Set gActiveElement = document.activeElement 
       End If
       //Call SetFocusToDocument("M")
       //frm1.txtArReqNo.focus
       
End Function

'======================================================================================================
' 재의뢰내역 참조 
'======================================================================================================
Function OpenReReqRef()

        Dim arrRet
       Dim arrParam(5), arrField(10)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Or UCase(frm1.txtArReqNo.className) = UCase(parent.UCN_PROTECTED) Then
           Exit Function
        End If   
        
        If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
           Call  DisplayMsgBox("900002","x","x","x")                                
           Exit Function
        End If
    
       IsOpenPop = True
       
       arrParam(0) = Trim(frm1.htxtInternalCd.value)        
                
        iCalledAspName = AskPRAspName("B82101RA1")
        
        arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

        IsOpenPop = False
                        
        If strRet = "" Then
           If Err.Number <> 0 Then
              Err.Clear 
           End If
           Exit Function
        End If
        
End Function

'======================================================================================================
' 기준품목선택 PopUp
'======================================================================================================
Function OpenBasicItem()

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
       
       arrParam(0) = "NEW"
          
       iCalledAspName = AskPRAspName("B82101PA2")
       
       If Trim(iCalledAspName) = "" Then
          IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82101PA2", "X")
          IsOpenPop = False
          Exit Function
       End If      
      // frm1.txtArReqNo.focus
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
          frm1.txtBasicItem.Value   = arrRet(0)
          frm1.txtBasicItemNm.Value = arrRet(1)
          frm1.txtItemNm.Value      = arrRet(1)
          frm1.txtSpec.Value        = arrRet(2)
          frm1.cboItemAcct.Value    = arrRet(3)
          frm1.txtItemKind.Value    = arrRet(5)
          frm1.txtItemKindNm.Value  = arrRet(6)
          frm1.txtItemLvl1.Value    = arrRet(7)
          frm1.txtItemLvl1NM.Value  = arrRet(8)
          frm1.txtItemLvl2.Value    = arrRet(9)
          frm1.txtItemLvl2Nm.Value  = arrRet(10)
          frm1.txtItemLvl3.Value    = arrRet(11)
          frm1.txtItemLvl3Nm.Value  = arrRet(12)
          frm1.txtSerialNo.Value    = arrRet(13)
          frm1.cboItemVer.Value     = arrRet(14)
          frm1.txtItemNm2.Value     = arrRet(16)          
          frm1.txtSpec2.Value       = arrRet(17)    
          frm1.txtItemUnit.Value    = arrRet(18)
          frm1.cboPurType.Value     = arrRet(19)
          frm1.txtPurGroup.Value    = arrRet(23)
          frm1.txtPurGroupNm.Value  = arrRet(24)
          frm1.txtPurVendor.Value   = arrRet(25)
          frm1.txtPurVendorNm.Value = arrRet(26)
          If arrRet(27) = "Y"  Then
             frm1.rdoUnifyPurFlg1.Checked = True
             frm1.rdoUnifyPurFlg2.Checked = False
          Else
             frm1.rdoUnifyPurFlg2.Checked = True
             frm1.rdoUnifyPurFlg1.Checked = False
          End If   
          frm1.txtNetWeight.Value= arrRet(28)
          frm1.txtNetWeightUnit.Value= arrRet(29)
          frm1.txtGrossWeight.Value= arrRet(30)
          frm1.txtGrossWeightUnit.Value= arrRet(31)
          frm1.txtCBM.Value= arrRet(32)
          frm1.txtCBMInfo.Value= arrRet(33)
          frm1.txtHSCd.Value= arrRet(34)
          frm1.txtHSNm.Value= arrRet(35)

          Set gActiveElement = document.activeElement 
          
          Call SetColumn(arrRet(3) , arrRet(5))
          
          lgBlnFlgChgValue = True
                          
       End If
              
       Call SetFocusToDocument("M")
       frm1.txtBasicItem.focus
       
End Function

'======================================================================================================
' 표준규격 PopUp
'======================================================================================================
Function OpenCategory()
       Dim arrRet
       Dim arrParam(10), arrField(6)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Or UCase(frm1.txtSpec.className) = UCase(parent.UCN_PROTECTED) Then
           Exit Function
        End If   

       IsOpenPop = True
       
       arrParam(0) = Trim(frm1.cboItemAcct.value)
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
       	  Call DisplayMsgBox("971012", "X", frm1.cboItemAcct.Alt, "X")
          frm1.cboItemAcct.focus
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
          frm1.txtSpec.Value = arrRet(0)          
          Set gActiveElement = document.activeElement 
          lgBlnFlgChgValue = True
       End If
       
       Call SetFocusToDocument("M")
       frm1.txtSpec.focus

End Function


'======================================================================================================
' 재의뢰버튼 실행시...
'======================================================================================================
Function RunReReq()
    Dim IntRetCD
    Dim strReqNo , strSerialNo , strItemCd

    Err.Clear                                                                    '☜: Clear err status
    
    If frm1.htxtStatus.Value <> "D" Then
       Exit Function
    End If
    
    //If lgBlnFlgChgValue = False Then 
    //    IntRetCD =  DisplayMsgBox("900001","x","x","x")                          '☜:There is no changed data. 
    //    Exit Function
    //End If
   
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                    '☜: Please do Display first. 
       Call  DisplayMsgBox("900002","x","x","x")                                
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
    
    Call DisableToolBar(parent.TBC_SAVE)
    
    frm1.htxtStatus.Value = "R" 
       
    If DbSave = False Then
    	'실패하면 복구한다.
        frm1.htxtStatus.Value = "D"  
        Call RestoreToolBar()        
        Exit Function
    End If
     
End Function

'======================================================================================================
' 도면파일관리 첨부및 변경...
'======================================================================================================
Function OpenDocFile()
       Dim arrRet
       Dim arrParam(10), arrField(6)
       Dim iCalledAspName, IntRetCD

       If IsOpenPop = True Or UCase(frm1.button1.className) = UCase(parent.UCN_PROTECTED) Then
           Exit Function
       End If   

       IsOpenPop = True
       
       arrParam(0) = Trim(frm1.htxtInternalCd.value)
       arrParam(1) = Trim(frm1.txtarItemCd.value)
       arrParam(2) = Trim(frm1.txtarItemNm.value)
       arrParam(3) = Trim(frm1.txtarReqNo.value)
       
       
       'R:의뢰 A : 접수 D : 반려 E : 완료 S : 중단 T : 이관 
       Select Case Trim(frm1.htxtStatus.Value)
              Case "R","A","D"
                  arrParam(9) = "0" 
              Case Else
                  '첨부파일 PopUp에서 수정 못하게 한다.
                  arrParam(9) = "1" 
       End Select
       
       iCalledAspName = AskPRAspName("B82101PA4")
       
       If Trim(iCalledAspName) = "" Then
          IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B82101PA4", "X")
          IsOpenPop = False
          Exit Function
       End If
       
       arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
               "dialogWidth=350px; dialogHeight=200px; center: Yes; help: No; resizable: No; status: No;")
   
       IsOpenPop = False
    
		if not isArray(arrRet) then

		  If arrRet= True Then
				call dbsaveOK()
				MyBizASP.location.reload	
			else 
				MyBizASP.location.reload	
				exit function						
			End If

		   Call SetFocusToDocument("M")
		    lgBlnFlgChgValue=true
		 end if
		
      	       
End Function