<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 


    Dim FSO, TextStream
    Dim arrMenu
    Dim strFileName
    Dim ADOConn
    Dim ADORs
    Dim StrSql
    Dim UsrMenu
    
	On Error Resume Next    

    Set FSO = Server.CreateObject("Scripting.FileSystemObject")

	call GetGlobalVar
	
    strFileName = Request.ServerVariables("PATH_INFO")
    MyPos       = InStrRev(strFileName, "/", -1, 1)
    strFileName = Left(strFileName, MyPos - 1)
    MyPos       = InStrRev(strFileName, "/", -1, 1)
    strFileName = Left(strFileName, MyPos)  & "Menu"
    strFileName = Server.MapPath(strFileName)

    'strFileName = strFileName & "\" & Request.Cookies("unierp")("gCompany") & "_" & Request.Cookies("unierp")("gUsrId") & ".Dat"

'	strFileName = strFileName & "\" & gCompany + "_" & gUsrId & ".Dat"
	strFileName = strFileName & "\" & gCompany + "_" & gDBServer + "_" + gDatabase + "_" & gUsrId & ".Dat"
    
    
    Set TextStream = FSO.OpenTextFile(strFileName, 1)									'OpenFileForReading(1) 유저별 전역별수로 

    arrMenu = TextStream.ReadAll

    TextStream.Close
    Set TextStream = Nothing
    Set FSO        = Nothing


    Call SubOpenDB(ADOConn)                                                        '☜: Make  a DB Connection
    
    strSql =		  "SELECT UPPER_MNU_ID, MNU_ID, MNU_NM, MNU_TYPE "
	strSql = strSql & "FROM Z_USR_MNU "
    strSql = strSql & "WHERE USR_ID = '" & gUsrID & "' "
    strSql = strSql & "AND LANG_CD  = '" & gLang & "' "
	strSql = strSql & "ORDER BY MNU_TYPE,SYS_LVL ASC, MNU_SEQ ASC"    
	
    
    If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
        UsrMenu =  ""
    Else
	    While Not ADORs.EOF        
           UsrMenu = UsrMenu & ADORs("UPPER_MNU_ID") & chr(11) & ADORs("MNU_ID") & chr(11) & ADORs("MNU_NM") & chr(11) & ADORs("MNU_TYPE") & chr(11) & Chr(12)
           ADORs.MoveNext
	    WEnd        
    End If

    Call SubCloseRs(ADORs)                                                          '☜: Release RecordSSet
    Call SubCloseDB(ADOConn)                                                       '☜: Colse a DB Connection

%>
<Script Language=VBScript src="../inc/incUni2KTV.vbs"></Script>
<Script Language=VBScript>
    Const C_SEP       = "::"
    Const C_MNU_ID    = 0
    Const C_MNU_UPPER = 1
    Const C_MNU_LVL   = 2
    Const C_MNU_TYPE  = 3
    Const C_MNU_NM    = 4
    Const C_MNU_AUTH  = 5

    Const C_UNDERBAR  = "_"
    
	Dim NodX

    Dim UsrMenu, arrMenu, i, arrLine
    Dim strImg, strImgC, strImgO
	
	On Error Resume Next
	
	arrMenu = "<%=replace(arrMenu,vbCrLf,chr(12))%>"

	If Trim(arrMenu) = "" Then
       MsgBox "메뉴 데이타가 없습니다.", vbCritical, gLogoName
    Else
       Call MakeMenuTreeView()
	End If
	
	Call MakeUserMenuTreeView()
	

'=======================================================================================================================
' Name : 
' Desc : 
'=======================================================================================================================
Function GetImageC(arrLine)
	Dim strImg

	Select Case arrLine(C_MNU_AUTH)
		Case "A", "Q", "E"
			If arrLine(C_MNU_TYPE) = "M" Then
                Select Case Left(arrLine(C_MNU_ID),1)
                    Case "A"
                        strImg = C_AC
                    Case "B"
                        strImg = C_BC
                    Case "C"
                        strImg = C_CC
                    Case "D"
                        strImg = C_DC
                    Case "G"
                        strImg = C_GC                                                                            
                    Case "H"
                        strImg = C_HC
                    Case "I"
                        strImg = C_IC
                    Case "J"
                        strImg = C_JC
                    Case "M"
                        strImg = C_MC
                    Case "O"
                        strImg = C_OC
                    Case "P"
                        strImg = C_PC
                    Case "Q"
                        strImg = C_QC
                    Case "R"
                        strImg = C_RC
                    Case "S"
                        strImg = C_SC
                    Case "U"
                        strImg = C_UC                    
                    Case "Z"
                        strImg = C_ZC
					Case Else
						strImg = C_USFolder
                End Select
			Else
				strImg = C_USURL
			End If
		Case "I"
			strImg = C_USConst
		Case "N"
			strImg = C_USNone
	End Select

	GetImageC = strImg
End Function

'=======================================================================================================================
' Name : 
' Desc : 
'=======================================================================================================================
Function GetImageO(arrLine)
	Dim strImg

	Select Case arrLine(C_MNU_AUTH)
		Case "A", "Q", "E"
			If arrLine(C_MNU_TYPE) = "M" Then
                Select Case Left(arrLine(C_MNU_ID),1)
                    Case "A"
                        strImg = C_AO
                    Case "B"
                        strImg = C_BO
                    Case "C"
                        strImg = C_CO
                    Case "D"
                        strImg = C_DO
                    Case "G"
                        strImg = C_GO                                                                              
                    Case "H"
                        strImg = C_HO
                    Case "I"
                        strImg = C_IO
                    Case "J"
                        strImg = C_JO
                    Case "M"
                        strImg = C_MO
                    Case "O"
                        strImg = C_OO
                    Case "P"
                        strImg = C_PO
                    Case "Q"
                        strImg = C_QO
                    Case "R"
                        strImg = C_RO
                    Case "S"
                        strImg = C_SO
                    Case "U"
                        strImg = C_UO                    
                    Case "Z"
                        strImg = C_ZO
					Case Else
						strImg = C_USFolder
                End Select
			Else
				strImg = C_USURL
			End If
		Case "I"
			strImg = C_USConst
		Case "N"
			strImg = C_USNone
	End Select

	GetImageO = strImg
End Function

'=======================================================================================================================
' Name : 
' Desc : 
'=======================================================================================================================
Sub Document_onReadyStateChange()
	parent.frm2.uniTree1.MousePointer = 0
End Sub
'=======================================================================================================================
' Name : MakeMenuTreeView
' Desc : Make menu tree
'=======================================================================================================================
Sub MakeMenuTreeView()

	On Error Resume Next

    arrMenu = Split(arrMenu, chr(12))
	
	With parent.frm2
       For i = 0 To UBound(arrMenu, 1)
           If arrMenu(i) = "" Then 
    	      Exit For
	       End If   

           arrLine = Split(arrMenu(i), C_SEP)

           If Left(arrLine(C_MNU_ID), 1) <> "Z" Then
              strImgC = GetImageC(arrLine)'Folder
              strImgO = GetImageO(arrLine)                      
              If arrLine(C_MNU_UPPER) = "*" Then
                 If strImgC = C_USURL Then
                    Set NodX = .uniTree1.Nodes.Add(, tvwChild, arrLine(C_MNU_ID) & Chr(20) & arrLine(C_MNU_UPPER), arrLine(C_MNU_NM), strImgC, strImgC)
                    NodX.ExpandedImage = strImgO
                 Else
                    Set NodX = .uniTree1.Nodes.Add(, tvwChild, arrLine(C_MNU_ID), arrLine(C_MNU_NM), strImgC, strImgC)
                    NodX.ExpandedImage = strImgO                    
                 End If
              Else
                 If strImgC = C_USURL Then
                    Set NodX = .uniTree1.Nodes.Add(arrLine(C_MNU_UPPER), tvwChild, arrLine(C_MNU_ID) & Chr(20) & arrLine(C_MNU_UPPER), arrLine(C_MNU_NM), C_USURL, C_USURL)
                 Else		
					if arrLine(C_MNU_Type) ="M" then 
						if strImgC = C_USNone then 
							Set NodX = .uniTree1.Nodes.Add(arrLine(C_MNU_UPPER), tvwChild, arrLine(C_MNU_ID), arrLine(C_MNU_NM), strImgC, strImgC)							
							NodX.ExpandedImage = strImgO 
							'NodX.ExpandedImage = C_USOpen 
						else
							Set NodX = .uniTree1.Nodes.Add(arrLine(C_MNU_UPPER), tvwChild, arrLine(C_MNU_ID), arrLine(C_MNU_NM), C_USFolder, C_USFolder)
							NodX.ExpandedImage = C_USOpen
						end if 
					else
						if strImgC = C_USNone then 
							Set NodX = .uniTree1.Nodes.Add(arrLine(C_MNU_UPPER), tvwChild, arrLine(C_MNU_ID), arrLine(C_MNU_NM), C_USNone, C_USNone)							
						else
							Set NodX = .uniTree1.Nodes.Add(arrLine(C_MNU_UPPER), tvwChild, arrLine(C_MNU_ID), arrLine(C_MNU_NM), C_USURL, C_USURL)							
						end if 
					end if 
                 
                 End If
	          End If
		
              'If strImgC = C_USConst Or strImgC = C_USNone Then 				
              '  NodX.ForeColor = &H808080
              'End If	
		
              If Not (NodX.parent Is Nothing) Then              
                 If NodX.Parent.Image = C_USConst Or NodX.Parent.Image = C_USNone Then
                    NodX.Image         = NodX.Parent.Image
                    NodX.ExpandedImage = NodX.Parent.Image
                    NodX.SelectedImage = NodX.Parent.SelectedImage
                    'NodX.ForeColor     = &H808080
                 End If		
              End If			
		      Set NodX = Nothing		
           End If	
           If Err.number > 0 Then
              MsgBox "메뉴 구성정보가 올바르지 않습니다." & vbCrLf & "메뉴 구성정보을 다시 확인하세요."_
                      & vbCrLf & "상위메뉴:"      & arrLine(C_MNU_UPPER) _
                      & vbCrLf & "점검대상메뉴:"  & arrLine(C_MNU_ID) , vbCritical, gLogoName
              Exit For
           End If
       Next
	End With

End Sub
'=======================================================================================================================
' Name : MakeUserMenuTreeView
' Desc : Make user menu tree
'=======================================================================================================================
Sub MakeUserMenuTreeView()
    Dim UsrMenuCol
    Dim UsrMenuRow
    Dim iDx
	Dim ndNode
	
    On Error Resume Next
    
    UsrMenu =  "<%=UsrMenu%>"
    
    If Trim(UsrMenu) = "" Then
       Exit Sub
    End If
    
    UsrMenuRow = Split(UsrMenu,Chr(12))
    For iDx = 0 To UBound(UsrMenuRow) - 1 
    
        UsrMenuCol = Split(UsrMenuRow(iDx),Chr(11))
        
		If UsrMenuCol(3) = "P" Then
			strImg = C_USURL
			strKey = UsrMenuCol(1) & Chr(20) & UsrMenuCol(0)			
		Else
			strImg = C_USFolder
			strKey = UsrMenuCol(1)
		End If
				
		
		Set ndNode = parent.frm2.uniTree1.Nodes.Add(UsrMenuCol(0), tvwChild,strKey, UsrMenuCol(2), strImg)
		
		ndNode.Tag = UsrMenuCol(3)
        
    Next
    
    
End Sub

</Script>