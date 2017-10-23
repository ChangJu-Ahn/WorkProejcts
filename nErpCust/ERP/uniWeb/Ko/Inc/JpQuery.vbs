'=======================================================================================
Function CommonQueryRs2(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    On Error Resume Next
    
    CommonQueryRs2 = False
    
    lgF0 = ""
    lgF1 = ""
    lgF2 = ""
    lgF3 = ""
    lgF4 = ""
    lgF5 = ""
    lgF6 = ""

    If gRdsUse = "T" Then
       CommonQueryRs2 = RDSQueryMain2(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    Else
       CommonQueryRs2 = HTTPQueryMain2(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    End If
    
End Function

'=======================================================================================
Function HTTPQueryMain2(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    Dim iOutData
    Dim arrRow, arrCol
    Dim ii
    Dim iiMax, jjMax
    
    Dim Tmp(6)

    On Error Resume Next
    
    HTTPQueryMain2 = False
    
    If HTTPQuery2(SelectList, FromList, WhereList, iOutData) = False Then
       Exit Function
    End If
    
    If IsEmpty(iOutData) Then
       Exit Function
    End If
    
    If Trim(iOutData) = "" Then
       Exit Function
    End If
    
    arrRow = Split(iOutData, Chr(12))
    For ii = 0 To UBound(arrRow) - 1
        arrCol = Split(arrRow(ii), Chr(11))
        lgF0 = lgF0 & arrCol(0) & Chr(11)
        If UBound(arrCol) > 0 Then
           lgF1 = lgF1 & arrCol(1) & Chr(11)
           If UBound(arrCol) > 1 Then
              lgF2 = lgF2 & arrCol(2) & Chr(11)
              If UBound(arrCol) > 2 Then
                 lgF3 = lgF3 & arrCol(3) & Chr(11)
                 If UBound(arrCol) > 3 Then
                    lgF4 = lgF4 & arrCol(4) & Chr(11)
                    If UBound(arrCol) > 4 Then
                       lgF5 = lgF5 & arrCol(5) & Chr(11)
                       If UBound(arrCol) > 5 Then
                          lgF6 = lgF6 & arrCol(6) & Chr(11)
                       End If
                    End If
                 End If
              End If
           End If
       End If
    Next
    
    HTTPQueryMain2 = True

End Function

'=======================================================================================
Function RDSQueryMain2(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    Dim rs0
    On Error Resume Next
    
    RDSQueryMain2 = False

    If RDSQuery2(SelectList, FromList, WhereList, rs0) = False Then
       Exit Function
    End If
    
    If (IsNull(rs0)) Or (rs0 Is Nothing) Or (rs0.EOF And rs0.BOF) Then
       rs0.Close
       Set rs0 = Nothing
       Exit Function
    End If
    
    While Not rs0.EOF
          If rs0.Fields.Count > 0 Then
             lgF0 = lgF0 & rs0(0) & Chr(11)
             If rs0.Fields.Count > 1 Then
                lgF1 = lgF1 & rs0(1) & Chr(11)
                If rs0.Fields.Count > 2 Then
                   lgF2 = lgF2 & rs0(2) & Chr(11)
                   If rs0.Fields.Count > 3 Then
                      lgF3 = lgF3 & rs0(3) & Chr(11)
                      If rs0.Fields.Count > 4 Then
                         lgF4 = lgF4 & rs0(4) & Chr(11)
                         If rs0.Fields.Count > 5 Then
                            lgF5 = lgF5 & rs0(5) & Chr(11)
                            If rs0.Fields.Count > 6 Then
                               lgF6 = lgF6 & rs0(6) & Chr(11)
                            End If  ' 6
                         End If  ' 5
                      End If  ' 4
                   End If  ' 3
                End If  ' 2
             End If  ' 1
          End If  ' 0
          rs0.MoveNext
    Wend

    rs0.Close
    Set rs0 = Nothing
    
    RDSQueryMain2 = True

End Function

'=======================================================================================
Function HTTPQuery2(ByVal SelectList, ByVal FromList, ByVal WhereList, prData)
    Dim iStrSQL
    Dim iXmlHttp
    Dim iRetByte 
    
    On Error Resume Next
    Err.Clear
    
    HTTPQuery = False

    iStrSQL = "Select " & SelectList
    
    If Trim(FromList) > "" Then
       iStrSQL = iStrSQL & " From  " & FromList
       If Trim(WhereList) > "" Then
          iStrSQL = iStrSQL & " Where  " & WhereList
       End If
       
    End If

    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP.4.0")
    
    iXmlHttp.open "POST", GetComaspFolderPath & "RequestCommonQry.asp", False
    iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    iStrSQL = Escape(iStrSQL)
    iStrSQL = Replace(iStrSQL, "+", "%2B")
    iStrSQL = Replace(iStrSQL, "/", "%2F")

    iXmlHttp.send "LangCD=" & gLang & "&ADODBConnString=" & Escape(gADODBConnString) & "&StrSQL=" & iStrSQL


    If gCharSet = "D" Then
       prData   = ConnectorControl.CStrConv(iXmlHttp.responseBody)
    Else
       prData   =                  iXmlHttp.responseText
    End If   
  
    Set iXmlHttp = Nothing
    If prData <> "" Then
        HTTPQuery2 = True
    End If
End Function

'=======================================================================================
Function RDSQuery2(SelectList, FromList, WhereList, rs0)
    Dim ADF                                                                    '☜ : declaration Variable indicating ActiveX Data Factory
    Dim lgStrSQL
    Dim strRetMsg                                                              '☜ : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag                                   '☜ : declaration DBAgent Parameter

    On Error Resume Next

    Err.Clear
    
    ReDim UNISqlId(0)
    ReDim UNIValue(0, 0)
    
    RDSQuery2 = False
    
    lgStrSQL = "Select " & SelectList
    
    If Trim(FromList) > "" Then
       lgStrSQL = lgStrSQL & " From  " & FromList
       
       If Trim(WhereList) > "" Then
          lgStrSQL = lgStrSQL & " Where  " & WhereList
       End If
       
    End If

    UNISqlId(0) = "commonqry"
    UNIValue(0, 0) = lgStrSQL
    UNILock = DISCONNREAD: UNIFlag = "1"

    If Trim(gDsnNo) = "" Then
       Exit Function
    End If

    If Trim(gServerIP) = "" Then
       Exit Function
    End If

    Set ADF = ADS.CreateObject("prjPublic.cCtlTake", gServerIP)
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    If Err.Number <> 0 Then
       Set ADF = Nothing
       Exit Function
    End If

    RDSQuery2 = True

    Set ADF = Nothing

End Function
	  
Sub Jump_Pgm(ByVal pvPgmid, ByVal pvFB_fg,ByVal pvNext,ByVal pvValue)
	Dim iNextInfo
	Dim iNextArr
	Dim iNextPgmid
	Dim iActMethod,StrNVar,StrNPgm
	Dim NPgmId,NPgmVar,NPgminval,StrCk,lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6
	
	' 다음화면 아이디 추출 
	Call Rtn_NextPgmId(pvPgmid,pvFB_fg,pvValue,NPgmId)

	if NPgmId <> "" then
	
		iNextArr = Split(NPgmId,Chr(11))
	
		If  iNextArr(0) <> "" Then	
			iNextPgmid = iNextArr(0)
			iActMethod = iNextArr(1)

			'*******************
			'	권한체크 
			'*******************
			
			StrCk =  CommonQueryRs2("mnu_id", "z_usr_role_mnu_authztn_asso a inner join z_usr_mast_rec_usr_role_asso b on a.usr_role_id=b.usr_role_id", "b.usr_id= '" & Parent.gUsrId & "' and a.mnu_id= '" & iNextPgmid & "' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
				

			If 	StrCk = False Then
				Call DisplayMsgBox("990016", "X", "X","X") 'msgbox "이동할 화면에 대해 권한이 없습니다. "
				Exit Sub	
			End If
		        
			'	다음화면에서 조회에 필요한 키값 추출 
			Call  Rtn_NexKeyValue(pvPgmid,iNextPgmid,pvValue,pvNext,NPgmVar,NPgminval)

			'***********************************************************************
			'	다음화면 아이디와 키값으로 화면이동 및 조회실행--- iActMethod 고려 
			'***********************************************************************

			StrNPgm = iNextPgmid
			StrNVar = NPgmVar 
		Else
			Call DisplayMsgBox("801005", "X", "X","X")
			Exit Sub
		End If
		If StrNPgm <> "" And StrNVar <> "" Then
			Call Jump_MvPgm (StrNPgm,StrNVar,StrCookie,NPgminval)
		End If
	Else
	Call DisplayMsgBox("801003", "X", "X","X")
	End if	
End Sub

Function Rtn_NexKeyValue(ByVal pvPgmid,ByVal iNextPgmid,ByVal pvValue,ByVal pvNext,ByRef NPgmVar,ByRef NPgminval)
	Dim iStrSql
	Dim iAct_Meth
	Dim iNext_Value,StrSel,StrForm,StrNeed
	Dim iCPgmid,StrCk,lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6
	

	
	'***************************************
	'	화면아이디 존재여부 검증 
	'***************************************
	
	StrCk =  CommonQueryRs2(" KEY_OBJECT_VALUE ", "Z_MOVE_PGM_KEY_OBJECT", "PGM_ID = '" & iNextPgmid & "'" , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
	
	If StrCk = False Then
		Call DisplayMsgBox("801003", "X", "X","X")
		Exit Function
	Else
		NPgminval = lgf0 

		StrCk =  CommonQueryRs2(" MOVE_ITEM_QRY ", "Z_MOVE_PGM_ITEM_QRY", "PGM_ID = '" & iNextPgmid & "' and MOVE_ITEM_CD = '" & pvPgmid & "' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
	
		If 	StrCk = False Then
		
		NPgmVar = pvValue & chr(11) & pvNext
		
	    Else
		
					iStrSql = replace(lgf0,chr(11),"")
					If 	iStrSql <> "" Then
						'********************************
						'	키값 추출 쿼리문 세팅 
						'********************************
						istrsql = Replace(iStrSql,"?", "'" & pvValue & "'" )
						StrCk = ""
						lgf0 = ""
						StrSel = "distinct *"
						StrForm = ""
						StrForm = StrForm & "("
						StrForm = StrForm & istrsql
						StrForm = StrForm & ")A"
											
						StrCk =  CommonQueryRs2(StrSel, StrForm , "" , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)

						If 	StrCk = False Then
							Call DisplayMsgBox("801003", "X", "X","X")
							NPgmVar = ""
							Exit Function    		
						Else
							   if lgf0 <> "" then
									iNext_Value = lgf0&lgf1&lgf2&lgf3&lgf4&lgf5&lgf6 'replace(lgf0,chr(11),"")
									If 	iNext_Value = "" Then
										Call DisplayMsgBox("801003", "X", "X","X")
										Exit Function
									End If
									NPgmVar = iNext_Value
								End If
						End If	
					Else
						NPgmVar = pvValue & chr(11)
					End If

				End If	
		End If
	
End Function

Function Rtn_NextPgmId (ByVal pvPgmid, ByVal pvFB_fg,ByVal pvValue,Byref NPgmId) 
	Dim iStrSql
	Dim iAct_Meth
	Dim iChk_Value,j,i,StrForm,StrSel,StrWhere
	Dim iNPgmid,StrCk,iNextArr,iNPgminval,StrCnt
	Dim rs0,lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6,StrValue

    On Error Resume Next

    Err.Clear

	'**********************************************************
	'	이동할 다음화면의 존재여부 체크 및 화면 아이디 추출 
	'**********************************************************
	
	StrCk =  CommonQueryRs2("top 1 DATA_COLM_ID,TBL_ID,KEY_COLM_ID1", "Z_MOVE_ITEM", "MOVE_ITEM_CD = '" & pvPgmid & "' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)  'AND MOVE_ITEM_NM = '"& pvColm_id &"' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
	
    If 	StrCk = False Then
		Call DisplayMsgBox("801003", "X", "X","X")
		Exit Function    
    Else
		if lgf0 <> "" then
			
			StrWhere = Replace(lgf2,chr(11),"") & "=" & "'" & pvValue & "'"
			
			StrCk =  CommonQueryRs2(Replace(lgf0,chr(11),""), Replace(lgf1,chr(11),""), StrWhere , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
			
			If 	StrCk = False Then
				Call DisplayMsgBox("801004", "X", "X","X")
				NPgmId = "" & chr(11) & ""
				Exit Function    		
			Else
				
				if Replace(lgf0,chr(11),"") <> pvValue Then
				StrVal = Replace(lgf0,chr(11),"")
				Else
				StrVal = ""
				End if
				
				CALL CommonQueryRs2("COUNT(NEXT_PGM_ID)", "Z_MOVE_PGM_INF", " MOVE_ITEM_CD = '" & pvPgmid & "' and TYPE_VALUE = '" & StrVal &"' and MOVE_DIRECTION = '" & pvFB_fg & "' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
				StrCnt = Replace(lgf0,chr(11),"") 
				
				IF StrCnt > 1 THEN
				
				CALL OPENPGM(pvPgmid,StrVal,pvFB_fg,NPgmId)
				
				Else
				
					StrCk =  CommonQueryRs2("NEXT_PGM_ID", "Z_MOVE_PGM_INF", " MOVE_ITEM_CD = '" & pvPgmid & "' and TYPE_VALUE = '" & StrVal &"' and MOVE_DIRECTION = '" & pvFB_fg & "' " , lgf0,lgf1,lgf2,lgf3,lgf4,lgf5,lgf6)
				
					if StrCk = False Then
						Call DisplayMsgBox("801003", "X", "X","X")
						NPgmId = "" & chr(11) & ""
						Exit Function
					Else
						NPgmId = lgf0
					End if
				End If
			End if
		End IF	
	End if	
	
End Function


Function Jump_MvPgm(ByVal StrNPgm,ByVal StrNVar ,Byref StrCookie,ByVal NPgminval)
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim StrUrl
	Const CookieSplit	= 4877 
		'****************************
		'	Cookie에 넘겨줄 형식 맞춤	
		'****************************	

		Call CommonQueryRs2(" top 1 a.mnu_id, a.upper_mnu_id,  b.mnu_nm, c.called_frm_id "," z_auth_gen a, z_lang_co_mast_mnu b , z_co_mast_mnu c","a.mnu_id   = b.mnu_id and a.mnu_id   = c.mnu_id and  a.mnu_id   = " & FilterVar(Trim(StrNPgm) ,"''","S") & " and b.lang_cd  = " & FilterVar(Trim(parent.gLang) ,"''","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		StrUrl = "Module/" + Trim(Replace(lgF1,Chr(11),"")) + "/" + UCase(StrNPgm)  + ".asp?strRequestMenuID="& UCase(StrNPgm) & "&strRequestUpperMenuID=" & Trim(Replace(lgF1,Chr(11),"")) & "&strASPMnuMnuNm="& Trim(Replace(lgF2,Chr(11),""))

		StrCookie = StrUrl & chr(14) & StrNVar & chr(14) & NPgminval & chr(14) & StrNPgm

		
		Call WriteCookie2 (CookieSplit , StrCookie)
		Set objConn = CreateObject("uniConnector.cGlobal")                
			PostString = objConn.GetAspPostString
			window.open "../../SessionTrans2.asp?"& PostString 
		
		
End Function

Function WriteCookie2(varCookie, varValue)	
        varValue = Replace(varValue,";","<:::>") 
	Document.Cookie = varCookie & "=" & Escape(varValue) & "; path=" & "/"
End Function


Function OpenPGM(ByVal pvPgmid,ByVal StrVal,ByVal pvFB_fg,ByRef NPgmId)

	Dim arrRet
	Dim arrParam(7), arrField(8), arrHeader(8)

	
	        arrParam(0) = "NEXT_PGM_ID"
			arrParam(1) = "Z_MOVE_PGM_INF A,Z_LANG_CO_MAST_MNU B"
			arrParam(2) = ""
			arrParam(3) = ""
			arrParam(4) = "NEXT_PGM_ID = MNU_ID AND LANG_CD = '"& parent.gLang &"' AND MOVE_ITEM_CD = '"& pvPgmid &"' AND TYPE_VALUE = '"& StrVal &"' AND MOVE_DIRECTION = '"& pvFB_fg &"' "
			arrParam(5) = "NEXT_PGM_ID"

			arrField(0) = "ED08" & Chr(11) & "A.NEXT_PGM_ID"
			arrField(1) = "ED18" & Chr(11) & "B.MNU_NM"
			
			arrHeader(0) = "NEXT_PGM_ID"
			arrHeader(1) = "NEXT_PGM_ID명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	
	If arrRet(0) = "" Then
		Exit Function
	Else
		NPgmId = arrRet(0) & chr(11)
	End If	
	
End Function

