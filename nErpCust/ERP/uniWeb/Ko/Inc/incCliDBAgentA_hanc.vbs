Const mDISCONNREAD = "2"
Const mgAllowDragDropSpread = "T"
Dim mgColSep, mgRowSep
mgColSep = Chr(11)
mgRowSep = Chr(12)

Function mIsBetween(ByVal iFrom,ByVal iTo,ByVal iIt)
    mIsBetween =  False
    If iIt >= iFrom And iIt <= iTo Then
       mIsBetween = True
    End If
End Function

'==============================================================================
' Name  :  GetAdoFiledInf
' Desc  :  Get Information for Creation of Query Window using ADO
'==============================================================================
Sub GetZAdoFieldInf(ByVal iPgmId,ByVal iTypeCD,ByVal iSpdNo,ByVal iVersion,iSpread)
MsgBox "GetZAdoFieldInf "
    Dim lgstrRetMsg                                '☜ : declaration Variable indicating Record Set Return Message
    Dim iResultData, iStrSQL, iRowData, iColData, i
    Dim ADF
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
    
    On Error Resume Next
    
    Select Case  iSpdNo
	     Case "A"
                gTypeCD     = ""
                gFieldCD    = ""
                gFieldNM    = ""
                gFieldLen   = ""
                gFieldType  = ""
                gDefaultT   = ""
                gNextSeq    = ""
                gKeyTag     = ""
                gHidden     = ""
                gSortDirection = ""                
	     Case "B"
                gTypeCD1    = ""
                gFieldCD1   = ""
                gFieldNM1   = ""
                gFieldLen1  = ""
                gFieldType1 = ""
                gDefaultT1  = ""
                gNextSeq1   = ""
                gKeyTag1    = ""
                gHidden1    = ""
                gSortDirection1 = ""
	     Case "C"
                gTypeCD2    = ""
                gFieldCD2   = ""
                gFieldNM2   = ""
                gFieldLen2  = ""
                gFieldType2 = ""
                gDefaultT2  = ""
                gNextSeq2   = ""
                gKeyTag2    = ""
                gHidden2    = ""
                gSortDirection2 = ""
	     Case "D"
                gTypeCD3    = ""
                gFieldCD3   = ""
                gFieldNM3   = ""
                gFieldLen3  = ""
                gFieldType3 = ""
                gDefaultT3  = ""
                gNextSeq3   = ""
                gKeyTag3    = ""
                gHidden3    = ""
                gSortDirection3 = ""                
    End Select   

    Err.Clear

    ggoSpread.Source = iSpread
    If ggoSpread.VersionCheck(iVersion) Then
MsgBox "ver A"
       Select Case  iSpdNo
	        Case "A"
	           Call ggoSpread.GetDBAgentData(gTypeCD,gFieldCD,gFieldNM,gFieldLen,gFieldType,gDefaultT,gNextSeq,gKeyTag,gSortDirection,gHidden)
	        Case "B"
	           Call ggoSpread.GetDBAgentData(gTypeCD1,gFieldCD1,gFieldNM1,gFieldLen1,gFieldType1,gDefaultT1,gNextSeq1,gKeyTag1,gSortDirection1,gHidden1)
            Case "C"
	           Call ggoSpread.GetDBAgentData(gTypeCD2,gFieldCD2,gFieldNM2,gFieldLen2,gFieldType2,gDefaultT2,gNextSeq2,gKeyTag2,gSortDirection2,gHidden2)
            Case "D"
	           Call ggoSpread.GetDBAgentData(gTypeCD3,gFieldCD3,gFieldNM3,gFieldLen3,gFieldType3,gDefaultT3,gNextSeq3,gKeyTag3,gSortDirection3,gHidden3)
       End Select   
    Else   
MsgBox "ver B : " & iPgmId

        If gRdsUse = "T" Then
            Redim UNISqlId(0)
            Redim UNIValue(0, 3)
            UNISqlId(0)    = "z100001"
            UNIValue(0, 0) = iPgmId
            UNIValue(0, 1) = iTypeCD
            UNIValue(0, 2) = iSpdNo
            UNIValue(0, 3) = gLang

            UNILock = mDISCONNREAD   :	UNIFlag = "1"
            If Trim(gDsnNo) ="" Then
                Call SetADOFeild(iSpdNo,iTypeCD)
                Call SplitADOFieldVar(iSpdNo)
                Exit Sub
            End If
            Set ADF = ADS.CreateObject("prjPublic.cCtlTake",gServerIP )
            lgstrRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)	

            Set ADF = Nothing			

            If rs0.EOF And rs0.BOF Then
                Call SetADOFeild(iSpdNo,iTypeCD)
                Call SplitADOFieldVar(iSpdNo)
                rs0.Close
                Set rs0 = Nothing
                Exit Sub
            End If
	
            Do While Not (rs0.EOF Or rs0.BOF)
                Select Case  iSpdNo
                    Case "A"
                        gTypeCD    = gTypeCD    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                        gFieldCD   = gFieldCD   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                        gFieldNM   = gFieldNM   &       rs0("FIELD_NM")    & Chr(11)
                        gFieldLen  = gFieldLen  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                        gFieldType = gFieldType & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                        gDefaultT  = gDefaultT  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                        gNextSeq   = gNextSeq   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                        gKeyTag    = gKeyTag    & UCASE(rs0("KEY_TAG"))    & Chr(11)
                        gHidden    = gHidden    & "0"                      & Chr(11)
                        gSortDirection = gSortDirection & "ASC"            & Chr(11)		              
                    Case "B"
                        gTypeCD1    = gTypeCD1    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                        gFieldCD1   = gFieldCD1   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                        gFieldNM1   = gFieldNM1   &       rs0("FIELD_NM")    & Chr(11)
                        gFieldLen1  = gFieldLen1  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                        gFieldType1 = gFieldType1 & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                        gDefaultT1  = gDefaultT1  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                        gNextSeq1   = gNextSeq1   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                        gKeyTag1    = gKeyTag1    & UCASE(rs0("KEY_TAG"))    & Chr(11)
                        gHidden1    = gHidden1    & "0"                      & Chr(11)
                        gSortDirection1 = gSortDirection1 & "ASC"            & Chr(11)
                    Case "C"
                        gTypeCD2    = gTypeCD2    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                        gFieldCD2   = gFieldCD2   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                        gFieldNM2   = gFieldNM2   &       rs0("FIELD_NM")    & Chr(11)
                        gFieldLen2  = gFieldLen2  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                        gFieldType2 = gFieldType2 & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                        gDefaultT2  = gDefaultT2  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                        gNextSeq2   = gNextSeq2   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                        gKeyTag2    = gKeyTag2    & UCASE(rs0("KEY_TAG"))    & Chr(11)
                        gHidden2    = gHidden2    & "0"                      & Chr(11)
                        gSortDirection2 = gSortDirection2 & "ASC"            & Chr(11)
                    Case "D"
                        gTypeCD3    = gTypeCD3    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                        gFieldCD3   = gFieldCD3   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                        gFieldNM3   = gFieldNM3   &       rs0("FIELD_NM")    & Chr(11)
                        gFieldLen3  = gFieldLen3  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                        gFieldType3 = gFieldType3 & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                        gDefaultT3  = gDefaultT3  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                        gNextSeq3   = gNextSeq3   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                        gKeyTag3    = gKeyTag3    & UCASE(rs0("KEY_TAG"))    & Chr(11)
                        gHidden3    = gHidden3    & "0"                      & Chr(11)
                        gSortDirection3 = gSortDirection3 & "ASC"            & Chr(11)
                End Select
                rs0.MoveNext
            Loop
            Set rs0 = Nothing
        Else
            iStrSQL = "SELECT type_cd, field_cd, field_nm, field_len, field_type, default_t, next_seq, key_tag " & _
		              "FROM Z_ADO_FIELD_INF WHERE PGM_ID = '" & iPgmId & "' AND TYPE_CD = '" & iTypeCD & _
		              "' AND SPD_NO = '" & iSpdNo & "' AND UPPER(LANG_CD) = '" & gLang & "' ORDER BY SEQ_NO"
            lgstrRetMsg = RequestZADO(iStrSQL, iResultData)
		
            If lgstrRetMsg = False Then
                Call SetADOFeild(iSpdNo,iTypeCD)
                Call SplitADOFieldVar(iSpdNo)
                Exit Sub
            End If
		
            iRowData = Split(iResultData,mgRowSep)
        
            For i = 0 To UBound(iRowData) - 1
                iColData = Split(iRowData(i),mgColSep)
                Select Case  iSpdNo
                    Case "A"
                        gTypeCD    = gTypeCD    & UCASE(iColData(0))   & Chr(11)	     
                        gFieldCD   = gFieldCD   & UCASE(iColData(1))   & Chr(11)
                        gFieldNM   = gFieldNM   &       iColData(2)    & Chr(11)
                        gFieldLen  = gFieldLen  & UCASE(iColData(3))   & Chr(11)
                        gFieldType = gFieldType & UCASE(iColData(4))   & Chr(11)
                        gDefaultT  = gDefaultT  & UCASE(iColData(5))   & Chr(11)
                        gNextSeq   = gNextSeq   & UCASE(iColData(6))   & Chr(11)
                        gKeyTag    = gKeyTag    & UCASE(iColData(7))   & Chr(11)
                        gHidden    = gHidden    & "0"                  & Chr(11)
                        gSortDirection = gSortDirection & "ASC"        & Chr(11)		              
                    Case "B"
                        gTypeCD1    = gTypeCD1    & UCASE(iColData(0)) & Chr(11)	     
                        gFieldCD1   = gFieldCD1   & UCASE(iColData(1)) & Chr(11)
                        gFieldNM1   = gFieldNM1   &       iColData(2)  & Chr(11)
                        gFieldLen1  = gFieldLen1  & UCASE(iColData(3)) & Chr(11)
                        gFieldType1 = gFieldType1 & UCASE(iColData(4)) & Chr(11)
                        gDefaultT1  = gDefaultT1  & UCASE(iColData(5)) & Chr(11)
                        gNextSeq1   = gNextSeq1   & UCASE(iColData(6)) & Chr(11)
                        gKeyTag1    = gKeyTag1    & UCASE(iColData(7)) & Chr(11)
                        gHidden1    = gHidden1    & "0"                & Chr(11)
                        gSortDirection1 = gSortDirection1 & "ASC"      & Chr(11)
                    Case "C"
                        gTypeCD2    = gTypeCD2    & UCASE(iColData(0)) & Chr(11)	     
                        gFieldCD2   = gFieldCD2   & UCASE(iColData(1)) & Chr(11)
                        gFieldNM2   = gFieldNM2   &       iColData(2)  & Chr(11)
                        gFieldLen2  = gFieldLen2  & UCASE(iColData(3)) & Chr(11)
                        gFieldType2 = gFieldType2 & UCASE(iColData(4)) & Chr(11)
                        gDefaultT2  = gDefaultT2  & UCASE(iColData(5)) & Chr(11)
                        gNextSeq2   = gNextSeq2   & UCASE(iColData(6)) & Chr(11)
                        gKeyTag2    = gKeyTag2    & UCASE(iColData(7)) & Chr(11)
                        gHidden2    = gHidden2    & "0"                & Chr(11)
                        gSortDirection2 = gSortDirection2 & "ASC"      & Chr(11)
                    Case "D"
                        gTypeCD3    = gTypeCD3    & UCASE(iColData(0)) & Chr(11)	     
                        gFieldCD3   = gFieldCD3   & UCASE(iColData(1)) & Chr(11)
                        gFieldNM3   = gFieldNM3   &       iColData(2)  & Chr(11)
                        gFieldLen3  = gFieldLen3  & UCASE(iColData(3)) & Chr(11)
                        gFieldType3 = gFieldType3 & UCASE(iColData(4)) & Chr(11)
                        gDefaultT3  = gDefaultT3  & UCASE(iColData(5)) & Chr(11)
                        gNextSeq3   = gNextSeq3   & UCASE(iColData(6)) & Chr(11)
                        gKeyTag3    = gKeyTag3    & UCASE(iColData(7)) & Chr(11)
                        gHidden3    = gHidden3    & "0"                & Chr(11)
                        gSortDirection3 = gSortDirection3 & "ASC"      & Chr(11)
                End Select
            Next
        End If
        
        Select Case  iSpdNo
            Case "A"
                    Call ggoSpread.SetDBAgentData(gTypeCD,gFieldCD,gFieldNM,gFieldLen,gFieldType,gDefaultT,gNextSeq,gKeyTag,gSortDirection,gHidden,mC_MaxSelList)
	        Case "B"
	           Call ggoSpread.SetDBAgentData(gTypeCD1,gFieldCD1,gFieldNM1,gFieldLen1,gFieldType1,gDefaultT1,gNextSeq1,gKeyTag1,gSortDirection1,gHidden1,mC_MaxSelList)
            Case "C"
	           Call ggoSpread.SetDBAgentData(gTypeCD2,gFieldCD2,gFieldNM2,gFieldLen2,gFieldType2,gDefaultT2,gNextSeq2,gKeyTag2,gSortDirection2,gHidden2,mC_MaxSelList)
            Case "D"
	           Call ggoSpread.SetDBAgentData(gTypeCD3,gFieldCD3,gFieldNM3,gFieldLen3,gFieldType3,gDefaultT3,gNextSeq3,gKeyTag3,gSortDirection3,gHidden3,mC_MaxSelList)
        End Select   
    End If
    'dumy Data For Export
    Select Case  iSpdNo
	     Case "A"
                  gTypeCD     = gTypeCD     & iTypeCD & Chr(11)	     
                  gFieldCD    = gFieldCD    & "1"       & Chr(11)
                  gFieldNM    = gFieldNM    & "1"       & Chr(11)
                  gFieldLen   = gFieldLen   & "2"       & Chr(11)
                  gFieldType  = gFieldType  & "HH"      & Chr(11)
                  gDefaultT   = gDefaultT   & "L"       & Chr(11)
                  gNextSeq    = gNextSeq    & "0"       & Chr(11)
                  gKeyTag     = gKeyTag     & "0"       & Chr(11)
                  gHidden     = gHidden     & "0"       & Chr(11)
                  gSortDirection = gSortDirection & "ASC" & Chr(11)
	     Case "B"
                  gTypeCD1    = gTypeCD1    & iTypeCD & Chr(11)	     
                  gFieldCD1   = gFieldCD1   & "1"       & Chr(11)
                  gFieldNM1   = gFieldNM1   & "1"       & Chr(11)
                  gFieldLen1  = gFieldLen1  & "2"       & Chr(11)
                  gFieldType1 = gFieldType1 & "HH"      & Chr(11)
                  gDefaultT1  = gDefaultT1  & "L"       & Chr(11)
                  gNextSeq1   = gNextSeq1   & "0"       & Chr(11)
                  gKeyTag1    = gKeyTag1    & "0"       & Chr(11)
                  gHidden1    = gHidden1    & "0"       & Chr(11)
                  gSortDirection1 = gSortDirection1 & "ASC" & Chr(11)
	     Case "C"
                  gTypeCD2    = gTypeCD2    & iTypeCD & Chr(11)	     
                  gFieldCD2   = gFieldCD2   & "1"       & Chr(11)
                  gFieldNM2   = gFieldNM2   & "1"       & Chr(11)
                  gFieldLen2  = gFieldLen2  & "2"       & Chr(11)
                  gFieldType2 = gFieldType2 & "HH"      & Chr(11)
                  gDefaultT2  = gDefaultT2  & "L"       & Chr(11)
                  gNextSeq2   = gNextSeq2   & "0"       & Chr(11)
                  gKeyTag2    = gKeyTag2    & "0"       & Chr(11)
                  gHidden2    = gHidden2    & "0"       & Chr(11)
                  gSortDirection2 = gSortDirection2 & "ASC" & Chr(11)
	     Case "D"
                  gTypeCD3    = gTypeCD3    & iTypeCD & Chr(11)	     
                  gFieldCD3   = gFieldCD3   & "1"       & Chr(11)
                  gFieldNM3   = gFieldNM3   & "1"       & Chr(11)
                  gFieldLen3  = gFieldLen3  & "2"       & Chr(11)
                  gFieldType3 = gFieldType3 & "HH"      & Chr(11)
                  gDefaultT3  = gDefaultT3  & "L"       & Chr(11)
                  gNextSeq3   = gNextSeq3   & "0"       & Chr(11)
                  gKeyTag3    = gKeyTag3    & "0"       & Chr(11)
                  gHidden3    = gHidden3    & "0"       & Chr(11)
                  gSortDirection3 = gSortDirection3 & "ASC" & Chr(11)
    End Select
    If UCase(iTypeCD) = "G" Then
	   gMethodText = "집계"
	Else   
	   gMethodText = "정렬"
	End If   

    Call SplitADOFieldVar(iSpdNo)

End Sub

Function RequestZADO(ByVal pStrSQL,prData)
    Dim iXmlHttp,iRetByte
    
	On Error Resume Next
	Err.Clear
	
    RequestZADO = False
    
    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP")		
    
    iXmlHttp.open "POST", GetComaspFolderPath & "RequestCommonQry.asp", False        
    iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    pStrSQL = escape(pStrSQL)
    pStrSQL = Replace(pStrSQL, "+", "%2B")
    pStrSQL = Replace(pStrSQL, "/", "%2F")

    iXmlHttp.send "LangCD=" & glang & "&ADODBConnString=" & escape(gADODBConnString) & "&StrSQL=" & pStrSQL

    If gCharSet = "D" Then 'U : unicode, D:DBCS
       prData   = ConnectorControl.CStrConv(iXmlHttp.responseBody)
    Else
       prData   = iXmlHttp.responseText
    End If   
          
    If prData <> "" Then
        RequestZADO = True
    End If
    Set iXmlHttp = Nothing
End Function


Sub SetADOFeild(ByVal iSpdNo,ByVal iTypeCD)

    Select Case  iSpdNo
              Case "A"
                     gTypeCD     = gTypeCD      & iTypeCD  & Chr(11)
                     gFieldCD    = gFieldCD     & "Error"  & Chr(11)
                     gFieldNM    = gFieldNM     & "N/A"    & Chr(11)
                     gFieldLen   = gFieldLen    & "30"     & Chr(11)
                     gFieldType  = gFieldType   & "ST"     & Chr(11)
                     gDefaultT   = gDefaultT    & "L"      & Chr(11)
                     gNextSeq    = gNextSeq     & "0"      & Chr(11)
                     gKeyTag     = gKeyTag      & "0"      & Chr(11)
                     gHidden     = gHidden      & "0"      & Chr(11)
                     gSortDirection = gSortDirection & "ASC" & Chr(11)
              Case "B"
                     gTypeCD1    = gTypeCD1     & iTypeCD  & Chr(11)
                     gFieldCD1   = gFieldCD1    & "Error"  & Chr(11)
                     gFieldNM1   = gFieldNM1    & "N/A"    & Chr(11)
                     gFieldLen1  = gFieldLen1   & "30"     & Chr(11)
                     gFieldType1 = gFieldType1  & "ST"     & Chr(11)
                     gDefaultT1  = gDefaultT1   & "L"      & Chr(11)
                     gNextSeq1   = gNextSeq1    & "0"      & Chr(11)
                     gKeyTag1    = gKeyTag1     & "0"      & Chr(11)
                     gHidden1    = gHidden1     & "0"      & Chr(11)
                     gSortDirection1 = gSortDirection1 & "ASC" & Chr(11)
              Case "C"
                     gTypeCD2    = gTypeCD2     & iTypeCD  & Chr(11)
                     gFieldCD2   = gFieldCD2    & "Error"  & Chr(11)
                     gFieldNM2   = gFieldNM2    & "N/A"    & Chr(11)
                     gFieldLen2  = gFieldLen2   & "30"     & Chr(11)
                     gFieldType2 = gFieldType2  & "ST"     & Chr(11)
                     gDefaultT2  = gDefaultT2   & "L"      & Chr(11)
                     gNextSeq2   = gNextSeq2    & "0"      & Chr(11)
                     gKeyTag2    = gKeyTag2     & "0"      & Chr(11)
                     gHidden2    = gHidden2     & "0"      & Chr(11)
                     gSortDirection2 = gSortDirection2 & "ASC" & Chr(11)
              Case "D"
                     gTypeCD3    = gTypeCD3     & iTypeCD  & Chr(11)
                     gFieldCD3   = gFieldCD3    & "Error"  & Chr(11)
                     gFieldNM3   = gFieldNM3    & "N/A"    & Chr(11)
                     gFieldLen3  = gFieldLen3   & "30"     & Chr(11)
                     gFieldType3 = gFieldType3  & "ST"     & Chr(11)
                     gDefaultT3  = gDefaultT3   & "L"      & Chr(11)
                     gNextSeq3   = gNextSeq3    & "0"      & Chr(11)
                     gKeyTag3    = gKeyTag3     & "0"      & Chr(11)
                     gHidden3    = gHidden3     & "0"      & Chr(11)
                     gSortDirection3 = gSortDirection3 & "ASC" & Chr(11)
     End Select 

End Sub

Sub SplitADOFieldVar(ByVal iSpdNo)
    Select Case  iSpdNo
	     Case "A"
                 gTypeCD     = Split (gTypeCD    ,Chr(11))                           
                 gFieldCD    = Split (gFieldCD   ,Chr(11))                           
                 gFieldNM    = Split (gFieldNM   ,Chr(11))                           
                 gFieldLen   = Split (gFieldLen  ,Chr(11))                           
                 gFieldType  = Split (gFieldType ,Chr(11))                           
                 gDefaultT   = Split (gDefaultT  ,Chr(11))                           
                 gNextSeq    = Split (gNextSeq   ,Chr(11))                           
                 gKeyTag     = Split (gKeyTag    ,Chr(11))  
                 gHidden     = Split (gHidden    ,Chr(11))  
                 gSortDirection = Split (gSortDirection ,Chr(11))  
	     Case "B"
                 gTypeCD1    = Split (gTypeCD1   ,Chr(11))                           
                 gFieldCD1   = Split (gFieldCD1  ,Chr(11))                           
                 gFieldNM1   = Split (gFieldNM1  ,Chr(11))                           
                 gFieldLen1  = Split (gFieldLen1 ,Chr(11))                           
                 gFieldType1 = Split (gFieldType1,Chr(11))                           
                 gDefaultT1  = Split (gDefaultT1 ,Chr(11))                           
                 gNextSeq1   = Split (gNextSeq1  ,Chr(11))                           
                 gKeyTag1    = Split (gKeyTag1   ,Chr(11))
                 gHidden1    = Split (gHidden1   ,Chr(11))  
                 gSortDirection1 = Split (gSortDirection1 ,Chr(11))                             
	     Case "C"
                 gTypeCD2    = Split (gTypeCD2   ,Chr(11))                           
                 gFieldCD2   = Split (gFieldCD2  ,Chr(11))                           
                 gFieldNM2   = Split (gFieldNM2  ,Chr(11))                           
                 gFieldLen2  = Split (gFieldLen2 ,Chr(11))                           
                 gFieldType2 = Split (gFieldType2,Chr(11))                           
                 gDefaultT2  = Split (gDefaultT2 ,Chr(11))                           
                 gNextSeq2   = Split (gNextSeq2  ,Chr(11))                           
                 gKeyTag2    = Split (gKeyTag2   ,Chr(11))
                 gHidden2    = Split (gHidden2   ,Chr(11))  
                 gSortDirection2 = Split (gSortDirection2 ,Chr(11))                             
	     Case "D"
                 gTypeCD3    = Split (gTypeCD3   ,Chr(11))                           
                 gFieldCD3   = Split (gFieldCD3  ,Chr(11))                           
                 gFieldNM3   = Split (gFieldNM3  ,Chr(11))                           
                 gFieldLen3  = Split (gFieldLen3 ,Chr(11))                           
                 gFieldType3 = Split (gFieldType3,Chr(11))                           
                 gDefaultT3  = Split (gDefaultT3 ,Chr(11))                           
                 gNextSeq3   = Split (gNextSeq3  ,Chr(11))                           
                 gKeyTag3    = Split (gKeyTag3   ,Chr(11))
                 gHidden3    = Split (gHidden3   ,Chr(11))  
                 gSortDirection3 = Split (gSortDirection3 ,Chr(11))                             
    End Select
    
End Sub

Sub SetCellTypeOfSpreadSheet(pObject,ByVal iCol,ByVal pFieldType,ByVal pFieldNM,ByVal pFieldLen,ByVal pHidden)
   Dim iAlign
   
   ggoSpread.Source = pObject
   
   iAlign = Trim(Mid(pFieldType,3,1))
   
   If iAlign = "" Then
      Select Case Mid(pFieldType,1,1)
         Case "D"  : iAlign = "2"
         Case "F"  : iAlign = "1"
         Case "T"  : iAlign = "2"
         Case Else : iAlign = "0"
      End Select   
   End If   

'   If pHidden <> "0" Then      
'      ggoSpread.Source.Col = iCol
'      ggoSpread.Source.ColHidden = True
'   End If      
   Select Case Mid(pFieldType,1,2) 
     Case "BT" 'Button
		    ggoSpread.SSSetButton iCol
     Case "CB" 'Combo
            ggoSpread.SSSetCombo  iCol , pFieldNM , pFieldLen,iAlign
     Case "CK" 'Check
            ggoSpread.SSSetCheck  iCol , pFieldNM , pFieldLen,iAlign, "", True, -1 
     Case "DD"   '날짜 
            ggoSpread.SSSetDate   iCol , pFieldNM , pFieldLen,iAlign,gDateFormat
     Case "D5"   '편집(Year,Month) - 2003/01/26 europen
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign
     Case "ED"   '편집 
'           ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign                          '2003/04/29 LEE JINSOO
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign             ,     , 200  '2003/04/29 LEE JINSOO

     Case "F2"  ' 금액 
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "2")
     Case "F3"  ' 수량 
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "3")
     Case "F4"  ' 단가 
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "4")
     Case "F5"   ' 환율 
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "5")
     Case "F6"   ' user-defined
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "6")
     Case "F7"   ' user-defined
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "7")
     Case "F8"   ' user-defined
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "8")
     Case "F9"   ' user-defined
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "9")

     Case "FA"   ' money default
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "A")
     Case "FB"   ' qty default
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "B")
     Case "FC"   ' unitcost default
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "C")
     Case "FD"   ' exchrate default
            Call SetSpreadFloat (iCol , pFieldNM , pFieldLen,iAlign, "D")

     Case "MK"   ' Mask
            ggoSpread.SSSetMask   iCol , pFieldNM , pFieldLen,iAlign
     Case "ST"   ' Static
            ggoSpread.SSSetStatic iCol , pFieldNM , pFieldLen,iAlign
     Case "TT"   ' Time
            ggoSpread.SSSetTime   iCol , pFieldNM , pFieldLen,iAlign,1,1
	 Case "HH"	 ' Hidden
    '           ggoSpread.SSSetEdit   iCol , "" , 0,iAlign                          '2003/04/29 LEE JINSOO
                ggoSpread.SSSetEdit   iCol , "" , 0,iAlign             ,     , 200  '2003/04/29 LEE JINSOO

	        Call ggoSpread.SSSetColHidden(iCol,iCol,True)
     Case Else
'           ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign                          '2003/04/29 LEE JINSOO
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign             ,     , 200  '2003/04/29 LEE JINSOO
   End Select
   If pHidden = "-1" Then      
       Call ggoSpread.SSSetColHidden(iCol,iCol,True,"D")
   End If      
   
End Sub


Function MakeSQLGroupOrderByList(ByVal pSpdNo)
    Dim iStr,jStr
    Dim ii,jj,kk      
    Dim tmpPopUpR   
    Dim iMark
    Dim iFirst    
    Dim pMaxColCnt
    Dim pPopUpR
    
    Select Case UCase(pSpdNo)
      Case "A"
              tmpTypeCD    = gTypeCD
              tmpFieldCD   = gFieldCD
              tmpFieldNM   = gFieldNM
              tmpFieldLen  = gFieldLen
              tmpFieldType = gFieldType
              tmpDefaultT  = gDefaultT
              tmpNextSeq   = gNextSeq
              tmpKeyTag    = gKeyTag
              pPopUpR      = gPopUpR_A
      Case "B"
              tmpTypeCD    = gTypeCD1
              tmpFieldCD   = gFieldCD1
              tmpFieldNM   = gFieldNM1
              tmpFieldLen  = gFieldLen1
              tmpFieldType = gFieldType1
              tmpDefaultT  = gDefaultT1
              tmpNextSeq   = gNextSeq1
              tmpKeyTag    = gKeyTag1
              pPopUpR      = gPopUpR_B
      Case "C"
              tmpTypeCD    = gTypeCD2
              tmpFieldCD   = gFieldCD2
              tmpFieldNM   = gFieldNM2
              tmpFieldLen  = gFieldLen2
              tmpFieldType = gFieldType2
              tmpDefaultT  = gDefaultT2
              tmpNextSeq   = gNextSeq2
              tmpKeyTag    = gKeyTag2
              pPopUpR      = gPopUpR_C
      Case "D"
              tmpTypeCD    = gTypeCD3
              tmpFieldCD   = gFieldCD3
              tmpFieldNM   = gFieldNM3
              tmpFieldLen  = gFieldLen3
              tmpFieldType = gFieldType3
              tmpDefaultT  = gDefaultT3
              tmpNextSeq   = gNextSeq3
              tmpKeyTag    = gKeyTag3
              pPopUpR      = gPopUpR_D
    End Select        

    pMaxColCnt = UBound(tmpTypeCD)
    Redim tmpPopUpR(pMaxColCnt)    
    
    For kk = 0 to pMaxColCnt - 1
        tmpPopUpR(kk) = pPopUpR(kk,0)
    Next
    
    MakeSQLGroupOrderByList = "" 
    
    iStr   = ""
    jStr   = ""      
    iFirst = "N"

    Redim  iMark(pMaxColCnt) 

    For ii = 0 to pMaxColCnt - 1
        If tmpPopUpR(ii) <> "" Then     
           If tmpTypeCD(0) = "G" Then
              For jj = 0 To pMaxColCnt - 1                                            
                  If iMark(jj) <> "X" Then
                     If pPopUpR(ii,0) = Trim(tmpFieldCD(jj)) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If                        
               
                        If CInt(Trim(tmpNextSeq(jj))) >= 1 And CInt(Trim(tmpNextSeq(jj))) <= pMaxColCnt Then
                           iStr = iStr & pPopUpR(ii,0) & " " & pPopUpR(ii,1) & "," & tmpFieldCD(CInt(tmpNextSeq(jj)) - 1)
                           jStr = jStr & pPopUpR(ii,0) & " " &                 "," & tmpFieldCD(CInt(tmpNextSeq(jj)) - 1)
                           
                           If (ii + 1) <  pMaxColCnt   Then
                              For kk = ii + 1 to pMaxColCnt - 1
                                  If  pPopUpR(kk,0) = Trim(tmpFieldCD(CInt(tmpNextSeq(jj)) - 1))  Then
                                      iStr = iStr & " " & pPopUpR(kk,1) 
                                      tmpPopUpR(kk) = ""
                                  End If    
                              Next    
                           End If                              
                           iMark(CInt(tmpNextSeq(jj)) - 1) = "X"
                        Else
                          iStr = iStr & pPopUpR(ii,0) & " " & pPopUpR(ii,1)
                          jStr = jStr & pPopUpR(ii,0)
                        End If
                        iFirst = "Y"
                        iMark(jj) = "X"
                     End If
                     
                  End If
              Next
           Else
              If iFirst = "Y" Then
                 iStr = iStr & " , "
                 jStr = jStr & " , " 
              End If                         
           
              iStr = iStr & pPopUpR(ii,0) & " " & pPopUpR(ii,1)
              iFirst = "Y"
           End If
              
        End If
    Next     
    
    If tmpTypeCD(0) = "G" Then
       MakeSQLGroupOrderByList =  "Group By " & jStr  & " Order By " & iStr 
    Else
       If Trim(iStr) <> "" Then
          MakeSQLGroupOrderByList = "Order By " & iStr
       End If
    End If   

End Function

Sub MakePopData(pPopUpR,pSortFieldNm,pSortFieldCD,pSortDirection,ByVal pMaxSelList)
	Dim ii,kk	
	Dim iCast

    pSortFieldNm  = ""
    pSortFieldCD  = ""
    
    For ii = 0 To UBound(tmpFieldNM) - 1                                       
        iCast = tmpDefaultT(ii)
        If  IsNumeric(iCast) Or Trim(tmpDefaultT(ii)) = "V" Then
            If IsNumeric(iCast) Then 
'               If mIsBetween(1,pMaxSelList,CInt(iCast)) Then    'Sort정보default값 저장 
                  pPopUpR(CInt(tmpDefaultT(ii)) - 1,0) = Trim(tmpFieldCD(ii))
                  pPopUpR(CInt(tmpDefaultT(ii)) - 1,1) = Trim(pSortDirection(ii))
'               End If
            End If
            pSortFieldNm  = pSortFieldNm  & Trim(tmpFieldNM(ii)) & Chr(11)
            pSortFieldCD  = pSortFieldCD  & Trim(tmpFieldCD(ii))  & Chr(11)
        End If
    Next
    
    pSortFieldNm     = Split (pSortFieldNm ,Chr(11))
    pSortFieldCD     = Split (pSortFieldCD ,Chr(11))


End Sub

Function InitSpreadSheetFieldOfZADO(pObject,pPopUpR,pSelectList,pSelectListDT,pKeyPos,ByVal pMaxKey,ByVal pMaxSelList)

    Dim iMark
    Dim iSeq
    Dim iFieldCount
    Dim iAggregateList
    Dim iMaxList
    Dim iIntTemp
    
    iFieldCount = UBound(tmpFieldNM)
    Redim  iMark(iFieldCount) 

    For ii = 0 to iFieldCount - 1
        For jj = 0 to iFieldCount - 1
            If iMark(jj) <> "X" Then
               If pPopUpR(ii,0) = Trim(tmpFieldCD(jj)) Then
                  iSeq = iSeq + 1

                  Call InitSpreadSheetRow(pObject,iSeq,jj,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                  
                  If mIsBetween(1,iFieldCount,CInt(tmpNextSeq(jj))) Then 
                     kk = CInt(tmpNextSeq(jj)) 
                     iSeq = iSeq + 1
                     Call InitSpreadSheetRow(pObject,iSeq,kk-1,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                  End If    
               End If 
            End If 
        Next      
    Next      

    If tmpTypeCD(0) = "G" Then
        Redim iAggregateList(iFieldCount)
        iMaxList = -1
        For ii = 0 to iFieldCount - 1
            If iMark(ii) <> "X" Then
                If Left(tmpDefaultT(ii),1) = "L" And Len(tmpDefaultT(ii)) > 1 Then
                    iIntTemp = CInt(Mid(tmpDefaultT(ii),2)) - 1
                    iAggregateList(iIntTemp) = ii
                    If iIntTemp > iMaxList Then iMaxList = iIntTemp
                End If
            End If
        Next
        If iMaxList > -1 Then
            For ii = 0 to iMaxList
                iIntTemp = iAggregateList(ii)
                iSeq = iSeq + 1
                Call InitSpreadSheetRow(pObject,iSeq,iIntTemp,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
            Next      
        End If
    End If         

    For ii = 0 to iFieldCount - 1
        If iMark(ii) <> "X" Then
           If tmpTypeCD(0) = "S" Or (tmpTypeCD(0) = "G" And (tmpDefaultT(ii) = "L" Or tmpDefaultT(ii) = "HH")) Then
              iSeq = iSeq + 1
              
              Call InitSpreadSheetRow(pObject,iSeq,ii,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
              If tmpDefaultT(ii) = "HH" Then
                  ggoSpread.Source = pObject
                  Call ggoSpread.SSSetColHidden(iSeq,iSeq,True)
              End If
              
              If mIsBetween(1,UBound(tmpFieldNM),CInt(tmpNextSeq(ii))) Then 
                 kk = CInt(tmpNextSeq(ii)) 
                 iSeq = iSeq + 1
                 Call InitSpreadSheetRow(pObject,iSeq,kk-1,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                 If tmpDefaultT(kk-1) = "HH" Then
                     ggoSpread.Source = pObject
                     Call ggoSpread.SSSetColHidden(iSeq,iSeq,True)
                 End If    
              End If   
           End If   
        End If 
   Next      
   
   InitSpreadSheetFieldOfZADO =  iSeq
   
	If Trim(pSelectList) <> "" Then
	   If Mid(pSelectList,Len(pSelectList),1) = "," Then
	      pSelectList = Mid(pSelectList,1,Len(pSelectList)-1)
       End If   
	End If   
   
End Function


'========================================================================================================
' Function Name : InitSpreadSheetRow
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheetRow(pObject,ByVal iCol,ByVal iDx,pMark,pSelectList,pSelectListDT,pKeyPos,ByVal pMaxKey)
   pMark(iDx) = "X"
   Call SetCellTypeOfSpreadSheet(pObject,iCol,tmpFieldType(iDx),tmpFieldNM(iDx),tmpFieldLen(iDx),tmpHidden(iDx))
   
   pSelectList   = pSelectList   & tmpFieldCD(iDx)            & ","      'Make select part of SQL
   pSelectListDT = pSelectListDT & Mid(tmpFieldType(iDx),1,2) & mgColSep
   
   If CInt(tmpKeyTag(iDx)) > 0 And CInt(tmpKeyTag(iDx)) <= pMaxKey Then  
      pKeyPos(CInt(tmpKeyTag(iDx))) =  iCol
	End If
End Sub



Sub SetZAdoSpreadSheet(ByVal pPgmID ,ByVal pTypeCD, ByVal pSpdNo, ByVal pVersionNo, ByVal pScreenType,pSpdObjectName, ByVal pMaxKey, ByVal pFuture1, ByVal pFuture2)
   Dim iMaxColumn
   Dim iMaxFieldCount
   Dim tmpKeyPos
   
   Dim itmpSelectList
   Dim itmpSelectList_DT
   Dim itmpPopUpR
   
   Dim tmpSortFieldCD
   Dim tmpSortFieldNm

MsgBox "pVersionNo : " & pVersionNo   
   Set gSpdObject(Asc(UCase(pSpdNo))-65) = pSpdObjectName
   Call  GetZAdoFieldInf(pPgmID,pTypeCD,pSpdNo,pVersionNo,pSpdObjectName)
   Select Case UCase(pSpdNo)
      Case "A"
              tmpTypeCD    = gTypeCD
              tmpFieldCD   = gFieldCD
              tmpFieldNM   = gFieldNM
              tmpFieldLen  = gFieldLen
              tmpFieldType = gFieldType
              tmpDefaultT  = gDefaultT
              tmpNextSeq   = gNextSeq
              tmpKeyTag    = gKeyTag
              tmpHidden    = gHidden
              tmpSortDirection = gSortDirection              
              ReDim tmpKeyPos(pMaxKey)
      Case "B"
              tmpTypeCD    = gTypeCD1
              tmpFieldCD   = gFieldCD1
              tmpFieldNM   = gFieldNM1
              tmpFieldLen  = gFieldLen1
              tmpFieldType = gFieldType1
              tmpDefaultT  = gDefaultT1
              tmpNextSeq   = gNextSeq1
              tmpKeyTag    = gKeyTag1
              tmpHidden    = gHidden1
              tmpSortDirection = gSortDirection1
              ReDim tmpKeyPos(pMaxKey)
      Case "C"
              tmpTypeCD    = gTypeCD2
              tmpFieldCD   = gFieldCD2
              tmpFieldNM   = gFieldNM2
              tmpFieldLen  = gFieldLen2
              tmpFieldType = gFieldType2
              tmpDefaultT  = gDefaultT2
              tmpNextSeq   = gNextSeq2
              tmpKeyTag    = gKeyTag2
              tmpHidden    = gHidden2
              tmpSortDirection = gSortDirection2
              ReDim tmpKeyPos(pMaxKey)
      Case "D"
              tmpTypeCD    = gTypeCD3
              tmpFieldCD   = gFieldCD3
              tmpFieldNM   = gFieldNM3
              tmpFieldLen  = gFieldLen3
              tmpFieldType = gFieldType3
              tmpDefaultT  = gDefaultT3
              tmpNextSeq   = gNextSeq3
              tmpKeyTag    = gKeyTag3
              tmpHidden    = gHidden3
              tmpSortDirection = gSortDirection3
              ReDim tmpKeyPos(pMaxKey)
    End Select

    iMaxFieldCount = UBound(tmpFieldCD)
    ReDim itmpPopUpR(iMaxFieldCount - 1,1)

    Call MakePopData(itmpPopUpR,tmpSortFieldNm,tmpSortFieldCD,tmpSortDirection,mC_MaxSelList)

    ggoSpread.Source = pSpdObjectName

    With pSpdObjectName
        .MaxCols = 0
        .MaxCols = iMaxFieldCount
        .MaxRows = 0
        .ReDraw = False
    End With

    ggoSpread.Spreadinit pVersionNo , pScreenType

    iMaxColumn = InitSpreadSheetFieldOfZADO(pSpdObjectName, itmpPopUpR, itmpSelectList, itmpSelectList_DT, tmpKeyPos, pMaxKey, mC_MaxSelList)

    ggoSpread.Source = pSpdObjectName
    pSpdObjectName.ColsFrozen = 0
    Call ggoSpread.SplitData()

    With pSpdObjectName
         .MaxCols = iMaxColumn
         .ReDraw = True
    End With

   Select Case UCase(pSpdNo)
      Case "A"
              gSelectList_A   = itmpSelectList
              gSelectListDT_A = itmpSelectList_DT

              gPopUpR_A       = itmpPopUpR
              gSortFieldCD_A  = tmpSortFieldCD
              gSortFieldNm_A  = tmpSortFieldNm
              
              gKeyPos_A       = tmpKeyPos
              ReDim gKeyPosVal_A(pMaxKey)  
      Case "B"
              gSelectList_B   = itmpSelectList
              gSelectListDT_B = itmpSelectList_DT
              gPopUpR_B       = itmpPopUpR

              gSortFieldCD_B  = tmpSortFieldCD
              gSortFieldNm_B  = tmpSortFieldNm

              gKeyPos_B       = tmpKeyPos
              ReDim gKeyPosVal_B(pMaxKey)   
      Case "C"
              gSelectList_C   = itmpSelectList
              gSelectListDT_C = itmpSelectList_DT
              gPopUpR_C       = itmpPopUpR

              gSortFieldCD_C  = tmpSortFieldCD
              gSortFieldNm_C  = tmpSortFieldNm

              gKeyPos_C       = tmpKeyPos
              ReDim gKeyPosVal_C(pMaxKey)   
      Case "D"
              gSelectList_D   = itmpSelectList
              gSelectListDT_D = itmpSelectList_DT
              gPopUpR_D       = itmpPopUpR

              gSortFieldCD_D  = tmpSortFieldCD
              gSortFieldNm_D  = tmpSortFieldNm

              gKeyPos_D       = tmpKeyPos
              ReDim gKeyPosVal_D(pMaxKey)   
   End Select
End Sub

'=================================================================================
'
'=================================================================================
Function GetSQLSelectList(ByVal pSpdNo)
    Select Case UCase(pSpdNo)
      Case "A"
            GetSQLSelectList = gSelectList_A
      Case "B"
            GetSQLSelectList = gSelectList_B
      Case "C"
            GetSQLSelectList = gSelectList_C
      Case "D"
            GetSQLSelectList = gSelectList_D
    End Select  
End Function

'=================================================================================
'
'=================================================================================
Function GetPopUpR(ByVal pSpdNo)
    Select Case UCase(pSpdNo)
      Case "A"
            GetPopUpR = gPopUpR_A
      Case "B"
            GetPopUpR = gPopUpR_B
      Case "C"
            GetPopUpR = gPopUpR_C
      Case "D"
            GetPopUpR = gPopUpR_D
    End Select  
End Function

'=================================================================================
'
'=================================================================================
Function GetSQLSelectListDataType(ByVal pSpdNo)
    Select Case UCase(pSpdNo)
      Case "A"
            GetSQLSelectListDataType = gSelectListDT_A
      Case "B"
            GetSQLSelectListDataType = gSelectListDT_B
      Case "C"
            GetSQLSelectListDataType = gSelectListDT_C
      Case "D"
            GetSQLSelectListDataType = gSelectListDT_D
    End Select  
End Function

'=================================================================================
'
'=================================================================================
Sub SetSpreadColumnValue(ByVal pSpdNo, pSpdObject , ByVal Col, ByVal Row)
    Dim iLoop    
    Select Case UCase(pSpdNo)
      Case "A"
         For iLoop = 1 to UBound(gKeyPos_A)
             pSpdObject.Col = gKeyPos_A(iLoop)
             pSpdObject.Row = Row
             gKeyPosVal_A(iLoop) = pSpdObject.text
         Next    
      Case "B"
         For iLoop = 1 to UBound(gKeyPos_B)
             pSpdObject.Col = gKeyPos_B(iLoop)
             pSpdObject.Row = Row
             gKeyPosVal_B(iLoop) = pSpdObject.text
         Next
      Case "C"
         For iLoop = 1 to UBound(gKeyPos_C)
             pSpdObject.Col = gKeyPos_C(iLoop)
             pSpdObject.Row = Row
             gKeyPosVal_C(iLoop) = pSpdObject.text
         Next
      Case "D"
         For iLoop = 1 to UBound(gKeyPos_D)
             pSpdObject.Col = gKeyPos_D(iLoop)
             pSpdObject.Row = Row
             gKeyPosVal_D(iLoop) = pSpdObject.text
         Next
    End Select
End Sub

'=================================================================================
'
'=================================================================================
Function GetKeyPosVal(ByVal pSpdNo, ByVal iDx)
    Dim iVarTemp
    
    If gSpdObject(Asc(UCase(pSpdNo))-65).MaxRows = 0 Then
        Select Case UCase(pSpdNo)
            Case "A"
                iVarTemp = UBound(gKeyPosVal_A)
                ReDim gKeyPosVal_A(iVarTemp)
            Case "B"
                iVarTemp = UBound(gKeyPosVal_B)
                ReDim gKeyPosVal_B(iVarTemp)
            Case "C"
                iVarTemp = UBound(gKeyPosVal_C)
                ReDim gKeyPosVal_C(iVarTemp)
            Case "D"
                iVarTemp = UBound(gKeyPosVal_D)
                ReDim gKeyPosVal_D(iVarTemp)
        End Select
        Exit Function
    End If
    Select Case UCase(pSpdNo)
      Case "A"
         GetKeyPosVal = gKeyPosVal_A(iDx)
      Case "B"
         GetKeyPosVal = gKeyPosVal_B(iDx)
      Case "C"
         GetKeyPosVal = gKeyPosVal_C(iDx)
      Case "D"
         GetKeyPosVal = gKeyPosVal_D(iDx)
    End Select
End Function

'=================================================================================
'
'=================================================================================
Function GetKeyPos(ByVal pSpdNo, ByVal iDx)
    Select Case UCase(pSpdNo)
      Case "A"
         GetKeyPos = gKeyPos_A(iDx)
      Case "B"
         GetKeyPos = gKeyPos_B(iDx)
      Case "C"
         GetKeyPos = gKeyPos_C(iDx)
      Case "D"
         GetKeyPos = gKeyPos_D(iDx)
    End Select
End Function