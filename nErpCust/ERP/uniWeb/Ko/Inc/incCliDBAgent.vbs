'==============================================================================
' Name  :  GetAdoFiledInf
' Desc  :  Get Information for Creation of Query Window using ADO
'==============================================================================

Sub GetAdoFieldInf(ByVal iPgmId,ByVal iTypeCD,ByVal iSpdNo)
    Dim lgstrRetMsg                                '☜ : declaration Variable indicating Record Set Return Message
	Dim iResultData, iStrSQL, iRowData, iColData, i
	Dim ADF
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
	
    On Error Resume Next

    If UCase(iTypeCD) = "G" Then
	   gMethodText = "집계"
	Else   
	   gMethodText = "정렬"
	End If   
    
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
	     Case "B"
                gTypeCD1    = ""
                gFieldCD1   = ""
                gFieldNM1   = ""
                gFieldLen1  = ""
                gFieldType1 = ""
                gDefaultT1  = ""
                gNextSeq1   = ""
                gKeyTag1    = ""
	     Case "C"
                gTypeCD2    = ""
                gFieldCD2   = ""
                gFieldNM2   = ""
                gFieldLen2  = ""
                gFieldType2 = ""
                gDefaultT2  = ""
                gNextSeq2   = ""
                gKeyTag2    = ""
    End Select   

    Err.Clear

    If gRdsUse = "T" Then
        Redim UNISqlId(0)
        Redim UNIValue(0, 3)
        UNISqlId(0)    = "z100001"
        UNIValue(0, 0) = iPgmId
        UNIValue(0, 1) = iTypeCD
        UNIValue(0, 2) = iSpdNo
        UNIValue(0, 3) = gLang
	
        UNILock = DISCONNREAD :	UNIFlag = "1"
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
                Case "B"
                    gTypeCD1    = gTypeCD1    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                    gFieldCD1   = gFieldCD1   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                    gFieldNM1   = gFieldNM1   &       rs0("FIELD_NM")    & Chr(11)
                    gFieldLen1  = gFieldLen1  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                    gFieldType1 = gFieldType1 & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                    gDefaultT1  = gDefaultT1  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                    gNextSeq1   = gNextSeq1   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                    gKeyTag1    = gKeyTag1    & UCASE(rs0("KEY_TAG"))    & Chr(11)
                Case "C"
                    gTypeCD2    = gTypeCD2    & UCASE(rs0("TYPE_CD"))    & Chr(11)	     
                    gFieldCD2   = gFieldCD2   & UCASE(rs0("FIELD_CD"))   & Chr(11)
                    gFieldNM2   = gFieldNM2   &       rs0("FIELD_NM")    & Chr(11)
                    gFieldLen2  = gFieldLen2  & UCASE(rs0("FIELD_LEN"))  & Chr(11)
                    gFieldType2 = gFieldType2 & UCASE(rs0("FIELD_TYPE")) & Chr(11)
                    gDefaultT2  = gDefaultT2  & UCASE(rs0("DEFAULT_T"))  & Chr(11)
                    gNextSeq2   = gNextSeq2   & UCASE(rs0("NEXT_SEQ"))   & Chr(11)
                    gKeyTag2    = gKeyTag2    & UCASE(rs0("KEY_TAG"))    & Chr(11)
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

        iRowData = Split(iResultData,gRowSep)
     
        For i = 0 To UBound(iRowData) - 1
            iColData = Split(iRowData(i),gColSep)
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
                Case "B"
                    gTypeCD1    = gTypeCD1    & UCASE(iColData(0)) & Chr(11)	     
                    gFieldCD1   = gFieldCD1   & UCASE(iColData(1)) & Chr(11)
                    gFieldNM1   = gFieldNM1   &       iColData(2)  & Chr(11)
                    gFieldLen1  = gFieldLen1  & UCASE(iColData(3)) & Chr(11)
                    gFieldType1 = gFieldType1 & UCASE(iColData(4)) & Chr(11)
                    gDefaultT1  = gDefaultT1  & UCASE(iColData(5)) & Chr(11)
                    gNextSeq1   = gNextSeq1   & UCASE(iColData(6)) & Chr(11)
                    gKeyTag1    = gKeyTag1    & UCASE(iColData(7)) & Chr(11)
                Case "C"
                    gTypeCD2    = gTypeCD2    & UCASE(iColData(0)) & Chr(11)	     
                    gFieldCD2   = gFieldCD2   & UCASE(iColData(1)) & Chr(11)
                    gFieldNM2   = gFieldNM2   &       iColData(2)  & Chr(11)
                    gFieldLen2  = gFieldLen2  & UCASE(iColData(3)) & Chr(11)
                    gFieldType2 = gFieldType2 & UCASE(iColData(4)) & Chr(11)
                    gDefaultT2  = gDefaultT2  & UCASE(iColData(5)) & Chr(11)
                    gNextSeq2   = gNextSeq2   & UCASE(iColData(6)) & Chr(11)
                    gKeyTag2    = gKeyTag2    & UCASE(iColData(7)) & Chr(11)
            End Select
        Next
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
	     Case "B"
                  gTypeCD1    = gTypeCD1    & iTypeCD & Chr(11)	     
                  gFieldCD1   = gFieldCD1   & "1"       & Chr(11)
                  gFieldNM1   = gFieldNM1   & "1"       & Chr(11)
                  gFieldLen1  = gFieldLen1  & "2"       & Chr(11)
                  gFieldType1 = gFieldType1 & "HH"      & Chr(11)
                  gDefaultT1  = gDefaultT1  & "L"       & Chr(11)
                  gNextSeq1   = gNextSeq1   & "0"       & Chr(11)
                  gKeyTag1    = gKeyTag1    & "0"       & Chr(11)
	     Case "C"
                  gTypeCD2    = gTypeCD2    & iTypeCD & Chr(11)	     
                  gFieldCD2   = gFieldCD2   & "1"       & Chr(11)
                  gFieldNM2   = gFieldNM2   & "1"       & Chr(11)
                  gFieldLen2  = gFieldLen2  & "2"       & Chr(11)
                  gFieldType2 = gFieldType2 & "HH"      & Chr(11)
                  gDefaultT2  = gDefaultT2  & "L"       & Chr(11)
                  gNextSeq2   = gNextSeq2   & "0"       & Chr(11)
                  gKeyTag2    = gKeyTag2    & "0"       & Chr(11)
    End Select

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

Sub SetADOFeild(iSpdNo,iTypeCD)

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
              Case "B"
                     gTypeCD1    = gTypeCD1     & iTypeCD  & Chr(11)
                     gFieldCD1   = gFieldCD1    & "Error"  & Chr(11)
                     gFieldNM1   = gFieldNM1    & "N/A"    & Chr(11)
                     gFieldLen1  = gFieldLen1   & "30"     & Chr(11)
                     gFieldType1 = gFieldType1  & "ST"     & Chr(11)
                     gDefaultT1  = gDefaultT1   & "L"      & Chr(11)
                     gNextSeq1   = gNextSeq1    & "0"      & Chr(11)
                     gKeyTag1    = gKeyTag1     & "0"      & Chr(11)
              Case "C"
                     gTypeCD2    = gTypeCD2     & iTypeCD  & Chr(11)
                     gFieldCD2   = gFieldCD2    & "Error"  & Chr(11)
                     gFieldNM2   = gFieldNM2    & "N/A"    & Chr(11)
                     gFieldLen2  = gFieldLen2   & "30"     & Chr(11)
                     gFieldType2 = gFieldType2  & "ST"     & Chr(11)
                     gDefaultT2  = gDefaultT2   & "L"      & Chr(11)
                     gNextSeq2   = gNextSeq2    & "0"      & Chr(11)
                     gKeyTag2    = gKeyTag2     & "0"      & Chr(11)
     End Select 

End Sub

Sub SplitADOFieldVar(iSpdNo)
    
    Select Case  iSpdNo
	     Case "A"
                 gTypeCD     = Split (gTypeCD    ,Chr(11))                           
                 gFieldCD    = Split (gFieldCD   ,Chr(11))                           
                 gFieldNM    = Split (gFieldNM   ,Chr(11))                           
                 gFieldLen   = Split (gFieldLen  ,Chr(11))                           
                 gFieldType  = Split (gFieldType ,Chr(11))                           
                 gDefaultT   = Split (gDefaultT  ,Chr(11))                           
                 gNextSeq    = Split (gNextSeq   ,Chr(11))                           
                 gkeyTag     = Split (gKeyTag    ,Chr(11))                           
	     Case "B"
                 gTypeCD1    = Split (gTypeCD1   ,Chr(11))                           
                 gFieldCD1   = Split (gFieldCD1  ,Chr(11))                           
                 gFieldNM1   = Split (gFieldNM1  ,Chr(11))                           
                 gFieldLen1  = Split (gFieldLen1 ,Chr(11))                           
                 gFieldType1 = Split (gFieldType1,Chr(11))                           
                 gDefaultT1  = Split (gDefaultT1 ,Chr(11))                           
                 gNextSeq1   = Split (gNextSeq1  ,Chr(11))                           
                 gkeyTag1    = Split (gKeyTag1   ,Chr(11))                           
	     Case "C"
                 gTypeCD2    = Split (gTypeCD2   ,Chr(11))                           
                 gFieldCD2   = Split (gFieldCD2  ,Chr(11))                           
                 gFieldNM2   = Split (gFieldNM2  ,Chr(11))                           
                 gFieldLen2  = Split (gFieldLen2 ,Chr(11))                           
                 gFieldType2 = Split (gFieldType2,Chr(11))                           
                 gDefaultT2  = Split (gDefaultT2 ,Chr(11))                           
                 gNextSeq2   = Split (gNextSeq2  ,Chr(11))                           
                 gkeyTag2    = Split (gKeyTag2   ,Chr(11))                           
    End Select
    
End Sub

Sub SetCellTypeOfSpreadSheet(pObject,iCol,pFieldType,pFieldNM,pFieldLen)
   Dim iAlign
   
   ggoSpread.Source = pObject
   
   iAlign = Trim(Mid(pFieldType,3,1))
   
   If iAlign = "" Then
      Select Case Mid(pFieldType,1,1)
         Case "F"  : iAlign = "1"
         Case "TT" : iAlign = "2"
         Case Else : iAlign = "0"
      End Select   
   End If
   
   Select Case Mid(pFieldType,1,2) 
     Case "BT" 'Button
		    ggoSpread.SSSetButton iCol
     Case "CB" 'Combo
            ggoSpread.SSSetCombo  iCol , pFieldNM , pFieldLen,iAlign
     Case "CK" 'Check
            ggoSpread.SSSetCheck  iCol , pFieldNM , pFieldLen,iAlign, "", True, -1
     Case "DD"   '날짜 
            ggoSpread.SSSetDate   iCol , pFieldNM , pFieldLen,iAlign,gDateFormat
     Case "D5"   '편집(Year,Month)
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign
     Case "ED"   '편집 
'           ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign                         '2003/04/29 LEE JINSOO
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign             ,     , 200 '2003/04/29 LEE JINSOO
     Case "F2"  ' 금액 
            Call SetSpreadFloat  (iCol , pFieldNM , pFieldLen,iAlign,ggAmtOfMoneyNo)
     Case "F3"  ' 수량 
            Call SetSpreadFloat  (iCol , pFieldNM , pFieldLen,iAlign,ggQtyNo)
     Case "F4"  ' 단가 
            Call SetSpreadFloat  (iCol , pFieldNM , pFieldLen,iAlign,ggUnitCostNo)
     Case "F5"   ' 환율 
            Call SetSpreadFloat  (iCol , pFieldNM , pFieldLen,iAlign,ggExchRateNo)
     Case "MK"   ' Mask
            ggoSpread.SSSetMask   iCol , pFieldNM , pFieldLen,iAlign
     Case "ST"   ' Static
            ggoSpread.SSSetStatic iCol , pFieldNM , pFieldLen,iAlign
     Case "TT"   ' Time
            ggoSpread.SSSetTime   iCol , pFieldNM , pFieldLen,iAlign,1
	 Case "HH"	 ' Hidden
			ggoSpread.Source.Col	=	iCol
			ggoSpread.Source.ColHidden	=	True                        
     Case Else
'           ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign                          '2003/04/29 LEE JINSOO
            ggoSpread.SSSetEdit   iCol , pFieldNM , pFieldLen,iAlign             ,     , 200  '2003/04/29 LEE JINSOO
   End Select

End Sub


Function MakeSQLGroupOrderByList(pMaxColCnt,pPopUpR,pFieldCD,pNextSeq,pTypeCD,pMaxSelList)
    Dim iStr,jStr
    Dim ii,jj,kk      
    Dim tmpPopUpR   
    Dim iMark
    Dim iFirst
        
    Redim tmpPopUpR(pMaxSelList - 1)    
    
    For kk = 0 to pMaxSelList - 1
        tmpPopUpR(kk) = pPopUpR(kk,0)
    Next
    
    MakeSQLGroupOrderByList = "" : iStr   = ""  :   jStr   = ""      
    iFirst = "N"
    
    Redim  iMark(pMaxColCnt) 

    For ii = 0 to pMaxSelList - 1
        If tmpPopUpR(ii) <> "" Then     
           If pTypeCD = "G" Then
              For jj = 0 To pMaxColCnt - 1                                            
                  If iMark(jj) <> "X" Then
                     If pPopUpR(ii,0) = Trim(pFieldCD(jj)) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If                        
                        If CInt(Trim(pNextSeq(jj))) >= 1 And CInt(Trim(pNextSeq(jj))) <= pMaxColCnt Then
                           iStr = iStr & pPopUpR(ii,0) & " " & pPopUpR(ii,1) & "," & pFieldCD(CInt(pNextSeq(jj)) - 1)
                           jStr = jStr & pPopUpR(ii,0) & " " &                 "," & pFieldCD(CInt(pNextSeq(jj)) - 1)
                           
                           If (ii + 1) <  pMaxSelList   Then
                              For kk = ii + 1 to pMaxSelList - 1
                                  If  pPopUpR(kk,0) = Trim(pFieldCD(CInt(pNextSeq(jj)) - 1))  Then
                                      iStr = iStr & " " & pPopUpR(kk,1) 
                                      tmpPopUpR(kk) = ""
                                  End If    
                              Next    
                           End If                              
                           iMark(CInt(pNextSeq(jj)) - 1) = "X"
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
    
    If pTypeCD = "G" Then
       MakeSQLGroupOrderByList =  "Group By " & jStr  & " Order By " & iStr 
    Else
       If Trim(iStr) <> "" Then
          MakeSQLGroupOrderByList = "Order By " & iStr
       End If
    End If   

End Function

Sub MakePopData(pDefaultT,pFieldNM,pFieldCD,pPopUpR,pSortFieldNm,pSortFieldCD,pMaxSelList)
	Dim ii,kk	
	Dim iCast

    pSortFieldNm  = ""
    pSortFieldCD  = ""
    
    For ii = 0 To UBound(pFieldNM) - 1                                       
        iCast = pDefaultT(ii)
        If  IsNumeric(iCast) Or Trim(pDefaultT(ii)) = "V" Then
            If IsNumeric(iCast) Then 
               If IsBetween(1,pMaxSelList,CInt(iCast)) Then    'Sort정보default값 저장 
                  pPopUpR(CInt(pDefaultT(ii)) - 1,0) = Trim(pFieldCD(ii))
                  pPopUpR(CInt(pDefaultT(ii)) - 1,1) = "ASC"
               End If
            End If
            pSortFieldNm  = pSortFieldNm  & Trim(pFieldNM (ii)) & Chr(11)
            pSortFieldCD  = pSortFieldCD  & Trim(pFieldCD(ii))  & Chr(11)
        End If
    Next
    
    pSortFieldNm     = Split (pSortFieldNm ,Chr(11))
    pSortFieldCD     = Split (pSortFieldCD ,Chr(11))


End Sub

Sub CopyToTmpBuffer(pTypeCD,pFieldCD,pFieldNM,pFieldLen,pFieldType,pDefaultT,pNextSeq,pKeyTag)
    tmpTypeCD    = pTypeCD
    tmpFieldCD   = pFieldCD
    tmpFieldNM   = pFieldNM
    tmpFieldLen  = pFieldLen
    tmpFieldType = pFieldType
    tmpDefaultT  = pDefaultT
    tmpNextSeq   = pNextSeq
    tmpKeyTag    = pKeyTag    
End Sub


Function InitSpreadSheetFieldOfZADO(pObject,pPopUpR,pSelectList,pSelectListDT,pKeyPos,pMaxKey,pMaxSelList)

    Dim iMark
    Dim iSeq
    
    Redim  iMark(UBound(tmpFieldNM)) 

    For ii = 0 to pMaxSelList - 1
        For jj = 0 to UBound(tmpFieldNM) - 1
            If iMark(jj) <> "X" Then
               If pPopUpR(ii,0) = Trim(tmpFieldCD(jj)) Then
                  iSeq = iSeq + 1

                  Call InitSpreadSheetRow(pObject,iSeq,jj,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                  
                  If IsBetween(1,UBound(tmpFieldNM),CInt(tmpNextSeq(jj))) Then 
                     kk = CInt(tmpNextSeq(jj)) 
                     iSeq = iSeq + 1
                     Call InitSpreadSheetRow(pObject,iSeq,kk-1,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                  End If    
               End If 
            End If 
        Next      
    Next      
         
    For ii = 0 to UBound(tmpFieldNM) - 1
        If iMark(ii) <> "X" Then
           If tmpTypeCD(0) = "S" Or (tmpTypeCD(0) = "G" And tmpDefaultT(ii) = "L") Then
              iSeq = iSeq + 1
              
              Call InitSpreadSheetRow(pObject,iSeq,ii,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
              
              If IsBetween(1,UBound(tmpFieldNM),CInt(tmpNextSeq(ii))) Then 
                 kk = CInt(tmpNextSeq(ii)) 
                 iSeq = iSeq + 1
                 Call InitSpreadSheetRow(pObject,iSeq,kk-1,iMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)
                 
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
Sub InitSpreadSheetRow(pObject,iCol,iDx,pMark,pSelectList,pSelectListDT,pKeyPos,pMaxKey)

   pMark(iDx) = "X"

   Call SetCellTypeOfSpreadSheet(pObject,iCol,tmpFieldType(iDx),tmpFieldNM(iDx),tmpFieldLen(iDx))
   
   pSelectList   = pSelectList   & tmpFieldCD(iDx)            & ","      'Make select part of SQL
   pSelectListDT = pSelectListDT & Mid(tmpFieldType(iDx),1,2) & gColSep
   
   If CInt(tmpKeyTag(iDx)) > 0 And CInt(tmpKeyTag(iDx)) <= pMaxKey Then  
     pKeyPos(CInt(tmpKeyTag(iDx))) =  iCol
	End If
End Sub