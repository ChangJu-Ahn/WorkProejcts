<%
'==============================================================================
' Name  :  GetAdoFiledInf
' Desc  :  Get Information for Creation of Query Window using ADO
'==============================================================================

Sub GetAdoFiledInf(ByVal iPgmId,ByVal iTypeCD,ByVal iSpdNo)
    Dim ADF                                                                    '☜ : declaration Variable indicating ActiveX Data Factory
    Dim lgstrRetMsg                                                            '☜ : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : declaration DBAgent Parameter 
	Dim iSaveType
	
	On Error Resume Next
    
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
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

	UNISqlId(0)    = "z100001"
	UNIValue(0, 0) = iPgmId
	UNIValue(0, 1) = iTypeCD
	UNIValue(0, 2) = iSpdNo
	UNIValue(0, 3) = gLang
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	If Trim(gDsnNo) ="" Then
        Call SetADOFeild(iSpdNo,iTypeCD)
        Exit Sub
	End If
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)	

	If rs0.EOF And rs0.BOF Then
        Call SetADOFeild(iSpdNo,iTypeCD)
 		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing			
		Exit Sub
	End If
	
    If rs0("TYPE_CD") = "G" Then
	   gMethodText = "집계"
	Else   
	   gMethodText = "정렬"
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
       iSaveType = UCASE(rs0("TYPE_CD"))
       rs0.MoveNext
	Loop

       'dumy Data For Export
	   Select Case  iSpdNo
	     Case "A"
                  gTypeCD     = gTypeCD     & iSaveType & Chr(11)	     
                  gFieldCD    = gFieldCD    & "1"       & Chr(11)
                  gFieldNM    = gFieldNM    & "1"       & Chr(11)
                  gFieldLen   = gFieldLen   & "2"       & Chr(11)
                  gFieldType  = gFieldType  & "HH"      & Chr(11)
                  gDefaultT   = gDefaultT   & "L"       & Chr(11)
                  gNextSeq    = gNextSeq    & "0"       & Chr(11)
                  gKeyTag     = gKeyTag     & "0"       & Chr(11)
	     Case "B"
                  gTypeCD1    = gTypeCD1    & iSaveType & Chr(11)	     
                  gFieldCD1   = gFieldCD1   & "1"       & Chr(11)
                  gFieldNM1   = gFieldNM1   & "1"       & Chr(11)
                  gFieldLen1  = gFieldLen1  & "2"       & Chr(11)
                  gFieldType1 = gFieldType1 & "HH"      & Chr(11)
                  gDefaultT1  = gDefaultT1  & "L"       & Chr(11)
                  gNextSeq1   = gNextSeq1   & "0"       & Chr(11)
                  gKeyTag1    = gKeyTag1    & "0"       & Chr(11)
	     Case "C"
                  gTypeCD2    = gTypeCD2    & iSaveType & Chr(11)	     
                  gFieldCD2   = gFieldCD2   & "1"       & Chr(11)
                  gFieldNM2   = gFieldNM2   & "1"       & Chr(11)
                  gFieldLen2  = gFieldLen2  & "2"       & Chr(11)
                  gFieldType2 = gFieldType2 & "HH"      & Chr(11)
                  gDefaultT2  = gDefaultT2  & "L"       & Chr(11)
                  gNextSeq2   = gNextSeq2   & "0"       & Chr(11)
                  gKeyTag2    = gKeyTag2    & "0"       & Chr(11)
       End Select
	Set rs0 = Nothing
	Set ADF = Nothing			
End Sub

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

Function FormatRsString(ByVal pSelectListDT,ByVal pRsValue)

     Select Case  pSelectListDT
        Case "D5"   :  If Len(Trim(pRsValue)) > 4 Then    'D5는 DB에YYYYMM(type이 string 형식)으로 들어 있을때 Company형식으로 바꾸어주는 역할을 한다.
                          FormatRsString = MakeDateTo("YM" ,gDateFormatYYYYMM ,gComDateType ,Mid(pRsValue,1,4),Mid(pRsValue,5,2),"00")    '날짜 '2003/01/27 구주요청 
                       Else
                          FormatRsString = ""
                       End If   
        Case "DD"   :  FormatRsString = UNIDateClientFormat(pRsValue)                            '날짜 
        Case "F2"   :  FormatRsString = UNINumClientFormat(pRsValue, ggAmtOfMoney.DecPoint, 0)  '금액        
        Case "F3"   :  FormatRsString = UNINumClientFormat(pRsValue, ggQty.DecPoint       , 0)  '수량 
        Case "F4"   :  FormatRsString = UNINumClientFormat(pRsValue, ggUnitCost.DecPoint  , 0)  '단가 
        Case "F5"   :  FormatRsString = UNINumClientFormat(pRsValue, ggExchRate.DecPoint  , 0)  '환율 
        Case "F6"   :  FormatRsString = UNINumClientFormat(pRsValue, gUDF6                , 0)  '사용자 정의 2003/02/10
        Case "F7"   :  FormatRsString = UNINumClientFormat(pRsValue, gUDF7                , 0)  '사용자 정의 2003/02/10
        Case "F8"   :  FormatRsString = UNINumClientFormat(pRsValue, gUDF8                , 0)  '사용자 정의 2003/02/10
        Case "F9"   :  FormatRsString = UNINumClientFormat(pRsValue, gUDF9                , 0)  '사용자 정의 2003/02/10

' 2003-02-19 Kim In Tae
        Case "FA"   :  FormatRsString = UniConvNumDBToCompanyWithOutChange(pRsValue, 0)  '금액 2003/03/06
        Case "FB"   :  FormatRsString = UniConvNumDBToCompanyWithOutChange(pRsValue, 0)  '수량 2003/03/06
        Case "FC"   :  FormatRsString = UniConvNumDBToCompanyWithOutChange(pRsValue, 0)  '단가 2003/03/06
        Case "FD"   :  FormatRsString = UniConvNumDBToCompanyWithOutChange(pRsValue, 0)  '환율 2003/03/06

'        Case "FA"   :  FormatRsString = UNINumClientFormat(pRsValue, ggAmtOfMoneyDefault.DecPoint , 0)  '금액 default 2003/02/26
'        Case "FB"   :  FormatRsString = UNINumClientFormat(pRsValue, ggQty.DecPointDefault        , 0)  '수량 default 2003/02/26
'        Case "FC"   :  FormatRsString = UNINumClientFormat(pRsValue, ggUnitCostDefault.DecPoint   , 0)  '단가 default 2003/02/26
'        Case "FD"   :  FormatRsString = UNINumClientFormat(pRsValue, ggExchRateDefault.DecPoint   , 0)  '환율 default 2003/02/26
        
        Case Else   :  FormatRsString = ConvSPChars(pRsValue)
    End Select
    
End Function

%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        