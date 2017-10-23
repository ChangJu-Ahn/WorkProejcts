<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : ADO Template (Query)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/16
'*  7. Modified date(Last)  : 2002/12/16
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%  
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "*", "NOCOOKIE", "QB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT

Dim TmpBuffer
Dim iTotalStr

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPlantCd																'⊙ : 공장 
Dim strPltCd
Dim strResourceCd																'⊙ : 자원 
Dim strShiftCd				
Dim iOpt											'⊙ : Shift
Dim pPB6S101
Dim ADF
Dim strRetMsg

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Dim flgPlantCd
Dim flgShiftCd
flgPlantCd = True
flgShiftCd = True

Call HideStatusWnd
strPltCd = Request("txtPlantCd")

On Error Resume Next
Err.Clear

'======================================================================================================
'	조건컬럼의 이름 처리해주는 부분 
'======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 1)
	
UNISqlId(0) = "180000san"
UNISqlId(1) = "122700sab"
UNISqlId(2) = "180000sao"
	
strResourceCd = FilterVar(Request("txtResourceCd"), "''", "S")
strPlantCd = FilterVar(Request("txtPlantCd"),"''","S")

IF Request("txtShiftCd") = "" Then
   strShiftCd = "|"
ELSE
   strShiftCd = FilterVar(UCase(Request("txtShiftCd")),"''","S")
END IF

UNIValue(0, 0) = strPlantCd
UNIValue(0, 1) = strResourceCd
UNIValue(1, 0) = strPlantCd
UNIValue(2, 0) = strPlantCd
UNIValue(2, 1) = strShiftCd
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

Response.Write "<Script Language=VBScript>" & vbCrLf
If rs0.EOF And rs0.BOF Then
	Response.Write "parent.frm1.txtResourceNm.value = """"" & vbCrLf			'☜: 화면 처리 ASP 를 지칭함 
Else
	Response.Write "parent.frm1.txtResourceNm.value = """ & ConvSPChars(rs0("Description")) & """" & vbCrLf		'☜: 화면 처리 ASP 를 지칭함 
End If
	
If rs1.EOF And rs1.BOF Then
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf		'☜: 화면 처리 ASP 를 지칭함 
	flgPlantCd = False
Else
	Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgPlantCd = True
End If
	
If rs2.EOF And rs2.BOF Then
	Response.Write "parent.frm1.txtShiftNm.value = """"" & vbCrLf		'☜: 화면 처리 ASP 를 지칭함 
	flgShiftCd = False
Else
	Response.Write "parent.frm1.txtShiftNm.value = """ & ConvSPChars(rs2("Description")) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgShiftCd = True
End If
Response.Write "</Script>" & vbCrLf

If flgPlantCd = False Then
	Call DisplayMsgBox(125000, vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1.txtPlantCd.focus() " & vbCrLf
	Response.Write "</Script>" & vbCr
	Response.End
End If

If flgShiftCd = False And strShiftCd <> "|" Then
	Call DisplayMsgBox(180400, vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1.txtShiftCd.focus() " & vbCrLf
	Response.Write "</Script>" & vbCr
	Response.End
End if

rs0.Close
rs1.Close
rs2.Close
		
Set rs0 = Nothing
Set rs1 = Nothing
Set rs2 = Nothing

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
								'☜: ActiveX Data Factory Object Nothing

'==================================================================================================

lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
lgMaxCount     = 30							                           '☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

Call TrimData()
Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

 Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    ReDim TmpBuffer(0)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)
               Case Else
                    iStr = iStr & Chr(11) & rs0(ColCnt) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
			ReDim Preserve TmpBuffer(iRCnt)
            TmpBuffer(iRCnt) = iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotalStr = Join(TmpBuffer, "")
	
    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,5)

    UNISqlId(0) = "181900saa"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = UCase(Trim(strPlantCd))
    UNIValue(0,2) = UCase(Trim(strPlantCd))
    UNIValue(0,3) = UCase(Trim(strResourceCd))
    UNIValue(0,4) = UCase(Trim(strShiftCd))
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("181900", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	 strPlantCd		= FilterVar(Request("txtPlantCd"),"' '","S")	 
     strResourceCd  = FilterVar(Request("txtResourceCd"),"' '","S")                   '자원 
     strShiftCd     = FilterVar(Request("txtShiftCd"),"' '","S")                   '자원그룹 
      iOpt			= Request("iOpt")
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowDataByClip  "<%=ConvSPChars(iTotalStr)%>"                          '☜: Display data 
         .lgStrPrevKey_A      = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
         .DbQueryOk("<%=iOpt%>")
	End with
</Script>	
