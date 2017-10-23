<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Query)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/15
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
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPlantCd																'⊙ : 공장 
Dim strPltCd
Dim strResourceCd																'⊙ : 자원 
Dim strResourceGroupCd1																'⊙ : 자원그룹 
Dim strResourceGroupCd2																'⊙ : 자원그룹 
Dim strToDt1
Dim strToDt2
Dim ADF
Dim iOpt

Dim TmpBuffer
Dim iTotalStr

Dim strRetMsg
Dim pPB6S101

Dim flgPlantCd
Dim flgResourceGroup

flgPlantCd = True	
flgResourceGroup = True

Dim R1_P_Plant
Const EA_b_plant_plant_cd = 0
Const EA_b_plant_plant_nm = 1

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
Call HideStatusWnd 

strPltCd = Trim(Request("txtPlantCd"))

On Error Resume Next
Err.Clear

'======================================================================================================
'	조건컬럼의 이름 처리해주는 부분 
'======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 1)
	
UNISqlId(0) = "180000san"
UNISqlId(1) = "122700sab"
UNISqlId(2) = "181800sad"
	
	
strResourceCd = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "
strPlantCd = FilterVar(UCase(Request("txtPlantCd")),"''","S")

IF Request("txtResourceGroupCd1") = "" Then
   strResourceGroupCd1 = "|"
ELSE
   strResourceGroupCd1 = FilterVar(UCase(Request("txtResourceGroupCd1")),"''","S")
END IF

UNIValue(0, 0) = strPlantCd
UNIValue(0, 1) = strResourceCd
UNIValue(1, 0) = strPlantCd
UNIValue(2, 0) = strResourceGroupCd1
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

Response.Write "<Script Language=VBScript>" & vbCrLf
If rs0.EOF And rs0.BOF Then
	Response.Write "parent.frm1.txtResourceNm.value = """"" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
Else
	Response.Write "parent.frm1.txtResourceNm.value = """ & ConvSPChars(rs0("Description")) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
End If

If rs1.EOF And rs1.BOF Then
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgPlantCd = False
Else
	Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgPlantCd =True
End If
	
If rs2.EOF And rs2.BOF Then
	Response.Write "parent.frm1.txtResourceGroupNm.value = """"" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgResourceGroup = False
Else
	Response.Write "parent.frm1.txtResourceGroupNm.value = """ & ConvSPChars(rs2(0)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	flgResourceGroup = True
End If
Response.Write "</Script>" & vbCrLf

If flgPlantCd = False Then
	Call DisplayMsgBox(125000, vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1.txtPlantCd.focus() " & vbCrLf
	Response.Write "</Script>" & vbCr
	Response.End
End If

If flgResourceGroup = False And strResourceGroupCd1 <> "|" Then
	Call DisplayMsgBox(181704, vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1.txtResourceGroupCd.focus() " & vbCrLf
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
'======================================================================================================
'	품목이름 처리해주는 부분 
'======================================================================================================

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
                           '소숫점 이하 표시하지 않기 위해 수정(p1503qa1도 같이 수정)
                           iStr = iStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(ColCnt), 0)
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
            TmpBuffer(iRCnt)      = iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
  	iTotalStr = Join(TmpBuffer, "")
  	
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

    Redim UNIValue(0,7)

    UNISqlId(0) = "181800saa"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = UCase(Trim(strPlantCd))
    UNIValue(0,2) = UCase(Trim(strResourceCd))
    UNIValue(0,3) = UCase(Trim(strResourceGroupCd1))
    UNIValue(0,4) = UCase(Trim(strResourceGroupCd2))
    UNIValue(0,5) = UCase(Trim(strToDt1))
    UNIValue(0,6) = UCase(Trim(strToDt2))
    
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
        Call DisplayMsgBox(900014, vbOKOnly, "", "", I_MKSCRIPT)
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
	 strPlantCd    = FilterVar(Request("txtPlantCd"),"' '","S")	 
     strResourceCd     = FilterVar(Request("txtResourceCd"),"' '","S")                   '자원 
     strResourceGroupCd1     = FilterVar(Request("txtResourceGroupCd1"),"' '","S")                   '자원그룹 
     strResourceGroupCd2     = FilterVar(Request("txtResourceGroupCd2"),"' '","S")                   '자원그룹 
     strToDt1		= FilterVar(UniConvDate(Request("txtToDt1")),"' '","S")
     strToDt2		= FilterVar(UniConvDate(Request("txtToDt2")),"' '","S")  
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

