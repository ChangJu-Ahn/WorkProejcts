<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111RA2
'*  4. Program Name         : 이전매출채권참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim BlankchkFlg

Dim iFrPoint
iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

    lgPageNo   = Request("lgPageNo")                               '☜ : Next key flag
    lgMaxCount     = 30							                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
	on error resume next 
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt

    lgstrData = ""

	If IsNumeric(lgPageNo) Then 
		If CLng(lgPageNo) > 0 Then
		   rs0.Move = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		End If
	Else
		lgPageNo = 0
	End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

       If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
 
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim iStrVal
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "S5111ra201"									'* : 데이터 조회를 위한 SQL문 
 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	If Len(Request("txtSoldtoParty")) Then
	    UNISqlId(1) = "s0000qa002"
	    UNIValue(1,0) = FilterVar(Trim(Request("txtSoldtoParty")), "''", "S")
		iStrVal = " AND A.SOLD_TO_PARTY = " & FilterVar(Trim(Request("txtSoldtoParty")), "''", "S") & ""
	Else
		iStrVal = ""
	End If

	If Len(Request("txtBillToParty")) Then
	    UNISqlId(2) = "s0000qa002"
	    UNIValue(2,0) = FilterVar(Trim(Request("txtBillToParty")), "''", "S")
		iStrVal =  iStrVal & " AND A.BILL_TO_PARTY = " & FilterVar(Trim(Request("txtBillToParty")), "''", "S") & ""
	End If

	If Len(Request("txtSalesGrp")) Then
	    UNISqlId(3) = "s0000qa005"
	    UNIValue(3,0) = FilterVar(Trim(Request("txtSalesGrp")), "''", "S")
		iStrVal =  iStrVal & " AND A.SALES_GRP = " & FilterVar(Trim(Request("txtSalesGrp")), "''", "S") & ""
	End If

    If Len(Trim(Request("txtBillFrDt"))) Then
		iStrVal = iStrVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Request("txtBillFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtBillToDt"))) Then
		iStrVal = iStrVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Request("txtBillToDt")), "''", "S") & ""		
	End If

    UNIValue(0,1) = iStrVal
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    

	Call BeginScriptTag()

	'주문처 존재여부 
	If Trim(Request("txtSoldtoParty")) <> "" Then
		If rs1.EOF And rs1.BOF Then
			Call CloseAdoObject(rs1)
			Call WriteConDesc("txtSoldtoPartyNm", "")
			Call ConNotFound("txtSoldtoParty")			
			Exit Sub
		Else
			Call WriteConDesc("txtSoldtoPartyNm", rs1(1))		
			Call CloseAdoObject(rs1)
		End If
	Else
		Call WriteConDesc("txtSoldtoPartyNm", "")
	End If

	' 발행처 존재여부 
	If Trim(Request("txtBillToParty")) <> "" Then
		If rs2.EOF And rs2.BOF Then
			Call CloseAdoObject(rs2)
			Call WriteConDesc("txtBillToPartyNm", "")
			Call ConNotFound("txtBillToParty")			
			Exit Sub
		Else	
			Call WriteConDesc("txtBillToPartyNm", rs2(1))		
			Call CloseAdoObject(rs2)
		End If
	Else
		Call WriteConDesc("txtBillToPartyNm", "")
	End If

	' 영업그룹 존재여부 
	If Trim(Request("txtSalesGrp")) <> "" Then
		If rs3.EOF And rs3.BOF Then
			Call CloseAdoObject(rs3)
			Call WriteConDesc("txtSalesGrpNm", "")
			Call ConNotFound("txtSalesGrp")			
			Exit Sub
		Else	
			Call WriteConDesc("txtSalesGrpNm", rs3(1))		
			Call CloseAdoObject(rs3)
		End If
	Else
		Call WriteConDesc("txtSalesGrpNm", "")
	End If

    If  rs0.EOF And rs0.BOF Then	
		Call CloseAdoObject(rs0)
        Call DataNotFound("txtSoldToParty")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
		Call CloseAdoObject(rs0)
		If lgPageNo = "1" Then Call SetConditionData()
        Call WriteResult()
    End If

End Sub

' Recordset 객체 Release
Sub CloseAdoObject(ByRef prObjRs)
	If VarType(prObjRs) <> vbObject Then Exit Sub
	
    If Not (prObjRs Is Nothing) Then
       If prObjRs.State = 1 Then		' adStateOpen
          prObjRs.Close
       End If
       Set prObjRs = Nothing
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHSoldToParty.value	= """ & ConvSPChars(Request("txtSoldToParty")) & """" & vbCr
	Response.Write ".txtHBillToParty.value	= """ & ConvSPChars(Request("txtBillToParty")) & """" & vbCr
	Response.Write ".txtHBillFrDt.value = """ & Request("txtBillFrDt") & """" & vbCr
	Response.Write ".txtHBillToDt.value	= """ & Request("txtBillToDt") & """" & vbCr
	Response.Write ".txtHSalesGrp.value	= """ & ConvSPChars(Request("txtSalesGrp")) & """" & vbCr
	Response.Write "End with" & vbCr
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성(조회조건 포함)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
 	Response.Write ".vspdData.Redraw = False " & vbCr      
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData  & """ ,""F""" & vbCr
	Response.Write "parent.lgPageNo	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
 	Response.Write ".vspdData.Redraw = True " & vbCr      
	Response.Write "End with" & vbCr
	Call EndScriptTag()
End Sub
%>
