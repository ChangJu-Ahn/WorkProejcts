<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4116PA4
'*  4. Program Name         : 출고상세현황 
'*  5. Program Desc         : 출고상세현황 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "RB")
On Error Resume Next

Call HideStatusWnd

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data

Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList													       '☜ : select 대상목록 
Dim lgSelectListDT														   '☜ : 각 필드의 데이타 타입	
Dim lgPageNo

Const C_SHEETMAXROWS_D  = 30   

Dim iStrDNNo
Dim iStrDNType
Dim iStrShipToParty
Dim iStrFromDt
Dim iStrToDt
Dim iStrConfFlag

iStrDNNo = Request("txtConDNNo")
iStrDNType = Request("txtConDnType")
iStrShipToParty = Request("txtConShipToParty")
iStrFromDt = Request("txtConFromDt")
iStrToDt = Request("txtConToDt")
iStrConfFlag = Request("txtConConfFlag")

lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Order by value

Call TrimData()
Call FixUNISQLData()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim iStrVal					
								
    Redim UNISqlId(0)           
								
    Redim UNIValue(0,2)			

    UNISqlId(0) = "S4116PA401" 
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	iStrVal = " "
   
	
	'조회기간시작=========================================================================================
	If Len(iStrFromDt) Then
		If iStrConfFlag = "Y" Then
			iStrVal = iStrVal & " DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""		
		Else
			iStrVal = iStrVal & " DH.PROMISE_DT >=  " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""		
		End If
	End If		
	
	'조회기간끝===========================================================================================
	If Len(iStrToDt) Then
		If iStrConfFlag = "Y" Then
			iStrVal = iStrVal & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""		
		Else 
			iStrVal = iStrVal & " AND DH.PROMISE_DT <=  " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""		
		End If
	End If
		
    '출하번호=============================================================================================    	
	If Len(iStrDNNo) Then
		iStrVal = iStrVal & " AND DH.DN_NO =  " & FilterVar(iStrDNNo, "''", "S") & ""				
	End If
	
	'출하형태=============================================================================================    	
	If Len(iStrDNType) Then
		iStrVal = iStrVal & " AND DH.MOV_TYPE =  " & FilterVar(iStrDNType, "''", "S") & ""				
	End If
	
	'납품처=============================================================================================    	
	If Len(iStrShipToParty) Then
		iStrVal = iStrVal & " AND DH.SHIP_TO_PARTY =  " & FilterVar(iStrShipToParty, "''", "S") & ""				
	End If	 
	
	'확정여부===========================================================================================
	If Len(iStrConfFlag) Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =  " & FilterVar(iStrConfFlag , "''", "S") & ""		
	End If
		
    UNIValue(0,1) = iStrVal   
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'☜:ADO 객체를 생성 
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End     
    Else   
        Call  MakeSpreadSheetData()
    End If
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>
<Script Language=vbscript>
    With parent  
        .ggoSpread.Source = .frm1.vspdData
        .frm1.vspdData.Redraw = False
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
        .lgPageNo = "<%=lgPageNo%>"		
		.frm1.vspdData.Redraw = True
		Call .DbQueryOk
   	End with
</Script>	 	
<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
