<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2313QB1
'*  4. Program Name         : 불량유형조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPlantCd                                               '   공장 
Dim strDtFr       	                                     '   기간(From)
Dim strDtTo		  				'   기간(From)
Dim strInspReqNo                                        '   검사의뢰번호 
Dim strLotNo						'  로트번호 
Dim strItemCd                                             '   품목 
Dim strInspItemCd					'검사항목 

Dim FilterPlantCd
Dim FilterDtFr
Dim FilterDtTo
Dim FilterInspReqNo
Dim FilterLotNo
Dim FilterItemCd
Dim FilterInspItemCd

Dim strFlag

'Header의 Name부분에 대한 변수 
Dim strPlantNm
Dim strItemNm
Dim strInspItemNm
Dim iOpt
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
Call HideStatusWnd 

lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
iOpt		= Request("iOpt")

Call TrimData()
Call  HeaderData()                                                '☜ : Header의 Name부분 불러오기 
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
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
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
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub HeaderData()
	Dim iStr
	
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(0,0)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNISqlId(0) = "Q2313QA121"
	UNIValue(0,0) = FilterPlantCd		'---공장 
	
    	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  	iStr = Split(lgstrRetMsg,gColSep)
    
    	If iStr(0) <> "0" Then
        		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    	End If    
        
    	If  rs0.EOF And rs0.BOF Then
        		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        		rs0.Close
        		Set rs0 = Nothing
        		Response.End													'☜: 비지니스 로직 처리를 종료함 
    	Else    
        		strPlantNm=rs0(0)
        		rs0.Close
        		Set rs0 = Nothing
    	End If
    	
	'품목명 
	If strItemCd <> "" Then
		UNISqlId(0) = "Q2313QA122"
		UNIValue(0,0) = FilterItemCd		'---품목 
		
    		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'☜: 비지니스 로직 처리를 종료함 
    		Else    
        			strItemNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
		
	'검사항목 
	If strInspItemCd <> "" Then
		Redim UNIValue(0,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
		
		UNISqlId(0) = "Q2313QA123"
		UNIValue(0,0) = FilterPlantCd		'---공장 
		UNIValue(0,1) = FilterItemCd		'---품목 
		UNIValue(0,2) = FilterInspItemCd	'---검사항목 
		
		lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  		iStr = Split(lgstrRetMsg,gColSep)
    
    		If iStr(0) <> "0" Then
        			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    		End If    
        
    		If  rs0.EOF And rs0.BOF Then
        			Call DisplayMsgBox("220201", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        			rs0.Close
        			Set rs0 = Nothing
        			Response.End													'☜: 비지니스 로직 처리를 종료함 
    		Else    
        			strInspItemNm=rs0(0)
        			rs0.Close
        			Set rs0 = Nothing
    		End If
	End If
	
	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------	
     	
End Sub

Sub FixUNISQLData()
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Select Case strFlag
		Case "N"
			Redim UNIValue(0,6)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
	                                                                      			'    parameter의 수에 따라 변경함 
			UNISqlId(0) = "Q2313QA101"
		Case "I"
			Redim UNIValue(0,7)    
			UNISqlId(0) = "Q2313QA102"
		Case "A"
			Redim UNIValue(0,9)             
			UNISqlId(0) = "Q2313QA103"
	End Select
	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = FilterPlantCd		'---공장 
    UNIValue(0,2) = FilterDtFr			'---기간 
    UNIValue(0,3) = FilterDtTo
	UNIValue(0,4) = FilterInspReqNo		'---검사의뢰번호 
    UNIValue(0,5) = FilterLotNo		'---로트번호 
	
	Select Case strFlag
	    	Case "N"
	
	    	Case "I"
	    		 UNIValue(0,6) = FilterItemCd					'---품목 
	    	Case "A"
	    		UNIValue(0,6) = FilterItemCd					'---품목 
	    		UNIValue(0,7) = FilterInspItemCd	    			'---검사항목 
	    End Select
	    
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)		'---	Sort By 조건 

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
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
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
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strPlantCd = Request("txtPlantCd")
    strDtFr = UNIConvDate(Request("txtDtFr"))
	strDtTo = UNIConvDate(Request("txtDtTo"))
	strInspReqNo = Request("txtInspReqNo")
	strLotNo = Request("txtLotNo")
	strItemCd = Request("txtItemCd")
	strInspItemCd = Request("txtInspItemCd")
	
    FilterPlantCd  = FilterVar(strPlantCd, "''", "S")
    FilterDtFr =FilterVar(strDtFr, "''", "S")
    FilterDtTo =FilterVar(strDtTo, "''", "S")
	FilterInspReqNo = FilterVar(strInspReqNo, "''", "S")
	FilterLotNo = FilterVar(strLotNo, "''", "S")
	FilterItemCd = FilterVar(strItemCd, "''", "S")
	FilterInspItemCd = FilterVar(strInspItemCd, "''", "S")
		
	If strItemCd = "" And strInspItemCd = "" Then
		strFlag = "N"
	ElseIf strItemCd <> "" And strInspItemCd = "" Then
		strFlag = "I"
	ElseIf strItemCd <> "" And strInspItemCd <> "" Then
		strFlag = "A"
	End If	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
%>
<Script Language=vbscript>
    With parent
		'헤더데이타 Display
		.frm1.txtPlantNm.Value = "<%=ConvSPChars(strPlantNm)%>"
		.frm1.txtItemNm.Value = "<%=ConvSPChars(strItemNm)%>"
		.frm1.txtInspItemNm.Value = "<%=ConvSPChars(strInspItemNm)%>"
			
		'Detail Data Display
         .ggoSpread.Source = .frm1.vspdData 
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                          '☜: Display data 
         .lgStrPrevKey_A = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
         .DbQueryOk("<%=iOpt%>")
	End with
</Script>	
<%
Response.End													'☜: 비지니스 로직 처리를 종료함 
%>