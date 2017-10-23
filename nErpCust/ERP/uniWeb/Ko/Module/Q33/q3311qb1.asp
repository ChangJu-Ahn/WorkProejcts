<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Quality
'*  2. Function Name        : ADO  (QUERY)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/27
'*  7. Modified date(Last)  : 2004/08/05
'*  8. Modifier (First)     : Koh Jae Woo
'*  9. Modifier (Last)      : Lee Seung Wook
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgStrPrevKey                                            '☜ : 이전 값 
Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
Dim strPlantCd
Dim strDtFr
Dim strDtTo
Dim strItemCd
Dim strRoutNo
Dim strOprNo

Dim FilterPlantCd
Dim FilterDtFr
Dim FilterDtTo
Dim FilterItemCd
Dim FilterRoutNo
Dim FilterOprNo

Dim strFlag

'Header의 Name부분에 대한 변수 
Dim strPlantNm
Dim strItemNm
Dim strDefectRatioUnit
Dim strLotRejUnit
'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

     Call HideStatusWnd 
     lgStrPrevKey     = Request("lgStrPrevKey")                           '☜ : Next key flag
     lgMaxCount       = CInt(Request("lgMaxCount"))                       '☜ : 한번에 가져올수 있는 데이타 건수 
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
     
     Call  TrimData()                                                     '☜ : Parent로 부터의 데이타 가공 
     Call  HeaderData()                                                   '☜ : Header의 Name부분 불러오기 
     Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
     Call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iCnt
    Dim iRCnt                                                                     
    Dim strTmpBuffer                                                              
    Dim iStr
    Dim ColCnt
     
    iCnt = 0
    lgstrData = ""
   
    If Len(Trim(lgStrPrevKey)) Then                                              '☜ : Chnage str into int
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                         '☜ : Discard previous data
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
               Case "F6"   '불량률, 불합격률 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), 2, 0)
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

    If  iRCnt < lgMaxCount Then                                     '☜: Check if next data exists
        lgStrPrevKey = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
    Set lgADF = Nothing                                             '☜: Release ADF

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
	UNISqlId(0) = "Q3311QA121"
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
		UNISqlId(0) = "Q3311QA122"
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
	
		
	'불량률 
	UNISqlId(0) = "Q3311QA124"
	UNIValue(0,0) = FilterPlantCd		'---공장 

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	If  rs0.EOF And rs0.BOF Then
			Call DisplayMsgBox("220401", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
			rs0.Close
			Set rs0 = Nothing
			Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else    
			strDefectRatioUnit=rs0(0)
			rs0.Close
			Set rs0 = Nothing
	End If
	
	'LOT불합격률 단위 
	strLotRejUnit = "%"
	
	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------	
     	
End Sub

Sub FixUNISQLData()
	Dim strQryCnd
	Redim UNISqlId(0) 
	Redim UNIValue(0,5)
	
	UNISqlId(0) = "Q3311QA101"
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	'Select Case strFlag
	'	Case "N"
	'		Redim UNIValue(0,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
	                                                                      			'    parameter의 수에 따라 변경함 
	'		UNISqlId(0) = "Q3311QA101"
	'	Case "I"
	'		Redim UNIValue(0,5)    
	'		UNISqlId(0) = "Q3311QA102"
	'	Case "W"
	'		Redim UNIValue(0,5)    
	'		UNISqlId(0) = "Q3311QA103"
	'	Case "A"
	'		Redim UNIValue(0,6)    
	'		UNISqlId(0) = "Q3311QA104"
	'End Select
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = FilterPlantCd		'---공장 
    UNIValue(0,2) = FilterDtFr			'---기간 
    UNIValue(0,3) = FilterDtTo
    
    If Trim(strItemCd) <> "" Then
		strQryCnd = "AND A.ITEM_CD = " & FilterItemCd
	End If
	
	If Trim(strRoutNo) <> "" Then
		strQryCnd = strQryCnd & "AND A.ROUT_NO = " & FilterRoutNo
	End If
	
	If Trim(strOprNo) <> "" Then
		strQryCnd = strQryCnd & "AND A.OPR_NO = " & FilterOprNo
	End If
	
	UNIValue(0,4) = (Trim(strQryCnd))
	
	'Select Case strFlag
	'    Case "N"
	
	'    Case "I"
	'    	 UNIValue(0,4) = FilterItemCd					'---품목 
	'    Case "W"
	'    	 UNIValue(0,4) = FilterWcCd	    					'---작업장 
	'    Case "A"
	'    	 UNIValue(0,4) = FilterItemCd					'---품목 
	'    	 UNIValue(0,5) = FilterWcCd	    					'---작업장 
	'End Select
     
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Group By 조건 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
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
    strDtFr = Request("txtDtFr")
	strDtTo = Request("txtDtTo")
	strItemCd = Request("txtItemCd")
	strRoutNo = Request("txtRoutNo")
	strOprNo = Request("txtOprNo")
	
	FilterPlantCd  = FilterVar(strPlantCd, "''", "S")
    FilterDtFr =FilterVar(strDtFr, "''", "S")
	FilterDtTo = FilterVar(strDtTo, "''", "S")
	FilterItemCd = FilterVar(strItemCd, "''", "S")
	FilterRoutNo = FilterVar(strRoutNo, "''", "S")
	FilterOprNo = FilterVar(strOprNo, "''", "S")
	
		
	'If strItemCd = "" And strWcCd = "" Then
	'	strFlag = "N"
	'ElseIf strItemCd <> "" And strWcCd = "" Then
	'	strFlag = "I"
	'ElseIf strItemCd = "" And strWcCd <> "" Then
	'	strFlag = "W"
	'ElseIf strItemCd <> "" And strWcCd <> "" Then
	'	strFlag = "A"
	'End If	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>
    
    With Parent
    	 '헤더데이타 Display
         .frm1.txtPlantNm.Value = "<%=ConvSPChars(strPlantNm)%>"
		.frm1.txtItemNm.Value = "<%=ConvSPChars(strItemNm)%>"
		.frm1.txtDefectRatioUnit.Value = "<%=ConvSPChars(strDefectRatioUnit)%>"
		.frm1.txtLotRejUnit.Value = "<%=ConvSPChars(strLotRejUnit)%>"
		'Detail Data Display
         .ggoSpread.Source  = .frm1.vspdData
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '☜ : Display data
         .lgStrPrevKey      =  "<%=ConvSPChars(lgStrPrevKey)%>"               '☜ : Next next data tag
         .DbQueryOk
	End with
</Script>	
<%
Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
