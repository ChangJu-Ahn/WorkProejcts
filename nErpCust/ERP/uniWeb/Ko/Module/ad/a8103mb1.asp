<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '☜ : DBAgent Parameter 선언
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFromDt																'⊙ : 시작일
Dim strToDt																	'⊙ : 종료일
Dim strBizCd																'⊙ : 사업장
Dim strFromAmt																'⊙ : 시작금액
Dim strToAmt																'⊙ : 끝금액
Dim striOpt																	'⊙ : 어느 Grid인지..
Dim strRdo																	'⊙ : '1' 미결조회, '2' 완결조회
Dim strTemphq

Dim strCond

Dim strMsgCd
Dim strMsg1

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "MB")	

    lgStrPrevKey   = Request("lgStrPrevKey")								'☜ : Next key flag
    lgSelectList   = Request("lgSelectList")								'☜ : select 대상목록
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'☜ : 각 필드의 데이타 타입
    lgTailList     = Request("lgTailList")									'☜ : Orderby value
	striOpt		   = Request("txtIOpt")										'어느 Grid인지..
	strRdo		   = Request("txtRdoFg")									''1'미결조회, '2' 완결조회                             	

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
 
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

	Const C_SHEETMAXROWS_D = 100 

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
		If Isnumeric(lgStrPrevKey) Then
			iCnt = CInt(lgStrPrevKey)
		End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do While Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
        
        For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub  FixUNISQLData()
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보
    Redim UNIValue(1,2)

    UNISqlId(0) = "A8103MA101"
	UNISqlId(1) = "ABIZNM"

    UNIValue(0,0) = " DISTINCT " & lgSelectList                                          '☜: Select list
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(FilterVar(strBizCd, "''", "S") )
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	If strRdo = "1" Then	    
		If striOpt = "B" Then       
		    If (rs1.EOF And rs1.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBIzArea_Alt")
				End If
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.frm1.txtBIzArea.value = "<%=Trim(rs1(0))%>"
					.frm1.txtBIzAreaNm.value = "<%=Trim(rs1(1))%>"
				End With
				</Script>
		<%
		    End If
		    
			rs1.Close
			Set rs1 = Nothing 	
		End If   
	Else
		If striOpt = "A" Then       
		    If (rs1.EOF And rs1.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBIzArea_Alt")
				End If
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.frm1.txtBIzArea.value = "<%=Trim(rs1(0))%>"
					.frm1.txtBIzAreaNm.value = "<%=Trim(rs1(1))%>"
				End With
				</Script>
		<%
		    End If

			rs1.Close
			Set rs1 = Nothing
		End If
	End If	

	If striOpt = "A" Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
			Set rs0 = Nothing
		    Set lgADF = Nothing		
		Else
			Call  MakeSpreadSheetData()		
		End If
	Else	
		If  rs0.EOF And rs0.BOF Then
			
		Else
			Call  MakeSpreadSheetData()			
		End If			
	End If			
	
'    If  rs0.EOF And rs0.BOF Then
'		If strRdo = "1" Then	    
'			If striOpt = "B" Then       
'			    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
'				rs0.Close
'			    Set rs0 = Nothing
'			    Set lgADF = Nothing	
'			End If	
'		Else
'			If striOpt = "A" Then       
'			    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
'			    rs0.Close
'				Set rs0 = Nothing
'			    Set lgADF = Nothing	
'			End If	
'		End If	
'	Else    
'       Call  MakeSpreadSheetData()
'   End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
    strFromDt  = UNIConvDate(Request("txtFromDt"))			'시작일자
    strToDt	   = UNIConvDate(Request("txtToDt"))			'종료일자
	strBizCd   = UCase(Trim(Request("txtBizArea")))         '사업장
    strFromAmt = Request("txtFromAmt")						'시작금액
    strToAmt   = Request("txtToAmt")						'종료금액	 
    strTemphq  = Request("hTemphq")  
     
	If strFromDt <> "" Then
		strCond = strCond & " and A.TEMP_GL_DT >= " & FilterVar(strFromDt, "''", "S") 
    End If
     
    If strToDt <> "" Then
		strCond = strCond & " and A.TEMP_GL_DT <= " & FilterVar(strToDt, "''", "S")
    End If
     
    If strBizCd <> "" Then
		strCond = strCond & " and  D.BIZ_AREA_CD = " & FilterVar(strBizCd, "''", "S")
    End If
     
    If strFromAmt <> "" Then
		strCond = strCond & " and A.DR_LOC_AMT >= " & UNIConvNum(strFromAmt,0)
    End If

    If strToAmt <> "" Then
		strCond = strCond & " and A.DR_LOC_AMT <= " & UNIConvNum(strToAmt,0)
    End If
     	
	If strRdo = "1" Then							'미결
		If striOpt = "A" Then						'차변
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "
		ElseIf striOpt = "B" Then					'대변
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  "
		End If
		
	 	strCond = strCond & " AND A.GL_INPUT_TYPE = " & FilterVar("TG", "''", "S") & "  AND isnull(A.HQ_BRCH_NO,'')  = '' "
	 	
	ElseIf strRdo = "2" Then						'완결
		If striOpt = "A" Then						'차변

			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  AND isnull(A.HQ_BRCH_NO,'')  <> '' "
		ElseIf striOpt = "B" Then					'대변
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  AND  A.HQ_BRCH_NO = " & FilterVar(strTemphq, "''", "S")
		End If	
	End If
End Sub
%>

<Script Language=vbscript>
    With parent
<% 
		If striOpt = "A" Then 
%>		
			.ggoSpread.Source    = .frm1.vspdData
			.lgStrPrevKey_1      = "<%=lgStrPrevKey%>"                       '☜: set next data tag
<%			
        ElseIf striOpt = "B" Then   
%>        
			.ggoSpread.Source    = .frm1.vspdData2
			.lgStrPrevKey_2      = "<%=lgStrPrevKey%>"                       '☜: set next data tag
<%			
        End If
%>        
        If Trim(.frm1.txtBizArea.value) = "" Then	
			.frm1.txtBizAreaNm.Value = ""
		End If     
        .ggoSpread.SSShowData  "<%=lgstrData%>"                          '☜: Display data 
        '.lgStrPrevKey_A      = "<%=lgStrPrevKey%>"                       '☜: set next data tag       
	    .DbQueryOk("<%=striOpt%>")
	End with
</Script>
