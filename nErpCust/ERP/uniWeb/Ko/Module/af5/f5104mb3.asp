<%
Option Explicit
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f5104ma
'*  4. Program 이름      : 만기어음일괄처리 
'*  5. Program 설명      : 만기어음일괄처리 
'*  6. Comproxy 리스트   : f5104ma
'*  7. 최초 작성년월일   : 2000/10/16
'*  8. 최종 수정년월일   : 2002/02/15
'*  9. 최초 작성자       : 김종환 
'* 10. 최종 작성자       : 오수민 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/10/16 : ..........
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next			' ☜: 

Dim lgADF																	'☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg																'☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2						'☜ : DBAgent Parameter 선언 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim LngMaxRow																' 현재 그리드의 최대Row
Dim LngRow

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim lgDataExist
Dim lgPageNo

Dim strNoteNo
Dim strBizAreaCd
Dim strBizAreaCd1
Dim strNoteFg
Dim strFrStsDT
Dim strToStsDT
Dim strMsgCd
Dim strData

Dim iPrevEndRow
'Dim iEndRow

Dim strCond

	lgDataExist = "NO"
'    iPrevEndRow = 0
'    iEndRow = 0	
	strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

    lgPageNo  = UNICInt(Trim(Request("lgPageNo")),0)					'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	LngMaxRow = UNICInt(Trim(Request("txtMaxRows")),0)

	Call TrimData()	
	Call FixUNISQLDATA()
	Call QueryData()

Sub TrimData()

	strNoteFg	  = UCase(Request("cboNoteFg"))                         	'어음구분 
	strFrStsDT	  = UNIConvDate(Request("txtStsDtStart"))					'조건시작일 
    strToStsDT    = UNIConvDate(Request("txtStsDtEnd"))						'조건종료일 
	strNoteNo	  = Trim(Ucase(Request("txtNoteNo1")))						'어음번호 
	strBizAreaCd  = Trim(Ucase(Request("txtBizAreaCd")))					'사업장From
	strBizAreaCd1 = Trim(Ucase(Request("txtBizAreaCd1")))					'사업장To


    strCond = "" 
	strCond = strCond & "     A.NOTE_FG = " & FilterVar(strNoteFg,"''","S")
	strCond = strCond & " AND B.STS_DT >= " & Filtervar(UNIConvDate(Request("txtStsDtStart"))	, "", "S")
	strCond = strCond & " AND B.STS_DT <= " & Filtervar(UNIConvDate(Request("txtStsDtEnd"))	, "", "S") 
	
	If strBizAreaCd <> "" Then
		strCond = strCond & " AND A.BIZ_AREA_CD >= " & FilterVar(strBizAreaCd,"''","S")
	Else
		strCond = strCond & " AND A.BIZ_AREA_CD >= " & FilterVar("0" , "''", "S") 
	End If
	
	If strBizAreaCd1 <> "" Then
		strCond = strCond & " AND A.BIZ_AREA_CD <= " & FilterVar(strBizAreaCd1,"''","S")
	Else
		strCond = strCond & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ" , "''", "S")  
	End If
	
	If strNoteNo <> "" Then
		strCond = strCond & " AND A.NOTE_NO >= " & FilterVar(strNoteNo,"''","S")
	End If	
End Sub

Sub FixUNISQLData()
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 

    UNISqlId(0) = "F5104MA102"
    UNISqlId(1) = "A_GETBIZ"
    UNISqlId(2) = "A_GETBIZ"

    Redim UNIValue(2,0)

	UNIValue(0,0)  = strCond
	UNIValue(1,0)  = FilterVar(strBizAreaCd,"''","SNM")
	UNIValue(2,0)  = FilterVar(strBizAreaCd1,"''","SNM")	
	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
		    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtfromBizAreaCd1.value = "<%=Trim(rs1(0))%>"
		.frm1.txtfromBizAreaNm1.value = "<%=Trim(rs1(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs1.Close
	Set rs1 = Nothing   
    
    
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txttoBizAreaCd1.value = "<%=Trim(rs2(0))%>"
		.frm1.txttoBizAreaNm1.value = "<%=Trim(rs2(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs2.Close
	Set rs2 = Nothing 
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Set lgADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else
		Call  MakeSpreadSheetData()
    End If				
		    
    Call ReleaseObj()
End Sub
		
Sub MakeSpreadSheetData()
    Dim  iLoopCount
    
    lgDataExist  = "Yes"
    strData      = ""

	Const C_SHEETMAXROWS_D = 50
    
    If CInt(lgPageNo) > 0 Then
		iPrevEndRow = C_SHEETMAXROWS_D * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   
    End If

    iLoopCount = -1
			
    Do While Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1    

        If  iLoopCount < C_SHEETMAXROWS_D Then
			strData = strData & Chr(11) & 0
			strData = strData & Chr(11) & ConvSPChars(rs0("NOTE_NO"))
			strData = strData & Chr(11) & ConvSPChars(rs0("TEMP_GL_NO"))
			strData = strData & Chr(11) & UNIDateClientFormat(rs0("TEMP_GL_DT"))
			strData = strData & Chr(11) & ConvSPChars(rs0("GL_NO"))
			strData = strData & Chr(11) & UNIDateClientFormat(rs0("GL_DT"))
			strData = strData & Chr(11) & UNINumClientFormat(rs0("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("BP_NM"))
			strData = strData & Chr(11) & ConvSPChars(rs0("DEPT_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("DEPT_NM"))
			strData = strData & Chr(11) & ConvSPChars(rs0("RCPT_TYPE"))
			strData = strData & Chr(11) & ConvSPChars(rs0("ORG_CHANGE_ID"))
			strData = strData & Chr(11) & ConvSPChars(rs0("DEPT_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("INTERNAL_CD"))
			strData = strData & Chr(11) & ConvSPChars(rs0("NOTE_ITEM_DESC"))
			strData = strData & Chr(11) & LngMaxRow + iLoopCount
			strData = strData & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If

		rs0.MoveNext
	Loop


    If  iLoopCount < C_SHEETMAXROWS_D Then								'☜: Check if next data exists
        lgPageNo = ""													'☜: 다음 데이타 없다.
    End If
End Sub	

Sub ReleaseObj()			
	Set rs0 = Nothing
	Set lgADF = Nothing													'☜: ActiveX Data Factory Object Nothing
End Sub				    
%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
		If "<%=lgPageNo%>" = "1" Then
			With parent.frm1
				.hProcFg.value		 = "<%=ConvSPChars(Request("cboProcFg"))%>"			
				.hNoteFg2.value      = Trim(.cboNoteFg2.value)
				.hFrStsDT1.value     = Trim(.txtStsDtStart.Text)
				.hToStsDT1.value     = Trim(.txtStsDtEnd.Text)
				.htxtNoteNo1.value   = Trim(.txtNoteNo1.value)
				.hFrBizAreaCd1.value = Trim(.txtfromBizAreaCd1.value)
				.hToBizAreaCd1.value = Trim(.txttoBizAreaCd1.value)
			End With
		End If
       
       'Show multi spreadsheet data from this line
		With Parent
			.ggoSpread.Source  = .frm1.vspdData2
			.frm1.hProcFg.value		 = "<%=ConvSPChars(Request("cboProcFg"))%>"						
			.frm1.vspdData2.Redraw = False
			.ggoSpread.SSShowData "<%=strData%>"						'☜ : Display data
			.frm1.vspdData2.Redraw = True	
			.lgPageNo      =  "<%=lgPageNo%>"							'☜ : Next next data tag
			.DbQueryOk			
		End With
	End If	
</script>	

