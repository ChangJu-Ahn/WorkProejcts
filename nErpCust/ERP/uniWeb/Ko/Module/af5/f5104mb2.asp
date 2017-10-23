<%
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

On Error Resume Next			' ☜: 

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3	           '☜ : DBAgent Parameter 선언 


Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

'Dim LngMaxRow																' 현재 그리드의 최대Row
Dim LngRow

Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt																		'☜: Group Count

Dim strMode						'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKeyNoteNo			' NoteNO 다음 값 
Dim StrNextKeyGlNo				' GLNO 다음 값 
Dim lgStrPrevKeyNoteNo			' Note NO 이전 값 
Dim lgStrPrevKeyGlNo

Dim strNoteFg, strNoteSts, strDueDtStart, strDueDtEnd, strBankCd, strNoteNo
Dim strWhere0
Dim strMsgCd, strMsg1, strMsg2

Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Const GroupCount = 30

strMode = Request("txtMode")	'☜ : 현재 상태를 받음 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	lgStrPrevKeyNoteNo = "" & UCase(Trim(Request("lgStrPrevKeyNoteNo")))
	lgStrPrevKeyGlNo = "" & UCase(Trim(Request("lgStrPrevKeyGlNo")))
		
Call TrimData()
Call FixUNISQLDATA()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Function FixUNISQLData()
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim intI
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
		    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "F5104MA103"
    UNISqlId(1) = "ABANKNM"
    UNISqlId(2) = "A_GETBIZ"
    UNISqlId(3) = "A_GETBIZ"

    Redim UNIValue(3,1)
	
	UNIValue(0,0) = strWhere0
	UNIValue(1,0) = FilterVar(strbankcd, "''", "S")
	UNIValue(2,0) = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(3,0) = FilterVar(strBizAreaCd1, "''", "S")	
		
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
		 
End Sub

'----------------------------------------------------------------------------------------------------------
' Function QueryData()
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strBankCd <> "" Then			
			strMsgCd = "970000"
			strMsg1 = Request("txtBankCd_Alt")
			If strMsgCd <> "" Then
				Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
				Response.End 
			End If	
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.txtBankNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%
	End If

If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtfromBizAreaCd.value = "<%=Trim(rs2(0))%>"
		.frm1.txtfromBizAreaNm.value = "<%=Trim(rs2(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs2.Close
	Set rs2 = Nothing   
    
    
If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txttoBizAreaCd.value = "<%=Trim(rs3(0))%>"
		.frm1.txttoBizAreaNm.value = "<%=Trim(rs3(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs3.Close
	Set rs3 = Nothing 
		
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
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

	strNoteFg = Request("cboNoteFg")								'어음구분 
	strNoteSts = Request("cboNoteSts")								'어음상태 
	strDueDtStart = UNIConvDate(Request("txtDueDtStart"))			'시작만기일 
	strDueDtEnd = UNIConvDate(Request("txtDueDtEnd"))				'종료만기일 
	strBankCd = UCase(Request("txtBankCd"))							'은행코드 
	'2003/12/12 Oh Soo Min 추가 
	strNoteNo = UCase(Request("txtNoteNo"))							'기준어음번호 
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))          '사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))         '사업장To
	
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))	
		    		
	strWhere0 = ""
	strWhere0 = strWhere0 & " and A.NOTE_FG =  " & FilterVar(strNoteFg , "''", "S") & " "	
	strWhere0 = strWhere0 & " AND A.ENDORSE_FG <> " & FilterVar("D4","''","S") & " "	
	strWhere0 = strWhere0 & " and  A.DUE_DT between  " & FilterVar(strDueDtStart, "''", "S") & " and  " & FilterVar(strDueDtEnd, "''", "S") & " "
	If strNoteSts <> "" Then 
		strWhere0 = strWhere0 & " and A.NOTE_STS =  " & FilterVar(strNoteSts , "''", "S") & " "				'어음상태 
'		strWhere0 = strWhere0 & " and D.MINOR_CD = '" & strNoteSts & "' "	
	Else		
		strWhere0 = strWhere0 & " and A.NOTE_STS in (" & FilterVar("OC", "''", "S") & " , " & FilterVar("DC", "''", "S") & " , " & FilterVar("SE", "''", "S") & " ," & FilterVar("MV", "''", "S") & " ) " ''''2004.04.08 KHJ 'MV'추가 
	End If
	
	If strBankCd <> "" Then 
		strWhere0 = strWhere0 & " and A.BANK_CD = " & Filtervar(strBankCd	, "''", "S")
		strWhere0 = strWhere0 & " and E.BANK_CD = " & Filtervar(strBankCd	, "''", "S")
	End If
			
	If strNoteNo <> "" Then 
		strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(strNoteNo	, "''", "S")		
	End If
	
	if strBizAreaCd <> "" then
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	strWhere0 = strWhere0 & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
	If lgStrPrevKeyNoteNo <> "" Then strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(lgStrPrevKeyNoteNo	, "''", "S")
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim intLoopCnt
%>
<Script Language=vbscript>
Option Explicit

	Dim LngMaxRow       
	Dim strData
	Const C_SHEETMAXROWS_D = 30
	
	LngMaxRow = parent.frm1.vspdData.MaxRows										'Save previous Maxrow                                         	
<%
	
	If rs0.recordcount > GroupCount Then
		intLoopCnt = GroupCount
	Else
		intLoopCnt = rs0.recordcount
	End If
			
	For LngRow = 1 To intLoopCnt

%>		
		strData = strData & Chr(11) & 0
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DUE_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BANK_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BANK_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>" 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_NM"))%>"		
		strData = strData & Chr(11) & ""        
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("note_no"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> 
		strData = strData & Chr(11) & Chr(12)		
<%      
		rs0.MoveNext
    Next
		    
    If Not rs0.EOF Then
%>    
		parent.lgStrPrevKeyNoteNo = "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		parent.lgStrPrevKeyGlNo = ""
<%	Else	%>
		parent.lgStrPrevKeyNoteNo = ""
		parent.lgStrPrevKeyGlNo = ""
<%	End If	%>
		
	With parent
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
				
		If .frm1.vspdData.MaxRows < C_SHEETMAXROWS_D And .lgStrPrevKeyNoteNo <> "" Then
			.DbQuery					
		Else
			.frm1.hProcFg.value		= "<%=ConvSPChars(Request("cboProcFg"))%>"
			.frm1.hNoteFg1.value	= "<%=ConvSPChars(Request("cboNoteFg"))%>"
			.frm1.hNoteSts.value	= "<%=ConvSPChars(Request("cboNoteSts"))%>"								
			.frm1.hDueDtStart.value	= "<%=Request("txtDueDtStart")%>"
			.frm1.hDueDtEnd.value	= "<%=Request("txtDueDtEnd")%>"						
			.frm1.hBankCd.value		= "<%=ConvSPChars(Request("txtBankCd"))%>"
			.frm1.hfromtxtBizAreaCd.value		= "<%=strBizAreaCd%>"
			.frm1.htotxtBizAreaCd.value			= "<%=strBizAreaCd1%>"					
			.DbQueryOK
		End If
	End With
					
</script>
<%      
End Sub
		
Sub ReleaseObj()			
	Set rs0 = Nothing
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub			
		
%>		
		


