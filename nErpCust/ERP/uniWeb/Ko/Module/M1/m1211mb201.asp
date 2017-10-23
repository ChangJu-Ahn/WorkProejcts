<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MB201
'*  4. Program Name         : 공급처별배분비등록 
'*  5. Program Desc         : 공급처별배분비등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%	
Call HideStatusWnd
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim istrData
	Dim iStrPlantCd, iStrItemCd
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
    Dim sRow
    Dim lglngHiddenRows
    DIM MaxRow2
    Dim MaxCount
    Dim arrRsVal(11)
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	sRow           = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))

	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M1211MA202" 											' header
 
    iStrPlantCd = Trim(Request("txtPlantCd"))
    iStrItemCd = Trim(Request("txtItemCd"))
    
	UNIValue(0,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	UNIValue(0,1) = "  " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & "  "
	
	    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By 조건 
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
	Response.Write "    .lglngHiddenRows(" & sRow - 1 & ") = """ & MaxRow2 & """" & vbCr  
    Response.Write "    .DbQueryOk2(" & MaxCount & ")" & vbCr
	Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr        

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    DIM i
    Dim PvArr

	iLoopCount = -1
   ReDim PvArr(rs0.RecordCount)
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(0))		                        '0 bp_cd
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(1))								'1 pur_grp
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(7),ggExchRate.DecPoint,0)	'7 a.quota_rate
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(4))								'4 pur_priority
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(6))								'6 a.def_flg
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(5))								'5 sppl_dlvy_lt
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(2))								'2 pur_grp
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(3))								'3 pur_grp_nm
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPlantCd)	                        '8 상위 PlantCd
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrItemCd)	                        '9 상위 ItemCd
        iRowStr = iRowStr & Chr(11) & sRow
		iRowStr = iRowStr & Chr(11) & iLoopCount+1
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount+1

        istrData = istrData & iRowStr & Chr(11) & Chr(12)
        PvArr(iLoopCount) = istrData	
		istrData = ""
		
        rs0.MoveNext
   Loop

	istrData = Join(PvArr, "")    
    MaxRow2 = iLoopCount+1
    MaxCount = iLoopCount
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub




%>
