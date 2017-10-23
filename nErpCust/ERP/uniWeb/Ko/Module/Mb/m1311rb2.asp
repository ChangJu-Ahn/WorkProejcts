<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m2111rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Purchase Order Detail 참조 PopUp ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/23																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************

%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3     		   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim iTotstrData
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strPlantNm												  ' 공장명 
Dim strItemNm												  ' 품목명 
Dim strSpplNm											      ' 공급처명 

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB") 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
	Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			if ColCnt = 2 or ColCnt = 4 then
				iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(ColCnt),4,0)	
			else
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			end if
            
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
           lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    
    SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPlantNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strItemNm =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtItemCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strSpplNm =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtSpplCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If				
    End If
    
    SetConditionData = True
          
End Function 


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,5)

    UNISqlId(0) = "M1311RA201"
    UNISqlId(1) = "M2111QA302"								              '공장명 
	UNISqlId(2) = "M2111QA303"											  '품목명         
    UNISqlId(3) = "M3111QA102"								              '공급처명 
	
	strVal = ""
    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
	UNIValue(0,1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "		'---공장 
    UNIValue(0,2) = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "		'---품목 
    UNIValue(0,3) = " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "		'---공급처 
    UNIValue(0,4) = strVal   & "" '"ORDER BY    B.PL_SEQ_NO "
    
    UNIValue(1,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(2,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(2,1)  = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    UNIValue(3,0)  = " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "     
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList)) 
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    if SetConditionData = False then Exit sub
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>" 
		.frm1.txtItemNm.value = "<%=ConvSPChars(strItemNm)%>" 
        .frm1.txtSpplNm.value = "<%=ConvSPChars(strSpplNm)%>"
       		
		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hdnPlantCd.Value 		= "<%=Request("txtPlantCd")%>"
				.frm1.hdnItemCd.Value 		= "<%=Request("txtItemCd")%>"
				.frm1.hdnSpplCd.Value 		= "<%=Request("txtSpplCd")%>"
			End If    
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			
             if .frm1.vspdData.Maxrows >= 1 then
                .frm1.txtFrDt.text = .GetSpreadText(.frm1.vspdData,.GetKeyPos("A",10),1,"X","X")
                .frm1.txtToDt.text = .GetSpreadText(.frm1.vspdData,.GetKeyPos("A",11),1,"X","X")
             end if
			.DbQueryOk
		End If
	End with
</Script>	
