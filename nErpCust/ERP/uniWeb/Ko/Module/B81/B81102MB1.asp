<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81102MB1
'*  4. Program Name         : 품목구성코드등록 
'*  5. Program Desc         : 품목구성코드등록()
'*  6. Component List       : PM1G121.cMMntSpplItemPriceS
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="./B81COMM.ASP" -->


<%	
call LoadBasisGlobalInf()
'call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
'call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

   ' Dim lgOpModeCRUD
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim GroupCount  
    Dim lgPageNo
	Dim iErrorPosition
	Dim arrRsVal(11)
	Dim strSpread
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode") 
	 strSpread = Request("txtSpread")										                                              '☜: Read Operation 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
   On Error Resume Next

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")
	Call FixUNISQLData()
	Call QueryData()	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr   
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & ",""F""" & vbCr
    Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
    Response.Write "	.DbQueryOk " & vbCr 
    Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr   
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr    
End Sub    
	    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,4)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "B81102MA101" 											' header
     UNISqlId(1) = "B81QB_MINOR" 	
     UNISqlId(2) = "B81QB_MINOR" 
     
     UNIValue(0,0)=" A.ITEM_LVL,CASE A.ITEM_LVL WHEN 'L1' THEN '대분류' WHEN 'L2' THEN '중분류' "
     UNIValue(0,0)= UNIValue(0,0) & " WHEN 'L3' THEN '소분류' END, LEN( A.CLASS_CD) ,A.CLASS_CD,A.CLASS_NAME,A.PARENT_CLASS_CD,'',"
     UNIValue(0,0)= UNIValue(0,0) & " dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.PARENT_CLASS_CD,'', A.ITEM_LVL),A.REMARK "
     UNIValue(0,1)="=" & FilterVar(Request("txtItem_acct"), "''", "S") & ""
     UNIValue(0,2)="=" & FilterVar(Request("txtItem_kind"), "''", "S") & ""
     if Request("cboItem_lvl")="*" then 
      UNIValue(0,3)=""
     else
      UNIValue(0,3)="AND A.ITEM_LVL =" & FilterVar(Request("cboItem_lvl"), "''", "S") & ""
     end if
    
     UNIValue(0,3)=  UNIValue(0,3) & "ORDER BY A.ITEM_LVL,  A.PARENT_CLASS_CD,A.CLASS_CD"
     UNIValue(1,0) ="'P1001'"
     UNIValue(1,1) ="" & FilterVar(Request("txtItem_acct"), "''", "S") & ""
     
     UNIValue(2,0) ="'Y1001'"
     UNIValue(2,1) ="" & FilterVar(Request("txtItem_kind"), "''", "S") & ""
  
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

	Set lgADF   = Nothing
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
 
    '----- UI 각 항목 체크 ----
    call fnCheckItem (rs1,"txtItem_acct","품목계정"   ) 
    call fnCheckItem (rs2,"txtItem_kind","품목구분"   ) 
    
  If  rs0.EOF And rs0.BOF  Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        
        rs0.Close
        Set rs0 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		'Response.end
    ELSE
        Call  MakeSpreadSheetData()
        goFocus("txtItem_acct")
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  =100
    Dim iLoopCount                                                                     
    Dim iRowStr
	Dim PvArr
	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
   
   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)

   Do while Not (rs0.EOF Or rs0.BOF)
		
        iLoopCount =  iLoopCount + 1
        iRowStr = ""

		iRowStr = Chr(11) & ConvSPChars(Trim(rs0(0)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(1)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(2)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(3)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(4)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(5)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(6)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(7)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(8)))
		
		iRowStr = iRowStr &	Chr(11) & iLngMaxRow + iLoopCount + 1                             
		iRowStr = iRowStr &	Chr(11) & Chr(12)                          
        
        If iLoopCount < C_SHEETMAXROWS_D Then
	        PvArr(iLoopCount) = iRowStr
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop
	
	istrData = Join(PvArr, "")
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================

	
Sub SubBizSaveMulti()
   
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
                                           '☜: Clear Error status
    '----- UI 각 항목 체크 ----
    Call SubOpenDB(lgObjConn) 
	call GetNameChk("minor_nm","b_minor","major_cd='P1001' and minor_cd="&filterVar(Request("txtItem_acct"),"''","S"),	Request("txtItem_acct"),"txtItem_acct","품목계정","Y") '품목계정
	call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&filterVar(Request("txtItem_kind"),"''","S"),	Request("txtItem_kind"),"txtItem_kind","품목구분","Y") '품목구분
	if chkGridCd=1 then
	Response.End 
	end if
   
	Call SubCloseDB(lgObjConn)  
 
    Call ObjPY1G102.B_CIS_CTRL(gStrGlobalCollection,strSpread)
  
    
    If CheckSYSTEMError(Err,True) = True Then                                              
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>

<%
End Sub

'----------------------------------------------------------------------------------------------------------
' chkGridCd
' Grid CD Value check.
'----------------------------------------------------------------------------------------------------------
function chkGridCd()
  
    dim RowStr,ColStr
    Dim i,tSql
	RowStr=split(strSpread,"")
	
	CONST C_ITEM_ACCT =2
	CONST C_ITEM_KIND =3
	CONST C_LEVEL = 4
	CONST C_CLASSCD = 5
	CONST C_CLASS_PARENTCD = 7
	chkGridCd =0
    Call SubOpenDB(lgObjConn) 
		for i=0 to uBound(RowStr)-1
			ColStr=split(RowStr(i),"")
			if ColStr(0)="C" or ColStr(0)="U" then
				
				if ColStr(4)="L1" then
					if ColStr(7)="*" then
					else
						Call DisplayMsgBox("971012", vbInformation, "상위코드", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
						call goFocusGRid("parent.frm1.vspdData",ColStr(1),6)
						Response.End 
					end if
					
				else 
					call GetNameChkGrid("CLASS_NAME","B_CIS_ITEM_CLASS","CLASS_CD="&filtervar(ColStr(7),"''","S")&" AND ITEM_ACCT="&filtervar(ColStr(2),"''","S")&" AND ITEM_KIND="&filtervar(ColStr(3),"''","S")&" " ,ColStr(1),6,"parent.frm1.vspdData","상위코드") '
				end if
				
			elseif ColStr(0)="D" then	 '하위코드 존재시 삭제 못하도록함.
			
				dim stmp
				if ColStr(4)="L1" then
					stmp="L2"
				elseif ColStr(4)="L2" then
					stmp="L3"
				end if
				
					tSql=" select CLASS_CD from B_CIS_ITEM_CLASS "
					tSql = tSql & "where ITEM_ACCT =" & filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & "and ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & "and ITEM_LVL="& filtervar(stmp,"''","S")
					tSql = tSql & "AND parent_class_cd ="& filtervar(ColStr(5),"''","S")
					
					If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") <> False Then 
					
						Call DisplayMsgBox("127929", vbInformation, "하위코드", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
						call goFocusGRid("parent.frm1.vspdData",ColStr(1),4)
						Response.End 
			
						
					end if
				

				if ColStr(C_LEVEL)="L3" then '소분류체크 
					'=================================================
					'B_CIS_NEW_ITEM_REQ 와 B_CIS_ITEM_MASTER를 체크함 
					'=================================================
					
					tSql =  " SELECT TOP 1 1"
					tSql = tSql & " FROM B_CIS_ITEM_MASTER"
					tSql = tSql & " WHERE ITEM_ACCT= " & filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & " AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & " AND ITEM_LVL1 IN ( SELECT PARENT_CLASS_CD FROM B_CIS_ITEM_CLASS "
					tSql = tSql & "		WHERE ITEM_ACCT="& filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & "		AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & "		AND CLASS_CD="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S") & " AND ITEM_LVL='L2'"
					tSql = tSql & "		)"
					tSql = tSql & " AND ITEM_LVL2="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S")
					tSql = tSql & " AND ITEM_LVL3="& filtervar(ColStr(C_CLASSCD),"''","S")
					
					tSql = tSql & "UNION SELECT TOP 1 1"
					tSql = tSql & " FROM B_CIS_NEW_ITEM_REQ"
					tSql = tSql & " WHERE ITEM_ACCT= " & filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & " AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & " AND ITEM_LVL1 IN ( SELECT PARENT_CLASS_CD FROM B_CIS_ITEM_CLASS "
					tSql = tSql & "		WHERE ITEM_ACCT="& filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & "		AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & "		AND CLASS_CD="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S") & " AND ITEM_LVL='L2'"
					tSql = tSql & "		)"
					tSql = tSql & " AND ITEM_LVL2="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S")
					tSql = tSql & " AND ITEM_LVL3="& filtervar(ColStr(C_CLASSCD),"''","S")



					If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") = TRUE Then 
						chkGridCd =1
						Call DisplayMsgBox("900020", vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
						call goFocusGRid("parent.frm1.vspdData",ColStr(1),4)
						Call SubCloseDB(lgObjConn) 
						exit function
						Response.End 
			
				
						
					end if
					
				elseif ColStr(C_LEVEL)="L2" then
					'=================================================
					'B_CIS_NEW_ITEM_REQ 와 B_CIS_ITEM_MASTER를 체크함 
					'=================================================
					
					tSql =  " SELECT TOP 1 1"
					tSql = tSql & " FROM B_CIS_ITEM_MASTER"
					tSql = tSql & " WHERE ITEM_ACCT= " & filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & " AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & " AND ITEM_LVL1="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S")
					tSql = tSql & " AND ITEM_LVL2="& filtervar(ColStr(C_CLASSCD),"''","S")
					
					tSql = tSql & "UNION SELECT TOP 1 1"
					tSql = tSql & " FROM B_CIS_NEW_ITEM_REQ"
					tSql = tSql & " WHERE ITEM_ACCT= " & filtervar(ColStr(C_ITEM_ACCT),"''","S")
					tSql = tSql & " AND ITEM_KIND="& filtervar(ColStr(C_ITEM_KIND),"''","S")
					tSql = tSql & " AND ITEM_LVL1="& filtervar(ColStr(C_CLASS_PARENTCD),"''","S")
					tSql = tSql & " AND ITEM_LVL2="& filtervar(ColStr(C_CLASSCD),"''","S")



					If 	FncOpenRs("R",lgObjConn,lgObjRs,tSql,"X","X") = TRUE Then 
						chkGridCd =1
						Call DisplayMsgBox("900020", vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
						call goFocusGRid("parent.frm1.vspdData",ColStr(1),4)
						Call SubCloseDB(lgObjConn) 
						exit function
						Response.End 

					end if
				end if	
				
					
			
			end if
		next
    Call SubCloseDB(lgObjConn) 
   
	
End function

%>






<OBJECT RUNAT=server PROGID="PY1G102.cBCtrlBiz" id=ObjPY1G102></OBJECT>