<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next
'Err.Clear

Dim lgDataExist

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strCond
Dim strGlCtrlFld

Dim strMsgCd
Dim strMsg1

    Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")         
	
    lgDataExist    = "No"
   
	strCond = ""
	strGlCtrlFld = Trim(UCase(Request("txtGlCtrlFld")))								'전표관리항목 
	
    Call SubOpenDB(lgObjConn)                                                       '☜: Make a DB Connection
    Call SubBizQueryMulti()
    Call SubCloseDB(lgObjConn)														'☜: Close DB Connection    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx    
    Dim iGLCTRLFLD , iGLCTRLNM
    
    iGLCTRLFLD = ""
    iGLCTRLNM  = ""
    
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT GL_CTRL_FLD,GL_CTRL_NM "
	lgStrSQL = lgStrSQL & " FROM A_SUBLEDGER_CTRL "
	lgStrSQL = lgStrSQL & " WHERE GL_CTRL_FLD >= " & FilterVar(strGlCtrlFld , "''", "S")
	lgStrSQL = lgStrSQL & " ORDER BY SUBLG_CD ASC "
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("110307", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found.
        lgErrorStatus     = "YES"
        Exit Sub
	Else
		iDx = 1
		lgstrData = ""
        
		Do While Not lgObjRs.EOF
			If iDx = 1 And Trim(strGlCtrlFld) <> "" Then
				Response.Write "HERE"
				iGLCTRLFLD = ConvSPChars(lgObjRs("GL_CTRL_FLD"))
				iGLCTRLNM  = ConvSPChars(lgObjRs("GL_CTRL_NM"))
			End If				
		
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_CTRL_FLD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_CTRL_NM"))
			lgstrData = lgstrData & Chr(11) & iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext

			iDx =  iDx + 1
		Loop		

		Response.Write  " <Script Language=vbscript>                                " & vbCr
		Response.Write  "    Parent.frm1.txtGlCtrlFld.value = """ & iGLCTRLFLD & """" & vbCr
		Response.Write  "    Parent.frm1.txtGlCtrlNm.value  = """ & iGLCTRLNM  & """" & vbCr
		Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData     " & vbCr
		Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData     & """" & vbCr
		Response.Write  "    Parent.DBQueryOk( """ & 1 & """)						" & vbCr      
		Response.Write  " </Script>												    " & vbCr
	End If
End Sub
    
%>
