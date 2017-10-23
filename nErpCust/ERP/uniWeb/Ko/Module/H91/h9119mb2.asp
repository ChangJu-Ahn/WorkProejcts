<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                        '��: Clear Error status

    Dim AlgObjRs,BlgObjRs,ClgObjRs,DlgObjRs
    Dim BiDx,CiDx
    Dim strFilePath,strMode
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm    
    Call HideStatusWnd                                                              '��: Hide Processing message
    BiDx = 1

    lgErrorStatus     = "NO"													    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '��: "M"(Spread #1) "S"(Spread #2)
    
    strMode      = Request("txtMode")                                               '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
    lgstrData = ""
 
    Call SubOpenDB(lgObjConn)      
            
    Select Case strMode
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")  
                Fnm = Fso.GetFileName(lgKeyStream(4))                
                FPnm = Server.MapPath("../../files/u2000/" & Fnm)  '2002.02.01 /files ���� ���� u2000�� ����:���߿� ������ ����Ǹ� �����ؾ� ��.
                
                Call SubBizQuery("")

            If UCase(Trim(lgErrorStatus)) <> "YES" Then
                Set CTFnm = Fso.CreateTextFile (FPnm,True)                              'text�� ������ File�� ����            
                
                CTFnm.Write lgstrData                                                   'Text ����κ�                       
                DFnm = Fso.GetFileName(FPnm)            
                CTFnm.close    
                Set CTFnm = nothing
            Else
                Call DisplayMsgBox("800004", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
                Call SetErrorStatus() 
            End If
            Set Fso = nothing           
            
            Call HideStatusWnd           
            
%>
    <SCRIPT LANGUAGE=VBSCRIPT>
				parent.subVatDiskOK("<%=DFnm%>")
	</SCRIPT>
<%
    Case CStr(UID_M0002)

	    Err.Clear 

	    Call HideStatusWnd

	    strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
	    			   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
	    strFilePath = strFilePath  & "files/u2000/"    '2002.02.01 /files ���� ���� u2000�� ����:���߿� ������ ����Ǹ� �����ؾ� ��.
	    strFilePath = strFilePath & Request("txtFileName")

End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(CQuery)        
Dim strWhere    
Dim pComp
 
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            pComp = "="
            strWhere = FilterVar(lgKeyStream(6), "''", "S")
            strWhere = strWhere & CQuery
            
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '��: Make sql statements       
            Call SubBizQueryMulti()
        Case "B"
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year," & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '�ٲ� Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery
            strWhere = strWhere & " GROUP BY hfa100t.year_area_cd,"
            strWhere = strWhere & " hfa100t.tax_biz_cd,"
            strWhere = strWhere & " hfa100t.own_rgst_no,"
            strWhere = strWhere & " hfa100t.year_area_nm,"
            strWhere = strWhere & " hfa100t.repre_nm,"
            strWhere = strWhere & " hfa100t.co_own_rgst_no"
            
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)      
        Case "C"   
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year, " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '�ٲ� Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)
        Case "D"
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year, " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")    
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '�ٲ� Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)
    End Select     
End Sub	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call ASubBizQueryMulti()        
End Sub    
'============================================================================================================
' Name : ASubBizQueryMulti()
' Desc : Query ASheet Data from Db
'============================================================================================================
Sub ASubBizQueryMulti()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
        
    lgstrData = ""
    If FncOpenRs("R",lgObjConn,AlgObjRs,lgStrSQL,"X","X") = False Then
       Call SetErrorStatus("")
    Else        
        Call CommonQueryRs("count(*) ","HFA100T","year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        Do While Not AlgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("present_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("present_type"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("mag_no"),"","",6,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("hometax_id"),"","",20,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("taxpgm_cd"),"","",4,"")
            lgstrData = lgstrData & SetFixSrting(replace(AlgObjRs("biz_own_rgst_no"),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("biz_area_nm"),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_DEPT_NM"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_NAME"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_TEL"),"","",15,"")
'            lgstrData = lgstrData & SetFixSrting(replace(AlgObjRs("co_own_rgst_no"),"-",""),"","",13,"")
 '           lgstrData = lgstrData & SetFixSrting(AlgObjRs("repre_nm"),"","",30,"")
  '          lgstrData = lgstrData & SetFixSrting(AlgObjRs("tel_no"),"","",15,"")
            lgstrData = lgstrData & SetFixSrting(Replace(lgF0, Chr(11), ""),"","0",5,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("kr_code"),"","0",3,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("present_gigan"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("mod_code"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("empty"),"","",551,"")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            If Cdbl(AlgObjRs("b_count")) > 0 Then                
                lgCurrentSpd = "B"
                Call BSubBizQueryMulti()
                lgCurrentSpd = "A"
            End If
            AlgObjRs.MoveNext
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,AlgObjRs,Err)
    Call SubCloseRs(AlgObjRs)    
End Sub
'============================================================================================================
' Name : BSubBizQueryMulti()
' Desc : Query BSheet Data from Db
'============================================================================================================
Sub BSubBizQueryMulti()

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status    
    
    Call SubBizQuery("")
    If 	FncOpenRs("R",lgObjConn,BlgObjRs,lgStrSQL,"X","X") = False Then
        Call SetErrorStatus()
    Else    
        Do While Not BlgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(BiDx,"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(BlgObjRs("biz_own_rgst_no"),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("biz_area_nm"),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("repre_nm"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(replace(BlgObjRs("co_own_rgst_no"),"-",""),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("com_no"),"","0",7,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("old_com_no"),"","0",7,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tot"),"","0",14,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("dec_income_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tot_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("dec_res_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("dec_farm_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("dec_tot"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("empty"),"","",532,"")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            IF Cdbl(ConvSPChars(BlgObjRs("com_no"))) > 0 Then                
                lgCurrentSpd = "C"
                Call CSubBizQueryMulti()
                lgCurrentSpd = "B"
            End If
            BiDx =  BiDx + 1
            BlgObjRs.MoveNext
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,BlgObjRs,Err)
    Call SubCloseRs(BlgObjRs)
End Sub
'============================================================================================================
' Name : CSubBizQueryMulti()
' Desc : Query CSheet Data from Db
'============================================================================================================
Sub CSubBizQueryMulti()
    Dim iDx
    Dim c_per_sub, c_spouse_sub, c_fam_sub, c_old_sub, c_paria_sub, c_lady_sub, c_chl_rear_sub ,c_old_sub2
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    Call SubBizQuery("")
    If 	FncOpenRs("R",lgObjConn,ClgObjRs,lgStrSQL,"X","X") = False Then
        Call SetErrorStatus()
    Else  
         Call CommonQueryRs(" old_sub2 ","HFA020T"," 1=1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        c_old_sub2     = Replace(lgF0, Chr(11), "")      
        Call CommonQueryRs("per_sub, spouse_sub, fam_sub, old_sub, paria_sub, lady_sub, chl_rear_sub ","HFA020T"," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        c_per_sub      = Replace(lgF0, Chr(11), "") 
        c_spouse_sub   = Replace(lgF1, Chr(11), "")
        c_lady_sub     = Replace(lgF5, Chr(11), "")
 
        iDx = 1
        Do While Not ClgObjRs.EOF
            c_fam_sub      = CInt(ClgObjRs("hfa050t_supp_cnt"))  * Replace(lgF2, Chr(11), "")
            c_old_sub      = CInt(ClgObjRs("old_cnt1"))   * Replace(lgF3, Chr(11), "") + CInt(ClgObjRs("old_cnt2"))   * c_old_sub2
            c_paria_sub    = CInt(ClgObjRs("hfa050t_paria_cnt")) * Replace(lgF4, Chr(11), "")
            c_chl_rear_sub = CInt(ClgObjRs("hfa050t_chl_rear"))  * Replace(lgF6, Chr(11), "")                   

            lgstrData = lgstrData & SetFixSrting(ClgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("data_type"),"","",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(iDx,"","0",6,"RIGHT")
            CiDx = iDx
            lgstrData = lgstrData & SetFixSrting(replace(ClgObjRs("biz_own_rgst_no"),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_com_no"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hdf020t_res_flag"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_nat_cd"),"","",2,"")   '2002 ���������ڵ� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("FOREIGN_SEPARATE_TAX_YN"),"","",1,"")   '2004 �ܱ��δ��ϼ��� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_entr_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_retire_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_name"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("for_type"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(left(replace(ClgObjRs("res_no"),"-",""),13),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("start_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("end_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_new_pay_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_new_bonus_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa030t_after_bonus_amt"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("new_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_non_tax5"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_non_tax1"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("non_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("non_tax_sum"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_income_tot_amt"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_income_sub_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_income_amt"),"","0",11,"RIGHT")
' 2001 �������� 
'            lgstrData = lgstrData & SetFixSrting(c_per_sub,"","0",8,"RIGHT")                     
 '           IF ClgObjRs("hfa050t_spouse_sub_amt") > 0 Then
  '              lgstrData = lgstrData & SetFixSrting(c_spouse_sub,"","0",8,"RIGHT")              
   '         Else
    '            lgstrData = lgstrData & SetFixSrting(0,"","0",8,"RIGHT")                    
     '       End if                     
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_per_sub_amt")   ,"","0",8,"RIGHT"  )
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_spouse_sub_amt") ,"","0",8,"RIGHT")  
                         
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_supp_cnt"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(c_fam_sub,"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_cnt"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(c_old_sub,"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_paria_cnt"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(c_paria_sub,"","0",8,"RIGHT")
            IF ClgObjRs("hfa050t_lady_sub_amt") > 0 Then
                lgstrData = lgstrData & SetFixSrting(c_lady_sub,"","0",8,"RIGHT")               
            Else
                lgstrData = lgstrData & SetFixSrting(0,"","0",8,"RIGHT")               
            End if                     
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_chl_rear"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(c_chl_rear_sub,"","0",8,"RIGHT")
            
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_small_sub_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_national_pension_sub_amt"),"","0",10,"RIGHT") '2002
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_insur_sub_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_med_sub_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_edu_sub_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_house_fund_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_contr_sub_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_ceremony_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_std_sub_tot_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_std_sub_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_sub_income_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_indiv_anu_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_indiv_anu2_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_invest_sub_sum_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_card_sub_sum_amt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_our_stock_amt"),"","0",10,"RIGHT")   '2002 �츮���������⿬�� �߰� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("special_tax_sum"),"","0",10,"RIGHT")  '2004
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_tax_std_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_calu_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_income_redu_amt"),"","0",10,"RIGHT") '2002 ���װ���:�ҵ漼�� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_taxes_redu_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_redu_sum_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_income_tax_sub_amt"),"","0",8,"RIGHT") '���װ���:�ٷμҵ� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty_4"),"","0",8,"RIGHT")   '/* �������� */
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_house_repay_amt"),"","0",8,"RIGHT")   '�������� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_fore_pay_amt"),"","0",8,"RIGHT")      '�ܱ����� 
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_poli_tax_sub"),"","0",10,"RIGHT")         '2004            
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty_3"),"","0",8,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty_3"),"","0",8,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_tax_sub_amt"),"","0",8,"RIGHT")         '2004   
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_dec_income_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_dec_res_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_dec_farm_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("dec_tot"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_income_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_res_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_farm_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_tot"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_new_income_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_new_res_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_new_farm_tax_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("new_sum_tot"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty_2"),"","",26,"")  'by 20040210 cyc  
            
            lgstrData = lgstrData & Chr(13) & Chr(10)
            If Cdbl(ClgObjRs("old_com_no")) > 0 Then                
                lgCurrentSpd = "D"
                Call DSubBizQueryMulti()
                lgCurrentSpd = "C"
            End If
            ClgObjRs.MoveNext
            iDx = iDx + 1
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,ClgObjRs,Err)
    Call SubCloseRs(ClgObjRs)
End Sub
'============================================================================================================
' Name : DSubBizQueryMulti()
' Desc : Query DSheet Data from Db
'============================================================================================================
Sub DSubBizQueryMulti()
Dim Oldres_no
Dim Cwork_no
Dim ConWhere

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    ConWhere = " And haa010t.emp_no = " & FilterVar(ConvSPChars(ClgObjRs("haa010t_emp_no")), "''", "S")    
    Call SubBizQuery(ConWhere)
    If 	FncOpenRs("R",lgObjConn,DlgObjRs,lgStrSQL,"X","X") = False Then
        Call SetErrorStatus()
    Else        
        Do While Not DlgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(CiDx,"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(DlgObjRs("biz_own_rgst_no"),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("empty"),"","",50,"")
            If Trim(Oldres_no) = "" Then
               Oldres_no = ConvSPChars(DlgObjRs("res_no"))
               Cwork_no = 1
            Else
               If Oldres_no = ConvSPChars(DlgObjRs("res_no")) Then                            
                   Cwork_no = Cwork_no + 1
               Else
                   Oldres_no = ConvSPChars(DlgObjRs("res_no"))
                   Cwork_no = 1
               End If
            End If
            lgstrData = lgstrData & SetFixSrting(left(DlgObjRs("res_no"),13),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("comp_nm"),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("comp_no"),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("hfa040t_pay_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("bonus_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("bonus_amt"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("pay_tot"),"","0",11,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(Cwork_no,"","0",2,"RIGHT")                     
'            lgstrData = lgstrData & SetFixSrting(DlgObjRs("mod_code"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(DlgObjRs("empty_2"),"","",549,"")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            DlgObjRs.MoveNext            
        Loop
    End If
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
    Dim Submit_dt

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Submit_dt = UniConvDateToYYYYMMDD(lgKeyStream(2),gDateFormat,"")

    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))
                Case "A"
                    lgStrSQL = " SELECT " & FilterVar("A", "''", "S") & "   record_type,"                              '/* ���ڵ屸�� : A�� ���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"                           '/* �ڷᱸ�� : 20���� ���� */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(Submit_dt, "''", "S") & "  present_dt,"      '/* ���⿬���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "  present_type,"'/* ������(�븮��)���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & "  mag_no," '/* �����븮�ΰ�����ȣ */   
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  " '/* 2004 hometax id */ 
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  " 						 '/* 2004 �������α׷��ڵ� ��Ÿ */  
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_own_rgst_no,"   '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm)  biz_area_nm,"  '/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM,  " '����� �μ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_NAME,  " '����ڸ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_TEL,  " '����� ��ȭ��ȣ 2004                    
'                    lgStrSQL = lgStrSQL & " co_own_rgst_no,"     '/* �ֹ�(����)��Ϲ�ȣ */
 '                   lgStrSQL = lgStrSQL & " repre_nm,"           '/* ��ǥ��(����) */
  '                  lgStrSQL = lgStrSQL & " addr  address,"      '/* �ּ� -> 2000�� �������� ���� */
   '                 lgStrSQL = lgStrSQL & " tel_no,"             '/* ��ȭ��ȣ */
                    lgStrSQL = lgStrSQL & " " & FilterVar("101", "''", "S") & "  kr_code,"                            '/* ����ѱ��ڵ� : 101�� ���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & "   present_gigan,"                        '/* ������Ⱓ�ڵ� -> �Էº��� */
'                    lgStrSQL = lgStrSQL & " SPACE(1)  mod_code,"                        '/* �ڷ�����ڵ� */
                    lgStrSQL = lgStrSQL & " SPACE(551)  empty"                          '/* ���� */
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd" & pComp & pCode
                Case "B"                
                    lgStrSQL = " SELECT hfa100t.year_area_cd  year_area_cd,"           '/* �Ű����� */========='�ٲ� Query
                    lgStrSQL = lgStrSQL & " " & FilterVar("B", "''", "S") & "   record_type,"                          '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"                           '/* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"  '/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"             '/* ��ǥ��(����) */
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  co_own_rgst_no,"     '/* �ֹ�(����)��Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " COUNT(hfa050t.emp_no) com_no,"              '/* ��(��)�ٹ�ó(���ڵ�) �� */
                    lgStrSQL = lgStrSQL & " SUM(t.emp_no_cnt)  old_com_no,"          '/* ��(��)�ٹ�ó(���ڵ�) ��*/========='�ٲ� Query
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.income_tot_amt))  tot,"   '/* �ҵ�ݾ��Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_income_tax))  dec_income_tax,"  '/* �ҵ漼 �������� �Ѱ� */
                    lgStrSQL = lgStrSQL & " 0 tot_tax,"                                  '/* ���μ� �������� �Ѱ� -> 0 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_res_tax))  dec_res_tax,"        '/* �ֹμ� �������� �Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_farm_tax))  dec_farm_tax,"      '/* ��Ư�� �������� �Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_income_tax) + FLOOR(hfa050t.dec_farm_tax) + FLOOR(dec_res_tax))  dec_tot,"  '/* �������� �Ѱ� */
                    lgStrSQL = lgStrSQL & " SPACE(471) empty"                             '/* ���� */                                        
                    lgStrSQL = lgStrSQL & " FROM haa010t," 
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hfa100t,"
                    lgStrSQL = lgStrSQL & " (SELECT emp_no, COUNT(*) emp_no_cnt"                    '�ٲ� Query
                    lgStrSQL = lgStrSQL & " FROM hfa040t"                                           '�ٲ� Query
                    lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")                                 '�ٲ� Query
                    lgStrSQL = lgStrSQL & " GROUP BY emp_no) AS t"                                  '�ٲ� Query
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= t.emp_no"                         '�ٲ� Query
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"              '/* ������ ���� */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode
                Case "C"                
                    lgStrSQL = " SELECT " & FilterVar("C", "''", "S") & "   record_type,"                  '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"               ' /* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " ISNULL(before.emp_before_count,0)  old_com_no,"  '/* ��(��)�ٹ�ó(���ڵ�) ��*/
                    lgStrSQL = lgStrSQL & " CASE WHEN (hdf020t.res_flag IS NULL OR hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " ) THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("2", "''", "S") & " END  hdf020t_res_flag,"  '/* �����ڱ����ڵ� */
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd = " & FilterVar("KR", "''", "S") & ") THEN " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE haa010t.nat_cd END  haa010t_nat_cd,"  '2002 ���������ڵ� 
                    lgStrSQL = lgStrSQL & " CASE WHEN hdf020t.FOREIGN_SEPARATE_TAX_YN ='Y' THEN 1 ELSE 2 END FOREIGN_SEPARATE_TAX_YN,"   '/* �ܱ��δ��ϼ��� */ 2004                    
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.entr_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") 
                    lgStrSQL = lgStrSQL & " + " & FilterVar("0101", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE CONVERT(VARCHAR(8), haa010t.entr_dt, 112) END  haa010t_entr_dt,"  '/*  */
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.retire_dt IS NULL THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN haa010t.retire_dt > " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
                    lgStrSQL = lgStrSQL & " + " & FilterVar("1231", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE CONVERT(VARCHAR(8), haa010t.retire_dt, 112) END  haa010t_retire_dt,"  '/*  */ 
                    lgStrSQL = lgStrSQL & " haa010t.name  haa010t_name,"  '/*  */
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.nat_cd = " & FilterVar("KR", "''", "S") & " THEN 1 ELSE 9 END  for_type,"  '/*  */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"  '/* �ֹε�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " haa010t.zip_cd,"          '/* 2000�� �������� ���� */
                    lgStrSQL = lgStrSQL & " haa010t.addr,"            '/* 2000�� �������� ���� */
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.entr_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(2),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.entr_dt >= " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & " THEN CONVERT(VARCHAR(8), haa010t.entr_dt, 112)"
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("00000000", "''", "S") & " END  start_dt,"  '/* ����Ⱓ���ۿ����� */ 
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NULL THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NOT NULL"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt >= " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(2),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NOT NULL"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & " THEN CONVERT(VARCHAR(8), haa010t.retire_dt, 112)"
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("00000000", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " END  end_dt,"  '/* ����Ⱓ���Ῥ���� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_pay_tot)  hfa050t_new_pay_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_bonus_tot)  hfa050t_new_bonus_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa030t.after_bonus_amt)  hfa030t_after_bonus_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.new_pay_tot) + FLOOR(hfa050t.new_bonus_tot) + FLOOR(hfa030t.after_bonus_amt)) new_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.non_tax5)  hfa050t_non_tax5,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.non_tax1)  hfa050t_non_tax1,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.non_tax2) + FLOOR(hfa050t.non_tax3) + FLOOR(hfa050t.non_tax4) + FLOOR(hfa050t.non_tax6))  non_tax,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.non_tax1) + FLOOR(hfa050t.non_tax2) + FLOOR(hfa050t.non_tax3) + FLOOR(hfa050t.non_tax4) + FLOOR(hfa050t.non_tax5) + FLOOR(hfa050t.non_tax6))  non_tax_sum,"  '/*  */                     
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_tot_amt)  hfa050t_income_tot_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_sub)  hfa050t_income_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_amt)  hfa050t_income_amt,"  '/*  */
                    '2001 ��������                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.per_sub)  hfa050t_per_sub_amt,"  '/* ���ΰ��� */
                    'lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(hfa050t.spouse," & FilterVar("N", "''", "S") & " ) = " & FilterVar("Y", "''", "S") & "  THEN 1000000"
                    'lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_spouse_sub_amt,"  '/* ����ڰ��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.SPOUSE_SUB)  hfa050t_spouse_sub_amt,"  '/* ����ڰ��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.supp_cnt)  hfa050t_supp_cnt,"  '/* �ξ��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_cnt +hfa050t.old_cnt2)  hfa050t_old_cnt,hfa050t.old_cnt old_cnt1,hfa050t.old_cnt2 old_cnt2, "  '/* ��ο�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.paria_cnt)  hfa050t_paria_cnt,"  '/* ����� */
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(hfa050t.lady," & FilterVar("N", "''", "S") & " ) = " & FilterVar("Y", "''", "S") & "  THEN 500000"
                    lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_lady_sub_amt,"  '/* �γ��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.chl_rear)  hfa050t_chl_rear,"  '/* �ڳ���� */
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.small_sub)  hfa050t_small_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.national_pension_sub_amt)  hfa050t_national_pension_sub_amt," '2002 ���ο��� 
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.insur_sub)  hfa050t_insur_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.med_sub)  hfa050t_med_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.edu_sub)  hfa050t_edu_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.house_fund)  hfa050t_house_fund_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.contr_sub)  hfa050t_contr_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.ceremony_amt)  hfa050t_ceremony_amt,"  '2004

                    lgStrSQL = lgStrSQL & " CASE WHEN FLOOR(hfa050t.std_sub) <= 600000 and (hdf020t.FOREIGN_SEPARATE_TAX_YN <> " & FilterVar("Y", "''", "S") & ") THEN 0"
                    lgStrSQL = lgStrSQL & " ELSE FLOOR(hfa050t.std_sub) END  hfa050t_std_sub_tot_amt,"  '/* ��(Ư������) */
                    lgStrSQL = lgStrSQL & " CASE WHEN FLOOR(hfa050t.std_sub) <= 600000 and (hdf020t.FOREIGN_SEPARATE_TAX_YN <> " & FilterVar("Y", "''", "S") & ") THEN 600000"
                    lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_std_sub_amt,"     '/* ǥ�ذ��� */
           
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.sub_income_amt)  hfa050t_sub_income_amt,"  '/* �����ҵ�ݾ� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu)  hfa050t_indiv_anu_amt,"  ' ���ο������� 
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu2)  hfa050t_indiv_anu2_amt,"  ' �������� 
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.invest_sub_sum)  hfa050t_invest_sub_sum_amt,"  
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.card_sub_sum)  hfa050t_card_sub_sum_amt," 
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.our_stock_amt)  hfa050t_our_stock_amt," ' 2002 �츮���������⿬�� 
					lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu) +FLOOR(hfa050t.indiv_anu2)+FLOOR(hfa050t.invest_sub_sum)+FLOOR(hfa050t.card_sub_sum)+FLOOR(hfa050t.our_stock_amt) special_tax_sum, " '2004
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd <> " & FilterVar("KR", "''", "S") & ") THEN FLOOR(hfa050t.fore_edu_sub_amt)" 
                    lgStrSQL = lgStrSQL & " ELSE FLOOR(hfa050t.our_stock_amt)  END  hfa050t_our_stock_amt," '2003 (������: �츮���������⿬��) & (�ܱ��� :�ܱ��α�����������)
                                      
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.tax_std)  hfa050t_tax_std_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.calu_tax)  hfa050t_calu_tax_amt," '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_tax_sub)  hfa050t_income_tax_sub_amt,"  '/* �ٷμҵ� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & "  empty_4,"  '/* �������� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.house_repay)  hfa050t_house_repay_amt,"  '/* �������� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.stock_save)  hfa050t_stock_save_amt,"  '/* �ٷ����ֽ����� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.fore_pay)  hfa050t_fore_pay_amt,"  '/* �ܱ����� */
'                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.long_stock_save_amt)  hfa050t_long_stock_save_amt,"  '/* �������� */
					lgStrSQL = lgStrSQL & " hfa050t.POLI_TAX_SUB hfa050t_poli_tax_sub ,"	'2004
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & "  empty_3,"
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.tax_sub_sum)  hfa050t_tax_sub_amt,"  '/* ���װ����� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_redu)  hfa050t_income_redu_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.taxes_redu)  hfa050t_taxes_redu_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000", "''", "S") & " empty,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.redu_sum)  hfa050t_redu_sum_amt,"  '/* ���װ���� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_income_tax)  hfa050t_dec_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_res_tax)  hfa050t_dec_res_tax_amt,"  '/*  */                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_farm_tax)  hfa050t_dec_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.dec_income_tax) + FLOOR(hfa050t.dec_farm_tax) + FLOOR(hfa050t.dec_res_tax))  dec_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_income_tax)  hfa050t_old_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_res_tax)  hfa050t_old_res_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_farm_tax)  hfa050t_old_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.old_income_tax) + FLOOR(hfa050t.old_farm_tax) + FLOOR(hfa050t.old_res_tax))  old_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_income_tax)  hfa050t_new_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_res_tax)  hfa050t_new_res_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_farm_tax)  hfa050t_new_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.new_income_tax) + FLOOR(hfa050t.new_farm_tax) + FLOOR(hfa050t.new_res_tax))  new_sum_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(Submit_dt, "''", "S") & "  prov_dt,"  '/* �������� -> �Էº��� */
                    lgStrSQL = lgStrSQL & " SPACE(1)  mod_code,"   '/*  */
                    lgStrSQL = lgStrSQL & " SPACE(4)  empty_2,"    '/* ���� */
                    lgStrSQL = lgStrSQL & " haa010t.emp_no  haa010t_emp_no" '/* ��� */                '�ٲ� Query
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa030t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hfa100t,"
                    lgStrSQL = lgStrSQL & " (SELECT emp_no, COUNT(*) emp_before_count"
                    lgStrSQL = lgStrSQL & " FROM hfa040t"                    
                    lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
                    lgStrSQL = lgStrSQL & " GROUP BY emp_no) AS before,"
                    lgStrSQL = lgStrSQL & " (SELECT a.emp_no, a.pay_yymm, a.pay_tot_amt"
                    lgStrSQL = lgStrSQL & " FROM hdf070t a"
                    lgStrSQL = lgStrSQL & " WHERE a.pay_yymm = (SELECT MAX(b.pay_yymm)"
                    lgStrSQL = lgStrSQL & " FROM hdf070t b"
                    lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no"  
                    lgStrSQL = lgStrSQL & " AND b.pay_yymm LIKE " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("%", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND b.prov_type = " & FilterVar("1", "''", "S") & " "  '/* �����ױ޿� */
                    lgStrSQL = lgStrSQL & " AND b.pay_tot_amt <> 0)"
                    lgStrSQL = lgStrSQL & " AND a.prov_type = " & FilterVar("1", "''", "S") & " ) AS t"                    
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= t.emp_no"                                 '�ٲ� Query
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= before.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.year_yy *= hfa030t.yy"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= hfa030t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"  '/* ������ ���� */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode
                Case "D"
                    lgStrSQL = " SELECT " & FilterVar("D", "''", "S") & "   record_type,"                                  '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"                               '/* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " SPACE(50)  empty,"                              '/* ���� */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"                        '/* �ҵ����ֹε�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " hfa040t.a_comp_nm  comp_nm,"                    '/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " hfa040t.a_comp_no  comp_no,"                    '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_pay_tot_amt)  hfa040t_pay_tot,"         '/* �޿��Ѿ� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_bonus_tot_amt)  bonus_tot,"     '/* ���Ѿ� */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_after_bonus_amt) bonus_amt,"    '/* ������ */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa040t.a_pay_tot_amt) + FLOOR(hfa040t.a_bonus_tot_amt) + FLOOR(hfa040t.a_after_bonus_amt)) pay_tot," '/* �� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("", "''", "S") & "  work_no,"                                  '/* work_no : �����ٹ�ó�Ϸù�ȣ */
                    lgStrSQL = lgStrSQL & " SPACE(1)  mod_code,"                            '/* �ڷ�����ڵ� */
                    lgStrSQL = lgStrSQL & " SPACE(488)  empty_2"                            '/* ���� */   
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hfa040t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.year_yy = hfa040t.year_yy"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no = hfa040t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"                 '/* ������ ���� */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode
            End Select
    End Select             
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub
'============================================================================================================
' Name : SetFixSrting(�Է°�,�񱳹���,��ü����,��������,�������Ĺ���)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '�Է°��� ��������������� �Է°��� ���̸� 0���� �Ѵ�.
        Cnt = 0     
    Else
        
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                If Trim(strMid) = ")" Or Trim(strMid) = "(" Then
                    MCnt = MCnt + 1
                Else
                    MCnt = MCnt + 2                                                  '�ѱۺκи� ���̸� ���� 2���Ѵ�.
                End If                                                '�ѱۺκи� ���̸� ���� 2���Ѵ�.
            Else
                MCnt = MCnt + 1
            End If
        Next
        Cnt = MCnt
                         
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '�񱳹��ڰ� ������� 
        Else                                                                         '�񱳹��ڰ� �����Ұ�� �񱳹��ڸ� �� �������� �Է°������Ѵ�.
            ExSymbol = Split(InValue,ComSymbol)
            If UBound(ExSymbol) > 0 Then
                iDx = UBound(ExSymbol)
                For i = 0 To iDx
                    strSplit = strSplit & ExSymbol(i)
                Next
                InValue = strSplit
            End If               
        End If        
    End If        
    
    If InPos = "" Then                                                              '�������̰� �������� ���� ��� �Է¹��� ���̰� �������̰� �ȴ�.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '���������ϰ��(defalut) �������� ���� ���� ������ ���ڰ� �ԷµǸ� ������ ����(default)�κ��� ��ü���ڷ� ü���.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '���������� 
        If InPos > Cnt Then                                                           ' ex:     hi
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = 1 To (InPos - Cnt)
                strTemp = strTemp & strFix         
            Next
            InValue = strTemp & InValue
        End If
    End If
    SetFixSrting = InValue 
End Function

%>

<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>
