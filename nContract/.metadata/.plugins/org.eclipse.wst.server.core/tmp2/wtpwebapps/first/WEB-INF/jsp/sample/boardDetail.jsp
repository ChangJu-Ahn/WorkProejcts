<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
//뒤로가기 버튼을 사용하기 위해 캐쉬 제거
    response.setDateHeader("Expires", 0);
    response.setHeader("Pragma", "no-cache");
    if(request.getProtocol().equals("HTTP/1.1")) {
        response.setHeader("Cache-Control", "no-cache");
    }
%>
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>체결 계약서</H2>
    <table class="board_view">
			<colgroup>
				<col width="15%"/>
				<col width="25%"/>
				<col width="15%"/>
				<col width="15%"/>
				<col width="15%"/>
				<col width="15%"/>
			</colgroup>
			<tbody>
				<tr>
					<th scope="row">계약번호</th>
					<td colspan="2">${map.CONTRACT_NO }</td>
					<th scope="row">계약번호 최종이력(No)</th>
					<td colspan="2">${map.CONTRACT_NO }</td>
				</tr>
				<tr>
					<th scope="row">사업부</th>
					<td colspan="2">${map.BU_CODE }</td>
					<th scope="row">해지조건</th>
					<td colspan="2">${map.EXPIRE_CONDITION }</td>
				</tr>
				<tr>
					<th scope="row">고객사(업체)명</th>
					<td colspan="2">${map.BP_CD1 }</td>
					<th scope="row">구분</th>
					<td colspan="2">${map.BP_TYPE }</td>
				</tr>
				<tr>
					<th scope="row">고객사(업체)명2</th>
					<td colspan="2">${map.BP_CD2 }</td>
					<th scope="row">계약상태</th>
					<td colspan="2">
						<c:choose>
                       		<c:when test="${map.EXPIRE_FLAG == 'Y' && map.MODIFY_FLAG == 'Y'}">
                       			해지 / 해지일 :${map.END_DT }
                       		</c:when>
                       		<c:when test="${map.EXPIRE_FLAG == 'N' && map.MODIFY_FLAG == 'Y'}">
                       			변경
                       		</c:when>
                       		<c:otherwise>
                       			등록
                       		</c:otherwise>
	                       </c:choose>
	                </td>
				</tr>
				<tr>
					<th scope="row">계약서 명</th>
					<td colspan="5">${map.CONTRACT_NM }
						<input type="hidden" id="CONTRACT_NO" name ="CONTRACT_NO" value="${map.CONTRACT_NO }">
					</td>
				</tr>
				<tr>
					<th scope="row">계약 구분</th>
					<td>${map.CONTRACT_TYPE }</td>
					<th scope="row">목적사업(제품)</th>
					<td colspan="3">${map.PURPOSE }</td>
				</tr>
				<tr>
					<th scope="row">효력발생일(최초)</th>
					<td>${map.FIRST_VALID_DT }</td>
					<th scope="row">기간만료일</th>
					<td>${map.EXPIRE_DT }</td>
					<th scope="row">자동연장</th>
					<td>${map.EXTEND_FLAG } 
						<c:if test="${not empty map.EXTEND_TERM }">
						/ ${map.EXTEND_TERM }
						</c:if>
					</td>
				</tr>
				<tr>
					<th scope="row">부속계약서</th>
					<td colspan="3">${map.P_CONTRACT }</td>
					<th scope="row">해지통지기간</th>
					<td>${map.EXPIRE_TERM }</td>
				</tr>
				<tr>
					<th scope="row">비고</th>
					<td colspan="5">${map.REMARK }</td>
				</tr>
				<tr>
					<th scope="row">변경이력</th>
					<td colspan="5">
						<c:forEach var="row" items="${h_list }" varStatus="var">
						 <div>
	                       <input type="hidden" id="HST_SEQ" value="${row.HST_SEQ }">
	                       <a href="#this" name="CONTRACT_NM">${row.HST_SEQ }.${row.CONTRACT_NM }</a>
	                       <c:choose>
	                       		<c:when test="${row.EXPIRE_FLAG == 'Y' && row.MODIFY_FLAG == 'Y'}">
	                       			(해지)
	                       		</c:when>
	                       		<c:when test="${row.EXPIRE_FLAG == 'N' && row.MODIFY_FLAG == 'Y'}">
	                       			(변경)
	                       		</c:when>
	                       		<c:otherwise>
	                       			(등록)
	                       		</c:otherwise>
	                       </c:choose>
	                       <br/>
	                     </div>
                    	</c:forEach></td>
				</tr>
				<tr>
					<th scope="row">변경 내용(최종)</th>
					<td colspan="5">${map.MODIFY_CONTENT }</td>
				</tr>
				<tr>
					<th scope="row">파일</th>
					<td colspan="5">
						<c:forEach var="row" items="${list }">
						<div>
	                       <input type="hidden" id="SEQ" value="${row.SEQ }">
	                       <a href="#this" name="file">${row.ORIGINAL_FILE_NAME }</a> 
	                       (${row.FILE_SIZE }kb)
	                       <br/>
	                    </div>
                    	</c:forEach>	
					</td>
				</tr>
				<tr>
					<th scope="row">기타</th>
					<td colspan="5">${map.SYS_REMARK }</td>
				</tr>
				<tr>
					<th scope="row">관리자 의견</th>
					<td colspan="5">${map.CONTENTS }</td>
                </tr>
			</tbody>
	</table>
    <br/>
     
    <!-- Role 불러 오기 -->
    <sec:authentication property="principal.username"  var ="currentUsername"/>
    <sec:authorize access="hasRole('ROLE_ADMIN')" var="ROLE_ADMIN"></sec:authorize>
    
    <div style="width: 60%;list-style: inside; text-align: center; margin: auto;"> 	
    	<div id="CKEditor">
		</div>    		
    	<br/>
    	<div style="text-align: left; margin:auto; padding-bottom: 50px;">
    		<c:if test="${map.INSERT_USER == currentUsername || ROLE_ADMIN == true}">
    			<a href = "#this" class="btn" id="delete">삭제하기</a>
	    		<c:if test="${map.EXPIRE_FLAG != 'Y'}">	
	    			<a href = "#this" class="btn" id="update">수정하기</a>
	   			</c:if>
    		</c:if>
    		<a href = "#this" class="btn" id="list">목록으로</a>
    	</div>    	    	
    </div>
 	<img alt="Loading" id="imgLoading" src="../images/loading_spinner.gif" class="loadingSpinner_Layer">
     
    <%@ include file="/WEB-INF/include/include-body.jspf" %>
    <script type="text/javascript">
    	var CONTRACT_NO = $("#CONTRACT_NO").val();
    	var comSubmit;
    	
         $(document).ready(function(){
        	
        	if('${ROLE_ADMIN}' == 'true'){
        		var str = "";
        		
        		str  = "<h1 style='text-align: left;'>관리자 의견 입력</h1>";
        		str += "<textarea title='내용' id='CONTENTS' name='CONTENTS'></textarea>";
        		str += "<p style='text-align: left;'><input type='button' id ='save' value='의견입력'></p>";
        		
        		$("#CKEditor").append(str);
        		
        		CKEDITOR.replace("CONTENTS",{width:'100%',height:'350px'});
        	}
        	
            $("#list").on("click", function(e){ 
                e.preventDefault();
                fn_openBoardList();
            });
             
            $("#update").on("click", function(e){            	
                e.preventDefault();
                fn_openBoardUpdate();
            }); 
            
            $("#delete").on("click", function(e){
            	Lobibox.confirm({
            		msg: "해당 계약서의 모든 정보가 삭제됩니다. 삭제하시겠습니까?",
        		    callback: function ($this, type, ev) {
        		        if(type == "yes"){
        		            e.preventDefault();
        		            $("#imgLoading").show();
            				$("#delete").attr("disabled",true);
        	                fn_deleteBoard();
        		        }
        		    }
            	});  
            });
                        
            $("a[name='file']").on("click", function(e){
                e.preventDefault();
                fn_downloadFile($(this));
            });
            
            $("a[name='CONTRACT_NM']").on("click", function(e){ 
                e.preventDefault();
                fn_openBoardHst($(this));
            });   
            
            $("input[id='save']").on("click", function(e){ 
                e.preventDefault();
                fn_saveContents();
            });
                     
        });
        
        function fn_saveContents(){      
            var CONTENTS = CKEDITOR.instances.CONTENTS.getData();
            comSubmit = new ComSubmit();
            
            comSubmit.setUrl("<c:url value='/sample/openUpdateContent.do' />");
            comSubmit.addParam("CONTENTS",CONTENTS);
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.submit();
        }
         
        function fn_openBoardList(){
        	comSubmit = new ComSubmit();
        	
            comSubmit.setUrl("<c:url value='/sample/openBoardList.do' />");
            comSubmit.submit();
        }
         
        function fn_openBoardUpdate(){
        	comSubmit = new ComSubmit();
        	
            comSubmit.setUrl("<c:url value='/sample/openBoardUpdate.do' />");
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.submit();
        }
        
        function fn_deleteBoard(){
        	comSubmit = new ComSubmit();
        	
            comSubmit.setUrl("<c:url value='/sample/deleteBoard.do' />");
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.addParam("userid", '${currentUsername}');
            comSubmit.submit();
        }
        
        function fn_openBoardHst(obj){
        	var HST_SEQ = obj.parent().find("#HST_SEQ").val();        
        	comSubmit = new ComSubmit();
        	
      		comSubmit.setUrl("<c:url value='/sample/openBoardHst.do' />");
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.addParam("HST_SEQ", HST_SEQ);
            comSubmit.submit();
        }
        
        function fn_downloadFile(obj){
	        var SEQ = obj.parent().find("#SEQ").val();
       		comSubmit = new ComSubmit();
        
        	comSubmit.setUrl("<c:url value='/common/downloadFile.do' />");
	        comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
	        comSubmit.addParam("SEQ", SEQ);
	        comSubmit.submit();	
    	}
    
    </script>
</body>
</html>