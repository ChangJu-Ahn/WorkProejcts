
//종합 현황 (체결 계약 구분)
function fn_NtotalContract(title,value){
		$.ajax({
			url : "../report/openReportView.do",
			type : "post",
			data : value,
			dataType : "json",
				success:function(responseData){
                  var data = "";
                  var data1 = [];
                
                  data = responseData.graph;
                  //다시 가공 해줘야됨 ㅠㅠㅠ 왜 그러지?
                  $.each(responseData.grid, function(index, value){                	  
					  data1.push({year: value.year, count : value.count});					  
				  });
	                
	      		  $("#title").html(title);
	      		  $("#Graph").css("float","left");
	      		  $("#Grid").css("float","right");
	      		  
   				  //바 차트
   		          var margin = {top: 30, right: 10, bottom: 100, left: 60},
   		               width = 600 - margin.left - margin.right,
   		               height = 600 - margin.top - margin.bottom;
   			      
   		          //Canvas 그리기.
   		          var svg = d3.select("#Graph").append("svg")
   		                .attr("width", width + margin.left + margin.right)
   		                .attr("height", height + margin.top + margin.bottom)
   		                .append("g")
   		                .attr("transform", "translate(" + margin.left + "," + margin.top + ")");
   				  
   		          //축 셋팅
   		          var x = d3.scale.ordinal().rangeRoundBands([0, width], .1);
   		          var y = d3.scale.linear().range([height, 0]);
   		                    

     			  var xAxis = d3.svg.axis()
   				                .scale(x)
   				                .orient("bottom")
   				                .outerTickSize(1);


   				  var yAxis = d3.svg.axis()
   				                .scale(y)
   				                .orient("left")
   				                .outerTickSize(0)
   				                .innerTickSize(-width)
   				                .ticks(5);
   		          
   		          
   		          x.domain(data.map(function(d) { return d.GUBUN; }));
   		          y.domain([0, d3.max(data.map(function(d) { return d.G_COUNT; }))+10]);	          
   					
   		          svg.append("g").attr("class", "x axis")
   		                .attr("transform", "translate(0," + height + ")")
   		                .call(xAxis);

   		          svg.append("g").attr("class", "y axis")
   		          		.call(yAxis)
   		          		.append("text")
   		          		.attr("x", -15)
   		          		.attr("y", -15)
   		          		.text("건수");
   				  
   		          //바 셋팅
   		          svg.selectAll(".bar")
   		          		.data(data)
   		          		.enter()
   		                .append("rect")
   		                .attr("class", "bar")
   		                .attr("fill", "#2ECCFA")
   		                .attr("x", function(d) { return x(d.GUBUN); })
   		                .attr("width", x.rangeBand())
   		                .attr("y", function(d) { return y(d.G_COUNT); })
   		                .attr("height", function(d) { return height - y(d.G_COUNT); });
   				  
   		          
   		          //바 위에 건수 표시
   		          svg.append("g")
   		          		/*.attr("transform", "translate(" + margin.left + "," + margin.top + ")")*/
   		          		.selectAll(".textlabel")
   		         		.data(data)
   		          		.enter()
   		          		.append("text")
   		          		.attr("class", "textlabel")
   		         		.attr("x", function(d){ console.log(data.length); return x(d.GUBUN) + x.rangeBand()/2 - margin.right;})
   		          		.attr("y", function(d){ return y(d.G_COUNT) - 5; })
   		         	    .text(function(d){ return d.G_COUNT; });
   		          
   		          //x축 텍스트 로테이션
   		          svg.selectAll(".x.axis text")
			   		      .style("text-anchor", "end")
				          .attr("dx", "-.8em")
				          .attr("dy", ".15em")
				          .attr("transform", "rotate(-45)" );
   		          
   		          //차트 제목 표시
   		          /* svg.append("g")
   		          		.attr("transform", "translate(" + (width/2) + ", 15)")
   		          		.append("text")
   		         		.text("체결 계약 구분")
   		          		.style({"text-anchor":"middle", "font-family":"Arial", "font-weight":"800"}); */
   				  		
   		          //그리드		
   			      var thead = d3.select("#Grid")
   			          	  .append("table")
   			          	  .attr("class", "table")
   			          	  .append("thead").append("tr")
   			          	  .selectAll("th")
   			          	  .data(["연도","체결건수"])
   			          	  .enter()
   			          	  .append("th")
   			          	  .text(function(d) {  
   			          	    return d;
   			          	  });
   		
   			      var tbody = d3.select("#Grid table").append("tbody");

   	              tbody.selectAll("tbody")
   		         	  .data(data1)
   		         	  .enter()
   		         	  .append("tr")
   		         	  .selectAll("td")
   		         	  .data(function(d) {   		         		
   		         		  return d3.values(d);
   		         	   })
   		         	  .enter()
   		         	  .append("td")
   		         	  .text(function(d, i) {
   		         		  return d;
   		         	  });		         	  
   	          		
   		          var tfoot = d3.select("#Grid table").append("tfoot");
   		          
   		          var total = d3.sum(data1,function(d){return parseInt(d.count)});
   		          
   		          var data2 = [
   		            {
   		            	"TITLE" : "합계",  
   		            	"TOTAL" : total
   		          	}
   		          ];
   	      		  
   		          tfoot.selectAll("tfoot")           	
   			         	  .data(data2)
   			         	  .enter()
   			         	  .append("tr")
   			         	  .selectAll("td")
   			         	  .data(function(d) {
   			         	    return d3.values(d);
   			         	   })
   			         	  .enter()
   			         	  .append("td")
   			         	  .text(function(d, i) {
   			         	    return d;
   			         	  });
	            }		
		});
	}



//기간(연도)별 계약 현황
function fn_NperiodContract(title,value){
		/*var data = [
		            {"group":"15_1Q","gubun":"정리대상","value":15},
		       	    {"group":"15_2Q","gubun":"정리대상","value":4},
		            {"group":"15_3Q","gubun":"정리대상","value":8},
		            {"group":"15_4Q","gubun":"정리대상","value":5},
		            {"group":"15_1Q","gubun":"신규체결","value":7},
		            {"group":"15_2Q","gubun":"신규체결","value":8},
		            {"group":"15_3Q","gubun":"신규체결","value":10},
		            {"group":"15_4Q","gubun":"신규체결","value":3}
		        ];*/
	    $.ajax({
			url : "../report/openReportView.do",
			type : "post",
			data : value,
			dataType : "json",
				success:function(responseData){
					
					var data = "";
					
					data = responseData.graph;					
					$("#title").html(value.year+'년도 계약현황');
					
					$('#year').val("");
					
					var margin = {top: 30, right: 10, bottom: 40, left: 10},
			   		                width = 600 - margin.left - margin.right,
			   		                height = 500 - margin.top - margin.bottom;

				    var color = d3.scale.ordinal()
				    			  .range(["#1f4e79", "#e51843", "#73ad21", "#6b486b", "#a05d56", "#d0743c", "#ff8c00"]);
					// svg 생성
					// id = "chart" 인 div 태그의 하위 element로 svg 생성
					// 차트의 크기에다 margin값을 더해서 svg 전체의 면적을 결정한다
					// x축은 margin.left, y축은 margin.top만큼 평행이동 시켜서 이후의 차트는 margin값과 상관없이 그릴 수 있도록 한다
					var svg = d3.select("#Graph")
					        .append("svg")
					        .attr("width", width + margin.right + margin.left)
					        .attr("height", height + margin.top + margin.bottom)
					        .append("g")
					        .attr("transform", "translate("+margin.left+","+margin.top+")");

					// 각 그룹의 위치를 결정하는 scale 함수
					var x_scale_group = d3.scale.ordinal()
					                    .rangeRoundBands([0, width], 0.1);

					// 그룹 내에서 막대의 위치를 결정하는 scale 함수
					var x_scale_bar = d3.scale.ordinal();

					// 막대의 높이를 결정하는 scale 함수
					// y축의 값이 위에는 높은값, 아래에는 0이 있어야 한다. 그렇기 때문에 range에 들어가는 배열에서 값의 순서가 바뀐다
					var y_scale = d3.scale.linear()
					                .range([height, 0]);

					// x축을 생성하는 함수
					var xAxis = d3.svg.axis()
					                .scale(x_scale_group)
					                .orient("bottom")

					// y축을 생성하는 함수
					// 막대 뒤에 선을 표시하기 위해서 innerTickSize에 -width 값을 부여한다
					// 원래는 y축 눈금을 만드는 함수인데 -값을 주면 반대방향(차트 안쪽)으로 눈금이 길어진다
					var yAxis = d3.svg.axis()
					                .scale(y_scale)
					                .orient("left")
					                .outerTickSize(0)
					                .innerTickSize(-width)
					                .ticks(5);

				    // data를 group 값을 기준으로 그룹 지정한다
				    var nested = d3.nest()
				            .key(function(d){return d.group})
				            .entries(data)
				
				    // 그룹명(json 데이터에서 group 값)만 추출한다
				    var group_name = nested.map(function(d){return d.key});

				    // x_scale_group의 domain에 group_name 변수를 지정한다
				    // 그룹 수에 맞게 막대 그룹 개수 설정
				    x_scale_group.domain(group_name);

				    // x_scale_bar의 domain에 각 막대에 해당하는 값을 넣는다
				    // x_scale_group처럼 배열에서 값을 추출해서 쓰는 것이 더 좋다
				    // rangeRoundBands([interval], padding)의 형태로 사용했다
				    // padding의 값은 0과 1 사이의 값이 들어가는데 0.5일 때 막대의 너비와 막대 사이의 간격이 같아진다
				    x_scale_bar.domain(["자동연장", "신규체결"])
				               .rangeRoundBands([0, x_scale_group.rangeBand()], 0.1);

				    // y_scale의 domain 범위를 지정한다
				    // 막대 그래프의 아래쪽 눈금은 0부터 시작하는 것이 바람직하기 때문에 배열의 첫 번째 값은 0으로 한다
				    // 두 번째 값은 data의 value 값중에서 가장 큰 값을 찾고 그 값에서 임의의 값(여기서는 10)을 더해서 그래프 상단의 공간이 여유가 있도록 한다
				    y_scale.domain([0, d3.max(data.map(function(d){
				                return +d.value + 10; }))
				            ]);
			    
				    // x축을 생성한다
				    // 그냥 만들면 (0,0) 위치 (화면 좌측 상단)에 생기기 때문에 margin.left 만큼 x축을 이동시키고 차트의 높이만큼 아래로 내려준다
				     svg.append("g")
				        .attr("class", "x axis")
				        .attr("transform", "translate("+ margin.left +","+ height +")")
				        .call(xAxis);

				    // y축을 생성한다
				    // margin.left만큼 우측으로 평행이동 시킨다
				    svg.append("g")
				        .attr("class", "y axis")
				        .attr("transform", "translate("+ margin.left +","+ 0 +")")		        
				        .call(yAxis)
				        .append("text")
			      		.attr("x", -15)
			      		.attr("y", -15)
			      		.text("건수");
				        
				    // 새로 생성되는 부분(enter)에 대한 코드
				    // 막대의 x축 위치는 x_scale_group으로 생성되는 그룹 위치와 x_scale_bar로 생성되는 막대 위치를 더하고 margin.left를 더한다
				    // 막대의 y축 위치(막대가 시작되는 좌표)는 일반 막대그래프와 마찬가지로 value값을 y_scale로 변형시키면 된다
				    // width의 값은 x_scale_bar.rangeBand()값을 그래도 사용한다
				    // 막대의 높이는 y_scale(d.value)부터 시작해서 y축의 0 위치까지 와야 한다. 
				    // 따라서 막대가 시작되는 지점의 y좌표와 막대의 길이를 더하면 차트의 높이가 되어야 한다
				    // y의 값이 y_scale(d.value) 이니까 height값은 height - y_scale(d.value)
				    svg.append("g")
				        .selectAll(".bar")
			      		.data(data)
			      		.enter()
			            .append("rect")
			            .attr("class", "bar")
				        .attr("x", function(d){ 
				        	return margin.left + 
				                x_scale_group(d.group) + 
				                x_scale_bar(d.gubun); 
				            })
				        .attr("y", function(d){ return y_scale(d.value)})
				        .attr("width", function(d){ return x_scale_bar.rangeBand();})
				        .attr("height", function(d){ return height - y_scale(d.value); })
				        .style("fill", function(d) { return color(d.gubun); });
				    
				    //바 위에 건수 표시
			        svg.append("g")
			      		.selectAll(".textlabel")
			     		.data(data)
			      		.enter()
			      		.append("text")
			      		.attr("class", "textlabel")
			     		.attr("x", function(d){ 
			     			return margin.left + 
					                x_scale_group(d.group) + 
					                x_scale_bar(d.gubun) + 
					                x_scale_bar.rangeBand()/2 - margin.right;
				            })
			      		.attr("y", function(d){ return y_scale(d.value) - 5; })
			     	    .text(function(d){ return d.value; });
				    
				    
				    
					var nested = d3.nest()
			        			   .key(function(d){return d.gubun})
			        			   .entries(data)
				    
				    var group_gubun = nested.map(function(d){return d.key});
						    
				    
				    var legend = svg.selectAll(".legend")
				          .data(group_gubun)
					      .enter().append("g")
					      .attr("class", "legend")
					      .attr("transform", function(d, i) { return "translate(0," + i * 20 + ")"; });

					  legend.append("rect")
					      .attr("x", width - 18)
					      .attr("width", 18)
					      .attr("height", 18)
					      .style("fill",color);

					  legend.append("text")
					      .attr("x", width - 24)
					      .attr("y", 9)
					      .attr("dy", ".35em")
					      .style("text-anchor", "end")
					      .text(function(d) { return d; });
				
				}
	    });				    
}

function fn_AtotalContract(title,value){
	$.ajax({
		url : "../report/openReportView.do",
		type : "post",
		data : value,
		dataType : "json",
			success:function(responseData){
				console.log(responseData.grid);
				var data = [];
				
				$("#title").html(title);
				
				$.each(responseData.grid, function(index, value){                	  
					  data.push({"대분류"    		: value.CONTRACT_NM, 
						  		 "SEMI"    		: value.SEMI,
						  		 "AMC"     		: value.AMC,
						  		 "DISPLAY" 		: value.DISPLAY,
						  		 "ENC"     		: value.ENC,
						  		 "EM"      		: value.EM,
						  		 "LED"     		: value.LED,
						  		 "Rigmah"  		: value.Rigmah,
						  		 "기술원"    		: value.기술원,
						  		 "SOLVR"   		: value.SOLVR,
						  		 "관리지원"		: value.관리지원,
						  		 "TOTAL"   		: value.TOTAL
					  });					  
				  });
				
				var thead = d3.select("#Grid")
		          	  .append("table")
		          	  .attr("class", "table")
		          	  .append("thead").append("tr")
		          	  .selectAll("th")
		          	  .data(d3.keys(data[0]))
		          	  .enter()
		          	  .append("th")
		          	  .text(function(d) {  
		          	    return d;
		          	  });
	
				var tbody = d3.select("#Grid table").append("tbody");
				
				
	            tbody.selectAll("tbody")
		         	  .data(data)
		         	  .enter()
		         	  .append("tr")
		         	  .selectAll("td")
		         	  .data(function(d) {   		         		
		         		  return d3.values(d);
		         	   })
		         	  .enter()
		         	  .append("td")
		         	  .text(function(d, i) {
		         		  return d;
		         	  });
	            
	            var tfoot = d3.select("#Grid table").append("tfoot");
 		          
	            var s_total = d3.sum(data,function(d){return parseInt(d.SEMI)});
	            var a_total = d3.sum(data,function(d){return parseInt(d.AMC)});
	            var d_total = d3.sum(data,function(d){return parseInt(d.DISPLAY)});
	            var e_total = d3.sum(data,function(d){return parseInt(d.ENC)});
	            var m_total = d3.sum(data,function(d){return parseInt(d.EM)});
	            var l_total = d3.sum(data,function(d){return parseInt(d.LED)});
	            var r_total = d3.sum(data,function(d){return parseInt(d.Rigmah)});
	            var g_total = d3.sum(data,function(d){return parseInt(d.기술원)});
	            var sol_total = d3.sum(data,function(d){return parseInt(d.SOLVR)});
	            var nc_total = d3.sum(data,function(d){return parseInt(d.관리지원)});
	            var t_total = d3.sum(data,function(d){return parseInt(d.TOTAL)});	            
	            
	          
	            var data2 = [
	              {
	            	"TITLE" : "합계",  
	            	"s_total" : s_total,
	            	"a_total" : a_total,
	            	"d_total" : d_total,
	            	"e_total" : e_total,
	            	"m_total" : m_total,
	            	"l_total" : l_total,
	            	"r_total" : r_total,
	            	"g_total" : g_total,
	            	"sol_total" : sol_total,
	            	"nc_total" : nc_total,
	            	"t_total" : t_total
	          	  }
	            ];
      		  
	            tfoot.selectAll("tfoot")           	
		         	  .data(data2)
		         	  .enter()
		         	  .append("tr")
		         	  .selectAll("td")
		         	  .data(function(d) {
		         	    return d3.values(d);
		         	   })
		         	  .enter()
		         	  .append("td")
		         	  .text(function(d, i) {
		         	    return d;
		         	  }); 
	         	            
			}
		});
}

function fn_AperiodContract(title,value){
	$.ajax({
		url : "../report/openReportView.do",
		type : "post",
		data : value,
		dataType : "json",
			success:function(responseData){
				console.log(responseData.grid);
				
				var data = [];
				
				$.each(responseData.grid, function(index, value){                	  
					  data.push({"사업부"   : value.CODE_NM, 
						  		 "구분"     : value.gubun,
						  		 "Q1"      : value.Q1,
						  		 "Q2"      : value.Q2,
						  		 "Q3"      : value.Q3,
						  		 "Q4"      : value.Q4
					  });					  
				  });
				
				$("#title").html(value.year+'년도 계약현황');
				$('#year').val("");
				
				var thead = d3.select("#Grid")
				          	  .append("table")
				          	  .attr("class", "table")
				          	  .append("thead").append("tr")
				          	  .selectAll("th")
				          	  .data(d3.keys(data[0]))
				          	  .enter()
				          	  .append("th")
				          	  .text(function(d) {  
				          	    return d;
				          	  });

				
			    var tbody = d3.select("#Grid table").append("tbody");
			   
			    var nested = d3.nest()
				 			   .key(function(d){return d.사업부;})
				 			   .entries(data);

			    nested.forEach(function (d) {
				    var rowspan = d.values.length;
				    d.values.forEach(function (val, index) {
				        var tr = tbody.append("tr");
				        if (index == 0) {
				            tr.append("td")
				                .attr("rowspan", rowspan)
				                .text(val.사업부);
				        }
				        tr.append("td")
				            .text(val.구분);
				        tr.append("td")
				            .text(val.Q1);
				        tr.append("td")
				            .text(val.Q2);
				        tr.append("td")
			            	.text(val.Q2);
				        tr.append("td")
			            	.text(val.Q2);
				    });
				 			   
			   });   
				 			 
	           var tfoot = d3.select("#Grid table").append("tfoot");
	           
	           /*
	            	하드 코딩 시작ㅠㅠㅠㅠㅠ 고쳐보자!!!!!!
	           */
	           var q1_Atotal = d3.sum(data,function(d){ 						
						if(d.구분 == '자동연장'){
							return parseInt(d.Q1)
						}else
							return 0
						});
	           var q2_Atotal = d3.sum(data,function(d){ 						
						if(d.구분 == '자동연장'){
							return parseInt(d.Q2)
						}else
							return 0
						});
	           var q3_Atotal = d3.sum(data,function(d){ 
						if(d.구분 == '자동연장'){
							return parseInt(d.Q3)
						}else
							return 0
						});
	           var q4_Atotal = d3.sum(data,function(d){ 						
						if(d.구분 == '자동연장'){
							return parseInt(d.Q4)
						}else
							return 0
						});
	                    
	           var q1_Btotal = d3.sum(data,function(d){ 						
						if(d.구분 == '신규체결'){
							return parseInt(d.Q1)
						}else
							return 0
						});
	           var q2_Btotal = d3.sum(data,function(d){ 						
			 			if(d.구분 == '신규체결'){
							return parseInt(d.Q2)
						}else
							return 0
						});
	           var q3_Btotal = d3.sum(data,function(d){ 
						if(d.구분 == '신규체결'){
							return parseInt(d.Q3)
						}else
							return 0
						});
	          var q4_Btotal = d3.sum(data,function(d){ 						
						if(d.구분 == '신규체결'){
							return parseInt(d.Q4)
						}else
							return 0
						});
	           
	          
	           var data2 = [
	          	  {
	            	"TITLE" : "합계",
	            	"GUBUN" : "신규체결",
	            	"Q1" : q1_Btotal,
	            	"Q2" : q2_Btotal,
	            	"Q3" : q3_Btotal,
	            	"Q4" : q4_Btotal
		          },
		          {
	            	"TITLE" : "",
	            	"GUBUN" : "자동연장",
	            	"Q1" : q1_Atotal,
	            	"Q2" : q2_Atotal,
	            	"Q3" : q3_Atotal,
	            	"Q4" : q4_Atotal
		          }
	            ];
     		   	           
	           tfoot.selectAll("tfoot")           	
	         	  .data(data2)
	         	  .enter()
	         	  .append("tr")
	         	  .selectAll("td")
	         	  .data(function(d) {
	         	    return d3.values(d);
	         	   })
	         	  .enter()
	         	  .append("td")
	         	  .text(function(d, i) {
	         	    return d;
	         	  }); 
			}
	});
}
