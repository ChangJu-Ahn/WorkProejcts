using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web; 
using NEPES_CONTEXT;

namespace NEPES_CONTEXT
{
    //class example
    public class example
    {
        public void ExampleFunction()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            int rtnCnt = 0;

            string connectionString = "Data Source=192.168.10.15;Initial Catalog=nepes;User ID=sa;Password=nepes01!";
            //string connectionString = ConfigurationManager.ConnectionStrings["nepes"].ConnectionString;                         //DB의 커넥션 정보
            string exampleQuery = "SELECT * FROM T_IF_RCV_PROD_RSLT_KO441 WHERE PLANT_CD = @data AND REPORT_DT >= @userNo";     //쿼리, 작성할 때 파라메터 자리에 "@이름" 순으로 적고 문자열이라고 하더라도, 따옴표를 쓰지 않는다.
            string exampleQuery1 = "SELECT * INTO T_IF_AHNCJ_TEST FROM T_IF_RCV_PROD_RSLT_KO441 WHERE PLANT_CD = @data AND REPORT_DT >= @userNo";     //쿼리, 작성할 때 파라메터 자리에 "@이름" 순으로 적고 문자열이라고 하더라도, 따옴표를 쓰지 않는다.

            userSqlParams[] uParams = new userSqlParams[2];  //SQL의 파라메터 변수 값이다. 파라메터를 몇개 쓸 것이냐에 따라 변경하여 사용한다.

            //첫번째 파라메터(@data)
            uParams[0] = new userSqlParams();               //객체를 생성하기 위해 생성자를 호출한다.
            uParams[0].KeyName = "data";                    //위 @data 파라메터와 이름이 같아야 한다. 
            uParams[0].ColumnType = SqlDbType.NVarChar;     //해당 컬럼의 Type
            uParams[0].ColunmSize = 4;                      //컬럼 size
            uParams[0].Value = "P09";                       //실제 파라메터에 대입되어야 할 값

            //두번째 파라메터(@userNo)
            uParams[1] = new userSqlParams();               //객체를 생성하기 위해 생성자를 호출한다.
            uParams[1].KeyName = "userNo";                  //위 @userNo 파라메터와 이름이 같아야 한다. 
            uParams[1].ColumnType = SqlDbType.NVarChar;     //해당 컬럼의 Type
            uParams[1].ColunmSize = 8;                      //컬럼 size
            uParams[1].Value = "20171001";                  //실제 파라메터에 대입되어야 할 값

            //DB 커넥션 이후 자동으로 Dispose(메모리 반환)을 하기 위해 using문 선언
            using (NepesDBContext context = new NepesDBContext(connectionString))
            {
                //DB작업이므로 예외구문 추가
                //내부 DB작업 구조에서 예외가 발생될 경우 예외를 전파해주기 때문에 여기에서의 catch문으로 
                //오류내용이 전달된다.
                try
                {
                    //반환을 데이터셋으로 반환하기를 원한다면 GetDataSet 사용
                    ds = context.GetDataSet(exampleQuery, uParams);

                    //반환을 데이터 테이블로 반한하기를 원한다면 GetDataTable 사용
                    dt = context.GetDataTable(exampleQuery, uParams);

                    rtnCnt = context.ActionSqlQuery(exampleQuery1, uParams);
                }
                catch (Exception ex)
                {
                    string msg = ex.ToString();
                    Console.Write(msg);
                }
            }
        }

    }
}

