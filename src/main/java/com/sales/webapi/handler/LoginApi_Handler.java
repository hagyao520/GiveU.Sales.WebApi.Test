package com.sales.webapi.handler;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.json.JSONObject;
import org.testng.Assert;
import org.testng.Reporter;

import com.sales.webapi.util.HttpRequest_Util;
import com.sales.webapi.util.MobileApiTools_Util;
import com.sales.webapi.util.DataBase_Util;
import com.sales.webapi.util.Excel_Util;

public class LoginApi_Handler {
    private static Logger logger = Logger.getLogger(LoginApi_Handler.class);

    public static void GetExcelInstance() {
        logger.info(LoginApi_Handler.class);
        System.out.println(LoginApi_Handler.class);
		Excel_Util.getInstance().setFilePath("Excel/GiveU.Sales.WebApi.xlsx");
		Excel_Util.getInstance().setSheetName("Login");
    }
    
    //LoginApi_Handler.GetExcelInstance("GiveU.Sales.WebApi.xlsx", "Login");
    public static void GetExcelInstance(String FILE_PATH,String SHEET_NAME) {
        logger.info(LoginApi_Handler.class);
        System.out.println(LoginApi_Handler.class);
		Excel_Util.getInstance().setFilePath("src/test/resources/Excel/"+FILE_PATH);
		Excel_Util.getInstance().setSheetName(SHEET_NAME);
    }

    public static void InitializeExcelData() { 
        try {
            logger.info("初始化: " + Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName());
            System.out.println("初始化: " + Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName());
            MobileApiTools_Util.initializeData(6-1, "Run", "", 4);
            MobileApiTools_Util.initializeData(6-1, "$id_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "$id_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "phone_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "phone_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "roleId_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "roleId_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "roleName_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "roleName_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "salesId_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "salesId_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "userName_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "userName_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "identification_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "identification_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "photoName_Exp", "", 4);
            MobileApiTools_Util.initializeData(6-1, "photoName_Act", "", 4);
            MobileApiTools_Util.initializeData(6-1, "ActualResult", "", 4);
            MobileApiTools_Util.initializeData(6-1, "ResultCode", "", 4);
            MobileApiTools_Util.initializeData(6-1, "TestResult", "", 4);
            MobileApiTools_Util.initializeData(6-1, "RunningTime", "", 4);
            MobileApiTools_Util.initializeData(6-1, "Json", "", 4);
            MobileApiTools_Util.initializeData(6-1, "FailHint", "", 4);
            logger.info(Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName() + "初始化完成");
            System.out.println(Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName() + "初始化完成");
            System.out.println("==============================================================================================================================================");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void WriteExcelExpected(String DataVersion,int ID){ 
    	
    	int TITLE_LINE_INDEX =6;
    	List<Map<String, String>> data = null;
        data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);
        
        try {
        	if (data != null) {
//                for (int i = 0; i < data.size(); i++) {
        		  int i = ID;
        	      Map<String, String> map = data.get(i);
                  String Sql =map.get("userId");

                  logger.info("根据数据库查询结果, 开始写入预期值【Waiting...】");
                  System.out.println("根据数据库查询结果, 开始写入预期值【Waiting...】");
                  
//                  String $id_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"$id_Exp");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "$id_Exp", "2");
        	      
        	      String phone_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"phone");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "phone_Exp", phone_Exp);

        	      String roleId_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"roleId");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleId_Exp", roleId_Exp);
        	
        	      String roleName_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"roleName");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleName_Exp", roleName_Exp);
        	      
        	      String salesId_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"salesId");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "salesId_Exp", salesId_Exp);
        	      
        	      String userName_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"userName");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "userName_Exp", userName_Exp);
        	      
//        	      String identification_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"identification");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "identification_Exp", "string");
        	      
        	      String photoName_Exp= DataBase_Util.GetSqlResult(DataVersion, Sql,"photoName");
        	      MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "photoName_Exp", "http://10.10.11.136/"+photoName_Exp);
        	      
                 logger.info("对应用例预期值,写入成功【OK】");
                 System.out.println("对应预期值,写入成功【OK】");
                 System.out.println("==============================================================================================================================================");
               }
//            } 
        }catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    //LoginApi_Handler.InitializeExcelData(6, "Run", "$id_Act", "phone_Act", "roleId_Act", "roleName_Act", "salesId_Act", "userName_Act", "identification_Act", "photoName_Act","ActualResult", "ResultCode", "TestResult", "RunningTime");
    public static void InitializeExcelData(int TITLE_LINE_INDEX,String Column1,String Column2,String Column3,String Column4,String Column5,String Column6,
    		                                   String Column7,String Column8,String Column9,String Column10,String Column11,String Column12,String Column13) {

        try {
            logger.info("初始化: " + Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName());
            System.out.println("初始化: " + Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName());
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column1, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column2, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column3, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column4, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column5, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column6, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column7, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column8, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column9, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column10, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column11, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column12, "", 4);
            MobileApiTools_Util.initializeData(TITLE_LINE_INDEX-1, Column13, "", 4);
            logger.info(Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName() + "初始化完成");
            System.out.println(Excel_Util.getInstance().getFilePath() + ", " + Excel_Util.getInstance().getSheetName() + "初始化完成");
            System.out.println("==============================================================================================================================================");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void Login1(int ID) throws Exception {
    	
    	int TITLE_LINE_INDEX =6;
    	String PremiseUrl = "";
        String ApiUrl = "";
        String Act = "";
        String Method = "";
        List<Map<String, String>> data = null;
        boolean Flag = false;

        PremiseUrl = Excel_Util.getInstance().readExcelCell(1-1, 2-1);
        ApiUrl = Excel_Util.getInstance().readExcelCell(2-1, 2-1);
        Act = Excel_Util.getInstance().readExcelCell(3-1, 2-1);
        Method = Excel_Util.getInstance().readExcelCell(4-1, 2-1);       
        Flag = MobileApiTools_Util.isArgEquals(5-1, 2-1, TITLE_LINE_INDEX-1);

        if (PremiseUrl.equals("") ||ApiUrl.equals("") || Act.equals("") || Method.equals("") || !Flag) {
            logger.error("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.out.println("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.exit(-1);
        }

        data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);

        if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
            int i = ID-1;
            	
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
             
                //指定请求参数
                String Param ="userId=" + map.get("userId") + "&" + "password=" + map.get("password") + "&" + "type=" + map.get("type")+ "&" + "version=" + map.get("version")+ "&" + "identification=" + map.get("identification");

                //指定请求方式,API地址,请求参数,Cookie值, 发起请求
                String CookieVal = HttpRequest_Util.getPostcookie(PremiseUrl,Param);  
                Map<String, String> result = HttpRequest_Util.sendPostcookie(Method, ApiUrl, Param, CookieVal); 
                             
                //获取接口返回的Code,结果内容
                String Code = result.get("code");
                String RsTmp = result.get("result");

                //将RsTmp字符串转换为JSON
                JSONObject object_result = new JSONObject(RsTmp);
                String Result_Data = object_result.getString("data");
                
                //将Result_Data字符串转换为JSON
                JSONObject object_data = new JSONObject(Result_Data);
                String id_Act = object_data.getString("$id");
                String phone_Act = object_data.getString("phone");
                String roleId_Act = object_data.getString("roleId");
                String roleName_Act = object_data.getString("roleName");
                String salesId_Act = object_data.getString("salesId");
                String userName_Act = object_data.getString("userName");
                String identification_Act = object_data.getString("identification");
                String photoName_Act = object_data.getString("photoName");
                    
                //期望结果与实际结果比较
                String $id_Exp = MobileApiTools_Util.assertResult(map.get("$id_Exp"),id_Act);
                String phone_Exp = MobileApiTools_Util.assertResult(map.get("phone_Exp"),phone_Act);
                String roleId_Exp = MobileApiTools_Util.assertResult(map.get("roleId_Exp"),roleId_Act);
                String roleName_Exp = MobileApiTools_Util.assertResult(map.get("roleName_Exp"),roleName_Act);
                String salesId_Exp = MobileApiTools_Util.assertResult(map.get("salesId_Exp"),salesId_Act);
                String userName_Exp = MobileApiTools_Util.assertResult(map.get("userName_Exp"),userName_Act);
                String identification_Exp = MobileApiTools_Util.assertResult(map.get("identification_Exp"),identification_Act);
                String photoName_Exp = MobileApiTools_Util.assertResult(map.get("photoName_Exp"),photoName_Act);

                //写入Run列, 执行纪录
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Run", "Y");

                //写入ResultCode列,返回结果代码
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "ResultCode", Code);

                //设置ResultCode单元格颜色
                if (Integer.parseInt(Code) == 200)
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 1);            
                else
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 0);

                //写入实际结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "$id_Act", id_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "phone_Act", phone_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleId_Act", roleId_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleName_Act", roleName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "salesId_Act", salesId_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "userName_Act", userName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "identification_Act", identification_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "photoName_Act", photoName_Act);
                
                //写入测试结果单元格背景色
                if ($id_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "$id_Act",TITLE_LINE_INDEX + i, 2);}

                //写入测试结果单元格背景色
                if (phone_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "phone_Act",TITLE_LINE_INDEX + i, 0);}

                //写入测试结果单元格背景色
                if (roleId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (roleName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (salesId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "salesId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (userName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "userName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (identification_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "identification_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (photoName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "photoName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试通过与否,设置测试结果单元格背景色
//              &是位与操作，一定会执行； &&是逻辑与操作，如果&&的前面为false了，后面的就不会执行了。
//              |是位或操作、一定会执行； || 是逻辑或操作，如果||的前面为true了，||的后面就不会执行了
                if ($id_Exp.equals("OK")&phone_Exp.equals("OK")&roleId_Exp.equals("OK")&roleName_Exp.equals("OK")&
                	salesId_Exp.equals("OK")&userName_Exp.equals("OK")&identification_Exp.equals("OK")&photoName_Exp.equals("OK")){
                	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "OK");
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 1);}
                else{
                	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "NG");
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 0);}

                //写入执行时间
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "RunningTime", MobileApiTools_Util.getDate());
                
                //写入转换的JSON结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Json", RsTmp);
            } 
//        }  
    }	

    public static void Login(int ID) throws Exception {

        int TITLE_LINE_INDEX =6;
        String PremiseUrl = "";
        String ApiUrl = "";
        String Act = "";
        String Method = "";
        List<Map<String, String>> data = null;
        boolean Flag = false;

        PremiseUrl = Excel_Util.getInstance().readExcelCell(1-1, 2-1);
        ApiUrl = Excel_Util.getInstance().readExcelCell(2-1, 2-1);
        Act = Excel_Util.getInstance().readExcelCell(3-1, 2-1);
        Method = Excel_Util.getInstance().readExcelCell(4-1, 2-1);
        Flag = MobileApiTools_Util.isArgEquals(5-1, 2-1, TITLE_LINE_INDEX-1);

        if (PremiseUrl.equals("") ||ApiUrl.equals("") || Act.equals("") || Method.equals("") || !Flag) {
            logger.error("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.out.println("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.exit(-1);
        }

        data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);

        if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
        	
            //根据Excel列名称,获取单元格内容
            Map<String, String> map1 = data.get(0);
            String userId1 = map1.get("userId");
            String password1 = map1.get("password");
            String type1 = map1.get("type");
            String version1 = map1.get("version");
            String identification1 = map1.get("identification");
 
            int i = ID;
            //根据Excel列名称,获取单元格内容
            Map<String, String> map = data.get(i);
            String userId = map.get("userId");
            String password = map.get("password");
            String type = map.get("type");
            String version = map.get("version");
            String identification = map.get("identification");
            
            //指定请求参数
//            final String Param1 = "{" +
//                    "\"userId\": \"666666\",\"password\": \"612426\",\"type\": \"string\",\"version\": \"string\",\"identification\": \"string\"}";
            
            final String Param1 = "{" +
                    "\"userId\": \""+userId1+"\",\"password\": \""+password1+"\",\"type\": \""+type1+"\",\"version\": \""+version1+"\",\"identification\": \""+identification1+"\"}";
            
            final String Param = "{" +
                    "\"userId\": \""+userId+"\",\"password\": \""+password+"\",\"type\": \""+type+"\",\"version\": \""+version+"\",\"identification\": \""+identification+"\"}";
            
            //指定请求方式,API地址,请求参数, 发起请求,获取Cookie值
            Map<String, String> CookieVal = HttpRequest_Util.getPostCcookie(PremiseUrl,Param1);

            String RsTmp = HttpRequest_Util.GetJsonResult(ApiUrl, Param,CookieVal);
//          int Code = HttpRequest_Util.GetStatusCode(ApiUrl, Param1, CookieVal);
            int Code = HttpRequest_Util.GetJsonIntValue(ApiUrl, Param, CookieVal,"code");  
            String message = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"message");

            //写入Run列, 执行纪录
            MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Run", "Y");

            //写入ResultCode列,返回的结果代码
            MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "ResultCode", String.valueOf(Code));

            //设置ResultCode单元格颜色
            if (Code == 1000){
                Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 1);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "OK");
                Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 1);
                
                String id_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.$id");
                String phone_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.phone");
                String roleId_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.roleId");
                String roleName_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.roleName");
                int salesId_Act = HttpRequest_Util.GetJsonIntValue(ApiUrl, Param, CookieVal,"data.salesId");
                String userName_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.userName");
                String identification_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.identification");
                String photoName_Act = HttpRequest_Util.GetJsonStringValue(ApiUrl, Param, CookieVal,"data.photoName");

                //期望结果与实际结果比较
                String $id_Exp = MobileApiTools_Util.assertResult(map.get("$id_Exp"),id_Act);
                String phone_Exp = MobileApiTools_Util.assertResult(map.get("phone_Exp"),phone_Act);
                String roleId_Exp = MobileApiTools_Util.assertResult(map.get("roleId_Exp"),roleId_Act);
                String roleName_Exp = MobileApiTools_Util.assertResult(map.get("roleName_Exp"),roleName_Act);
                String salesId_Exp = MobileApiTools_Util.assertResult(map.get("salesId_Exp"), String.valueOf(salesId_Act));
                String userName_Exp = MobileApiTools_Util.assertResult(map.get("userName_Exp"),userName_Act);
                String identification_Exp = MobileApiTools_Util.assertResult(map.get("identification_Exp"),identification_Act);
                String photoName_Exp = MobileApiTools_Util.assertResult(map.get("photoName_Exp"),photoName_Act);
                
                //写入实际结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "$id_Act", id_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "phone_Act", phone_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleId_Act", roleId_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleName_Act", roleName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "salesId_Act", String.valueOf(salesId_Act));
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "userName_Act", userName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "identification_Act", identification_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "photoName_Act", photoName_Act);

                //写入测试结果单元格背景色
                if ($id_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "$id_Act",TITLE_LINE_INDEX + i, 0);}

                //写入测试结果单元格背景色
                if (phone_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "phone_Act",TITLE_LINE_INDEX + i, 0);}

                //写入测试结果单元格背景色
                if (roleId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (roleName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (salesId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "salesId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (userName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "userName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (identification_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "identification_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (photoName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "photoName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试通过与否,设置测试结果单元格背景色
//              &是位与操作，一定会执行； &&是逻辑与操作，如果&&的前面为false了，后面的就不会执行了。
//              |是位或操作、一定会执行； || 是逻辑或操作，如果||的前面为true了，||的后面就不会执行了
                if ($id_Exp.equals("OK")&phone_Exp.equals("OK")&roleId_Exp.equals("OK")&roleName_Exp.equals("OK")&
                        salesId_Exp.equals("OK")&userName_Exp.equals("OK")&identification_Exp.equals("OK")&photoName_Exp.equals("OK")){
                    MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "OK");
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 1);}
                else{
                	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "NG");
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 0);}
                
                //写入执行时间
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "RunningTime", MobileApiTools_Util.getDate());

                //写入转换的JSON结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Json", RsTmp);
                
            }else{
                Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 0);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "NG");
            	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 0);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "FailHint", "【message】"+message+"");
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "RunningTime", MobileApiTools_Util.getDate());
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Json", RsTmp);
            }   
        }
//        }
    }
    
    //LoginApi_Handler.Login(1,1, 2, 2, 2,3, 2, 4, 2,5, 2, 6,"userId", "password", "type", "version", "identification");
    public static void Login(int ID,int PremiseUrl_Line,int PremiseUrl_Column,int ApiUrl_Line,int ApiUrl_Column,
    		                   int Act_Line,int Act_Column,int Method_Line,int Method_Column,
    		                   int ArgLineIndex,int ArgColumnIndex,int TITLE_LINE_INDEX,
    		                   String Param_Column1,String Param_Column2,String Param_Column3,String Param_Column4,String Param_Column5
//    		                   String result_Column1,result_Column2,result_Column3,result_Column4,result_Column5,result_Column6,result_Column7,result_Column8
    		                   ) throws Exception {
    	String PremiseUrl = "";
        String ApiUrl = "";
        String Act = "";
        String Method = "";
        List<Map<String, String>> data = null;
        boolean Flag = false;

        PremiseUrl = Excel_Util.getInstance().readExcelCell(PremiseUrl_Line-1, PremiseUrl_Column-1);
        ApiUrl = Excel_Util.getInstance().readExcelCell(ApiUrl_Line-1, ApiUrl_Column-1);
        Act = Excel_Util.getInstance().readExcelCell(Act_Line-1, Act_Column-1);
        Method = Excel_Util.getInstance().readExcelCell(Method_Line-1, Method_Column-1);       
        Flag = MobileApiTools_Util.isArgEquals(ArgLineIndex-1, ArgColumnIndex-1, TITLE_LINE_INDEX-1);

        if (PremiseUrl.equals("") ||ApiUrl.equals("") || Act.equals("") || Method.equals("") || !Flag) {
            logger.error("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.out.println("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.exit(-1);
        }

        data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);

        if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
            int i = ID-1;
            	
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
             
                //指定请求参数
                String Param ="userId=" + map.get(Param_Column1) + "&" + "password=" + map.get(Param_Column2) + "&" + "type=" + map.get(Param_Column3)+ "&" + "version=" + map.get(Param_Column4)+ "&" + "identification=" + map.get(Param_Column5);

                //指定请求方式,API地址,请求参数,Cookie值, 发起请求
                String CookieVal = HttpRequest_Util.getPostcookie(PremiseUrl,Param);  
                Map<String, String> result = HttpRequest_Util.sendPostcookie(Method, ApiUrl, Param, CookieVal); 
                             
                //获取接口返回的Code,结果内容
                String Code = result.get("code");
                String RsTmp = result.get("result");

                //将RsTmp字符串转换为JSON
                JSONObject object_result = new JSONObject(RsTmp);
                String Result_Data = object_result.getString("data");
                
                //将Result_Data字符串转换为JSON
                JSONObject object_data = new JSONObject(Result_Data);
                String id_Act = object_data.getString("$id");
                String phone_Act = object_data.getString("phone");
                String roleId_Act = object_data.getString("roleId");
                String roleName_Act = object_data.getString("roleName");
                String salesId_Act = object_data.getString("salesId");
                String userName_Act = object_data.getString("userName");
                String identification_Act = object_data.getString("identification");
                String photoName_Act = object_data.getString("photoName");
                    
                //期望结果与实际结果比较
                String $id_Exp = MobileApiTools_Util.assertResult(map.get("$id_Exp"),id_Act);
                String phone_Exp = MobileApiTools_Util.assertResult(map.get("phone_Exp"),phone_Act);
                String roleId_Exp = MobileApiTools_Util.assertResult(map.get("roleId_Exp"),roleId_Act);
                String roleName_Exp = MobileApiTools_Util.assertResult(map.get("roleName_Exp"),roleName_Act);
                String salesId_Exp = MobileApiTools_Util.assertResult(map.get("salesId_Exp"),salesId_Act);
                String userName_Exp = MobileApiTools_Util.assertResult(map.get("userName_Exp"),userName_Act);
                String identification_Exp = MobileApiTools_Util.assertResult(map.get("identification_Exp"),identification_Act);
                String photoName_Exp = MobileApiTools_Util.assertResult(map.get("photoName_Exp"),photoName_Act);

                //写入Run列, 执行纪录
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Run", "Y");

                //写入ResultCode列,返回结果代码
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "ResultCode", Code);

                //设置ResultCode单元格颜色
                if (Integer.parseInt(Code) == 200)
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 1);            
                else
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 0);

                //写入实际结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "$id_Act", id_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "phone_Act", phone_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleId_Act", roleId_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "roleName_Act", roleName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "salesId_Act", salesId_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "userName_Act", userName_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "identification_Act", identification_Act);
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "photoName_Act", photoName_Act);
                
                //写入测试结果单元格背景色
                if ($id_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "$id_Act",TITLE_LINE_INDEX + i, 0);}

                //写入测试结果单元格背景色
                if (phone_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "phone_Act",TITLE_LINE_INDEX + i, 0);}

                //写入测试结果单元格背景色
                if (roleId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (roleName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "roleName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (salesId_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "salesId_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (userName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "userName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (identification_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "identification_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试结果单元格背景色
                if (photoName_Exp.equals("OK")){}
                else{
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "photoName_Act",TITLE_LINE_INDEX + i, 0);}
                
                //写入测试通过与否,设置测试结果单元格背景色
//              &是位与操作，一定会执行； &&是逻辑与操作，如果&&的前面为false了，后面的就不会执行了。
//              |是位或操作、一定会执行； || 是逻辑或操作，如果||的前面为true了，||的后面就不会执行了
                if ($id_Exp.equals("OK")&phone_Exp.equals("OK")&roleId_Exp.equals("OK")&roleName_Exp.equals("OK")&
                	salesId_Exp.equals("OK")&userName_Exp.equals("OK")&identification_Exp.equals("OK")&photoName_Exp.equals("OK")){
                	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "OK");
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 1);}
                else{
                	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", "NG");
                	Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 0);}

                //写入执行时间
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "RunningTime", MobileApiTools_Util.getDate());
                
                //写入转换的JSON结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Json", RsTmp);
                
//                //打印
//                logger.info("CaseID: " + map.get("CaseID") + ", CaseName: " + map.get("CaseName") + ", $id_Exp: " +
//                map.get("$id_Exp") + ", $id_Act: " + id_Act + ", phone_Exp: " + map.get("phone_Exp") + ", phone_Act: " + phone_Act + ", ResultCode: " + Code +", TestResult: " + map.get("TestResult"));
//                System.out.println("CaseID: " + map.get("CaseID") + ", CaseName: " + map.get("CaseName") + ", $id_Exp: " +
//                map.get("$id_Exp") + ", $id_Act: " + id_Act + ", phone_Exp: " + map.get("phone_Exp") + ", phone_Act: " + phone_Act + ", ResultCode: " + Code +", TestResult: " + map.get("TestResult"));
            } 
//        }  
    }
    
    public static void GetCaseInfo(int ID) throws Exception {
    	
    	int TITLE_LINE_INDEX=6;
    	List<Map<String, String>> data = null;
    	data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);
    	if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
                int i = ID;
                
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
                String CaseID = map.get("CaseID");
                String CaseName = map.get("CaseName");
                String $id_Exp = map.get("$id_Exp");
                String $id_Act = map.get("$id_Act");
                String phone_Exp = map.get("phone_Exp");
                String phone_Act = map.get("phone_Act");
                String roleId_Exp = map.get("roleId_Exp");
                String roleId_Act = map.get("roleId_Act");
                String roleName_Exp = map.get("roleName_Exp");
                String roleName_Act = map.get("roleName_Act");
                String salesId_Exp = map.get("salesId_Exp");
                String salesId_Act = map.get("salesId_Act");
                String userName_Exp = map.get("userName_Exp");
                String userName_Act = map.get("userName_Act");
                String identification_Exp = map.get("identification_Exp");
                String identification_Act = map.get("identification_Act");
                String photoName_Exp = map.get("photoName_Exp");
                String photoName_Act = map.get("photoName_Act");
                String ResultCode = map.get("ResultCode");
                String TestResult = map.get("TestResult");
                String FailHint = map.get("FailHint");
                
                //打印Caseinfo
                logger.info("CaseID: " + CaseID + ", CaseName: " + CaseName + ", $id_Exp: " + $id_Exp + ", $id_Act: " + $id_Act + ", phone_Exp: " + phone_Exp + ", phone_Act: " + phone_Act +
                ", roleId_Exp: " + roleId_Exp + ", roleId_Act: " + roleId_Act + ", roleName_Exp: " + roleName_Exp + ", roleName_Act: " + roleName_Act + ", salesId_Exp: " + salesId_Exp + ", salesId_Act: " + salesId_Act +
                ", userName_Exp: " + userName_Exp +", userName_Act: " + userName_Act + ", identification_Exp: " + identification_Exp +", identification_Act: " + identification_Act + ", photoName_Exp: " + photoName_Exp +", photoName_Act: " + photoName_Act +
                ", ResultCode: " + ResultCode +", TestResult: " + TestResult+ ", FailHint: " + FailHint);
                logger.info("==============================================================================================================================================");

                
                System.out.println("CaseID: " + CaseID + ", CaseName: " + CaseName + ", $id_Exp: " + $id_Exp + ", $id_Act: " + $id_Act + ", phone_Exp: " + phone_Exp + ", phone_Act: " + phone_Act +
                ", roleId_Exp: " + roleId_Exp + ", roleId_Act: " + roleId_Act + ", roleName_Exp: " + roleName_Exp + ", roleName_Act: " + roleName_Act + ", salesId_Exp: " + salesId_Exp + ", salesId_Act: " + salesId_Act +
                ", userName_Exp: " + userName_Exp +", userName_Act: " + userName_Act + ", identification_Exp: " + identification_Exp +", identification_Act: " + identification_Act + ", photoName_Exp: " + photoName_Exp +", photoName_Act: " + photoName_Act +
                ", ResultCode: " + ResultCode +", TestResult: " + TestResult + ", FailHint: " + FailHint);
                System.out.println("==============================================================================================================================================");
            }
//        }
    }
    
    //LoginApi_Handler.GetCaseInfo(2,6);
    public static void GetCaseInfo(int ID,int TITLE_LINE_INDEX) throws Exception {
    	
    	List<Map<String, String>> data = null;
    	data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);
    	if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
                int i = ID-1;
                
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
                String CaseID = map.get("CaseID");
                String CaseName = map.get("CaseName");
                String $id_Exp = map.get("$id_Exp");
                String $id_Act = map.get("$id_Act");
                String phone_Exp = map.get("phone_Exp");
                String phone_Act = map.get("phone_Act");
                String roleId_Exp = map.get("roleId_Exp");
                String roleId_Act = map.get("roleId_Act");
                String roleName_Exp = map.get("roleName_Exp");
                String roleName_Act = map.get("roleName_Act");
                String salesId_Exp = map.get("salesId_Exp");
                String salesId_Act = map.get("salesId_Act");
                String userName_Exp = map.get("userName_Exp");
                String userName_Act = map.get("userName_Act");
                String identification_Exp = map.get("identification_Exp");
                String identification_Act = map.get("identification_Act");
                String photoName_Exp = map.get("photoName_Exp");
                String photoName_Act = map.get("photoName_Act");
                String ResultCode = map.get("ResultCode");
                String TestResult = map.get("TestResult");
                String FailHint = map.get("FailHint");
                
                //打印Caseinfo
                logger.info("CaseID: " + CaseID + ", CaseName: " + CaseName + ", $id_Exp: " + $id_Exp + ", $id_Act: " + $id_Act + ", phone_Exp: " + phone_Exp + ", phone_Act: " + phone_Act +
                ", roleId_Exp: " + roleId_Exp + ", roleId_Act: " + roleId_Act + ", roleName_Exp: " + roleName_Exp + ", roleName_Act: " + roleName_Act + ", salesId_Exp: " + salesId_Exp + ", salesId_Act: " + salesId_Act +
                ", userName_Exp: " + userName_Exp +", userName_Act: " + userName_Act + ", identification_Exp: " + identification_Exp +", identification_Act: " + identification_Act + ", photoName_Exp: " + photoName_Exp +", photoName_Act: " + photoName_Act +
                ", ResultCode: " + ResultCode +", TestResult: " + TestResult+ ", FailHint: " + FailHint);
                logger.info("==============================================================================================================================================");

                
                System.out.println("CaseID: " + CaseID + ", CaseName: " + CaseName + ", $id_Exp: " + $id_Exp + ", $id_Act: " + $id_Act + ", phone_Exp: " + phone_Exp + ", phone_Act: " + phone_Act +
                ", roleId_Exp: " + roleId_Exp + ", roleId_Act: " + roleId_Act + ", roleName_Exp: " + roleName_Exp + ", roleName_Act: " + roleName_Act + ", salesId_Exp: " + salesId_Exp + ", salesId_Act: " + salesId_Act +
                ", userName_Exp: " + userName_Exp +", userName_Act: " + userName_Act + ", identification_Exp: " + identification_Exp +", identification_Act: " + identification_Act + ", photoName_Exp: " + photoName_Exp +", photoName_Act: " + photoName_Act +
                ", ResultCode: " + ResultCode +", TestResult: " + TestResult + ", FailHint: " + FailHint);
                System.out.println("==============================================================================================================================================");
            }
//        }
    }
    
	public static void resultCheck(int ID) throws IOException{
		
		String ApiUrl = "";
		List<Map<String, String>> data = null;
		
		ApiUrl = Excel_Util.getInstance().readExcelCell(2-1, 2-1);
		
    	data = Excel_Util.getInstance().readExcelAllData(6-1);
    	  	
    	if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
            int i = ID;
            	
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
                String CaseID = map.get("CaseID");
                String CaseName = map.get("CaseName");
                String userId = map.get("userId");
                String password = map.get("password");
                String type = map.get("type");
                String version = map.get("version");
                String identification = map.get("identification");
                String $id_Exp = map.get("$id_Exp");
                String $id_Act = map.get("$id_Act");
                String phone_Exp = map.get("phone_Exp");
                String phone_Act = map.get("phone_Act");
                String roleId_Exp = map.get("roleId_Exp");
                String roleId_Act = map.get("roleId_Act");
                String roleName_Exp = map.get("roleName_Exp");
                String roleName_Act = map.get("roleName_Act");
                String salesId_Exp = map.get("salesId_Exp");
                String salesId_Act = map.get("salesId_Act");
                String userName_Exp = map.get("userName_Exp");
                String userName_Act = map.get("userName_Act");
                String identification_Exp = map.get("identification_Exp");
                String identification_Act = map.get("identification_Act");
                String photoName_Exp = map.get("photoName_Exp");
                String photoName_Act = map.get("photoName_Act");
                String Json = map.get("Json");
             
                Reporter.log("用例ID: "+ CaseID);        
		        Reporter.log("用例名称:"+ CaseName);
		        Reporter.log("请求地址: "+ ApiUrl);
		        Reporter.log("请求参数: "+ "userId: " + userId + ", password: " + password + ", type: " + type + 
		        		     ", version: " + version + ", identification: "+ identification);
		        Reporter.log("返回结果: "+ Json);
		        
		        checkEquals("$id",$id_Exp,$id_Act,ID);	
		        checkEquals("phone",phone_Exp,phone_Act,ID);
		        checkEquals("roleId",roleId_Exp,roleId_Act,ID);
		        checkEquals("roleName",roleName_Exp,roleName_Act,ID);
		        checkEquals("salesId",salesId_Exp,salesId_Act,ID);
		        checkEquals("userName",userName_Exp,userName_Act,ID);
		        checkEquals("identification",identification_Exp,identification_Act,ID);		        
		        checkEquals("photoName",photoName_Exp,photoName_Act,ID);
            }
//        }
	}

	//LoginApi_Handler.resultCheck(2,6,2,2,"$id_Exp","$id_Act");
    public static void resultCheck(int ID,int TITLE_LINE_INDEX,int ApiUrl_Line,int ApiUrl_Column,String Check1,String Check2) throws IOException{
		
		String ApiUrl = "";
		List<Map<String, String>> data = null;
		
		ApiUrl = Excel_Util.getInstance().readExcelCell(ApiUrl_Line-1, ApiUrl_Column-1);
		
    	data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);
    	  	
    	if (data != null) {
//            for (int i = 0; i < data.size(); i++) {
            int i = ID-1;
            	
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
                String CaseID = map.get("CaseID");
                String CaseName = map.get("CaseName");
                String userId = map.get("userId");
                String password = map.get("password");
                String type = map.get("type");
                String version = map.get("version");
                String identification = map.get("identification");
                String Check3 = map.get(Check1);
                String Check4 = map.get(Check2);
                String Json = map.get("Json");
             
                Reporter.log("用例ID: "+ CaseID);        
		        Reporter.log("用例名称:"+ CaseName);
		        Reporter.log("请求地址: "+ ApiUrl);
		        Reporter.log("请求参数: "+ "userId: " + userId + ", password: " + password + ", type: " + type + 
		        		     ", version: " + version + ", identification: "+ identification);
		        Reporter.log("返回结果: "+ Json);
		        
		        checkEquals(Check1,Check3,Check4,ID);	
            }
//        }
	}

	/**
	 * <br>检查预期与实际是否相等，不等则提示错误信息，并截图</br>
	 *
	 * @author    102051
	 * @date      2017年8月2日 下午6:01:04
	 * @param Actual 
	 * @param Expected
	 * @param FailHint
	 * @param FailedScreenshotCaseName
	 */
	public static void checkEquals(String FailHint,String Expected,String Actual,int ID){
		int i = ID;
		int TITLE_LINE_INDEX =6;
		try {
			Assert.assertEquals(Expected,Actual);
			Reporter.log("用例结果: 〖"+FailHint+"〗=>Expected: " + "【"+Expected+"】" + ", Actual: "+ "【"+Actual+"】");
	        System.out.println("用例结果: 【"+FailHint+"】=>Expected: " + "【"+Expected+"】" + ", Actual: "+ "【"+Actual+"】");
        }
 	    catch (Error e)  {
 	    	MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "FailHint", "【"+FailHint+"】预期结果和实际结果不一致");
 	    	Assert.fail(""+FailHint+" =>Expected 【"+ Expected +"】"+" "+"but found 【"+ Actual +"】");
 	    }
	}
	
    public static String GetPostcookie(int ApiUrl_Line,int ApiUrl_Column,
            int Act_Line,int Act_Column,int Method_Line,int Method_Column,
            int ArgLineIndex,int ArgColumnIndex,int TITLE_LINE_INDEX,
            String Param_Column1,String Param_Column2,String Param_Column3,String Param_Column4,String Param_Column5
            ) throws Exception {
    	
    	String ApiUrl = "";
        String Act = "";
        String Method = "";
        String CookieVal = "";
        List<Map<String, String>> data = null;
        boolean Flag = false;

        ApiUrl = Excel_Util.getInstance().readExcelCell(ApiUrl_Line-1, ApiUrl_Column-1);
        Act = Excel_Util.getInstance().readExcelCell(Act_Line-1, Act_Column-1);
        Method = Excel_Util.getInstance().readExcelCell(Method_Line-1, Method_Column-1);       
        Flag = MobileApiTools_Util.isArgEquals(ArgLineIndex-1, ArgColumnIndex-1, TITLE_LINE_INDEX-1);

//        &是位与操作，一定会执行； &&是逻辑与操作，如果&&的前面为false了，后面的就不会执行了。
//        |是位或操作、一定会执行； || 是逻辑或操作，如果||的前面为true了，||的后面就不会执行了。 
        
        if (ApiUrl.equals("") || Act.equals("") || Method.equals("") || !Flag) {
            logger.error("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.out.println("请检查 Excel 中 Interface、Act、Method、ArgName 是否设置正确...");
            System.exit(-1);
        }

        data = Excel_Util.getInstance().readExcelAllData(TITLE_LINE_INDEX-1);

        if (data != null) {
            for (int i = 0; i < data.size(); i++) {
            	
            	//根据Excel列名称,获取单元格内容
                Map<String, String> map = data.get(i);
             
                //指定请求参数
                String Param ="userId=" + map.get(Param_Column1) + "&" + "password=" + map.get(Param_Column2) + "&" + "type=" + map.get(Param_Column3)+ "&" + "version=" + map.get(Param_Column4)+ "&" + "identification=" + map.get(Param_Column5);

                //指定请求方式,API地址,请求参数, 发起请求

                Map<String, String> result = HttpRequest_Util.sendPost(Method, ApiUrl,Param); 
                        
                //获取接口返回的Code,结果内容
                String Code = result.get("code");
                String RsTmp = result.get("result");

                //将字符串转换为JSON
                JSONObject object = new JSONObject(RsTmp);
                String ActualResult = object.getString("data");
                System.out.println("213:"+ActualResult);
                
                //期望结果与实际结果比较
                String TestResult = MobileApiTools_Util.assertResult(map.get("ExpectedResult"),ActualResult);
       
                //写入Run列, 执行纪录
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "Run", "Y");

                //写入ResultCode列,返回结果代码
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "ResultCode", Code);

                //设置ResultCode单元格颜色
                if (Integer.parseInt(Code) == 200)
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 1);            
                else
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "ResultCode",TITLE_LINE_INDEX + i, 0);

                //写入实际结果
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "ActualResult", ActualResult);

                //写入测试通过与否
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "TestResult", TestResult);

                //设置测试结果单元格背景色
                if (TestResult.equals("OK"))
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 1);
                else
                    Excel_Util.getInstance().setCellBackgroundColor(TITLE_LINE_INDEX-1, "TestResult",TITLE_LINE_INDEX + i, 0);

                //写入执行时间
                MobileApiTools_Util.writeResult(TITLE_LINE_INDEX-1, TITLE_LINE_INDEX + i, "RunningTime", MobileApiTools_Util.getDate());
                
                //打印
                logger.info("CaseID: " + map.get("CaseID") + ", CaseName: " + map.get("CaseName") + ", ExpectedResult: " +
                map.get("ExpectedResult") + ", ActualResult: " + ActualResult + ", ResultCode: " + Code +", TestResult: " + TestResult);
                System.out.println("CaseID: " + map.get("CaseID") + ", CaseName: " + map.get("CaseName") + ", ExpectedResult: " +
                        map.get("ExpectedResult") + ", ActualResult: " + ActualResult + ", ResultCode: " + Code +", TestResult: " + TestResult);
                
                CookieVal = HttpRequest_Util.getPostcookie(ApiUrl,Param); 
            }
        }    	
        return CookieVal;
    }
}