package com.sales.webapi;

import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.sales.webapi.handler.LoginApi_Handler;

public class LoginTest {
     
    @BeforeTest
	public static void Stup() throws Exception{
    	LoginApi_Handler.GetExcelInstance();
//    	LoginApi_Handler.InitializeExcelData();
    	LoginApi_Handler.WriteExcelExpected("测试环境",12);
    }
    
    @Test
	public static void Case1() throws Exception{
    	LoginApi_Handler.Login(1);	
    	LoginApi_Handler.resultCheck(1);
	}
	
    @Test
	public static void Case2() throws Exception{
    	LoginApi_Handler.Login(2);	
    	LoginApi_Handler.resultCheck(2);
	}
    
    @Test
	public static void Case3() throws Exception{
    	LoginApi_Handler.Login(3);	
    	LoginApi_Handler.resultCheck(3);
	}
    
    @Test
	public static void Case4() throws Exception{
    	LoginApi_Handler.Login(4);	
    	LoginApi_Handler.resultCheck(4);
	}
    
    @Test
	public static void Case5() throws Exception{
    	LoginApi_Handler.Login(5);	
    	LoginApi_Handler.resultCheck(5);
	}
    
    @Test
	public static void Case6() throws Exception{
    	LoginApi_Handler.Login(6);	
    	LoginApi_Handler.resultCheck(6);
	}
    
    @Test
	public static void Case7() throws Exception{
    	LoginApi_Handler.Login(7);	
    	LoginApi_Handler.resultCheck(7);
	}
    
    @Test
	public static void Case8() throws Exception{
    	LoginApi_Handler.Login(8);	
    	LoginApi_Handler.resultCheck(8);
	}
    
    @Test
	public static void Case9() throws Exception{
    	LoginApi_Handler.Login(9);	
    	LoginApi_Handler.resultCheck(9);
	}
    
    @Test
	public static void Case_10() throws Exception{
    	LoginApi_Handler.Login(10);	
    	LoginApi_Handler.resultCheck(10);
	}
 
    @Test
	public static void Case_11() throws Exception{
    	LoginApi_Handler.Login(11);	
    	LoginApi_Handler.resultCheck(11);
	}

    @Test
	public static void Case_12() throws Exception{
    	LoginApi_Handler.Login(12);	
    	LoginApi_Handler.resultCheck(12);
	}

    @Test
	public static void Case_13() throws Exception{
    	LoginApi_Handler.Login(13);	
    	LoginApi_Handler.resultCheck(13);
	}
    
    @Test
	public static void Case_14() throws Exception{
    	LoginApi_Handler.Login(14);	
    	LoginApi_Handler.resultCheck(14);
	}
    
    @Test
	public static void Case_15() throws Exception{
    	LoginApi_Handler.Login(15);	
    	LoginApi_Handler.resultCheck(15);
	}
    
    @Test
	public static void Case_16() throws Exception{
    	LoginApi_Handler.Login(16);	
    	LoginApi_Handler.resultCheck(16);
	}
    
    @Test
	public static void Case_17() throws Exception{
    	LoginApi_Handler.Login(17);	
    	LoginApi_Handler.resultCheck(17);
	}
 
    @Test
	public static void Case_18() throws Exception{
    	LoginApi_Handler.Login(18);	
    	LoginApi_Handler.resultCheck(18);
	}
    
    @Test
	public static void Case_19() throws Exception{
    	LoginApi_Handler.Login(19);	
    	LoginApi_Handler.resultCheck(19);
	}
    
    @Test
	public static void Case_20() throws Exception{
    	LoginApi_Handler.Login(20);	
    	LoginApi_Handler.resultCheck(20);
	}
    
    @Test
	public static void Case_21() throws Exception{
    	LoginApi_Handler.Login(21);	
    	LoginApi_Handler.resultCheck(21);
	}
    
	@AfterTest
	public static void TearDown() throws Exception{
		LoginApi_Handler.GetCaseInfo(1);
		LoginApi_Handler.GetCaseInfo(2);
		LoginApi_Handler.GetCaseInfo(3);
		LoginApi_Handler.GetCaseInfo(4);
		LoginApi_Handler.GetCaseInfo(5);
		LoginApi_Handler.GetCaseInfo(6);
		LoginApi_Handler.GetCaseInfo(7);
		LoginApi_Handler.GetCaseInfo(8);
		LoginApi_Handler.GetCaseInfo(9);
		LoginApi_Handler.GetCaseInfo(10);
		LoginApi_Handler.GetCaseInfo(11);
		LoginApi_Handler.GetCaseInfo(12);
		LoginApi_Handler.GetCaseInfo(13);
		LoginApi_Handler.GetCaseInfo(14);
		LoginApi_Handler.GetCaseInfo(15);
		LoginApi_Handler.GetCaseInfo(16);
		LoginApi_Handler.GetCaseInfo(17);
		LoginApi_Handler.GetCaseInfo(18);
		LoginApi_Handler.GetCaseInfo(19);
		LoginApi_Handler.GetCaseInfo(20);
		LoginApi_Handler.GetCaseInfo(21);
	}
	
	public static void main(String[] args) {
		try {

			Stup();
			Case1();
			Case2();
			Case3();
			Case4();
			Case5();
			Case6();
			Case7();
			Case8();
			Case9();
			Case_10();
			Case_11();
			Case_12();
			Case_13();
			Case_14();
			Case_15();
			Case_16();
			Case_17();
			Case_18();
			Case_19();
			Case_20();
			Case_21();
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
}