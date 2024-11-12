package com.SocialSite.Signup;

public class SignupTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Signup s=new Signup();
		
		System.out.println( s.UserName("Charanya"));
		System.out.println( s.UserSurname("Kummara"));
		System.out.println( s.UserDOB("2002/10/12"));
		System.out.println( s.UserMobileNum(9014798545l));
		System.out.println( s.UserPassword(123454321));
		

	}

}
