package Capgemini.Training;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.junit.jupiter.api.Test;

/**
 * Unit test for simple App.
 */
public class AppTest {

   public static void main(String[] args) {
	   App app = new App();
	   String regex="^(.+)@(.+)$";
	   String email=app.UserEmail("charanya@gmail.com");
	   Pattern pattern=Pattern.compile(regex);
	   
	   Matcher match=pattern.matcher(email);
	   System.out.println("Is it match::"+match.matches());
	   System.out.println(email);
	   int password=app.UserPassword(2187369);
	   System.out.println(password);
	   app.login();
   }
}
