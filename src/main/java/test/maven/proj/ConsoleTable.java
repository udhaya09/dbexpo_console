package test.maven.proj;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.stream.Stream;

public class ConsoleTable {

	public static void main( String[] args ) throws IOException {
		// TODO Auto-generated method stub
		Scanner sc = new Scanner(System.in);  
		        String s = "";
		        ArrayList<String> in = new ArrayList<String>();
		        System.out.println("enter input:");
		        while (sc.hasNextLine()){ //no need for "== true"
		            String read = sc.nextLine();
		            if(read == null || read.isEmpty()){ //if the line is empty
		                break;  //exit the loop
		            }
		            in.add(read);
		            
		        }
		    
		        for (String val : in) {
					s += val;
				}
		        System.out.println("output: " + s);
	}
	
	
}