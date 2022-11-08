package in.ac.coep.excel;

import java.io.FileOutputStream;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ArithNewAss
{

	public static void main(String[] args) 
	{
		try
		{
		String filename = "E:\\practice\\excel\\QuestionsBank.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Instruction ");
        XSSFSheet sheet1 = workbook.createSheet("Questions");
        XSSFRow row = sheet1.createRow((short)0);
	    
        String[] header = {"Sr. No","Question Type","Answer Type","Topic Number","Question (Text Only)","Correct Answer 1","Correct Answer 2","Correct Answer 3","Correct Answer 4","Wrong Answer 1","Wrong Answer 2","Wrong Answer 3","Time in seconds","Difficulty Level","Question (Image/ Audio/ Video)","Contributor's Registered mailId","Solution (Text Only)","Solution (Image/ Audio/ Video)","Variation Number"};
        for(int count = 0 ; count<header.length ; count++)
        {
       	     row.createCell(count).setCellValue(header[count]);
      	  	       	
        }
        sheet1.setColumnWidth(4,2*5000);
        sheet1.setColumnWidth(16,2*5000);
        Scanner sc = new Scanner(System.in);
	    System.out.println("Enter the number of questions:");
		int n = sc.nextInt();
		
		XSSFRow row1 = sheet1.createRow((short)n);
		
		for(int i = 1; i<=n;i++)
		{
			
			int amin = 1 ; int amax = 20;
         	int dmin = 1 ; int dmax = 20;
         	int nmin = 1 ; int nmax = 50;
         	
         	  int aval = (int)(Math.random()*(amax-amin+1)+amin);
         	  int dval = (int)(Math.random()*(dmax-dmin+1)+dmin);
         	  int nval = (int)(Math.random()*(nmax-nmin+1)+nmin);
         	
         	 String name[] = {"Suraj"," Gopi", "Kishan", "Amit", "Amol", "Sandesh", "Rohit", "Raj", "Dev", "Pravin"};
      	   String name1[] = {"Soni","Moni","Riya","Sonia","Deepa","Deepika","Rakhi","Yogini","Jaya","Jayashree"};
         	  Random random = new Random();
         	           		                  
         	  String gender ="";
         		int arr = (int)(Math.random()*(1-0+1)+0); 
         		int str; 
         		String nam ="";
         	   if(arr==0 )
         	   {
         		   gender = "he";
         		    str =random.nextInt(name.length);
         		    nam = name[str];
         	   }
         	   else
         	   {    
         		   gender = "she";
         		   str =random.nextInt(name1.length);
         		   nam = name1[str];

         	   }
         	   	
		  row1 = sheet1.createRow(i);
		  row1.createCell(0).setCellValue(i); //Serial number
		  row1.createCell(1).setCellValue(1); //Question type	
		  row1.createCell(2).setCellValue(1); //Answer type
		
		  row1.createCell(3).setCellValue(2); // Topic number
         	
		  String st = new String("If "+ nam+" starts doing Surya Namaskar. On the first day "+ gender+" did "+ aval +" Namaskar. "+ gender + "  increased "+ dval+ " Namaskar daily. On "+ nval + " day how many Soorya Namaskar did "+gender+" perform?");
  		  row1.createCell(4).setCellValue(st); // Question	
  		  
  		   int res = aval+(nval-1)*dval;   	
     	  row1.createCell(5).setCellValue(res); //Correct answer1
     	  row1.createCell(6).setCellValue(""); // Correct answer2
     	  row1.createCell(7).setCellValue(""); // Correct answer3
     	  row1.createCell(8).setCellValue(""); // Correct answer4
//     	  row1.createCell(9).setCellValue("");
     	 
     	  int wr1 = aval+(nval*1)+dval;
     	 row1.createCell(9).setCellValue(wr1); // Wrong answer 1
     	                            
     	 int wr2 = aval+(nval+1)-dval;
     	row1.createCell(10).setCellValue(wr2); // Wrong answer 2
     	
     	int wr3 = aval+(nval-1)+dval;
     	row1.createCell(11).setCellValue(wr3); // Wrong answer 3
     	
     	row1.createCell(12).setCellValue("120");
     	row1.createCell(13).setCellValue(1);  
     	row1.createCell(14).setCellValue("");
     	row1.createCell(15).setCellValue("abc@gmail.com");
     	
     	//String s = new String("tn = a+(n-1)*d so a = tn-(n-1)d");
//     	    String s = new String("tn = a+(n-1)*d so a = tn-(n-1)*d");
//     	   res=aval+(nval-1)*dval;
     	    int res1 = aval+(nval-1)*dval;
     	   
     	    String dt = "We know tn = a+(n-1)*d so a = tn-(n-1)*d here we have d = "+ dval+" , tn = "+ nval+" and a = "+aval+" now substituting values we get result=>" +res1;
    	     row1.createCell(16).setCellValue(dt); //solution
    	     
    	    
    	     XSSFRow row2 = sheet1.createRow((short)n+1);
             row2.createCell(0).setCellValue("****"); 	 
     	
       	}		
		 FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			fileOut.close();
			fileOut.close();
        System.out.println("File generated successfully");     
		}catch (Exception e)
           {
	          e.printStackTrace();
	         }
		
		}
}
