//TI article ratings2
//PUBLIC KCSautoRatingsTI "Article Number,""Knowledge Article Title"",""Issue Resolved"",""Article Rating"""
//PUBLIC KCSautoRatingsKondor
//PUBLIC KCSautoUsage  "Case Number","Subject","Case Owner","Case Article: Created By","Article Version: Last Modified By","Article Version: Last Modified Date","Article Version: Title","Knowledge Article ID"
//PUBLIC KCSautoAuthor  "Article Number,""Title"",""Created By: Full Name"""


package kcs1;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import java.io.File;  
import java.io.FileInputStream;  
import java.io.FileNotFoundException;
import java.io.IOException;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  
import java.util.*;         
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
 

public class KCS1{ 
    
    static List<ConsultantScore> teamScores;
    private static void credit(String name, int authorpoints, String authorDetail, int  reusepoints, String reuseDetail,
                               int ratingpoints, String ratingDetails, int issuesolvedpoints, String solvedDetails)
    {
        if (name.equals("\"Carine Gheron\""))
        {
            System.out.println("Carine Gheron: author +" + authorpoints + " reuse +" + reusepoints + " rating+" +ratingpoints + " solved+" + issuesolvedpoints);
        }
        name = name.replaceAll("\"", "");
        boolean done=false;
            int size = teamScores.size();
            for (int i =0; i<size; i++)
            {
                ConsultantScore c1 = teamScores.get(i);
                if(c1.Name.equals(name))
                {
                    //increment the consultant's score
                    if (!authorDetail.isEmpty())
                    {   
                            authorDetail = "<br>" + authorDetail;
                    }
                    if (!reuseDetail.isEmpty())
                    { 
                        reuseDetail = "<br>" + reuseDetail;
                    }
                    
                    
                    if (ratingpoints == 0)
                    { 
                        ratingDetails = "" ;
                    }
                    else
                    {
                       ratingDetails = "<br>"  + ratingpoints + " " + ratingDetails; 
                    
                    }
                    
                    
                    if (issuesolvedpoints == 0)
                    { 
                        solvedDetails = "" ;
                    }
                    else
                    {
                            solvedDetails = "<br>" + issuesolvedpoints + " " + solvedDetails ;
                    }
                    
                    ConsultantScore c2 = new ConsultantScore(c1.Name,c1.totalpoints+ reusepoints + authorpoints + ratingpoints+ issuesolvedpoints,
                            c1.authorpoints+ authorpoints,
                            c1.authorDetail  + authorDetail, c1.reusepoints+ reusepoints,
                            c1.reuseDetail + reuseDetail, 
                            c1.ratingpoints + ratingpoints, c1.ratingDetails + ratingDetails,
                            c1.issuesolvedpoints + issuesolvedpoints, 
                            c1.solvedDetails + solvedDetails);
                    teamScores.set(i,c2);
                    done=true;
                }
            }
        
            if (done==false)
            {    
                ConsultantScore consultant = new ConsultantScore (name,authorpoints + reusepoints + ratingpoints + issuesolvedpoints ,authorpoints, authorDetail, 
                        reusepoints,reuseDetail,ratingpoints, ratingDetails, issuesolvedpoints, solvedDetails );
                teamScores.add(consultant);
             }
            
    }
    
    
    
    public static void main(String... args) 
    { 
        teamScores = new ArrayList<>();
            
        List<ArticleWithUsage> articles_withUsages = readArticlesWithUsageFromFile("");
           
        //sort by article and author, do one run of credits   - we have to detect manually the change of author
        Collections.sort(articles_withUsages, new ArticleAuthor_comparator());
        String newtitle="", previoustitle="";
        String author="";
        int authorpoints =0;
        for (ArticleWithUsage row : articles_withUsages) 
        { 
            newtitle=row.ArticleVersionTitle;
            if( !newtitle.equals(previoustitle) && !previoustitle.isEmpty() )
            {
                // before giving credit, we check how many times the article was referenced .... but we 
               credit(author,authorpoints,authorpoints + " " + previoustitle,0, "",0,"", 0,"");
               authorpoints =0;
            }
            author= row.CaseArticleCreatedBy;
            previoustitle= row.ArticleVersionTitle;
            authorpoints++;  // we get a point for own publishing and for each reuse of the article
        } 
        
        Collections.sort(articles_withUsages, new ArticleWithUsage_comparator());
        int reusepoints=0;
        String previoususer="", newuser ="";
        previoustitle="";
        for (ArticleWithUsage row1 : articles_withUsages) 
        { 
            newtitle=row1.ArticleVersionTitle;
            newuser = row1.SubjectCaseOwner;
            if( !newuser.equals(previoususer) && !previoususer.isEmpty() )
            {
                if (reusepoints >0)
                    {
                        credit(previoususer,0, "", reusepoints,reusepoints + " " + previoustitle,0,"",0,"");
                        reusepoints =0;
                    }
            }
            if (   !row1.SubjectCaseOwner.equals(row1.CaseArticleCreatedBy) )
            {
                    reusepoints= reusepoints + 2 ;  // we get 2 points reusing someoneelse's  article
            }
            previoususer = row1.SubjectCaseOwner;
            previoustitle=row1.ArticleVersionTitle;
        }
    
        
        
        //----------------------- RATINGS   ----------------------
        List<ArticleAuthor> Articleauthors = readArticleAuthorFromCSV("authors.txt");
        List<ArticleWithRating> ratings = readArticleWithRatingsFromFile("ratings.xls");
        String id ="";       
        String theauthor ="";
        for (ArticleWithRating r : ratings) 
        { 
            theauthor ="";
            id = r.ArticleId;
            char issuesolved_char =r.IssueResolved.toCharArray()[0];
            char rate_char =r.ArticleRating.toCharArray()[0];
            int rate = Character.getNumericValue(rate_char);
            int issuesolved = Character.getNumericValue(issuesolved_char);
            System.out.println(" rated article id: " + id + " rating: " + rate + " solved?: "+ issuesolved ); 
            //we can't count the rating if on top of that the issue was solved.
            //In that case we only count the issue solved
            if (issuesolved>0)
            {
                rate=0;
            }
            //we only count the ratings above 2
            if (rate <3)
            {
                rate =0;
            }
            
            for (ArticleAuthor a : Articleauthors)
            {
                if (a.Id.equals(id))
                {
                  theauthor = a.Author;
                }
            }
            if (theauthor.isEmpty())
            {
                  theauthor ="NOT FOUND";
            }
            if (theauthor.equals("\"Carine Gheron\""))
            {
                 //      System.out.println(" credits for author: " + theauthor);
            }    
             credit(theauthor,0, "", 0, " " ,rate, r.KnowledgeArticleTitle ,4*issuesolved, r.KnowledgeArticleTitle);
        }

        
        //-----------------------   consulatnt scores and output     -----------------
        Collections.sort(teamScores, new ConsultantScore_comparator());
        try
        {
            BufferedWriter writer = new BufferedWriter(new FileWriter("output.html"));
            String htmlheader= "<style>\n" +
"#grad1 {\n" +
"background-color: #0a0a0a;\n" +
"background-image: linear-gradient(147deg, mediumvioletred , #434343  , mediumvioletred 	); "+
"}\n" +
"#grad2 {\n" +
"background-color: #000000;\n" +
"background-image: linear-gradient(147deg, #000000 , #434343 , #000000);\n" +
"}\n" +
".myTable { background-color: #000000;background-image: linear-gradient(147deg, #000000 , #090808, #000000);\n" + //#434343
";border-collapse:collapse; }\n" +
".myTable td, .myTable th { padding:5px;border:1px solid #000;  color:grey;  font-weight:lighter; }\n" +
           
                    
"</style>\n" +
"<body  style=\"background-color:black\" >\n" +
"<table  style=\"width:100%;height:5%\">\n" +
"<colgroup   id=\"grad1\"> \n" +
"<col  width=\"10%\"  >\n" +
"<col  width=\"5%\"   >\n" +
"<col  width=\"20%\"  >\n" +
"<col  width=\"20%\"  >\n" +
"<col  width=\"20%\"  >\n" +
"<col  width=\"20%\"  >\n" +
"</colgroup>\n" +
"<th align=\"left\"style=\"color:grey\" > &nbsp; Name  </th>\n" +
"<th align=\"center\" style=\"color:grey\"> score </th>\n" +
"<th align=\"left\" > author of article</th>\n" +
"<th align=\"left\" > article reuse </th>\n" +
"<th align=\"left\" > article ratings </th>\n" +
"<th align=\"left\" > issue solved by article</th></table><br style=\"line-height: 50%\">";
            writer.write(htmlheader);
            
            for (ConsultantScore c : teamScores) 
            { 
                writer.write(c.toString());
            }
            writer.close();
        }
        catch (Exception e)
        {
        }
    } 
     
    private static List<ArticleWithUsage> readArticlesWithUsageFromFile(String fileName) 
    { 

       List<ArticleWithUsage> articles = new ArrayList<>();
       
        ArrayList<String> list=new ArrayList<String>();//Creating arraylist  
        list.add("Case Number");//Adding object in arraylist  
        list.add("Subject");
        list.add("Case Owner");
        list.add("Case Article: Created By");  
	list.add("Article Version: Last Modified By");
	list.add("Article Version: Last Modified Date");
	list.add("Article Version: Title");
	list.add("Knowledge Article ID");

	//Traversing list through Iterator  
        Iterator itr=list.iterator(); 
        int length = list.size();
        System.out.println(length); 
        while(itr.hasNext()){  
            System.out.println(itr.next());  
        }  
        
                try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("C:\\JavaProjects\\KCS1\\usage.xls"));  

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb=new HSSFWorkbook(fis);   
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet=wb.getSheetAt(0);  
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
            if (sheet.getRow(0).getPhysicalNumberOfCells()!=8)
            {//check the number of columns of the file
                AskForUsageFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
                    //additional author called author writing an article titled title... 
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                int i=0;
                String col[] = new String [8];
                for(Cell cell: row)    //iteration over cell using for each loop  
                {  
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())  
                    {  
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type  
                        //getting the value of the cell as a number  
                        int answer = (int) cell.getNumericCellValue();
                        col[i] = ""+answer; // all the logic was previosuly done on strings. plugging new prototype to old code for now
                        //to refafctor later
                        break;  
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                        //getting the value of the cell as a string  
                        col[i] = cell.getStringCellValue();
                        break;  
                    }
                  i++;
                }
                // do somethig with it
                ArticleWithUsage article = createArticleWithUsage(col);
                // adding article into ArrayList 
                 articles.add(article);

            }
        }
        
        catch( FileNotFoundException e    )
        {
            System.out.println("File not found");
            System.out.println(e.getMessage());
            AskForAuthorFile();
        }
        catch (IOException e)
        {
            String ioex = e.getMessage();
            if (ioex.contains("Invalid header signature"))
            {
                System.out.println("The excel file produced by Salesforce is sually corrupt. \n"
                        + "\t to fix it open it in Excel, chose Export\n"
                        + "\t Chose : change file type]n"
                        + "\t chose the *.xls\n"
                        + "\t and save onto your same file name, overwriting your file with itself after this dummy excel export\n"
                        + "\t it should fix the issue.\n");
            }
            else
            {
                System.out.println(e.getMessage());
            }
            AskForUsageFile();
        
        }
    

                

       /*
       Path pathToFile = Paths.get(fileName);

       // create an instance of BufferedReader 
       // using try with resource, Java 7 feature to close resources 
       try (BufferedReader br = Files.newBufferedReader(pathToFile, StandardCharsets.US_ASCII)) 
       { 
            // read the first line from the text file 
                String line = br.readLine();
            // loop until all lines are read 
            while (line != null) 
            { 
             // use string.split to load a string array with the values from 
             // each line of the file, using a comma as the delimiter 
                 String[] attributes = line.split(",");
                 ArticleWithUsage article = createArticleWithUsage(attributes);

             // adding article into ArrayList 
                 articles.add(article);

             // read next line before looping 
             // if end of file reached, line would be null 
                 line = br.readLine();

            } 

       }  
       catch (IOException ioe) 
       { 

         ioe.printStackTrace();

       } 
       */
       return articles;

   } 


   private static ArticleWithUsage createArticleWithUsage(String[] metadata) 
   { 
        String CaseNumber= metadata[0];
        String SubjectCaseOwner= metadata[2];
        String CaseArticleCreatedBy= metadata[3];
        String ArticleVersionLastModifiedBy= metadata[4];
        String ArticleVersionLastModifiedDate= metadata[5];
        String ArticleVersionTitle= metadata[6];
        String KnowledgeArticleID= metadata[7];

    /*
    CaseNumber,   SubjectCaseOwner, CaseArticleCreatedBy,    ArticleVersionLastModifiedBy, ArticleVersionLastModifiedDate,  ArticleVersionTitle= metadata[5],    KnowledgeArticleID
    */
       // create and return article of this metadata 
       return new ArticleWithUsage(CaseNumber,   SubjectCaseOwner, CaseArticleCreatedBy,    ArticleVersionLastModifiedBy, ArticleVersionLastModifiedDate,  ArticleVersionTitle,    KnowledgeArticleID);

   }
   
   
   
   
     private static List<ArticleWithRating> readArticleWithRatingsFromFile(String fileName) 
    { 

       List<ArticleWithRating> ratings = new ArrayList<>();
        ArrayList<String> list=new ArrayList<String>();//Creating arraylist  
        list.add("Article Number");//Adding object in arraylist  
        list.add("Title");  
        list.add("Created By: Full Name");  
        //Traversing list through Iterator  
        Iterator itr=list.iterator(); 
        int length = list.size();
        System.out.println(length); 
        while(itr.hasNext()){  
            System.out.println(itr.next());  
        }  
                
        try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("C:\\JavaProjects\\KCS1\\ratings.xls"));  

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb=new HSSFWorkbook(fis);   
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet=wb.getSheetAt(0);  
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
            if (sheet.getRow(0).getPhysicalNumberOfCells()!=4)
            {//check the number of columns of the file
                AskForAuthorFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
                    //additional author called author writing an article titled title... 
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                int i=0;
                String col[] = new String [4];
                for(Cell cell: row)    //iteration over cell using for each loop  
                {  
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())  
                    {  
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type  
                        //getting the value of the cell as a number  
                        //System.out.print(cell.getNumericCellValue()+ "HELLO\t\t");   
                        int answer = (int) cell.getNumericCellValue();
                        col[i] = ""+answer; // all the logic was previosuly done on strings. plugging new prototype to old code for now
                        //to refafctor later
                        break;  
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                        //getting the value of the cell as a string  
                        col[i] = cell.getStringCellValue();
                        break;  
                    }
                  i++;
              
                }
                  // do somethig with it
                 ArticleWithRating rating = createArticeWithRating(col);

                 // adding article into ArrayList 
                 ratings.add(rating);
            }

        }
        
        catch( FileNotFoundException e    )
        {
            System.out.println("File not found");
            System.out.println(e.getMessage());
            AskForRatingsFile();
        }
        catch (IOException e)
        {
            String ioex = e.getMessage();
            if (ioex.contains("Invalid header signature"))
            {
                System.out.println("The excel file produced by Salesforce is sually corrupt. \n"
                        + "\t to fix it open it in Excel, chose Export\n"
                        + "\t Chose : change file type]n"
                        + "\t chose the *.xls\n"
                        + "\t and save onto your same file name, overwriting your file with itself after this dummy excel export\n"
                        + "\t it should fix the issue.\n");
            }
            else
            {
                System.out.println(e.getMessage());
            }
            AskForAuthorFile();
        
        }

        
        
/*
       Path pathToFile = Paths.get(fileName);
       // create an instance of BufferedReader 
       // using try with resource, Java 7 feature to close resources 
       try (BufferedReader br = Files.newBufferedReader(pathToFile, StandardCharsets.US_ASCII)) 
       { 


            // read the first line from the text file 
                String line = br.readLine();

            // loop until all lines are read 
            while (line != null) 
            { 


             // use string.split to load a string array with the values from 
             // each line of 
             // the file, using a comma as the delimiter 
                 String[] attributes = line.split(",");
                 ArticleWithRating rating = createArticeWithRating(attributes);

             // adding article into ArrayList 
                 ratings.add(rating);

             // read next line before looping 
             // if end of file reached, line would be null 
                 line = br.readLine();

            } 

       }  
       catch (IOException ioe) 
       { 

         ioe.printStackTrace();

       }
       */
       return ratings;

   } 
 static void   AskForAuthorFile()
    {
        System.out.println("We are expectig a file containing information about the authors aof the articles");
        System.out.println("It should be on the local directory where the application is running");
        System.out.println("It should be called authors.txt");
        System.out.println("It should have the following 3 columns : Article Number,	Title,	Created By: Full Name");
 	System.out.println("It is comming from the excel extract of a Salesforce report probably called something like KCSautoAuthor....");
        System.out.println("Report Type: Troubleshooting articles");
        System.out.println("Filtered By:   \n" +
        "   	Product Family equals Kondor+ Clear \n" +
        "   	AND Is Latest Version equals True Clear \n" +
        "   	AND Product equals Trade Innovation Pre TI Plus versions,Trade Innovation TI PLUS 1,Trade Innovation TI PLUS 2,"
                + "Trade Innovation TI PLUS 2 Global processing,Trade Innovation Trade Portal Interface (TPI) Clear ");
        System.out.println("Time Frame : \n" +
        "\tDate Field : Created Date\n" +
        "\tRange : Current FY");
    }
 
 
 static void   AskForUsageFile()
    {
        System.out.println("We are expectig a file containing information about the authors aof the articles");
        System.out.println("It should be on the local directory where the application is running");
        System.out.println("It should be called authors.txt");
        System.out.println("It should have the following 3 columns : \n\tArticle Number\n" +
    "\tKnowledge Article\n" +
    "\tTitle\n" +
    "\tIssue Resolved\n" +
    "\tArticle Rating");
 	System.out.println("It is comming from the excel extract of a Salesforce report probably called something like KCSautoAuthor....");
        System.out.println("Report Type: Troubleshooting articles");
        System.out.println("Filtered By:   \n" +
        "   	Product Family equals Kondor+ Clear \n" +
        "   	AND Is Latest Version equals True Clear \n" +
        "   	AND Product equals Trade Innovation Pre TI Plus versions,Trade Innovation TI PLUS 1,Trade Innovation TI PLUS 2,"
                + "Trade Innovation TI PLUS 2 Global processing,Trade Innovation Trade Portal Interface (TPI) Clear ");
        System.out.println("Time Frame : \n" +
        "\tDate Field : Created Date\n" +
        "\tRange : from 01/09/2010 to end of current FY");
    }
 static void   AskForRatingsFile()
    {
        System.out.println("We are expectig a file containing information about the rating of the articles");
        System.out.println("It should be on the local directory where the application is running");
        System.out.println("It should be called ratings.xls");
        System.out.println("It should have the following  columns :");
 	System.out.println("It is comming from the excel extract of a Salesforce report probably called something like KCSautoRating....");
        System.out.println("Report Type: Community Article Comments");
        System.out.println("Filtered By:   \n" +
        "   	Product Name \n" );
        System.out.println("Time Frame : \n" +
        "\tDate Field : Created Date\n" +
        "\tRange : Current FY");
    }
      
   
    private static ArticleWithRating createArticeWithRating(String[] metadata) 
   { 
        String ArticleId = metadata[0];
        String KnowledgeArticleTitle= metadata[1];
        String IssueResolved= metadata[2];
        String ArticleRating= metadata[3];
     /*
   "Knowledge Article Title","Article Number","Product Name","Comment","Issue Resolved","Article Rating","Comment Date"
 */
       // create and return article of this metadata 
       return new ArticleWithRating(ArticleId,KnowledgeArticleTitle,IssueResolved,ArticleRating);
        
   }
   
    
    
    
    
   
     private static List<ArticleAuthor> readArticleAuthorFromCSV(String fileName) 
    { 
       List<ArticleAuthor> articleauthors = new ArrayList<>();
       
       
        ArrayList<String> list=new ArrayList<String>();//Creating arraylist  
        list.add("Article Number");//Adding object in arraylist  
        list.add("Title");  
        list.add("Created By: Full Name");

        //Traversing list through Iterator  
        Iterator itr=list.iterator(); 
        int length = list.size();
        System.out.println(length); 
        while(itr.hasNext()){  
            System.out.println(itr.next());  
        }  
        
        
      //  readInputFile("C:\\JavaProjects\\excelread2\\author.xls",list);

        try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("C:\\JavaProjects\\KCS1\\author.xls"));  

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb=new HSSFWorkbook(fis);   
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet=wb.getSheetAt(0);  
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
            if (sheet.getRow(0).getPhysicalNumberOfCells()!=3)
            {//check the number of columns of the file
                AskForAuthorFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
                    //additional author called author writing an article titled title... 
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                int i=0;
                String col[] = new String [3];
                for(Cell cell: row)    //iteration over cell using for each loop  
                {  
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())  
                    {  
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type  
                        //getting the value of the cell as a number  
                        System.out.print(cell.getNumericCellValue()+ "HELLO\t\t");   
                        break;  
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                        //getting the value of the cell as a string  
                        col[i] = cell.getStringCellValue();
                        break;  
                    }
                  i++;
              
                }
                  // do somethig with it
                 ArticleAuthor articleauthor = createArticleAuthor(col);

                 // adding article into ArrayList 
                 articleauthors.add(articleauthor);
     }

        }
        
        catch( FileNotFoundException e    )
        {
            System.out.println("File not found");
            System.out.println(e.getMessage());
            AskForAuthorFile();
        }
        catch (IOException e)
        {
            String ioex = e.getMessage();
            if (ioex.contains("Invalid header signature"))
            {
                System.out.println("The excel file produced by Salesforce is sually corrupt. \n"
                        + "\t to fix it open it in Excel, chose Export\n"
                        + "\t Chose : change file type]n"
                        + "\t chose the *.xls\n"
                        + "\t and save onto your same file name, overwriting your file with itself after this dummy excel export\n"
                        + "\t it should fix the issue.\n");
            }
            else
            {
                System.out.println(e.getMessage());
            }
            AskForAuthorFile();
        
        }
    


       
       /*
       Path pathToFile = Paths.get(fileName);

       // create an instance of BufferedReader 
       // using try with resource, Java 7 feature to close resources 
       try (BufferedReader br = Files.newBufferedReader(pathToFile, StandardCharsets.US_ASCII)) 
       { 
            // read the first line from the text file 
             String line = br.readLine();

            // loop until all lines are read 
            while (line != null) 
            { 
             // use string.split to load a string array with the values from 
             // each line of 
             // the file, using a comma as the delimiter 
                 String[] attributes = line.split(",");
                 ArticleAuthor articleauthor = createArticleAuthor(attributes);

             // adding article into ArrayList 
                 articleauthors.add(articleauthor);

             // read next line before looping 
             // if end of file reached, line would be null 
                line = br.readLine();
            } 

       }  
       catch (IOException ioe) 
        { 
           ioe.printStackTrace();
        }
       */
       return articleauthors;

   } 
   
   
    private static ArticleAuthor createArticleAuthor(String[] metadata) 
   { 
       String Id= metadata[0];
       String Author= metadata[2];
       return new ArticleAuthor(Id, Author);
    }
} 


class ArticleWithUsage 
{ 
    String CaseNumber;
    String SubjectCaseOwner;
    String CaseArticleCreatedBy;
    String ArticleVersionLastModifiedBy;
    String ArticleVersionLastModifiedDate;
    String ArticleVersionTitle;
    String KnowledgeArticleID;
    

    public ArticleWithUsage(String CaseNumber,   String SubjectCaseOwner,String CaseArticleCreatedBy,   String ArticleVersionLastModifiedBy,String ArticleVersionLastModifiedDate, String ArticleVersionTitle,  String  KnowledgeArticleID) 
    { 

        this.CaseNumber = CaseNumber;
        this.SubjectCaseOwner = SubjectCaseOwner;
        this.CaseArticleCreatedBy = CaseArticleCreatedBy;
        this.ArticleVersionLastModifiedBy =ArticleVersionLastModifiedBy;
        this.ArticleVersionLastModifiedDate = ArticleVersionLastModifiedDate;
        this.ArticleVersionTitle = ArticleVersionTitle;
        this.KnowledgeArticleID = KnowledgeArticleID;
    
        
   }
} 



class ArticleWithRating 
{ 
    String ArticleId;
    String KnowledgeArticleTitle;
    String IssueResolved;
    String ArticleRating;
        
    public ArticleWithRating(String ArticleId, String KnowledgeArticleTitle,   String  IssueResolved,String ArticleRating ) 
   { 

        this.ArticleId=ArticleId;
        this.KnowledgeArticleTitle = KnowledgeArticleTitle;
        this.IssueResolved = IssueResolved;
        this.ArticleRating = ArticleRating;
   }
} 




class ArticleAuthor   /// SPEC QUESTION  what happens if the rating is on an article reused?  Should it not count for the reuser too? 
{ 
    String Id;
    String Author;
        
    public ArticleAuthor(String Id,   String Author ) 
   { 

        this.Id = Id;
        this.Author = Author;
            
        if(this.Author.equals("pair quotes"))
        {
            int i=0;
        }
   } 
}



class ConsultantScore
{ 
    String Name;
    int totalpoints;
    int authorpoints;
    String authorDetail;
    int reusepoints;
    String reuseDetail;
    String detail;
    int ratingpoints;
    String ratingDetails;
    int issuesolvedpoints;
    String solvedDetails;
    
    
    public ConsultantScore( String Name, int totalpoints, int authorpoints, String authorDetails, int reusepoints, String reuseDetails,
    int ratingpoints, String ratingDetails,   int issuesolvedpoints,String solvedDetails) 
   { 
        this.Name = Name;
        this.totalpoints = totalpoints;
        this.authorpoints = authorpoints;
        this.authorDetail = authorDetails;
        this.reusepoints = reusepoints;
        this.reuseDetail = reuseDetails;
        this.detail= detail;
        this.ratingpoints= ratingpoints;
        this.ratingDetails= ratingDetails;
        this.issuesolvedpoints= issuesolvedpoints;
        this.solvedDetails= solvedDetails;
        
        if(this.Name.equals("pair quotes"))
        {
            int i=0;
        }
        
   }
    
   public int getTotalPoints() 
   {  return totalpoints;}

   
   @Override public String toString() 
   { 
       // TODO  sort the array of articles by title + author and count how many times each article was read, output the number + the title + <br>  ad next next
       // then same but for reuse, ordre by article nam ad reuser, count, produce the resut, and next
       //todo add somehwer the creation date <> last modif date when author = reuser to count double whne they made a chnace before reusing
        
       if(Name.equals("pair quotes"))
       {
           int i=0;
       }
       String sexystyle = "style=\"font-family: Arial,  \n" +
"            Helvetica, sans-serif; \n" +
"        background: linear-gradient( \n" +
"            to right, #999999, #454545); \n" +
"        -webkit-text-fill-color: transparent; \n" +
"        -webkit-background-clip: text; \n" +
"	position: relative;\n" +
"    z-index: 2;\"";
       
       
                String result =        
       /*          "<style>\n" +
"#grad1 {\n" +
"  height: 200px;\n" +
//"  background-color: blueviolet; \n" +
"  background-image: linear-gradient(red, blue);\n" +
"}\n" +
"</style>"      
                        + "<body id=\"grad1\">"
                        
                 +*/
                        "<table style=\"width:100%\" class=\"myTable\"><colgroup>"
                + " <col  width=\"10%\" >"
                + "<col  width=\"5%\"    >"  //
                + "<col  width=\"20%\"   >" //style=\"background-color:lavender\"
                + "<col  width=\"20%\"  >" // style=\"background-color:aliceblue\"
                + "<col  width=\"20%\"   >" // style=\"background-color:cornsilk\"
                + "<col  width=\"20%\"   >" //style=\"background-color:lightcyan\
                
                +        "</colgroup>" +
                  "<th align=\"left\"  style=\"background-image: linear-gradient(147deg, #200D54 2%, blue 74%\" >" + Name + "</th>"
                      
                        
                        + "<th align=\"center\" style=\"background-image: linear-gradient(147deg, blue, darkblue, navy);color:grey;font-weight:bold\" >" + totalpoints +"</th>";
                if (authorpoints !=0)
                {
                    
                    result = result  + "<th align=\"left\"" + sexystyle + ">"+
                             "<details>"+ authorDetail+"<summary>" + authorpoints + "</summary></details></th>";
                }
                else
                {
                     result = result + "<th align=\"left\"" + sexystyle + " >" + authorpoints + "</th>";
                
                }
                
                if (reusepoints !=0)
                {
                    result = result +  "<th align=\"left\"" + sexystyle + " ><details>"+ reuseDetail+"<summary>" + reusepoints  + "</summary></details></th>";
                } 
                else
                {
                   result = result + "<th align=\"left\"" + sexystyle + " >" + reusepoints + "</th>";
                }
                
                
                if (ratingpoints !=0)
                {
                    
                    result = result +  "<th align=\"left\"" + sexystyle + " ><details>"+/* ratingpoints + " "*/ "" + ratingDetails+"<summary>" + "<br>" +  ratingpoints +"</summary></details></th>";
                
                }
                else
                {
                   result = result + "<th align=\"left\"" + sexystyle + " >" + ratingpoints +"</th>"; 
                }
                
                
                if (issuesolvedpoints !=0)
                {
                    result = result +  "<th align=\"left\" ><details>"+ /*issuesolvedpoints + " " */ "" + solvedDetails+"<summary>" + "<br>" + issuesolvedpoints +"</summary></details></th>";
                } 
                else
                {
                    result = result + "<th align=\"left\" >"+ issuesolvedpoints + "</th>";
                }
                
                result = result + "</body>";
                return result;
        }
}
    


class ConsultantScore_comparator implements Comparator<ConsultantScore> {
    public int compare(ConsultantScore c1, ConsultantScore c2) {
        return c2.getTotalPoints() - c1.getTotalPoints();
    }
}


class ArticleAuthor_comparator implements Comparator<ArticleWithUsage> {
    public int compare(ArticleWithUsage a1, ArticleWithUsage a2) {
        int c=0;
        c = a1.KnowledgeArticleID.compareTo(a2.KnowledgeArticleID);
        
        if (c ==0)
        {
            c=a1.CaseArticleCreatedBy.compareTo(a2.CaseArticleCreatedBy);
        }            
        
        return c;
    }
}




class ArticleWithUsage_comparator implements Comparator<ArticleWithUsage> {
    public int compare(ArticleWithUsage a1, ArticleWithUsage a2) {
        int c=0;
        c = a1.KnowledgeArticleID.compareTo(a2.KnowledgeArticleID);
        
        if (c ==0)
        {
            c=a1.ArticleVersionLastModifiedBy.compareTo(a2.ArticleVersionLastModifiedBy);
        }            
        
        return c;
    }
}


