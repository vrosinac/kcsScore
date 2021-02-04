
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
        if (name.equals("Andrei Nicolae"))
        {
            int stop =1;
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
        List<ArticleAuthor> Articleauthors = readArticleAuthorFromFile("");// le nom du fichier  // les colones attendues
        
        teamScores = new ArrayList<>();
            
        List<ArticleWithUsage> articles_withUsages = readArticlesWithUsageFromFile("");
           
        //sort by article and author, do one run of credits   - we have to detect manually the change of author
        Collections.sort(articles_withUsages, new ArticleAuthor_comparator());
        String newtitle="", previoustitle="";
        String authorNew="";
        String author="";
        int authorpoints =0;
        for (ArticleWithUsage row : articles_withUsages) 
        { 
            newtitle=row.ArticleVersionTitle;
            String id2 = row.KnowledgeArticleID;
            
            if (newtitle.equals("Summary for UTZA25R Maker Checker EOD process **KEEP AS INTERNAL**"))
            {
                int stop=1;
            }
            if( !newtitle.equals(previoustitle) && !previoustitle.isEmpty() )
            {
                // before giving credit, we check how many times the article was referenced .... but we 
               //credit(author,authorpoints,authorpoints + " " + previoustitle,0, "",0,"", 0,"");
               credit(author,authorpoints,authorpoints + " " + previoustitle,0, "",0,"", 0,"");
               authorpoints =0;
            }
            author= row.CaseArticleCreatedBy;//  TODO HERE INSTEAD WE SHOULD LOOK UP ON AUTHORS.XLS FOR AHTOR OR HSITORICAL AUTHOR
                boolean found =false;
                for (ArticleAuthor a : Articleauthors)
                {
                    if (a.Title.equals(newtitle))
                    {
                       found = true;
                       author = a.Author;
                      /* if (!authorNew.equals(author))
                       {
                          found=false;
                       }*/
                    }
                }
                if (found == false) //found author for creation points
                {
                    for (ArticleAuthor a : Articleauthors)
                    {
                        if (a.Id.equals(id2))
                        {
                           if (!author.isEmpty())
                           {
                                author = a.Author;
                                found = true;
                           }
                           else
                           {
                               found=false;// in the end we keep the one from the usage.xls file 
                           }    
                        }
                        
                    }

                }
                // END OF TODO
            if (found==false)
            {
                int notfound =0;
            }
            previoustitle= row.ArticleVersionTitle;
            if (previoustitle.equals("CSummary for UTZA25R Maker Checker EOD process **KEEP AS INTERNAL**"))
            {
                int stop=1;
            }
            authorpoints++;  // we get a point for own publishing and for each reuse of the article
        }
        //last line of teh file
        credit(author,authorpoints,authorpoints + " " + previoustitle,0, "",0,"", 0,"");
        
        
        Collections.sort(articles_withUsages, new ArticleWithUsage_comparator());
        int reusepoints=0;
        String previoususer="", newuser ="";
        previoustitle="";
        for (ArticleWithUsage row1 : articles_withUsages) 
        { 
            newtitle=row1.ArticleVersionTitle;
             if (newtitle.equals("Summary for UTZA25R Maker Checker EOD process **KEEP AS INTERNAL**"))
            {
                int stop=1;
            }
             
             /* Wendy needs to skeep the following
                •00008533 –  Misys Customer Support Guide to Case Severity Usage -we attach when customers log a critical severity case which is not really critical 
                •00009325 – Process in Manually Archiving a Retail Loan **KEEP AS INTERNAL** - is a collection of database updates used in emergencies so the system can continue 
                •00018364 –  AHS Customer Support team - list of LIVE and CRITICAL issues - link to confluence **KEEP AS INTERNAL**  - is a link from SFDC to an old confluence spreadsheet 
                •00022685 –  Equation COP SWIFT Presentations ** KEEP AS INTERNAL ** -is a link from SFDC to CS training notes 
             */
             if (newtitle.equals("Misys Customer Support Guide to Case Severity Usage") ||
                 newtitle.equals("Process in Manually Archiving a Retail Loan **KEEP AS INTERNAL**") ||
                 newtitle.equals("AHS Customer Support team - list of LIVE and CRITICAL issues") ||
                 newtitle.equals("Equation COP SWIFT Presentations ** KEEP AS INTERNAL **")  )
            {
                int stop=1;
                continue;
                
            }
             
             
            newuser = row1.SubjectCaseOwner;
            
            String id3 = row1.CaseNumber;
            if (id3.equals("02277729"))
            {
                int stop=1;
            }
        
            if(  (  !newtitle.equals(previoustitle)     )       ||
                    (    newtitle.equals(previoustitle) &&   (!newuser.equals(previoususer) && !previoususer.isEmpty()) )  )
            {
                if (reusepoints >0)
                    {
                        credit(previoususer,0, "", reusepoints,reusepoints + " " + previoustitle,0,"",0,"");
                        reusepoints =0;
                    }
            }
                if (   !row1.FirstPublishedDate.equals(row1.CaseArticleCreatedDate) ) // TODO Carine
                // first published <> last published = article was reused
                {
                    //check if the user is also the author (in that case we add only one point for the usage, otherwise we add 2
                    
                    boolean found =false;
                    for (ArticleAuthor a : Articleauthors )
                    {
                        if (a.Title.equals(newtitle)&& a.Author.equals(newuser))
                        {
                           found = true;
                         
                        }
                    }
                    if (found == true)  //found author when checking reuse
                    {
                        reusepoints= reusepoints + 1 ;
                    }
                    else
                    {
                        reusepoints= reusepoints + 2 ;  // we get 2 points reusing someoneelse's  article
                    }
            }
            previoususer = row1.SubjectCaseOwner;
            previoustitle=row1.ArticleVersionTitle;
        }
        //for the last row of the file
        credit(previoususer,0, "", reusepoints,reusepoints + " " + previoustitle,0,"",0,"");
                       
        
        
        //----------------------- RATINGS   ----------------------
        //List<ArticleAuthor> Articleauthors = readArticleAuthorFromFile("");// le nom du fichier  // les colones attendues
        List<ArticleWithRating> ArticleRatings = readArticleWithRatingsFromFile("");
        Collections.sort(ArticleRatings, new ArticleWithRating_comparator());
        
        String number ="";
        String lastNumber="qwertyytrewq"; //for sure does not exist
        String lastRatedBy="qwertyytrewq"; //for sure does not exist
        String theauthor ="";
        for (ArticleWithRating r : ArticleRatings) 
        { 
            theauthor ="";
            String ratedBy = r.CommunityArticleCommentOwnerName;
            char issuesolved_char =r.IssueResolved.toCharArray()[0];
            char rate_char =r.ArticleRating.toCharArray()[0];
            int rate = Character.getNumericValue(rate_char);
            int issuesolved = Character.getNumericValue(issuesolved_char);
            number = r.ArticleNumber;// ISSUE HERE
            if(number.equals("Article Number"))
            {
                rate =0;
                issuesolved =0;
                        
            }
            //System.out.println(" rated article id: " + id + " rating: " + rate + " solved?: "+ issuesolved ); 
            //we can't count the rating if on top of that the issue was solved.
            //In that case we only count the issue solved

            if (number.equals("000016433"))
            {
            
                int a=1;
            }

            if (  !((ratedBy.equals(lastRatedBy) && (number == lastNumber)))   ) // we avoid duplicates
            {
               
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
                    if (a.Number.equals(number))
                    {
                      theauthor = a.Author;
                    }
                }
                if(r.KnowledgeArticleTitle.equals("Knowledge Article Title"))
                {
                    int stop=1;
                }
                if(r.KnowledgeArticleTitle.equals("Summary for UTZA25R Maker Checker EOD process **KEEP AS INTERNAL**"))
                {
                    int stop=1;
                }   
                   
                
                
                
                               
                if (theauthor.isEmpty())
                {
                      theauthor ="NOT FOUND";
                }
                if (theauthor.equals("\"Carine Gheron\""))
                {
                     //      System.out.println(" credits for author: " + theauthor);
                }    
                if (!theauthor.equals("Created By: Full Name") ) //the colum names are showing up as data rows
                {      
                 credit(theauthor,0, "", 0, " " ,rate, r.KnowledgeArticleTitle ,4*issuesolved, r.KnowledgeArticleTitle);
                }
            }
            else
            {
                System.out.println("we avoided one duplicate rating: \'" + r.KnowledgeArticleTitle + "\' rated by: " + ratedBy);
            }
            lastNumber = number;
            lastRatedBy = ratedBy;        
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
    
    
    final int articleAuthor = 1;
    final int articleWithUsage = 2;
    final int articleWithRatings = 3;

    /*
    private static List<Object> readExcelFile(String fileName, int type) 
    {
        try {
            FileInputStream fs;
            ArrayList<String> listCol = new ArrayList<String>();//Creating arraylist  
            Object itemLine;
            ArrayList<Object> list;
            switch (type)
            {
                case 1:
                    list = new ArrayList<ArticleAuthor>();
                    break;
                case 2:
                    list = new ArrayList<>();
                    listCol.add("Case Number");//Adding object in arraylist  
                    listCol.add("Subject");
                    listCol.add("Case Owner");
                    listCol.add("Case Article: Created By");
                    listCol.add("Article Version: Last Modified By");
                    listCol.add("Article Version: Last Modified Date");
                    listCol.add("Article Version: Title");
                    listCol.add("Knowledge Article ID");
                     //open the file (obtaining input bytes from a file)  
                    fs = new FileInputStream(new File("usage.xls"));

                    break;
                case 3:
                    list = new ArrayList<>();
                    break;
           }

        
        //Traversing list through Iterator  
        //Iterator itr = list.iterator();
        //int length = list.size();
        // System.out.println(length); 
        //   while(itr.hasNext()){  
         //   System.out.println(itr.next());  
        //}  
         
            //open the file (obtaining input bytes from a file)  
            //FileInputStream fs = new FileInputStream(new File("usage.xls"));

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet = wb.getSheetAt(0);
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            if (sheet.getRow(0).getPhysicalNumberOfCells() != 10) {//check the number of columns of the file
                AskForUsageFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
            //additional author called author writing an article titled title... 
            for (Row row : sheet) //iteration over row using for each loop  
            {
                int i = 0;
                String col[] = new String[10];
                for (Cell cell : row) //iteration over cell using for each loop  
                {
                    switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type  
                            //getting the value of the cell as a number  
                            int answer = (int) cell.getNumericCellValue();
                            col[i] = "" + answer; // all the logic was previosuly done on strings. plugging new prototype to old code for now
                            //to refafctor later
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                            //getting the value of the cell as a string  
                            col[i] = cell.getStringCellValue();
                            break;
                    }
                    i++;
                }
                
            switch (type)
            {
                case 1:
                    break;
                case 2:
                    itemLine = createArticleWithUsage(col);
                    break;
            }
                // do somethig with it
                //itemLine = createArticleWithUsage(col);
                // adding article into ArrayList 
                list.add(itemLine);

            }
        } catch (FileNotFoundException e) {
            System.out.println("File not found");
            System.out.println(e.getMessage());
            AskForAuthorFile();
        } catch (IOException e) {
            String ioex = e.getMessage();
            if (ioex.contains("Invalid header signature")) {
                System.out.println("The excel file produced by Salesforce is sually corrupt. \n"
                        + "\t to fix it open it in Excel, chose Export\n"
                        + "\t Chose : change file type]n"
                        + "\t chose the *.xls\n"
                        + "\t and save onto your same file name, overwriting your file with itself after this dummy excel export\n"
                        + "\t it should fix the issue.\n");
            } else {
                System.out.println(e.getMessage());
            }
            AskForUsageFile();

        }

        return list;

    }
*/
    private static List<ArticleWithUsage> readArticlesWithUsageFromFile(String fileName) 
    { 

       List<ArticleWithUsage> articles = new ArrayList<>();
       
        ArrayList<String> list=new ArrayList<String>();//Creating arraylist  
        list.add("Case Number");//Adding object in arraylist  
        list.add("Subject");
        list.add("Case Owner");
        list.add("Article Version: Created By");  
	list.add("Article Version: Last Modified By");
	list.add("Article Version: Last Modified Date");
	list.add("Article Version: Title");
	list.add("Knowledge Article ID");

	//Traversing list through Iterator  
        Iterator itr=list.iterator(); 
        int length = list.size();
       // System.out.println(length); 
     /*   while(itr.hasNext()){  
            System.out.println(itr.next());  
        }  
       */ 
                try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("usage.xls"));  

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb=new HSSFWorkbook(fis);   
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet=wb.getSheetAt(0);  
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
            if (sheet.getRow(0).getPhysicalNumberOfCells()!=10)
            {//check the number of columns of the file
                AskForUsageFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
                    //additional author called author writing an article titled title... 
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                int i=0;
                String col[] = new String [11];
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
    
       return articles;

   } 


   private static ArticleWithUsage createArticleWithUsage(String[] metadata) 
   { 
        String CaseNumber= metadata[0];
        String SubjectCaseOwner= metadata[2];
        String CaseArticleCreatedBy= metadata[3];
        String ArticleVersionLastModifiedBy= metadata[4];
        String ArticleVersionLastModifiedDate= metadata[5];
        String FirstPublishedDate= metadata[6];
        String CaseArticleCreatedDate = metadata[7];
        String ArticleVersionTitle= metadata[9];
        String KnowledgeArticleID= metadata[10];
        


       // create and return article of this metadata 
       return new ArticleWithUsage(CaseNumber,   SubjectCaseOwner, CaseArticleCreatedBy,    ArticleVersionLastModifiedBy, ArticleVersionLastModifiedDate, FirstPublishedDate, CaseArticleCreatedDate, ArticleVersionTitle,    KnowledgeArticleID);

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
       // System.out.println(length); 
        /*while(itr.hasNext()){  
            System.out.println(itr.next());  
        } */ 
                
        try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("ratings.xls"));  

            POIFSFileSystem fis = new POIFSFileSystem(fs);
            //creating workbook instance that refers to .xls file  
            HSSFWorkbook wb=new HSSFWorkbook(fis);   
            //creating a Sheet object to retrieve the object  
            HSSFSheet sheet=wb.getSheetAt(0);  
            //evaluating cell type   
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
            if (sheet.getRow(0).getPhysicalNumberOfCells()!=5)
            {//check the number of columns of the file
                AskForAuthorFile();
            }
            //TODO, we should get rid of the header row, but I suppose it will not harm us to have an 
                    //additional author called author writing an article titled title... 
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                int i=0;
                String col[] = new String [5];
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
        String CommunityArticleCommentOwnerName = metadata[4];
     /*
   "Knowledge Article Title","Article Number","Product Name","Comment","Issue Resolved","Article Rating","Comment Date"
 */
       // create and return article of this metadata 
       return new ArticleWithRating(ArticleId,KnowledgeArticleTitle,IssueResolved,ArticleRating, CommunityArticleCommentOwnerName);
        
   }
   
    
    
   // we need to pass the type of object and the filename, maybe the column names for validatin
     
    private static List<ArticleAuthor> readArticleAuthorFromFile(String fileName) 
    { 
       List<ArticleAuthor> articleauthors = new ArrayList<ArticleAuthor>();
       
       
        ArrayList<String> list=new ArrayList<String>();//Creating arraylist  
        list.add("Article Number");//Adding object in arraylist  
        list.add("Title");  
        list.add("Created By: Full Name");
        list.add("Troubleshooting ID");

        //Traversing list through Iterator  
        Iterator itr=list.iterator(); 
        int length = list.size();
        //System.out.println(length); 
       /* while(itr.hasNext()){  
            System.out.println(itr.next());  
        }*/  
      //  readInputFile("C:\\JavaProjects\\excelread2\\author.xls",list);

        try
        {
            //open the file (obtaining input bytes from a file)  
            FileInputStream fs=new FileInputStream(new File("author.xls"));  

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
            int rown=0;
            for(Row row: sheet)     //iteration over row using for each loop  
            {  
                rown= rown+1;
                if (rown==2212)
                {
                    int stop=1;
                }
                int i=0;
                String col[] = new String [5]; // 3 OR 4 OR 5
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
                        if (col[i].equals("000016433"))
                        {
                            int stop=2;
                        }
                        break;  
                    }
                  i++;
              
                }
                  // do somethig with
                 ArticleAuthor articleauthor =  new ArticleAuthor(col);

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
    
       return articleauthors;

   } 
   
   
    
} 


class ArticleWithUsage 
{ 
    String CaseNumber;
    String SubjectCaseOwner;
    String CaseArticleCreatedBy;
    String ArticleVersionLastModifiedBy;
    String ArticleVersionLastModifiedDate;
    String FirstPublishedDate;
    String CaseArticleCreatedDate;
    String ArticleVersionTitle;
    String KnowledgeArticleID;
    

    public ArticleWithUsage(String CaseNumber,   String SubjectCaseOwner,String CaseArticleCreatedBy,   String ArticleVersionLastModifiedBy,
                            String ArticleVersionLastModifiedDate, String FirstPublishedDate,String CaseArticleCreatedDate, 
                            String ArticleVersionTitle,  String  KnowledgeArticleID) 
    { 

        this.CaseNumber = CaseNumber;
        this.SubjectCaseOwner = SubjectCaseOwner;
        this.CaseArticleCreatedBy = CaseArticleCreatedBy;
        this.ArticleVersionLastModifiedBy =ArticleVersionLastModifiedBy;
        this.ArticleVersionLastModifiedDate = ArticleVersionLastModifiedDate;
        this.FirstPublishedDate = FirstPublishedDate;
        this.CaseArticleCreatedDate = CaseArticleCreatedDate;
        this.ArticleVersionTitle = ArticleVersionTitle;
        this.KnowledgeArticleID = KnowledgeArticleID;
    
        
   }
} 



class ArticleWithRating 
{ 
    String ArticleNumber;
    String KnowledgeArticleTitle;
    String IssueResolved;
    String ArticleRating;
    String CommunityArticleCommentOwnerName;
        
    public ArticleWithRating(String ArticleId, String KnowledgeArticleTitle,   String  IssueResolved,String ArticleRating, String CommunityArticleCommentOwnerName ) 
   { 

        this.ArticleNumber=ArticleId;
        this.KnowledgeArticleTitle = KnowledgeArticleTitle;
        this.IssueResolved = IssueResolved;
        this.ArticleRating = ArticleRating;
        this.CommunityArticleCommentOwnerName = CommunityArticleCommentOwnerName;
   }
} 




class ArticleAuthor   /// SPEC QUESTION  what happens if the rating is on an article reused?  Should it not count for the reuser too? 
{ 
    String Number;
    String Id;
    String Author;
    String Title;
        
    public ArticleAuthor(String Number,   String Author, String title, String Id ) 
   { 

        this.Id = Id;
        
        this.Number = Number;
        this.Author = Author;
        this.Title=title;
   }
    public ArticleAuthor(String[] metadata) 
   { 
       this.Number= metadata[0];
       if( (metadata[3] != null) && !metadata[3].isEmpty() )
       {
        this.Author= metadata[3];
       
       
       }
       else{
           if( (metadata[2] != null) && !metadata[2].isEmpty() )
           {
               this.Author= metadata[2];
           }
           else
           {
             this.Author= "NOT FOUND";
           }
       }
       
       //return new ArticleAuthor(Id, Author);
       this.Title= metadata[1];
       this.Id= metadata[4];
    }
    
    public ArticleAuthor createArticleAuthor(String metadata [])
    {
        return new ArticleAuthor(metadata);
    }
    
    public String fromFile()
    {
        return "authors.xls";
    }
    public String [] arguments()
    {
    
        String args[] = new String[3]; /*String[4];*/
        args[0]="Article Number";
        args[1]="Title";
        args[2]="Created By: Full Name";
        
        return args;
        
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

class  ArticleWithRating_comparator implements Comparator<ArticleWithRating> {
    public int compare(ArticleWithRating a1, ArticleWithRating a2) {
        int c=0;
        c = a1.CommunityArticleCommentOwnerName.compareTo(a2.CommunityArticleCommentOwnerName);
        
        if (c ==0)
        {
            c=a1.IssueResolved.compareTo(a2.IssueResolved);
            
            if (c ==0)
            {
                c=a1.ArticleRating.compareTo(a2.ArticleRating);
            }
            
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


