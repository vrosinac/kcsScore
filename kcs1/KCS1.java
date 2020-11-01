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
import java.util.List;
 

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
            
        List<ArticleWithUsage> articles_withUsages = readArticlesWithUsageFromCSV("usage.txt");
           
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
        List<ArticleWithRating> ratings = readArticleWIthRatingsFromCSV("ratings.txt");
        String id ="";       
        String theauthor ="";
        for (ArticleWithRating r : ratings) 
        { 
            theauthor ="";
            id = r.ArticleId;
            char issuesolved_char =r.IssueResolved.toCharArray()[1];
            char rate_char =r.ArticleRating.toCharArray()[1];
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
     
    private static List<ArticleWithUsage> readArticlesWithUsageFromCSV(String fileName) 
    { 

       List<ArticleWithUsage> articles = new ArrayList<>();
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
        if (ArticleVersionLastModifiedDate.equals("pair quotes"))
        {
            int a =0;

        }
        if (CaseArticleCreatedBy.equals("pair quotes"))
        {
            int i=0;
        }
        
        
         if (ArticleVersionLastModifiedDate.equals("pair quotes"))
        {
            int i=0;
        }
        String KnowledgeArticleID= metadata[7];

    /*
    CaseNumber,   SubjectCaseOwner, CaseArticleCreatedBy,    ArticleVersionLastModifiedBy, ArticleVersionLastModifiedDate,  ArticleVersionTitle= metadata[5],    KnowledgeArticleID
    */
       // create and return article of this metadata 
       return new ArticleWithUsage(CaseNumber,   SubjectCaseOwner, CaseArticleCreatedBy,    ArticleVersionLastModifiedBy, ArticleVersionLastModifiedDate,  ArticleVersionTitle,    KnowledgeArticleID);

   }
   
   
   
   
     private static List<ArticleWithRating> readArticleWIthRatingsFromCSV(String fileName) 
    { 

       List<ArticleWithRating> ratings = new ArrayList<>();
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
       return ratings;

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
        if(this.SubjectCaseOwner.equals("pair quotes"))
        {
            int i=0;
        
        }
        
         if(this.ArticleVersionLastModifiedBy.equals("pair quotes"))
        {
            int i=0;
        
        }
        
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


