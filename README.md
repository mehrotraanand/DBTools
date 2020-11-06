# DBTools
Dev Tools
1. ASP Page generation using PL/SQL
I presume that ASP code is supported in the web server such as IIS before you want to create a page

Following steps need to be performed before using this tool to generate the asp code

    Create a directory using command
         For eg; create directory ASP_DIR AS 'c:\asp'
    grant access to this directory to the schema where you want to create this package.
    grant read,write on directory DEBUG_DIR to 'SCOTT'
    ODBC connection to Database is to be set  in windows, I named it in code as ODBCTESTDB
    Copy DB Package ASP_PAGE_GEN_PKG and compile it in you oracle schema.
    Two types of ASP pages using by executing DB Package as below
      EXEC ASP_PAGE_GEN_PKG.create_ASP_list_page_prc(<TABLE_NAME or VIEW_NAME>) ;
        This will create a list of all record page for any table or view

      EXEC ASP_PAGE_GEN_PKG.create_ASP_maint_page_prc(<TABLE_NAME>);
        This will create an asp page that will allow user  
        a. To insert/update/delete and browse through records of table.
        b. It will alidate the data entered for different field value and field type validation, 
        c. It will display valid values for any foreign keys.
    Copy the ASP pages to your Index page in Webserver and it will allow you maintain and report from Table passed.     
    
    

    
     

