# R code for Getting Access data into R using RODBC.

#################
## Code Segment 1
#################
# Installing RODBC, Loading RODBC, 
# and getting access to the RODBC vignette.

install.packages("RODBC")
library(RODBC)
vignette("RODBC")



#################
## Code Segment 2
#################
# Opening a connection to an Excel 
# spreadsheet and importing data.

   # Assumes a new R session
library(RODBC)

   # Assumes that there is a DSN named MyTestBook
con = odbcConnect("MyTestBook")

   # You will need to modify the path to the location 
   # of the Excel workbook on your computer
wb = "C:/Users/Sandy/Documents/SST/Database/Test.xlsx"
con2 = odbcConnectExcel2007(wb)

   # Gets the names of the tables in the Excel workbook
tbls = sqlTables(con)
tbls$TABLE_NAME

   # Imports the Student table
student = sqlFetch(con, "Student$")
str(student)

   # Imports the Student table using an SQL SELECT command
qry = "SELECT * FROM `Student$`"
student = sqlQuery(con, qry)
str(student)

   # Close connections
odbcCloseAll()



#################
## Code Segment 3
#################
# Opening a connection to an Access 
# database and importing data

   # Assumes a new R session
library(RODBC)

   # Method 1: Assumes that there is a DSN named ToyDB
con = odbcConnect("ToyDB")

   # Method 2: You will need to modify the path to the 
   # location of the Access database on your computer
db = "C:/Users/Sandy/Documents/SST/Database/Database for playing.accdb"
con2 = odbcConnectAccess2007(db)

   # Gets the names of the tables in the Access database
sqlTables(con2, tableType = "TABLE")$TABLE_NAME

   # Imports tables using the sqlFetch function
   # and SQL queries in the sqlQuery function
school = sqlFetch(con2, "school")
str(school)

qry = "SELECT * FROM class"
class = sqlQuery(con2, qry)
str(class)

   # Tidy-up
odbcCloseAll()



#################
## Code Segment 4
#################
# Data manipulations, aggregations, and merging before importing into R

   # Asssumes a new session
library(RODBC)
db = "C:/Users/Sandy/Documents/SST/Database/Database for playing.accdb"
con = odbcConnectAccess2007(db)
sqlTables(con, tableType = "TABLE")$TABLE_NAME

   # Get variables names
sqlColumns(con, "student")$COLUMN_NAME
sqlColumns(con, "class")$COLUMN_NAME
sqlColumns(con, "school")$COLUMN_NAME

   # Using SQL to import part of a table
      # limiting the variables
qry1 = "SELECT Test1, Test2 FROM student"
TestScores = sqlQuery(con, qry1)
str(TestScores)
      # limiting the cases
qry2 = "SELECT * FROM student WHERE Test1 > 50"
TestScoresLimit = sqlQuery(con, qry2)
str(TestScoresLimit)

   # Using SQL to compute a new variable
qry3 = "SELECT Test1, Test2, Test2-Test1 AS Diff FROM student"
TestScoreDiff = sqlQuery(con, qry3)
str(TestScoreDiff)

   # Using SQL to aggregate data
qry4 = "SELECT
        AVG(Test1) AS mean1, AVG(Test2) AS mean2,
        STDEV(Test1) AS sd1, STDEV(Test2) AS sd2,
        COUNT(Test1) AS N1, COUNT(Test2) AS N2
        FROM student"
sqlQuery(con, qry4)

   # Using SQL to aggregate data by groups
qry5 = "SELECT Gender, classID,
        AVG(Test1) AS mean1, AVG(Test2) AS mean2,
        STDEV(Test1) AS sd1, STDEV(Test2) AS sd2,
        COUNT(Test1) as N1, COUNT(Test2) as N2
        FROM student
        GROUP BY Gender, classID"
sqlQuery(con, qry5)

   # Using SQL to merge two tables
qry6 = "SELECT * FROM 
   student LEFT OUTER JOIN class
   ON student.classID = class.classID"
student = sqlQuery(con, qry6)
str(student)

   # Using SQL to merge three tables
qry7 = "SELECT * FROM
      student LEFT OUTER JOIN
         (SELECT * FROM
         class LEFT OUTER JOIN school
         ON class.schID = school.schID) AS X
      ON student.classID = X.classID"
student = sqlQuery(con, qry7)
str(student)



#################
## Code Segment 5
#################
# Using R to merge three tables

student = sqlFetch(con, "student")
class = sqlFetch(con, "class")
school = sqlFetch(con, "school")

studentNew = merge(student, class, by = "classID", all.x = TRUE)
studentNew = merge(studentNew, school, by = "schID", all.x = TRUE)
str(studentNew)

odbcCloseAll()

   # The R merge using a merge within a merge
studentNew = merge(merge(student, class, by = "classID", all.x = TRUE),
           school, by = "schID", all.x = TRUE)
str(studentNew)



#################
## Code Segment 6
#################
# Using sqldf to merge three tables

install.packages("sqldf")  # To install � Run once only
library(sqldf)             # To load � Run each new R session

studentNew = sqldf("SELECT * FROM
      student LEFT OUTER JOIN
         (SELECT * FROM
         class LEFT OUTER JOIN school
         ON class.schID = school.schID) AS X
      ON student.classID = X.classID")
str(studentNew)
