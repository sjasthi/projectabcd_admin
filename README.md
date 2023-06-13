# projectabcd_admin
For generating the content from ABCD website into Power Points; For managing the data


Current Capability: main.py
===========================
This generates the PPTX using Beautiful Soup. 
This web scraping tool scrapes the data from the web page and converts that data into PPTX slide.

However, this is broken because the web page structure changed.

The limitation of this approach is: Each time the web page structure changes, we need to make corresponding changes in the main.py.


[1] Create pptx_from_web.py
Refactor main.py so that it talks to www.projectabcd.com/
(for example, the dress is 2 is here https://www.projectabcd.com/display_the_dress.php?id=2)

So, restore the functionality

[2] Create pptx_from_excel.py

Export the abcd data into excel.
And use this excel to generate the PPTX

[3] Create pptx_from_db.py

Let Python talk to MySQL database using MySQL adaptor / libary.
You can fetch the data from the database directly.

[4] Create pptx_from_api.py

This has dependency on the APIs Alligators are doing.
Pyhon calls the APIs
Gets the data in JSON format.
Uses the JSON data to create the PPTX.


[5] For all options 1, 2, 3 and 4, following inputs are required.
-- layouts (how do you want to generate the PPTX)
-- what to generate? (1 - 100; 1, 56, 78, 99...) or based on some query)
--- What option do you want to use? 1 or 2 or 3 or 4
-- Once we finalize the list of IDs, you use that list to retrieve the data using one of options 1,2,3, and 4.
-- No matter what option we chose, the final PPTX should be exactly the same.

[6] Mass update:

Using this option, we can use a CSV file to update the database. 
Basically, it backs up the database, wipes out the database, and reloads the database from CSV.







http://localhost/abcd2 pages
(or you ca 
