# Introduction
This project uses Power Query for ETL processes and Microsoft Excel for data insights on Emag's most featured laptop offerings by focusing on key metrics such as ratings, number of reviews and price. The information gathered and presented can serve a prospective customer looking for a new laptop as well as a brand representative of a laptop manufacturer.

# ETL process and data cleanup using Power Query
The data used for this analysis has been gathered by using Microsoft Power Query, a data ingestion, cleaning, and transformation engine which is built into Microsoft Excel and other MS products. In order to extract the data needed from emag.ro within excel we first have to navigate into the Power Query Editor by going through 

_**Data -> Get Data -> Launch Power Query Editor**_.

Once the power query editor has been opened the data sources can be added by going into 

_**Home -> New Source -> Other Srouces -> Web**_. 

The links used for the web source field have been the first five pages of the Emag's laptops section, sorted by the most relevant: 

- https://www.emag.ro/laptopuri/c
- https://www.emag.ro/laptopuri/p2/c
- https://www.emag.ro/laptopuri/p3/c
- https://www.emag.ro/laptopuri/p4/c
- https://www.emag.ro/laptopuri/p5/c

Each time has been selected the first option among the Suggested Tables in the Navigator (Table 1), giving a total of five tables (which have been named Page 1-5 for convenience's sake).

![image 1](https://github.com/AnghelG/Emag_Laptop_Data/blob/95119a7405aec3a37c9c7afca648cd96963aaf6a/Screenshots/1%20Initial%20View.png)

As it can be seen from the preview each table corresponds to one of the Emag's first five laptop category pages. The suggested table added a lot of empty columns and nonsensical data. These fields will have to be removed from each table first, leaving only with the ones with valuable data. Once this step is finished we can combine all the records from the five tables into one through appending. It is important to make sure that the fields from each table in order to properly append them. The tables will be appended into a new table called Data by going through

_**Home -> Append Queries -> Append Queries as New**_. 

![image 2](https://github.com/AnghelG/Emag_Laptop_Data/blob/95119a7405aec3a37c9c7afca648cd96963aaf6a/Screenshots/2%20Appended.png)

As it can be noticed we're left with only four columns of relevant data: full name, price, rating, and number of reviews. In order for the dataset to be loaded into a spreadsheet we first have to ensure data integrity through the transformation process. 

From the right to the left we first oberve that the last column (Column12) is of type text. Even though it represents the number of reviews, the entries are recorded as "_number_ review-uri" or "_number_ de review-uri". To get rid of this we can split the column by the space delimiter (" "), giving us 3 new columns with the first containing the actual number of reviews and then removing the latter two.

The second column (Column11) contains the rating given by the user to the specific laptop in the dataset, being of type decimal

The third column (Column2) contains the price of each laptop but is of type text as the entries are recorded as "_number_ lei". We can correct this by splitting the column by the delimiter (" ") giving us two columns: the actual price and the "Lei" coulumns latter of which will be removed. It's important to note that this alone won't change the remaining column into the desired data type (decimal) as the prices are recorded using the comma number formatting for decimals rather than the point based one which Power Query uses. This can be changed within Power Query by changing the datatype using locale setting, although another option would be changing the values within the entries through replacing them. We can change the "." to "," and "," to "." by first using a placeholder like "^" (as replacing one with the other instead of the placeholder would turn both of them of the same type).

The last column (Column1) Contains the full name of the product as it appears on the Emag web page. Although serving as a unique identifier for each row the names are too long and difficult to read in which case there will be an index column added for this purpose. One thing that can be observed though is that the brand (or manufacturer) of each laptop is mentioned within the first three words of the name. In order to extract this information the column is going to be sliptted by the space delimiter (" ") and all but the first free resulting columns will be removed.

![image 3](https://github.com/AnghelG/Emag_Laptop_Data/blob/76c8434c8f7714b4f8e074d90a447e142aa6f985/Screenshots/3%20Cleaned.png)



# Dashboarding, statistics and exploratory data analysis using MS Excel

# Conclusions

# Downloading and usage
