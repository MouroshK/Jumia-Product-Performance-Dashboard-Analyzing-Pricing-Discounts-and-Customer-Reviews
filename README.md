# Jumia Product Performance Dashboard: Analysing Pricing, Discounts and Customer Reviews


## Project Overview

In this project I'll be able to clean the Jumia dataset with excel ,create an interactive Excel dashboard that provides insights into the performance of products listed on Jumia.


The Dataset contains the following columns :

    • Product - Name of the product 
    • Current price - The current selling price of the product (in Ksh)
    • Old price - The original price of the product (in Ksh)
    • Discount - The discount percentage offered on the product 
    • Review - The number of reviews for the products 
    • Rating - The average customer rating of the product (out of 5)



    


## Data cleaning 

The raw data was messy and I performed the following data cleansing steps :

 * **Checked for missing values** - I checked missing values in all the columns and found out that the Product, Current price, Old Price and Discount columns had no blanks while reviews and Ratings have 59 blanks 


* **Removed Duplicates**  - Checked for duplicates and found 3 duplicates and removed them.
 

* **Data Standardization** - Removed Ksh in the old and current price by finding Ksh and replace with blank , Removed the comma in the prices column , Removed out of 5 by using find out of 5 and replace with nothing 



## Data Enrichment

* **New metrics** Used Excel formulas
  1. Created a new column of discount where I calculated by getting difference of ( Old price - Current price) that is =C2 - B2 then draggged to autofill
  2. Created a new column to categorize products based on their ratings . Used the IF statement
 = IF(G2<3,"Poor",IF(G2<=4.4,"Average",IF(A2>=4.5,"Excellent","N/A")))
  3. Created a new column to categorize products based on their discount percentage , Used IF statement 
=IF(E5<=20%,"Low Discount",IF(E5<=40%,"Medium Discount",IF(E5>40%,"High Discount","N/A")))


## Data Analysis

### Descriptive statistics

Calculate the average current price, old price ,discount percentage and rating across all the products

Average Current Price = AVERAGE(B2:B113)
                                        = 1181.369369

Average Old Price = AVERAGE(C2:C113)
                                = 1803.099

Average Discount Percentage = AVERAGE(E2:E113)
                                                     = 37%

AVERAGE RATING = AVERAGE(H2:H113)
