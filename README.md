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



Sample of the data is as shown below

![My Excel Dashboard](Screenshot%202026-06-06%20084548.png)
    


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



Identify the most expensive and least expensive products

Most expensive 
               = MAX(B2:B113)
               =32 PCS Portable Cordless Drill Set with Cyclic Battery Drive -26 variable speed


Least expensive
                = MIN(B2:B113)
                = 3 PCS Single Head Knitting Crochet Sweater Needle Set




### Trend Analysis

Identfying the top 5 products with the highest ratings and top 5 products with the highest and lowest ratings 

- Go to the column with the rating and select number filter top 5
  
  
The following are what we get


  *Anti-Skid Absorbent Insulation Coaster  For Home Office  - 5*
  
  *Peacock  Throw Pillow Cushion Case For Home Car - 5*
  
  *LASA Aluminum Folding Truck Hand Cart - 68kg Max  - 5*
  
  *DIY File Folder, Office Drawer File Holder, Pen Holder, Desktop Storage Rack  - 5*
  
  *Classic Black Cat Cotton Hemp Pillow Case For Home Car  - 5*


- Go to the column with the ratings and select number filter bottom 5
  


*Wall-mounted Sticker Punch-free Plug Fixer - 2*
  
  *5-PCS Stainless Steel Cooking Pot Set With Steamed Slices - 2.1*
  
  *Electric LED UV Mosquito Killer Lamp, Outdoor/Indoor Fly Killer Trap Light -USB - 2.1*
  
  *Artificial Potted Flowers Room Decorative Flowers (2 Pieces) - 2.2*
  
  *7-piece Set Of Storage Bags, Travel Storage Bags, Shoe Bags - 2.2*



Identify the top 10 products with the highest discounts , Top 10 products with most reviews

- Go to the column with products and select filter top 10

  *5-PCS Stainless Steel Cooking Pot Set With Steamed Slices  - 2585*
  
  *LASA 3 Tier Bamboo Shoe Bench Storage Shelf  -  2452*
  
  *32PCS Portable Cordless Drill Set With Cyclic Battery Drive -26 Variable Speed  - 2393*
  
  *LASA Aluminum Folding Truck Hand Cart - 68kg Max  - 1946*
   
  *MultiFunctional Storage Rack Multi-layer Bookshelf - 1880*
  
  *LASA Stainless Steel Double Wall Mount Soap Dispenser - 500ml -  1721*
  
  *LED Eye Protection  Desk Lamp , Study, Reading, USB Fan - Double Pen Holder - 1670*
  
  *LASA FOLDING TABLE SERVING STAND - 1526*
  
  *Electric LED UV Mosquito Killer Lamp, Outdoor/Indoor Fly Killer Trap Light -USB  - 1418*
  
  *Angle Measuring Tool Full Metal Multi Angle Measuring Tool  -   1399*


- Go to the colum with reviews and select top 10

  *120W Cordless Vacuum Cleaners Handheld Electric Vacuum Cleaner  -  69*
  
  *137 Pieces Cake Decorating Tool Set Baking Supplies  - 55*
  
  *Electronic Digital Display Vernier Caliper  - 49*
  
  *3D Waterproof EVA Plastic Shower Curtain 1.8*2Mtrs  - 44*
  
  *100 Pcs Crochet Hook Tool Set Knitting Hook Set With Box  - 39*
  
  *Punch-free Great Load Bearing Bathroom Storage Rack Wall Shelf-White  - 36*
  
  *53 Pieces/Set Yarn Knitting Crochet Hooks With Bag - Pansies  - 32*
  
  *Portable Mini Cordless Car Vacuum Cleaner - Blue   -  24*
  
  *53Pcs/Set Yarn Knitting Crochet Hooks With Bag - Fortune Cat   -  20*
  
  *52 Pieces Cake Decorating Tool Set Gift Kit Baking Supplies   -  20*


  ### DASHBOARD

  Here is how the dashboard looks like

