# Stock-Analysis

### 1. Overview of Project:
##### Taking a previously prepared stocks analysis file that summarizes tickers by daily volume and return and refactoring the code to run faster and more efficiently. We have stock data for 12 tickers from 2017 and 2018.

### 2. Analysis and Challenges:
##### The dataset needed to be formatted prior to any analysis. I added the 'year' column to the very end so I can use it inside of a pivot table. At that point, the dataset was complete and I could create a pivot table to summarize my data. In order to view theater outcomes by date, I dragged months into rows, outcomes into columns, and count of outcomes into values. This allowed me to visualize this information as a line chart. The last piece was to filter the visual to only see the theater parent category. See visual below
![Theater_Outcomes_vs_Launch](https://github.com/maldonado91/Fundraiser-Success-Analysis-in-Excel/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
##### The outcomes based on goals were populated via formulas. I converted the information into a table and graphed it as a line chart. See visual below
![Outcomes_vs_Goals](https://github.com/maldonado91/Fundraiser-Success-Analysis-in-Excel/blob/main/Resources/Outcomes_by_Goals_Range.png)
##### Any formatting, formulas, or pivot tables can be found in this [file](https://github.com/maldonado91/Fundraiser-Success-Analysis-in-Excel/blob/main/Kickstater_Challenge.xlsx)
##### The biggest challenges were in the graping piece of outcomes vs goals. I could not get the graph to display properly. After some research/trial and error, I realized I was hilghiting the entire table versus just the columns I was interested in graphing (Goal, Percentage Successful, Percentage Failed, Percentage Canceled).

### 3. Results:
##### After refactoring the code the analysis for all stock 2017 and 2018 displayed the exact same results. 

##### What can you conclude about the Outcomes based on Goals?
    -Outside of very low goals (Less Than $1,000), plays with goals between $35,000 to $40,000 and $40,000 to $45,000 have the highest success rate

##### What are some limitations of this dataset?
    -We only have two fields to break data down into smaller groups (parent category and subcategory). More categories can help drill down further into the dataset.

##### What are some other possible tables and/or graphs that we could create?
    -I thought a Histogram would be helpful to look at Goal vs Number Successful since most successful campaigns have a goal of under $20,000
