# Stock Performance & VBA Code Refactoring

## **Project Overview**

In this project we leveraged our Microsoft Visual Basic Application (VBA) expertise to perform analyses for Steve, a recent Finance Graduate. Steve has recently taken on his guardians as his first clients and they are passionate about green energy, believing fossil fuel depletion will trigger ascendancy of alternative energy practices. Therefore, Steve's guardians would like to invest in a high yielding Stock within the aforementioned industry. We were provided with a dataset including subcategories of alternative energy organizations including, Hydro-energy, Bioenergy, Geothermal-energy, and Wind-energy. Adversely, Steve's parents have little knowledge of the analysis and performance of these stocks. Steve's guardians selected DAQO New Energy Corp (ticker: DQ), a company that makes silicon wafers for solar panels, where they invested all their funds.  

Our team performed analysis on the stock provided in the data set, which enabled Steve’s ability to diversify his guardian's portfolio and provide analysis on their current choice DAQO. Steve feels empowered by the outcomes of the dataset and wants to expand his analysis over the last few years. 

Although our code was efficient for the sample data provided, our current VBA script may have functionality challenges with larger data sets. Our goal for this project is to refactor code to optimize our scripts functionality for a larger data sample. We then will compare our beginning script and draw conclusions on our findings.

## **Results**

### *Results From Stock Analysis Metrics*


We found that that DAQO's return for the year 2018 dropped 62.6% contrary to their 199.4% increase over the year 2017. Similarly, we found Jinko Solar Holding Co., Ltd. (Ticker: JKS) and SunPower Corp (Ticker:SPWR) dropped %60.5 and 44.6% respectively. We can infer based on 2017 and 2018 performance that Enphase Energy (Ticker: ENPH) and SunRun (Ticker: RUN) are realizing strong growth over 2017 and 2018 and they would yield a solid return if invested in during 2017-2018 timespan. 

##### *Example A: 2017 Stock Performance Analysis (Data Identical for Original and Refactored) *

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/VBA_Challenge_DATA_2017.PNG)

##### *Example B: 2018 Stock Performance Analysis (Data Identical for Original and Refactored) *

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/VBA_Challenge_DATA_2018.PNG)

### *Results From Refactored Code Timing Metrics*

We discovered that refactoring the code has some positive effects on runtime for the VBA Macro's performance. From the metrics below we can draw the conclusion that refactoring execution can have positive effects on code workflow efficiencies.

##### Example A-1: Speed of Execution time for original code (2017)

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/Original_VBA_Challenge_2017.PNG)

##### Example A-2: Speed of Execution time for original code (2018)

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/Original_VBA_Challenge_2018.PNG)

##### Example B-1: Speed of Execution time for refactored code (2017)

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/VBA_Challenge_2017.png)

##### Example B-2: Speed of Execution time for original code (2018)

![](https://github.com/88lventerprises/excel_challenge/blob/main/Resources/VBA_Challenge_2018.png)

## **Summary**

### Advantages and Disadvantages of Refactoring Code in General

#### ***Advantages***

- Having a baseline of code to help organize outcome patterns allows developers to strategize on predefined outcomes 
- Time and workflow optimization for timelines, organizational KPI, Quarterly Metrics of any other deliverables
- Removal of redundant code/ simplifying code makes for cleaner and more digestible workflows for internal and external partners
- Streamline testability with clean shorter code rationalizing outcomes and supplying more interactive code structure 
- Improved legibility and comprehensibility for other programmers who may work on team
- Understanding different workflows through restructured but identical functionality
- Removing code smells aid with cohesiveness
- Improved design of existing code
  
#### ***Disadvantages***

- New bugs and errors introduced to code may cause difficulty in pinpointing issues 
- Code standards may be compromised 
- Large code and improper testing of that code can trickle down to larger issues 
- Can be extremely time-consuming interfering with guidelines and deadlines
- More prone to mistakes due to code complexity 


### Advantages and Disadvantages of Refactoring Code in Our VBA Script

#### ***Advantages***

- Refactored code can withstand larger data sets allowing more granular analysis
- Redundant data removed from for loops throughout out code
- Faster more efficient code with identical outcomes for the given data set
- Code is easier to follow and more digestible with comments that structure outcomes: The code details a sound picture of what is happening within loops, what is being referenced when calling in code blocks, and demonstrates a more streamlined code structure
  
#### ***Disadvantages***

- Persistent debugging error that required rebooting of VBA: This was a huge roadblock as reimagining workflows and process can lead to gaps in logic causing debug errors in code blocks
- Code complexity led to stagnated workflows:  In order to combat this I understood that the outcome of my original code set would have operational similarities so I leveraged the preexisting code to refactor the current code.






