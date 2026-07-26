# Excel-Salary-Dashboard
# Data Jobs Salary Dashboard (Excel)

Interactive Excel dashboard for exploring 2023 data-job salaries by role, country, and contract type. The dataset and brief come from Luke Barousse's Excel course; the build is my own.

**File:** `Salary_Dashboard.xlsx`

## What it does

Three dropdowns (Job Title, Country, Contract Type) drive everything at once: median salary by role (bar), median salary by country (map), median salary by contract type (bar), plus KPIs for median salary, top hiring platform, and job count.


<img width="1366" height="479" alt="EXCEL_t155KwCqPe" src="https://github.com/user-attachments/assets/28fdcba5-90dc-413c-9e0c-76c0bd89b110" />



                                                                    
**Excel Skills used**:  

•	📉 Charts  

•	🧮 Formulas and Functions  

•	❎ Data Validation  



## Build notes

**One consistent population.** The raw `job_schedule_type` field contains compound values like `"Full-time and Part-time"`. Exact matching silently drops them, so I added one helper column to the fact table:

```excel
=ISNUMBER(SEARCH(Type,**jobs[job_schedule_type]**))
```
Every count on the dashboard filters through it, so the median salary, job count, and platform count all describe the same rows. Without it each one applies its own filter and the three KPIs stop agreeing.
SUMPRODUCT in one place, COUNTIFS in another. 

SUMPRODUCT materialises the full 32,672-row comparison per cell. Across 10 job titles that costs nothing and reads better.

```=SUMPRODUCT((jobs[job_title_short]=A3)*(jobs[job_country]=Country)*(jobs[Type_match]=TRUE)) ```

<img width="611" height="201" alt="EXCEL_H6WFMdG07y" src="https://github.com/user-attachments/assets/7564d046-a316-4a72-b79a-edf636729c13" />


Across 594 platforms it lags on every dropdown change, so that table uses COUNTIFS, a native aggregate that never builds the array. Both return the same number; the difference only shows up at scale.

```=COUNTIFS(jobs[job_via],A4,jobs[job_title_short],Title,jobs[job_country],Country,jobs[Type_match],TRUE)```

<img width="746" height="179" alt="EXCEL_Yt6XZfZdW7" src="https://github.com/user-attachments/assets/f15184d1-12be-4bae-8fbb-6e8c0b017f1a" />


The map changed measure mid-build. I originally had plotted job count but later I realised it was useless as the US dominates so heavily that everything else rendered the same pale shade. Median salary works because comparable roles pay very differently by region, and that has nothing to do with job volume

# Limitations  

•	2023 postings only, so none of this shows a trend.  

•	Country medians have no minimum-sample guard. Some are computed from three postings and render exactly like ones computed from thousands.  

•	Only postings that disclose salary can be counted, and firms that publish salaries aren't a random sample.  

•	The Type_match helper puts 32,672 dashboard-dependent formulas inside the fact table. That's the price of the consistency described above: every selection triggers a full recalc, and the Data sheet is no longer a reusable extract. In a larger build the logic belongs in Power Query.  
