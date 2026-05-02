## Detailed description of steps in data preprocessing

The data set for financial information was received as Excel files via e-mail from Proff AS (2026). This Norwegian firm is one of the official suppliers of such information, making financial reporting from the original source, The Brønnøysund Register Centre, publicly available. Included in the original data set is yearly financial reporting for the years 2019-2024, for all limited companies registered in Norway in 2018 and 2019. All 64 columns present in the original Excel files are listed in Table D1. Including header row, the data set contained 60,165 rows of information. 

After removing unused columns and combining the data in a single sheet, the key profitability measure is constructed: aggregated operating margin, defined as the sum of operating income over the first five years divided by the sum of total operating revenue over the same period. Pretax margin for year 5 is also calculated for the 2019 cohort. Observations with aggregated revenue is zero or negative are set to missing, since operating margin is undefined when revenue is zero and produces misleading values when revenue is negative. The same exclusion rule is applied to pretax margin.

Complementing the financial data is additional information on companies retrieved from The Brønnøysund Register Centre. As not all required information on deleted companies is available in the data received from Proff AS, nor openly available from The Brønnøysund Register Centre, this complementing data set is retrieved as Excel files from The Brønnøysund Register Centre (2026a). Information on the company objective as stated in the official memorandum of association for each company, as well as the size of the share capital and number of employees, are examples of data included in the data set from The Brønnøysund Register Centre. All 71 columns present in the original Excel files are listed in Table D2. Before merging this data with the data from Proff AS, columns not being used for analysis are removed. To simplify further processing, the data from Proff AS and The Brønnøysund Register Centre are combined in a single Excel file, linking data using Organization number, the unique identifier for all limited companies.

Next, a new column is created, where the trimmed, two-digit NACE-code for each company is added. Second, columns for survival status according to the definition of survival as described in the Conceptual Background section in this study are created. Two columns are used in the further analysis; One column contains the relative year for exit, and the other column indicates the reason for exit. Observations with no values in these columns indicate the company has survived. 

Starting with the 486 terms found in categories “Environment and energy” and “Environment glossary” in the EGSS classification system (European Union, n.d.-a), duplicates are removed and keywords are manually verified. When manually verifying keywords, some keywords are removed or simplified, especially to ensure identification of sustainability-oriented startups are not hindered by keywords consisting of several single words in exact phrases. For example, the two keywords “Energy balance - old methodology” and “Energy balance - new methodology” are simplified to one keyword: “Energy balance”, which ensures all companies using the keyword “Energy balance” are identified, which is the part of the original keywords relevant for sustainability identification. 

This results in 257 keywords from EGSS. Translating to Norwegian and adding 106 relevant keywords from sustainability literature, a total of 363 keywords are used for the final analysis. There are some overlaps between EGSS keywords and the added keywords, typically because EGSS keywords are more detailed. For example, EGSS has multiple keywords containing “environment”, like “Agri-environmental indicator - soil erosion”, while “environment” is one of the added keywords. Due to wildcard matching (e.g. “*Agri-environmental indicator - soil erosion*” and “*environment*”) the keyword list is not simplified further as this would not change the final results. Keeping both adds the possibility to see if a description is matched by an EGSS keyword, an added keyword, or both. 

False positives in the first step in the process is a consequence mainly of the wildcard approach, but this is remediated by manually assessing all positives and eliminating false positives in the next step. An example is the keyword “natur”, and the Norwegian term “...som naturlig faller sammen med..”, which is a commonly used phrase in company descriptions, but in itself has no relation to sustainability. Similarly, the keyword for “electric” will give a positive result for electricians and retailers of electronic cigarettes. Another example is “miljø”, a central keyword for environmentally sustainable firms, but can also be used when addressing social sustainability, which is not in scope for this study. The downside of this approach is that the manual step can be time-consuming, but it results in higher data quality and fewer false negatives.

The example companies in Table 7 are assessed as follows: Company B has 5 keyword hits, which is due to “hydro” (“vann” in Norwegian”) and “hydropower” (“vannkraft” in Norwegian) being separate keywords, and thus the word “hydropower” will give two keyword hits. In Norwegian, “electricity” and “electric” are written as “elektrisitet” and “elektrisk”, and since “elektrisk” is the keyword only “electric”, not “electricity” gives a keyword hit, which is intended. 

Looking at Company C, “waste” (“avfall” in Norwegian) gives keyword hits in both EGSS and own keywords, while “recycling” (“gjenvinning” in Norwegian) gives keyword hit both for EGSS keyword “recycling” and own keyword “recycl”, resulting in 7 keyword hits in total. The same word twice in the same text does only give one hit for that word per keyword list. 

Similarly, Company D gets keyword hits both in EGSS and own keywords for the word “waste”. For Company E, the word “organic” (“økologisk” in Norwegian) gives two hits due to the original Norwegian word (“økologiske”) contains both “øko”, the Norwegian word for “eco”, and “økologi”, the Norwegian word for “organic”. As for Company F, “forest” (“skog” in Norwegian) is found in both EGSS and own keywords.

---

## Excel formulas used to identify possible sustainability-oriented startups

### Commands that map keywords (EGSS and own) and company objectives and activity descriptions

Formulas are for row 2, and CM2/CL2 is adjusted for each row. Explanation of formula (using the first formula as an example):

```
SØK(Keywords!$A$2:$A$262;CL2)
```

For each keyword in the range A2:A262, search for it within the text in cell CL2. Returns the position number where the keyword is found, or an error if not found.

```
--ERTALL(...)
```

Check if the result from SØK is a number. Returns SANN (TRUE) if the keyword was found, USANN (FALSE) if SØK returned an error. The double minus converts SANN/USANN to 1/0, making it possible to do math with the results.

```
LENGDE(TRIMME(Keywords!$A$2:$A$262))>0
```

For each keyword cell, it trims whitespace (TRIMME) and checks the remaining length (LENGDE). Returns 1 if the cell has actual content, 0 if empty or only whitespace. This prevents empty keyword cells from producing false positives.

```
(LENGDE(...)>0) * --ERTALL(...)
```

Multiplies the two arrays together. A result of 1 means both conditions are met: the keyword cell is not empty AND the keyword was found in the text. If either condition fails, the result is 0.

```
SUMMERPRODUKT(...)
```

Sums up all the 1s and 0s. The final number is the count of how many non-empty keywords are found in the text in CL2.

---

### Formulas

```
=SUMMERPRODUKT((LENGDE(TRIMME(Keywords!$A$2:$A$262))>0)*--ERTALL(SØK(Keywords!$A$2:$A$262;CL2)))
=SUMMERPRODUKT((LENGDE(TRIMME(Keywords!$B$2:$B$108))>0)*--ERTALL(SØK(Keywords!$B$2:$B$108;CL2)))
=SUMMERPRODUKT((LENGDE(TRIMME(Keywords!$A$2:$A$262))>0)*--ERTALL(SØK(Keywords!$A$2:$A$262;CM2)))
=SUMMERPRODUKT((LENGDE(TRIMME(Keywords!$B$2:$B$108))>0)*--ERTALL(SØK(Keywords!$B$2:$B$108;CM2)))
```

---

### Adding count for EGSS and own keywords and company objective and activity in one column

```
=CN2+CO2+CP2+CQ2
```
