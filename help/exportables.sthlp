{smcl}
{* *! version 1.0.0 31aug2025 Ashikur Rahman}
{cmd:help exportables}
{hline}

{title:Title}

{phang}
{bf:exportables} {hline 2} Export frequency and multiple response tables to Excel.

{title:Syntax}

{p 8 15 2}
{cmd:exportables}
[{varlist}]
{cmd:,} {opt using(filename)}

{synoptset 20 tabbed}{...}
{synopthdr}
{synoptline}
{synopt:{opt using(filename)}}Specify the Excel file to which tables are exported. Required.{p_end}
{synopt:{it:varlist}}Optional list of variables to export. If omitted, all variables are processed.{p_end}
{synoptline}

{title:Description}

{pstd}
{cmd:exportables} creates clean frequency tables for single-select variables and multiple-response 
tables for select-multiple variables, and writes them directly to an Excel file.

{pstd}
The program distinguishes automatically between single- and multiple-select variables based 
on variable naming patterns (e.g., {it:var_1 var_2 var_3 ...} are grouped together).  

{pstd}
Percentages are reported with two decimal places, and totals are included for single-select variables.

{title:Options}

{phang}
{opt using(filename)} specifies the Excel file where output tables will be written. The file is created 
or replaced at the first export.

{phang}
{it:varlist} allows exporting only a subset of variables. If no variables are specified, the program 
processes all eligible variables in the dataset.

{title:Remarks}

{pstd}
This command is designed for survey datasets where select-one and select-multiple questions 
need to be summarized for reporting.  

{pstd}
By default:
{pmore}- Single-select variables → frequency + percent + total.
{pmore}- Multiple-select variables → cross-tab style multiple-response tables.

{title:Examples}

{phang}{cmd:. exportables, using("all_tables.xlsx")}{p_end}
{pmore}Exports all variables into {it:all_tables.xlsx}.{p_end}

{phang}{cmd:. exportables s2_5 s2_11, using("selected_tables.xlsx")}{p_end}
{pmore}Exports only {it:s2_5} and {it:s2_11} into {it:selected_tables.xlsx}.{p_end}

{title:Stored results}

{pstd}
{cmd:exportables} does not store results in memory. All output is written to Excel.

{title:Author}

{pstd}
Written by Ashikur Rahman ({browse "mailto:ashikur.rahman@email.com":ashikur.rahman@email.com})

{title:Acknowledgements}

{pstd}
This package builds on best practices in survey data cleaning and tabulation.  
Special thanks to collaborators and users who provided feedback.

{title:Also see}

{psee}
Manual: {bf:[R] putexcel}, {bf:[R] tabulate}, {bf:[R] table}

{psee}
Online: {manhelp putexcel R}, {manhelp tabulate R}
