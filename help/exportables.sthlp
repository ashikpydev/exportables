*! version 2.2.0
*! exportables.sthlp
*! Author: Ashikur Rahman
*! Contact: [Your email or GitHub link]
*! Description: Help file for exportables.ado

{\ttitle exportables — Export survey tables to Excel with totals and two-decimal percentages}

{\b Syntax}

\begin{verbatim}
exportables [, using(filename.xlsx)]
\end{verbatim}

{\b Description}

{\tt exportables} generates frequency tables for all variables in the current dataset and exports them to an Excel file. It supports both:

1. Single-select variables (numeric variables with value labels).
2. Multi-select variables (dummy-coded variables with names like `var_1`, `var_2`, …).

For single-select variables:

- Each option is listed with frequency and percent.
- A total row is added at the bottom.
- Percentages are rounded to 2 decimal places.

For multi-select variables:

- Each child variable is listed with frequency, percent of responses, and percent of cases.
- Percentages are rounded to 2 decimal places.
- A total row of responses is included.

{\b Options}

\begin{tabular}{lp{10cm}}
{\tt using(filename.xlsx)} & Name of the Excel file to create. The file will be replaced if it already exists. \\
\end{tabular}

{\b Examples}

Export all tables to an Excel file named `all_tables.xlsx`:

\begin{verbatim}
. exportables, using("all_tables.xlsx")
\end{verbatim}

This will create a sheet called `AllTables` with formatted frequency tables for all variables in the dataset.

{\b Notes}

- Single-select variables must have value labels assigned.
- Multi-select variables should be numeric dummy-coded (0/1) and follow a naming convention `var_1`, `var_2`, etc.
- Columns `_oth` or `_rank` are automatically excluded from multi-select tables.
- Totals and percentages are formatted for research reporting.

{\b Author and Credits}

Author: Ashikur Rahman  
Contact: [Your email or GitHub link]  

This ado was inspired by common survey reporting needs in humanitarian and social science research, aiming to automate the tedious process of creating clean Excel frequency tables with totals and percentages.  

Contributions welcome via GitHub or email.

{\b Version History}

2.2.0 — Added total row for single-select, two-decimal percentages, fixed rounding in Excel.  
2.1.0 — Initial release supporting both single-select and multi-select tables.  

{\b See Also}

\hyperlink{tabstat}{tabstat}, \hyperlink{tabulate}{tabulate}, \hyperlink{putexcel}{putexcel}, \hyperlink{levelsof}{levelsof}
