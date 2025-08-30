*! exportables.sthlp
*! Author: Ashikur Rahman
*! Description: Export survey tables to Excel

{\ttitle exportables — Export survey tables to Excel}

{\b Syntax}

\begin{verbatim}
exportables [, using(filename.xlsx)]
\end{verbatim}

{\b Description}

Exports frequency tables of all variables in the dataset to an Excel file. Supports:

- Single-select variables: frequency, percent (2 decimals), total row included.
- Multi-select variables: frequency, percent of responses, percent of cases, total row included.
- Automatically skips _oth and _rank columns in multi-select variables.

{\b Options}

\begin{tabular}{lp{10cm}}
{\tt using(filename.xlsx)} & Name of the Excel file to create (will replace if exists). \\
\end{tabular}

{\b Example}

\begin{verbatim}
. exportables, using("all_tables.xlsx")
\end{verbatim}

Exports all survey tables to `all_tables.xlsx` with formatted totals and percentages.

{\b Author}

Ashikur Rahman  

Credits: Developed for automated survey reporting in Stata.
