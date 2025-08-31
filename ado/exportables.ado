*! version 1.2.0
*! exportables.ado
*! Author: Ashiqur Rahman Rony
*! Email: ashiqurrahman.stat@gmail.com
*! Organization: Development Research Initiative (dRi)
*! Description: Export single-select and multi-select survey tables to Excel
*!              with totals and percentages (rounded to two decimals)
*!              Automatically handles split variables for multi-select questions.
*!              Optional variable selection added: by default exports all, or only selected.

capture program drop exportables
program define exportables
    version 17
    syntax [varlist] , using(string)

    * --- DETERMINE VARIABLES TO PROCESS ---
    if "`varlist'" == "" {
        ds
        local varlist `r(varlist)'
    }

    * --- ALIGN VALUE LABELS WITH VARIABLE NAMES ---
    foreach var of local varlist {
        capture confirm variable `var'
        if _rc continue
        local vlabel : value label `var'
        if "`vlabel'" != "" {
            capture label copy `vlabel' `var'
            label values `var' `var'
        }
    }

    * --- SETUP EXCEL ---
    putexcel set "`using'", replace sheet("AllTables")
    local row = 1
    local tablecount = 0

    * --- LOOP OVER ALL VARIABLES TO EXPORT ---
    foreach v of local varlist {
        capture confirm variable `v'
        if _rc continue

        * --- CHECK FOR MULTI-SELECT CHILDREN ---
        ds
        local allvars `r(varlist)'
        local children ""
        foreach c of local allvars {
            if strpos("`c'", "`v'_") == 1 & regexm("`c'", ".*(_oth|_rank.*)$")==0 {
                local children `children' `c'
            }
        }

        * --- MULTI-SELECT VARIABLE ---
        if "`children'" != "" {

            local vlabel : variable label `v'
            if "`vlabel'" == "" local vlabel = "`v'"

            putexcel A`row' = "Variable: `v' (`vlabel')", bold border(all)
            local ++row
            putexcel A`row' = "Option", bold border(all)
            putexcel B`row' = "Frequency", bold border(all)
            putexcel C`row' = "Percent of responses", bold border(all)
            putexcel D`row' = "Percent of cases", bold border(all)
            local ++row

            * total cases = at least one child ticked
            gen byte __tmp_case = 0
            foreach c of local children {
                quietly replace __tmp_case = 1 if `c'==1
            }
            quietly count if __tmp_case==1
            local total_cases = r(N)
            drop __tmp_case

            * total responses = sum across numeric dummies
            local total_resp = 0
            foreach c of local children {
                quietly count if `c'==1
                local total_resp = `total_resp' + r(N)
            }

            * loop over children
            foreach c of local children {
                local clabel : variable label `c'
                if "`clabel'" == "" local clabel = "`c'"

                quietly count if `c'==1
                local freq = r(N)
                local pct_resp = cond(`total_resp'>0, 100*`freq'/`total_resp', .)
                local pct_cases = cond(`total_cases'>0, 100*`freq'/`total_cases', .)

                putexcel A`row' = "`clabel'", border(all)
                putexcel B`row' = `freq', border(all)
                putexcel C`row' = `=round(`pct_resp',0.01)', border(all)
                putexcel D`row' = `=round(`pct_cases',0.01)', border(all)
                local ++row
            }

            * totals row
            putexcel A`row' = "Total", bold border(all)
            putexcel B`row' = `total_resp', bold border(all)
            putexcel C`row' = 100, bold border(all)
            putexcel D`row' = "", bold border(all)
            local ++row
            local ++row
            local ++tablecount

        }
        * --- SINGLE-SELECT VARIABLE ---
        else {
            local valuelabel : value label `v'
            if "`valuelabel'" != "" {

                local vlabel : variable label `v'
                if "`vlabel'" == "" local vlabel = "`v'"

                putexcel A`row' = "Variable: `v' (`vlabel')", bold border(all)
                local ++row
                putexcel A`row' = "Option", bold border(all)
                putexcel B`row' = "Frequency", bold border(all)
                putexcel C`row' = "Percent", bold border(all)
                local ++row

                levelsof `v', local(options)
                local total = 0
                foreach opt of local options {
                    quietly count if `v'==`opt'
                    local freq = r(N)
                    local total = `total' + `freq'

                    local txt = "`opt'"
                    if "`valuelabel'" != "" {
                        local lbl : label (`valuelabel') `opt'
                        if "`lbl'" != "" local txt = "`lbl'"
                    }

                    local pct = cond(`total'>0, 100*`freq'/`=_N', .)
                    putexcel A`row' = "`txt'", border(all)
                    putexcel B`row' = `freq', border(all)
                    putexcel C`row' = `=round(`pct',0.01)', border(all)
                    local ++row
                }

                * Total row for single-select
                putexcel A`row' = "Total", bold border(all)
                putexcel B`row' = `total', bold border(all)
                putexcel C`row' = 100, bold border(all)
                local ++row
                local ++row
                local ++tablecount
            }
        }
    }

    * --- Final Message ---
    di as txt "{hline 65}"
    di as txt "                 " as result "✔ EXPORT COMPLETED SUCCESSFULLY ✔"
    di as txt "{hline 65}"
    di as txt "   Number of tables created : " as result `tablecount'
    di as txt "   File saved as            : " as result "`using'"
    di as txt "{hline 65}"
    di as txt "        Thank you for using " as result "exportables" as txt "!"
    di as txt "        Created by: Ashiqur Rahman Rony | Data Analyst | Development Research Initiative (dRi) | ashiqurrahman.stat@gmail.com"
    di as txt "{hline 65}"
end
