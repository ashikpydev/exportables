*! version 1.2.1
*! exportables_fixed.ado
*! Author: Ashiqur Rahman Rony (fixed by ChatGPT)
*! Email: ashiqurrahman.stat@gmail.com
*! Organization: Development Research Initiative (dRi)
*! Description: Same as exportables.ado but fixes rounding error so that
*!              the printed percentages sum to exactly 100.00 for each table.
*!
*! Strategy used:
*!  - Compute raw percentages in Stata and round them to 2 decimals.
*!  - Compute the rounding residual (100 - sum(rounded percentages)).
*!  - Add the residual to the option with largest frequency (common tie-breaker).
*!    This guarantees the row-wise displayed percentages sum to exactly 100.00.
*!  - This approach avoids writing prematurely-rounded values to Excel and then
*!    showing a non-100 total (e.g., 99.22) caused by cumulative rounding.

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

        * --- KEEP ONLY NUMERIC CHILDREN ---
        local children_numeric ""
        foreach c of local children {
            capture confirm numeric variable `c'
            if !_rc local children_numeric `children_numeric' `c'
        }
        local children `children_numeric'

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
            capture drop __tmp_case
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

            * --------- First pass: collect freqs and rounded percentages ---------
            local freqslist ""
            local pctroundlist_resp ""
            local pctroundlist_cases ""
            local sum_round_resp = 0
            local sum_round_cases = 0
            local maxfreq = -1
            local maxidx = 1
            local idx = 1

            foreach c of local children {
                quietly count if `c'==1
                local freq = r(N)
                local freqslist `freqslist' `freq'

                if `total_resp' > 0 {
                    local pct_r_resp = round(100*`freq'/`total_resp', 0.01)
                }
                else {
                    local pct_r_resp = 0
                }

                if `total_cases' > 0 {
                    local pct_r_cases = round(100*`freq'/`total_cases', 0.01)
                }
                else {
                    local pct_r_cases = 0
                }

                local pctroundlist_resp `pctroundlist_resp' `pct_r_resp'
                local pctroundlist_cases `pctroundlist_cases' `pct_r_cases'
                local sum_round_resp = `sum_round_resp' + `pct_r_resp'
                local sum_round_cases = `sum_round_cases' + `pct_r_cases'

                if `freq' > `maxfreq' {
                    local maxfreq = `freq'
                    local maxidx = `idx'
                }
                local ++idx
            }

            local resid_resp = 100 - `sum_round_resp'
            local resid_cases = 100 - `sum_round_cases'

            * --------- Second pass: write values to Excel, adjusting the largest cell ---------
            local i = 1
            foreach c of local children {
                local clabel : variable label `c'
                if "`clabel'" == "" local clabel = "`c'"

                local freq_val : word `i' of `freqslist'
                local pct_write_resp : word `i' of `pctroundlist_resp'
                local pct_write_cases : word `i' of `pctroundlist_cases'

                if `i' == `maxidx' {
                    quietly capture confirm number `pct_write_resp'
                    if _rc local pct_write_resp = `pct_write_resp' + `resid_resp'
                    quietly capture confirm number `pct_write_cases'
                    if _rc local pct_write_cases = `pct_write_cases' + `resid_cases'
                }

                putexcel A`row' = "`clabel'", border(all)
                putexcel B`row' = `freq_val', border(all)
                putexcel C`row' = `pct_write_resp', border(all)
                putexcel D`row' = `pct_write_cases', border(all)
                local ++row
                local ++i
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

                * denominator = number of non-missing responses for this variable
                quietly count if !missing(`v')
                local denom = r(N)

                levelsof `v', local(options)

                * --------- First pass: compute freqs and rounded percentages ---------
                local freqslist ""
                local pctroundlist ""
                local sum_round = 0
                local maxfreq = -1
                local maxidx = 1
                local i = 1

                foreach opt of local options {
                    quietly count if `v'==`opt'
                    local freq = r(N)
                    local freqslist `freqslist' `freq'

                    if `denom' > 0 {
                        local pct_r = round(100*`freq'/`denom', 0.01)
                    }
                    else {
                        local pct_r = 0
                    }

                    local pctroundlist `pctroundlist' `pct_r'
                    local sum_round = `sum_round' + `pct_r'

                    if `freq' > `maxfreq' {
                        local maxfreq = `freq'
                        local maxidx = `i'
                    }
                    local ++i
                }

                local resid = 100 - `sum_round'

                * --------- Second pass: write rows, adjust the largest option ---------
                local i = 1
                foreach opt of local options {
                    local lbl = "`opt'"
                    if "`valuelabel'" != "" {
                        local lbl2 : label (`valuelabel') `opt'
                        if "`lbl2'" != "" local lbl = "`lbl2'"
                    }

                    local freq_val : word `i' of `freqslist'
                    local pct_val : word `i' of `pctroundlist'
                    if `i' == `maxidx' {
                        local pct_val = `pct_val' + `resid'
                    }

                    putexcel A`row' = "`lbl'", border(all)
                    putexcel B`row' = `freq_val', border(all)
                    putexcel C`row' = `pct_val', border(all)
                    local ++row
                    local ++i
                }

                * Total row for single-select
                putexcel A`row' = "Total", bold border(all)
                putexcel B`row' = `denom', bold border(all)
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
