*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF


*** Variables ***
${Login_URL}        https://robotsparebinindustries.com/
${Username}         maria
${Password}         thoushallnotpass
${Salesdata_URL}    https://robotsparebinindustries.com/SalesData.xlsx
${Workbook_Name}    SalesData.xlsx


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log In
    Download Excel file
    Fill and submit the form using data from excel file
    Collect the results
    Create PDF from the table
    [Teardown]    Log out and close the Browser


*** Keywords ***
Open the intranet website
    Open Available Browser    ${Login_URL}

Log in
    Input Text    username    ${Username}
    Input Password    password    ${Password}
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download Excel file
    Download    ${Salesdata_URL}    overwrite=true

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Fill and submit the form using data from excel file
    Open Workbook    ${Workbook_Name}
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Log    ${sales_rep}
        Fill and submit the form for one person    ${sales_rep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Create PDF from the table
    Wait Until Page Contains Element    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log out and close the Browser
    Click Button    Log out
    Close Browser
