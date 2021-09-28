*** Settings ***
Documentation     Orders robots from RobotSpareBin Industries Inc.
...               Saves the order HTML receipt as a PDF file.
...               Saves the screenshot of the ordered robot.
...               Embeds the screenshot of the robot to the PDF receipt.
...               Creates ZIP archive of the receipts and the images.
Library         Dialogs
Library         RPA.Browser.Selenium
Library         RPA.Word.Application
Library         OperatingSystem
Library         String
Library         Collections
Library         RPA.HTTP
Library         RPA.Tables
Library         RPA.SAP
Library         RPA.PDF
Library         RPA.Archive
Library         RPA.Robocorp.Vault


*** Variables ***
${PDF_TEMP_OUTPUT_DIRECTORY}=    ${CURDIR}${/}temp
${OUTPUT_DIRECTORY}=    ${CURDIR}${/}output
${PNG_TEMP_OUTPUT_DIRECTORY}=    ${CURDIR}${/}temp_img

*** Keywords ***
Get orders csv file
    ${orders_file_link}    Get Value From User   Input link to orders csv file    default
    [RETURN]    ${orders_file_link}

*** Keywords ***
Set up directories
    Create Directory    ${PDF_TEMP_OUTPUT_DIRECTORY}
    Create Directory    ${PNG_TEMP_OUTPUT_DIRECTORY}
    Create Directory    ${OUTPUT_DIRECTORY}

*** Keywords *** 
Open the robot order website
    Open Available Browser      https://robotsparebinindustries.com/#/robot-order
    Maximize Browser Window

*** Keywords *** 
Get orders 
    [ARGUMENTS]    ${orders_file_link}
    Download        ${orders_file_link}     ${CURDIR}/orders.csv        overwrite=True
    @{orders}=      Read Table From Csv     ${CURDIR}/orders.csv
    [Return]        ${orders}

*** Keywords *** 
Close the annoying model
    Click Button When Visible       xpath:/html/body/div/div/div[2]/div/div/div/div/div/button[1]

*** Keywords *** 
Fill the form
    [Arguments]   ${head}   ${body}     ${legs}     ${address}
    Click Element When Visible      id:head
    FOR     ${i}    IN RANGE       0    ${head}
        Press Keys      None        ARROW_DOWN
    END
    Press Keys      None        ENTER
    
    Click Element When Visible      id:id-body-${body}
    
    Click Element When Visible      xpath:/html/body/div/div/div[1]/div/div[2]
    FOR        ${i}     IN RANGE     10
        Press Keys      None        ARROW_DOWN
    END
    RPA.Browser.Selenium.Input Text      xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input        ${legs}
    
    RPA.Browser.Selenium.Input Text      id:address          ${address}


*** Keywords *** 
Preview the robot
    Click Button When Visible       id:preview


*** Keywords *** 
Submit the order
    FOR     ${i}       IN RANGE     9999
        Click Button When Visible       id:order
        ${present}=  Run Keyword And Return Status    Element Should Be Visible    id:receipt
        Exit For Loop If    ${present}
    END

*** Keywords *** 
Store the receipt as a PDF file
    [Arguments]     ${order_number}
    ${string}=      Get Text    id:receipt
    ${pdf_name}=    Catenate    ${order_number}_receipt
    Html To Pdf    ${string}    ${PDF_TEMP_OUTPUT_DIRECTORY}/${pdf_name}.pdf
    [RETURN]        ${pdf_name}

*** Keywords *** 
Take a screenshot of the robot
    [Arguments]     ${order_number}
    ${robot_name}=      Catenate    ${order_number}_robot
    Screenshot      id:robot-preview-image      filename=${PNG_TEMP_OUTPUT_DIRECTORY}/${robot_name}.png
    [RETURN]        ${robot_name}

*** Keywords *** 
Embed the robot screenshot to the receipt PDF file
    [Arguments]     ${screenshot}    ${pdf}
    
   Add Watermark Image To PDF
    ...             image_path=${PNG_TEMP_OUTPUT_DIRECTORY}/${screenshot}.png  
    ...             source_path=${PDF_TEMP_OUTPUT_DIRECTORY}/${pdf}.pdf
    ...             output_path=${PDF_TEMP_OUTPUT_DIRECTORY}/${pdf}.pdf

*** Keywords *** 
Go to order another robot
    Click Button When Visible       id:order-another

*** Keywords *** 
Create a ZIP file of the receipts
    ${zip_file_name}=    Set Variable    ${OUTPUT_DIRECTORY}/PDFs.zip
    Archive Folder With Zip
    ...    ${PDF_TEMP_OUTPUT_DIRECTORY}
    ...    ${zip_file_name}
    Sleep       10

*** Keywords ***
Cleanup temporary PDF directory
    Remove Directory    ${PDF_TEMP_OUTPUT_DIRECTORY}    True
    Sleep               10
    Remove Directory    ${PNG_TEMP_OUTPUT_DIRECTORY}    True

*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    #${orders_file_link}=    Get orders csv file
    Set up directories
    Open the robot order website  
    #${orders_file_link}=     Set variable           https://robotsparebinindustries.com/orders.csv
    ${orders_file_link}=        Get Secret            orders_file_url 
    @{orders}=      Get orders          ${orders_file_link}[link]
     
    FOR     ${row}     IN   @{orders}    
        Close the annoying model 
        Fill the form       ${row}[Head]    ${row}[Body]    ${row}[Legs]    ${row}[Address]
        Preview the robot
        Sleep       2
        Submit the order
        Sleep       2
        ${pdf}=           Store the receipt as a PDF file         ${row}[Order number]
        Sleep       2
        ${screenshot}=    Take a screenshot of the robot    ${row}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
        Sleep       3
        Go to order another robot
    END
    Create a ZIP file of the receipts
    #Cleanup temporary PDF directory
