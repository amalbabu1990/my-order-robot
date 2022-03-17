*** Settings ***
Documentation     Orders robots from RobotSpareBin Industries Inc.
...               Saves the order HTML receipt as a PDF file.
...               Saves the screenshot of the ordered robot.
...               Embeds the screenshot of the robot to the PDF receipt.
...               Creates ZIP archive of the receipts and the images.
Library           RPA.Browser.Selenium    auto_close=${FALSE}
Library           RPA.HTTP
Library           RPA.Excel.Files
Library           RPA.PDF
Library           RPA.Tables
Library           RPA.Desktop
Library           RPA.FileSystem
Library           RPA.Archive
Library           RPA.Dialogs
Library           RPA.Robocorp.Vault

*** Variables ***
${Index}          ${0}
${DataFolder}
${str_RootFolder}
${str_OutputFolderName}
${str_PDFSaveFolderName}
${str_ReceiptsSaveFolderName}

*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    # Config Variables
    # Path can be mentioned as follows : E:${/}SomeDrive{/}Documents${/}Robocorp${/}_DATA
    ${str_RootFolder}    Set Variable    ${OUTPUT_DIR}
    ${secret_FolderNames}=    Get Secret    FolderNames
    #str_OutputFolderName = Output
    #str_PDFSaveFolderName = PDFs
    #str_ReceiptsSaveFolderName = Receipts
    #str_ScreenshotFolderName = Screenshots
    ${str_OutputFolderName}    Set Variable    ${secret_FolderNames}[OutputFolderName]
    ${str_PDFSaveFolderName}    Set Variable    ${secret_FolderNames}[PDFSaveFolderName]
    ${str_ReceiptsSaveFolderName}    Set Variable    ${secret_FolderNames}[ReceiptsSaveFolderName]
    ${str_ScreenshotFolderName}    Set Variable    ${secret_FolderNames}[ScreenshotFolderName]
    #Construct Config paths
    ${str_ReceiptsSaveFolderPath}    Set Variable    ${str_RootFolder}${/}${str_OutputFolderName}${/}${str_ReceiptsSaveFolderName}
    ${str_ScreenshotSaveFolderPath}    Set Variable    ${str_RootFolder}${/}${str_OutputFolderName}${/}${str_ScreenshotFolderName}
    ${str_PdfSaveFolderPath}    Set Variable    ${str_RootFolder}${/}${str_OutputFolderName}${/}${str_PDFSaveFolderName}
    #_LOG_
    Log To Console    The Receipts Folder Location : ${str_ReceiptsSaveFolderPath}
    Log To Console    The Screenshot Folder Location : ${str_ScreenshotSaveFolderPath}
    Log To Console    The PDF Folder Location : ${str_PdfSaveFolderPath}
    Log To Console    The output ZIP File Path : ${str_RootFolder}${/}OutputPdfs.zip
    #_LOG_
    Download the Excel file    ${str_RootFolder}
    Launch the webportal
    Manage pop-up message
    Fill the form using the data from the Excel file
    ...    ${str_RootFolder}
    ...    ${str_ReceiptsSaveFolderPath}
    ...    ${str_ScreenshotSaveFolderPath}
    ...    ${str_PdfSaveFolderPath}
    Create an archive for the PDFs    ${str_PdfSaveFolderPath}    ${str_RootFolder}${/}OutputPdfs.zip
    [Teardown]    Close Browser

*** Keywords ***
Fill the form using the data from the Excel file
    [Arguments]
    ...    ${RootFolderPath}
    ...    ${str_ReceiptsSaveFolderPath}
    ...    ${str_ScreenshotSaveFolderPath}
    ...    ${str_PdfSaveFolderPath}
    # Body
    # Create Folder Paths If not already available
    Create Directory    ${str_ReceiptsSaveFolderName}    True    True
    Create Directory    ${str_ScreenshotSaveFolderPath}    True    True
    Create Directory    ${str_PdfSaveFolderPath}    True    True
    ${Orders}=    Read table from CSV    ${RootFolderPath}${/}orders.csv
    FOR    ${Order}    IN    @{Orders}
        #_LOG_
        ${Index}=    Evaluate    ${Index}+1
        Log To Console    Processing Row: ${Index}
        Run Keyword And Continue On Failure
        ...    Data entry for One Order
        ...    ${Order}    ${str_ReceiptsSaveFolderPath}    ${str_ScreenshotSaveFolderPath}
        Run Keyword And Continue On Failure
        ...    Save Receipts to PDF
        ...    ${str_ReceiptsSaveFolderPath}${/}${Order}[Order number].pdf
        Run Keyword And Continue On Failure
        ...    Take Screenshots
        ...    ${str_ScreenshotSaveFolderPath}${/}${Order}[Order number].png
        Run Keyword And Continue On Failure
        ...    Combine Receipt and Screenshot to create Pdf
        ...    ${str_ReceiptsSaveFolderPath}${/}${Order}[Order number].pdf
        ...    ${str_ScreenshotSaveFolderPath}${/}${Order}[Order number].png
        ...    ${str_PdfSaveFolderPath}${/}${Order}[Order number].pdf
        # Run this keyword regardless of any possible
        Order another robot
    END
    #_LOG_
    Log To Console    Processing of ${Index} rows completed successfully

Launch the webportal
    ${UserInput}=    Read CSV file URL from User
    Open Available Browser
    ...    ${UserInput.url}
    Maximize Browser Window

Manage pop-up message
    Wait Until Element Is Visible    css:button.btn.btn-dark    15
    Click Button    css:button.btn.btn-dark

Read CSV file URL from User
    Add heading    Provide Input CSV URL containing Order details
    Add text input    url    label=CSV URL string
    ${result}=    Run dialog
    [Return]    ${result}

Combine Receipt and Screenshot to create Pdf
    [Arguments]    ${ReceiptsFilePath}    ${ScreenshotsFilePath}    ${DestinationFilePath}
    Add Watermark Image To PDF
    ...    image_path=${ScreenshotsFilePath}
    ...    source_path=${ReceiptsFilePath}
    ...    output_path=${DestinationFilePath}

Create an archive for the PDFs
    [Arguments]    ${FolderPath}    ${OutputFilePath}
    Archive Folder With Zip
    ...    folder=${FolderPath}
    ...    archive_name=${OutputFilePath}

Data entry for One Order
    [Arguments]
    ...    ${Order}
    ...    ${str_ReceiptsSaveFolderPath}
    ...    ${str_ScreenshotSaveFolderPath}
    # Boby
    Wait Until Element Is Visible    head    15
    Select From List By Value    head    ${Order}[Head]
    Click Button    id:id-body-${Order}[Body]
    Input Text
    ...    css:input[placeholder='Enter the part number for the legs']
    ...    ${Order}[Legs]
    Input Text    address    ${Order}[Address]
    # Use the Wait Until Keyword Succeeds 10 times and at one second intervals
    Wait Until Keyword Succeeds
    ...    10x
    ...    0.5 sec
    ...    Preview and Order

Download the Excel file
    [Arguments]    ${RootFolder}
    Download
    ...    https://robotsparebinindustries.com/orders.csv
    ...    target_file=${RootFolder}${/}orders.csv
    ...    overwrite=True

Preview and Order
    Wait Until Element Is Visible    preview
    Click Button    preview
    Sleep    0.5s
    Wait Until Element Is Visible    css:#robot-preview-image
    Click Button    order
    Wait Until Element Is Visible    css:#receipt

Order another robot
    # Dont allow this keyword to error out
    Run Keyword And Continue On Failure
    ...    Wait Until Element Is Visible    css:#order-another    5
    ${visible}=    Is Element Visible    css:#order-another
    # If there is an error the element is not vissible
    IF    ${visible}
        Click Button    css:#order-another
    ELSE
        Reload Page
    END
    Manage pop-up message

Save Receipts to PDF
    [Arguments]    ${ReceiptsDownloadLocation}
    Wait Until Element Is Visible    css:#receipt    5
    ${ReceiptsHTML}=    Get Element Attribute    css:#receipt    outerHTML
    Html To Pdf    ${ReceiptsHTML}    ${ReceiptsDownloadLocation}

Take Screenshots
    [Arguments]    ${ScreenshotDownloadLocation}
    Wait Until Element Is Visible    css:#robot-preview-image    5
    Capture Element Screenshot    css:#robot-preview-image    ${ScreenshotDownloadLocation}
