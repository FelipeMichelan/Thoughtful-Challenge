import os
import re
import requests
from robocorp.tasks import task
from robocorp.tasks import get_output_dir
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from selenium.webdriver.common.keys import Keys
from datetime import datetime

@task
def minimal_task():
    # Initializing variables
    totalPhraseCount = 0
    containsNumber = False
    containsMoney = False
    imageCount = 1

    browser = Selenium()

    # [START INITIALIZE]
    # Open browser
    browser.open_available_browser("https://www.latimes.com")
    browser.maximize_browser_window()
    
    # Wait for the main page to load
    browser.wait_until_element_is_visible("xpath://button[@data-element='search-button']")

    # Click Search button (magnifying glass)
    browser.click_button_when_visible("xpath://button[@data-element='search-button']")
    
    # Input searched text and Submit the search
    browser.input_text_when_element_is_visible("xpath://input[@data-element='search-form-input']", "sports" + Keys.ENTER)

    # Wait for the searched page to load
    browser.wait_until_element_is_visible("xpath://ul[@class='search-results-module-results-menu']")
    
    #[END INITIALIZE]

    #[START RETRIEVE INFO]
    # HTML containing the results
    result_elementsPath = f"//main[@class='search-results-module-main']/ul/li"

    # Retrieve all elements
    results_returnedElements = browser.find_elements(result_elementsPath)

    # Create a Dictionary with the results
    results = {'Title':[], 'Description':[],'Date':[], 'PictureName':[], 'PhraseCount': [], 'ContainsMoney': [], 'PictureLink': [], "PictureFilePath":[]}

    # For each element found
    for index, element in enumerate(results_returnedElements, 1):
        
        # Resetting reusable variables
        get_TitlePhraseCount = 0
        get_DescriptionPhraseCount = 0
        totalPhraseCount = 0
        
        # Get Title
        results['Title'].append(browser.get_text(result_elementsPath+f"[{index}]//ps-promo/div/div/div/h3"))
        # Get Description
        results['Description'].append(browser.get_text(result_elementsPath+f"[{index}]//ps-promo/div/div/p[@class='promo-description']"))
        # Get Date
        results['Date'].append(browser.get_text(result_elementsPath+f"[{index}]//ps-promo/div/div/p[@class='promo-timestamp']"))
        # Get PictureName
        results['PictureName'].append(browser.get_element_attribute(result_elementsPath+f"[{index}]//ps-promo/div/div/a[@class='link promo-placeholder']","aria-label"))      

        # Retrieve Phrase Count:
        # [START]

        # Title
        get_TitlePhraseCount = results['Title'][index-1].upper().count("SPORTS")
        # Description
        get_DescriptionPhraseCount = results['Description'][index-1].upper().count("SPORTS")
        # Total Phrase Count
        totalPhraseCount = get_TitlePhraseCount + get_DescriptionPhraseCount
        
        results["PhraseCount"].append(totalPhraseCount)
        
        # [END]

        # Checks if contains money:
        # [START]
        
        # Defining Title for easier comprehension
        resultStr_checkIfMoney = results['Title'][index-1]

        # Title:
        # If contains any number, it will return true
        containsNumber = bool(re.search('\\d', resultStr_checkIfMoney))
        if containsNumber is True:
            # Checks if it includes any defined pattern
            if "USD" in resultStr_checkIfMoney or "$" in resultStr_checkIfMoney or "dollar" in resultStr_checkIfMoney:
                containsMoney = True

        # Defining Description for easier comprehension
        resultStr_checkIfMoney = results['Description'][index-1]

        # Description:
        containsNumber= bool(re.search('\\d', resultStr_checkIfMoney))
        if containsNumber is True:
            if "USD" in resultStr_checkIfMoney or "$" in resultStr_checkIfMoney or "dollar" in resultStr_checkIfMoney:
                containsMoney = True

        # Insert the result into the dictionary
        if containsMoney is True:
            results["ContainsMoney"].append("True")
        else:
            results["ContainsMoney"].append("False")

        #[END]
        
        # Downloading image:
        # [START]
        
        # Defining today
        today = datetime.now()
        if today.hour < 12:
            h = "00"
        else:
            h = "12"

        imageFolder = "output/pictures " + today.strftime('%m-%d-%Y') + "/"

        # Creating folder, if doesn't exists
        if not os.path.exists(imageFolder):
            os.makedirs(imageFolder)

        # Get image link (src)
        results['PictureLink'].append(browser.get_element_attribute(result_elementsPath+f"[{index}]//ps-promo/div/div/a/picture/img[@class='image']","src"))
        # Insert into dictionary
        pictureLink = results['PictureLink'][index-1]
        
        # Download image
        img_data = requests.get(pictureLink).content
        with open(imageFolder+str(imageCount)+'.jpg', 'wb') as handler:
            handler.write(img_data)
        
        # Insert into dictionary
        results["PictureFilePath"].append(imageFolder+str(imageCount)+'.jpg') 

        # Iterating index image name
        imageCount = imageCount + 1

        
        # [END]
    # [END RETRIEVE INFO]

    # [Write to Excel]
    # [START]

    # Create a new Excel workbook and add the search results
    excel = Files()
    workbook = excel.create_workbook()
    workbook.create_worksheet('Results')
    excel.append_rows_to_worksheet(results,header=True,name='Results')
    workbook.save(os.path.join(get_output_dir(),'Final result.xlsx'))
    
    # [END]