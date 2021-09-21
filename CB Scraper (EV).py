#imports
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time

#driver path
PATH = "C:/Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

#access crunchbase ui
driver.get("https://www.crunchbase.com/search/organizations/field/organization.companies/categories/electric-vehicle")
driver.maximize_window()
time.sleep(5)
print(driver.title)

#log-in
login = driver.find_element_by_xpath('/html/body/chrome/div/app-header/div[1]/div[2]/div/anon-nav-row/nav-action-item[2]/nav-header-action-item/a/span[1]/nav-action-item-header-layout/div/span')
login.click()

email = driver.find_element_by_name('email')
email.send_keys()

passw = driver.find_element_by_name('password')
passw.send_keys()
passw.send_keys(Keys.RETURN)
    
time.sleep(3)

#navigate crunchbase & add parameters
addcolumn = driver.find_element_by_xpath('//button[@class="mat-focus-indicator add-column-button mat-stroked-button mat-button-base mat-primary"]')
addcolumn.click()

# get items only from first list
all_categories = driver.find_elements_by_xpath('(//mat-nav-list)[1]//mat-list-item')
print('len(all_categories):', len(all_categories))

for category in all_categories:
    print('-----')
    
    # select category
    print('Category:', category.text.strip())
    
    # scroll it to make it visible and clickable
    #driver.execute_script("arguments[0].scrollIntoView(true);", category)
    # or
    ActionChains(driver).move_to_element(category).perform()
    
    # click category to display list of columns in this category
    category.click()
    time.sleep(0.5)

    # search columns ONLY in selected category

    # it selects item only if `mat-checkbox` doesn't have class `mat-checkbox-checked`
    # and it click `label` instead of `checkbox` because `label` is not hidden by `popup message`
    columns = driver.find_elements_by_xpath('(//mat-nav-list)[2]//mat-checkbox[not(contains(@class, "mat-checkbox-checked"))]//label')
    print('len(columns):', len(columns))

    for col in columns:
        print('click:', col.text.strip())
        ActionChains(driver).move_to_element(col).perform()
        
        col.click()
    
    # TODO: click subcategory, select checkboxes, click back button 
    subcategories = driver.find_elements_by_xpath('(//mat-nav-list)[2]//mat-list-item[.//icon[@key="icon_caret_right"]]')
    print('len(subcategories):', len(subcategories))
    for sub in subcategories:
        sub.click()
    
        subcolumns = driver.find_elements_by_xpath('(//mat-nav-list)[3]//mat-checkbox[not(contains(@class, "mat-checkbox-checked"))]//label')
    
        for subc in subcolumns:
            subc.click()

    if subcategories:               
        backbutton = driver.find_element_by_xpath('//*[@id="mat-dialog-0"]//mat-dialog-content//button')   
        backbutton.click() 
        
        
        
driver.find_element_by_xpath('//button[@aria-label="Apply Changes"]').click()
   
#await element location

WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, ('//a[@aria-label="Next"][@aria-disabled="false"][@type="button"]'))))
    
#next page
page = driver.find_element_by_xpath('/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/search/page-layout/div/div/form/div[2]/results/div/div/div[1]/div/results-info/h3/a[2]')

company_list = [] ###create dictionary
   
counter = 0
for _ in range(38):
    if counter == 37:
        break
    counter += 1
    
    if page.is_displayed():
        
        time.sleep(25)
        
        #webscrape through iterations
        all_rows = driver.find_elements_by_css_selector("grid-row")
                                                      
        for row in all_rows:
           cblist = {
                'company': row.find_element_by_xpath('.//*[@class="identifier-label"]').text.strip(),
                'industry': row.find_element_by_xpath('.//*[@data-columnid="categories"]//span').text.strip(),
                'hq': row.find_element_by_xpath('.//*[@data-columnid="location_identifiers"]//span').text.strip(),
                'description': row.find_element_by_xpath('.//*[@data-columnid="short_description"]//span').text.strip(),
                'cb rank': row.find_element_by_xpath('.//*[@data-columnid="rank_org"]').text.strip(),
                'cb rank (company)': row.find_element_by_xpath('.//*[@data-columnid="rank_org_company"]').text.strip(),
                'cb rank (school)': row.find_element_by_xpath('.//*[@data-columnid="rank_org_school"]').text.strip(),
                'contact job department': row.find_element_by_xpath('.//*[@data-columnid="contact_job_departments"]').text.strip(),
                'number of contacts': row.find_element_by_xpath('.//*[@data-columnid="num_contacts"]').text.strip(),
                'diversity spotlight (US only)': row.find_element_by_xpath('.//*[@data-columnid="diversity_spotlights"]').text.strip(),
                'founded date': row.find_element_by_xpath('.//*[@data-columnid="founded_on"]').text.strip(),
                'number of investments': row.find_element_by_xpath('.//*[@data-columnid="num_investments_funding_rounds"]').text.strip(),
                'number of diversity investments': row.find_element_by_xpath('.//*[@data-columnid="num_diversity_spotlight_investments"]').text.strip(),
                'number of exits': row.find_element_by_xpath('.//*[@data-columnid="num_exits"]').text.strip(),
                'industry groups': row.find_element_by_xpath('.//*[@data-columnid="category_groups"]').text.strip(),
                'founders': row.find_element_by_xpath('.//*[@data-columnid="founder_identifiers"]').text.strip(),
                'funding status': row.find_element_by_xpath('.//*[@data-columnid="funding_stage"]').text.strip(),
                'transaction name': row.find_element_by_xpath('.//*[@data-columnid="acquisition_identifier"]').text.strip(),
                'acquired by': row.find_element_by_xpath('.//*[@data-columnid="acquirer_identifier"]').text.strip(),
                'announced by': row.find_element_by_xpath('.//*[@data-columnid="acquisition_announced_on"]').text.strip(),
                'price': row.find_element_by_xpath('.//*[@data-columnid="acquisition_price"]').text.strip(),
                'ipo status': row.find_element_by_xpath('.//*[@data-columnid="ipo_status"]').text.strip(),
                'ipo date': row.find_element_by_xpath('.//*[@data-columnid="went_public_on"]').text.strip(),
                'delisted date': row.find_element_by_xpath('.//*[@data-columnid="delisted_on"]').text.strip(),
                'ipo amount raised': row.find_element_by_xpath('.//*[@data-columnid="ipo_amount_raised"]').text.strip(),
                'ipo valuation': row.find_element_by_xpath('.//*[@data-columnid="ipo_valuation"]').text.strip(),
                'stock exchange': row.find_element_by_xpath('.//*[@data-columnid="stock_exchange_symbol"]').text.strip(),
                'last leadership hiring date': row.find_element_by_xpath('.//*[@data-columnid="last_key_employee_change_date"]').text.strip(),
                'last layoff date': row.find_element_by_xpath('.//*[@data-columnid="last_layoff_date"]').text.strip()
           }
           company_list.append(cblist)
        print("next")
        page.click()
    
    
#create dataframe    
df = pd.DataFrame(company_list)

print(df)

#create excel writer object
writer=pd.ExcelWriter('crunchbase.xlsx')

#export to excel
df.to_excel(writer)

writer.save()



print("It's alive!")