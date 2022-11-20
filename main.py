import re
from multiprocessing.dummy import Pool as ThreadPool
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import requests, lxml

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.0.0 "
                  "Safari/537.36"
}

# ---------------- Selenium Code --------------------------------
options = Options()
options.add_argument('--headless')
options.add_argument(f'user-agent={headers["User-Agent"]}')


# ---------------- Selenium Code --------------------------------


def new_url(url):
    html = requests.get(url, headers=headers).text
    return html


def get_reviews(asin_code, driver):
    review_names = []
    review_dates = []
    review_titles = []
    review_contents = []
    count = 1

    for i in range(5):
        driver.get(
            f"https://www.amazon.com/product-reviews/{asin_code}/?ie=UTF8&reviewerType=all_reviews&pageNumber={count}")
        soup = BeautifulSoup(driver.page_source, 'lxml')
        for review_name in soup.select("div.a-profile-content > span.a-profile-name")[2:]:
            if review_name.text.strip() not in review_names:
                review_names.append(review_name.text.strip())

        for review_date in soup.select("span.review-date")[2:]:
            formatted_value = re.sub(".*on", '', review_date.text.strip())
            review_dates.append(formatted_value.strip())

        for review_title in soup.select("a.review-title-content"):
            review_titles.append(review_title.text.strip())

        for review_content in soup.select("span.review-text-content"):
            review_contents.append(review_content.text.replace("\n", "").strip())
        count += 1

    return review_names, review_dates, review_titles, review_contents


final_prods = []


def get_product_info(next_prod):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    prods = []
    print(f"Scraping Product No - {next_prod}")
    try:
        url = f"https://www.amazon.com/dp/{next_prod}?th=1"
        prods.append(next_prod)
        driver.get(url)
        driver.execute_script("window.scrollTo(1080, document.body.scrollHeight)")
        time.sleep(2)
        url_split = driver.current_url.split('/')[-1].split('?')[0]
        if url_split == next_prod:
            soup = BeautifulSoup(driver.page_source, 'lxml')
            title = soup.select_one("#title > span").text.strip()
            prods.append(title)
            prods.append(url)
            if len(soup.select_one("#reviews-medley-footer > div.a-spacing-medium > a")) > 0:
                print("Inside all reviews")
                r1, r2, r3, r4 = get_reviews(next_prod, driver)
                for jj in range(len(r1)):
                    prods.append(r1[jj])
                    prods.append(r2[jj])
                    prods.append(r3[jj])
                    prods.append(r4[jj])
            else:
                print("No reviews")
        else:
            print("Different asin code")
    except:
        pass
    final_prods.append(prods)
    driver.quit()


if __name__ == "__main__":
    print("Reading File.....")
    df = pd.ExcelFile(f'./Asin Codes.xlsx').parse('Sheet1')
    process = input("Enter Amount of Processes You want the script to run in? ")
    pool = ThreadPool(int(process))
    results = pool.map(get_product_info, df["AMAZON ASIN"])
    pool.close()
    pool.join()

    output_file = pd.DataFrame(final_prods)
    writer = pd.ExcelWriter(f'outputFile.xlsx', engine='xlsxwriter')
    output_file.to_excel(writer, sheet_name='products', startrow=1, index=False, header=False)
    workbook = writer.book
    worksheet = writer.sheets['products']

    header_format = workbook.add_format({'bold': True,
                                         'bottom': 2,
                                         'bg_color': '#F9DA04'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    writer.save()
    print("done")
