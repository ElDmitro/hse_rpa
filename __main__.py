### https://github.com/MazarsLabs/hse-rpa
import os
import time

import pandas as pd
import selenium.webdriver as webdriver
import smtplib
from email.message import EmailMessage
from selenium.webdriver.chrome.options import Options

from conf import QUERY, NPAGE, RECEIVER, EMAIL_SENDER, PASSWORD_SENDER, SLEEP


QUERY_LINK_FORMAT = "https://www.semanticscholar.org/search?q={}&sort=relevance&page={}"

EMAIL_SUBJECT = "Topics analysis"
EMAIL_CONTENT = "Hi!\n\nFind attached excel file with articles info.\n\nRegard,"

WORKING_DIR = os.path.dirname(os.path.realpath(__file__))
PDF_FOLDER = os.path.join(WORKING_DIR, "articles")
WEBDRIVER_PATH = os.path.join(WORKING_DIR, "chromedriver")

_chrome_options = Options()
_prefs = {"download.default_directory": PDF_FOLDER, "download.prompt_for_download": False}
_chrome_options.add_experimental_option('prefs', _prefs)
os.environ["webdriver.chrome.driver"] = WEBDRIVER_PATH


def fetch_article(driver, link, date):
    driver.get(link)
    time.sleep(SLEEP)

    title = driver.find_elements_by_xpath("//*[@data-selenium-selector='paper-detail-title']")[0].text
    author = driver.find_elements_by_class_name('paper-meta-item')[0].text

    try:
        loaded_pdfs_filenames = os.listdir(PDF_FOLDER)
        driver.find_element_by_xpath("//*[@class='alternate-sources__dropdown-wrapper']").click()
        time.sleep(SLEEP)

        new_pdfs_filenames = os.listdir(PDF_FOLDER)
        filename = list(set(current_dir) - set(initial_dir))[0]
    except Exception as exc:
        filename = ''

    return {
        'title': title,
        'date': date,
        'authors': author,
        'path_to_file': os.path.join(PDF_FOLDER, filename),
    }

def dump_info_into_xlsx(info, filename='data'):
    df = pd.DataFrame(info)
    filepath = f"{filename}.xlsx"
    df.to_excel(os.path.join(WORKING_DIR, filepath), index=False)
    return filepath

def create_email(attachment_path, filename):
    mail = EmailMessage()
    mail['From'], mail['To'], mail['Subject'] = EMAIL_SENDER, RECEIVER, EMAIL_SUBJECT
    mail.set_content(EMAIL_CONTENT)

    with open(attachment_path, 'rb') as input_stream:
        file_data = input_stream.read()
    mail.add_attachment(
        file_data,
        maintype='application', subtype='octet-stream',
        filename=f'{filename}.xlsx'
    )
    return mail

def send_email(email):
    server = smtplib.SMTP('smtp.office365.com')
    server.starttls()
    server.login(EMAIL_SENDER, PASSWORD_SENDER)
    server.send_message(email)
    server.quit()

def main():
    if not os.path.isdir(PDF_FOLDER):
        os.mkdir(PDF_FOLDER)

    print(f'QUERY: {QUERY}, Page Number: {NPAGE}')
    articles_info = list()
    driver = webdriver.Chrome(executable_path=WEBDRIVER_PATH, options=_chrome_options)
    for page in range(1, NPAGE + 1):
        query_link = QUERY_LINK_FORMAT.format(QUERY, page)
        driver.get(query_link)
        time.sleep(SLEEP)

        page_dates, articles_links = list(), list()
        articles_element = driver.find_elements_by_xpath("//*[@data-selenium-selector='title-link']")
        dates_element = driver.find_elements_by_class_name('cl-paper-pubdates')
        for data in dates_element:
            page_dates.append(data.text)
        for article in articles_element:
            try:
                articles_links.append(article.get_attribute("href"))
            except Exception as exc:
                pass
        for i, link in enumerate(articles_links):
            articles_info.append(fetch_article(driver, link, page_dates[i]))
    driver.quit()

    info_filepath = dump_info_into_xlsx(articles_info)
    email = create_email(info_filepath, 'articles_info')
    send_email(email)


if __name__ == '__main__':
    main()
