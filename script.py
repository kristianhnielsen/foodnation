from selenium import webdriver
from selenium.common import exceptions as seleniumException
from selenium.webdriver.common.by import By
import openpyxl
from time import sleep


def click_on_pagination_num(number: int):
    pagination_base_element = driver.find_element(By.CLASS_NAME, "pagination")
    pagination_a_tags = pagination_base_element.find_elements(By.TAG_NAME, "a")
    for a_tag in pagination_a_tags:
        try:
            a_tag_num = int(a_tag.text)
        except ValueError:
            continue

        if a_tag_num == number:
            a_tag.click()
            return


def find_case_links():
    case_holder_base = driver.find_element(By.CLASS_NAME, "case-holder-content")
    case_holder_section = case_holder_base.find_element(By.CLASS_NAME, "case-section")
    case_a_tags = case_holder_section.find_elements(By.TAG_NAME, "a")
    for a_tag in case_a_tags:
        case_links.append(a_tag.get_attribute("href"))
    return


def get_case_links():
    base_url = "https://foodnationdenmark.com/case-overview/"
    driver.get(base_url)

    # these page numbers are the numers of the pages to look through. may be changed.
    # for page_num in range(1, 2):
    for page_num in range(1, 16):
        sleep(10)
        find_case_links()
        click_on_pagination_num(page_num)


def save_to_excel():
    wb = openpyxl.load_workbook("foodnation.xlsx")
    sheet = wb.active

    # Headers
    col_names = [
        "title",
        "company",
        "website",
        "email",
        "phone",
        "address",
        "categories",
        "link",
        "introduction",
        "quote",
        "article",
    ]

    sheet[f"A1"] = col_names[0].capitalize()
    sheet[f"B1"] = col_names[1].capitalize()
    sheet[f"C1"] = col_names[2].capitalize()
    sheet[f"D1"] = col_names[3].capitalize()
    sheet[f"E1"] = col_names[4].capitalize()
    sheet[f"F1"] = col_names[5].capitalize()
    sheet[f"G1"] = col_names[6].capitalize()
    sheet[f"H1"] = col_names[7].capitalize()
    sheet[f"I1"] = col_names[8].capitalize()
    sheet[f"J1"] = col_names[9].capitalize()
    sheet[f"K1"] = col_names[10].capitalize()

    data_start_row = 2
    for index, case in enumerate(cases):
        sheet[f"A{index+data_start_row}"] = case["title"]
        sheet[f"B{index+data_start_row}"] = case["company"]

        if len(case["website"]) > 1:
            sheet[f"C{index+data_start_row}"].value = case["website"]
            sheet[f"C{index+data_start_row}"].hyperlink = case["website"]
            sheet[f"C{index+data_start_row}"].style = "Hyperlink"

        if len(case["email"]) > 1:
            sheet[f"D{index+data_start_row}"].value = case["email"]
            sheet[f"D{index+data_start_row}"].hyperlink = f"mailto:{case['email']}"
            sheet[f"D{index+data_start_row}"].style = "Hyperlink"

        sheet[f"E{index+data_start_row}"] = case["phone"]
        sheet[f"F{index+data_start_row}"] = case["address"]
        sheet[f"G{index+data_start_row}"] = case["categories"]

        sheet[f"H{index+data_start_row}"].value = case["link"]
        sheet[f"H{index+data_start_row}"].hyperlink = case["link"]
        sheet[f"H{index+data_start_row}"].style = "Hyperlink"

        sheet[f"I{index+data_start_row}"] = case["introduction"]
        sheet[f"J{index+data_start_row}"] = case["quote"]
        sheet[f"K{index+data_start_row}"] = case["article"]

    wb.save(filename="foodnation_data.xlsx")


def get_article_title():
    h1 = driver.find_element(By.TAG_NAME, "h1")
    title = h1.text

    return title


def get_company_name():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    company_name = aside_elem.find_element(By.CLASS_NAME, "sidebar-header").text
    return company_name


def get_article_website():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    a_tags = aside_elem.find_elements(By.TAG_NAME, "a")
    for a_tag in a_tags:
        href = a_tag.get_attribute("href")
        if "google" in href:
            continue
        if href.startswith("http"):
            return href


def get_article_email():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    a_tags = aside_elem.find_elements(By.TAG_NAME, "a")
    for a_tag in a_tags:
        href = a_tag.get_attribute("href")
        if href.startswith("mailto"):
            email = href.split("mailto:")[1]
            return email


def get_article_phone():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    a_tags = aside_elem.find_elements(By.TAG_NAME, "a")
    for a_tag in a_tags:
        href = a_tag.get_attribute("href")
        if href.startswith("tel"):
            phone_num = href.split("tel:")[1]

            return phone_num


def get_article_categories():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    categories_container = aside_elem.find_element(
        By.CLASS_NAME, "stronghold-container"
    )
    category_elements = categories_container.find_elements(By.TAG_NAME, "a")
    active_categories = []
    for category_element in category_elements:
        category_name = category_element.text

        # find out if category is active/inactive
        icon_constainer = category_element.find_element(By.CLASS_NAME, "icon-container")
        icon_container_classes = icon_constainer.get_attribute("class")
        if "inactive" not in icon_container_classes:
            active_categories.append(category_name)

    categories_str = ", ".join(active_categories)
    return categories_str


def get_article_intro_text():
    main_article = driver.find_element(By.CLASS_NAME, "main-article")
    all_elements = main_article.find_elements(By.XPATH, "*")

    for element in all_elements:
        if element.tag_name == "p":
            return element.text


def get_article_main_text():
    main_article = driver.find_element(By.CLASS_NAME, "main-article")
    all_elements = main_article.find_elements(By.XPATH, "*")

    article_text = ""
    for index, element in enumerate(all_elements):
        if element.tag_name == "h1":
            continue
        elif element.tag_name == "div":
            continue
        elif element.tag_name == "blockquote":
            continue

        if element.tag_name == "p" and index == 1:
            continue

        article_text += f"{element.text}\n"

    return article_text


def get_article_address():
    aside_elem = driver.find_element(By.CLASS_NAME, "cases-sidebar")
    address = aside_elem.find_element(By.TAG_NAME, "address")
    return address.text


def get_article_quote():
    main_article = driver.find_element(By.CLASS_NAME, "main-article")
    try:
        blockquote = main_article.find_element(By.TAG_NAME, "blockquote").text
        return blockquote
    except seleniumException.NoSuchElementException:
        return ""


# Global variables and configuration
case_links = []
cases = []
driver = webdriver.Firefox()


def main():
    print("Collecting website links")
    get_case_links()
    print("Successfully collected case links")
    print(f"Found {len(case_links)} cases")

    for index, case_link in enumerate(case_links):
        # Progress visualization
        percentage = index / len(case_links) * 100
        formatted_percentage = "{:.2f}".format(percentage)
        print(f"{formatted_percentage}% Complete")

        # Website data configuration
        case_article = {}
        driver.get(case_link)
        sleep(1)

        # Save the website data in variables
        case_article["title"] = get_article_title()
        case_article["website"] = get_article_website()
        case_article["company"] = get_company_name()
        case_article["email"] = get_article_email()
        case_article["phone"] = get_article_phone()
        case_article["address"] = get_article_address()
        case_article["categories"] = get_article_categories()
        case_article["link"] = case_link
        case_article["introduction"] = get_article_intro_text()
        case_article["article"] = get_article_main_text()
        case_article["quote"] = get_article_quote()

        # Bundle the website data variables together with the data from the previous website data
        cases.append(case_article)

    driver.quit()
    save_to_excel()
    print("Done!")


main()
