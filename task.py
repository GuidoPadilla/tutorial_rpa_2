# Guido Padilla python tutorial 2 for rpa robot.

import os

from Browser import Browser
from Browser.utils.data_types import SelectAttribute, ElementState
from RPA.Excel.Files import Files 
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Tables import Tables
from RPA.Robocorp.Vault import Vault

_secret = Vault().get_secret("credentials")

URL = _secret["url"]


browser = Browser()
pdf = PDF()
zip = Archive()


def open_the_intranet_website():
    browser.new_page(URL)

def download_the_excel_file():
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv",
        overwrite=True,
        target_file="orders.csv")


def fill_and_submit_the_form_for_one_order(order):
    browser.click("text=OK")
    browser.select_options_by(
        "css=#head",
        SelectAttribute["value"],
        str(order["Head"]))
    selector = f"css=#id-body-{order['Body']}"
    browser.click(selector)
    browser.type_text("css=.form-group:nth-child(3) > input", order["Legs"])
    browser.type_text("css=#address", order["Address"])
    browser.click("css=#preview")
    export_the_order_as_a_pdf(order["Order number"])
    browser.click("css=#order-another")


def fill_the_form_using_the_data_from_the_excel_file():
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv", True)
    for order in orders:
        fill_and_submit_the_form_for_one_order(order)

def export_the_order_as_a_pdf(number):
    while True:
        try:
            browser.click("css=#order")
            order_html = browser.get_property(
            selector="css=#receipt", property="outerHTML")
            break
        except AssertionError:
            print("The order failed to submit, trying again...")
    pdf_path = f"output/receipts/{number}.pdf"
    image_path = f"{os.getcwd()}/output/images/{number}.png"
    pdf.html_to_pdf(order_html, pdf_path)
    browser.take_screenshot(
        filename=image_path,
        selector="css=#robot-preview-image")
    pdf.add_files_to_pdf(
        files=[pdf_path, image_path+":align=center"],
        target_document=pdf_path
    )
    
def zip_receipts_folder():
    zip.archive_folder_with_zip(folder="output/receipts", archive_name="output/results.zip")



def main():
    try:
        open_the_intranet_website()
        download_the_excel_file()
        fill_the_form_using_the_data_from_the_excel_file()
        zip_receipts_folder()
    finally:
        browser.playwright.close()


if __name__ == "__main__":
    main()