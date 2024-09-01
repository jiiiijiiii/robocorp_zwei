from robocorp.tasks import task

from robocorp import browser
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from pathlib import Path
from RPA.PDF import PDF
from RPA.Archive import Archive



@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    download_excel_file()
    close_annoying_modal()
    orders = get_orders()

    for order in orders:
        print(f"Processing order: {order}") 
        order_number = order["Order number"]
        fill_the_form(order)
        pdf_file = store_receipt_as_pdf(order_number)
        screenshot = screenshot_robot(order_number)
        embed_screenshot_to_receipt(screenshot, pdf_file)

    archive_receipts()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")        
    
def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    tables = Tables()
    file_path = "/Users/jihyelee/Desktop/Robocorp/Neuer Ordner/orders.csv"

    orders = tables.read_table_from_csv(file_path, header=True)

    return orders

def fill_the_form(order):
    page = browser.page()
    page.reload()
    page.wait_for_load_state("networkidle")
    close_annoying_modal()

    page.select_option("#head", order["Head"])
    page.click(f'input[name="body"][value="{order["Body"]}"]')
    page.fill("[type='number']", order["Legs"])
    page.fill("#address", str(order["Address"]))

    page.click("#preview")
    page.click("#order")

    while page.is_visible(".alert-danger"):
        page.click("#order")    

def close_annoying_modal():
    page = browser.page()

    page.click(".btn.btn-dark")

def store_receipt_as_pdf(order_number):
    page = browser.page()

    receipts_html = page.locator('#receipt').inner_html()
    pdf_path = Path("output") / f"receipt_{order_number}.pdf"

    pdf = PDF()
    pdf.html_to_pdf(receipts_html, str(pdf_path))

    return pdf_path

def screenshot_robot(order_number):
    page = browser.page()
    screenshot_path = Path("output") / f"robot_{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=str(screenshot_path))
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[str(screenshot)],
        target_document=str(pdf_file),
        append=True
    )

def archive_receipts():
    archive = Archive()
    output_dir = Path("output")
    
    zip_path = output_dir / "receipts_archive.zip"
    
    output_dir_str = str(output_dir)
    zip_path_str = str(zip_path)

    archive.archive_folder_with_zip(
        folder=output_dir_str, 
        archive_name=zip_path_str
    )

    print(f"Created ZIP archive at: {zip_path}")