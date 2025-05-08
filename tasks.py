from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

MAX_SUBMIT_RETRIES = 10
pdf = PDF()

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
    orders = get_orders()
    fill_form_orders(orders)
    archive_receipts()


def open_robot_order_website():
    """Navigates to the robot order URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Download the orders file, read it as a table, and return the result"""
    download_orders_file()
    return read_csv_file()

def download_orders_file():
    """Downloads the orders csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_csv_file():
    """Read the orders file and return it as a table"""
    library = Tables()
    return library.read_table_from_csv(
        "orders.csv",
        columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    
def fill_form_orders(orders):
    """Fills form orders, submit them and process the receipt"""
    for order in orders:
        try:
            if page := fill_and_submit_order_form(order):
                process_order_checkout(order["Order number"], page)
            else:
                print(f"Unable to process order {order["Order number"]}")
        except Exception as e:
            print(f"Something went wrong while processing the order {order["Order number"]}: {e}")

def fill_and_submit_order_form(order):
    """Fills in the robot order, click the 'Submit' button"""
    page = browser.page()
    page.click("button:text('OK')")
    page.select_option("#head", order["Head"])
    page.check(f'input[name="body"][value="{str(order["Body"])}"]')
    page.fill('input[placeholder="Enter the part number for the legs"]', str(order["Legs"]))
    page.fill("#address", order["Address"])
    page.click("#preview")
    success = submit_order_form(page, order["Order number"])
    
    return page if success else None

def process_order_checkout(order_number, page):
    """Generates the receipt pdf including a robot screenshot"""
    pdf_path = store_receipt_as_pdf(order_number, page)
    image_path = screenshot_robot(order_number, page)
    embed_screenshot_to_receipt(image_path, pdf_path)
    # Go to process the next order
    page.click("#order-another")

def submit_order_form(page, order_number):
    """Submits a robot order form with a retry strategy"""
    successful_order = False
    order_tries = 0
    while not successful_order and order_tries < MAX_SUBMIT_RETRIES:
        page.click("#order")
        successful_order = page.query_selector("#order-another")
        order_tries += 1
        
        if order_tries > 1:
            print(f"Retrying order {order_number} (try {order_tries} of {MAX_SUBMIT_RETRIES})")

    return successful_order

def store_receipt_as_pdf(order_number, page):
    """Export the receipt data to a pdf file"""
    receipt_html = page.locator("#receipt").inner_html()
    file_path = f"output/receipts/order_{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, file_path)
    return file_path

def screenshot_robot(order_number, page):
    """Takes a screenshot from the robot"""
    file_path = f"output/robot_images/order_{order_number}_picture.png"
    page.locator("#robot-preview-image").screenshot(path=file_path)
    return file_path

def embed_screenshot_to_receipt(screenshot_path, pdf_file_path):
    """Embeds the screenshot from the robot into the reciept file"""
    # Using add_watermark_image_to_pdf instead of add_files_to_pdf
    # since adding images directly to PDF files is not working currently (May 8th 2025)
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot_path,
        source_path=pdf_file_path,
        output_path=pdf_file_path
    )

def archive_receipts():
    """Creates a ZIP file of receipt PDF files"""
    archiver = Archive()
    archiver.archive_folder_with_zip("./output/receipts", "output/receipts.zip")