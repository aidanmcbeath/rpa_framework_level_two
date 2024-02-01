#Initialisation
from robocorp.tasks import task 
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
import shutil
import time

csv_url = "https://robotsparebinindustries.com/orders.csv"
workbook_path = "C:\\Users\\bot1\\Documents\\Robots\\level_two\\orders.csv"
output_path = "C:\\Users\\bot1\\Documents\\Robots\\level_two\\output\\"
csv_columns = ["Order number", "Head", "Body", "Legs", "Address"]

@task
def order_processing():
    """Processes orders on the RSB website."""
    open_order_website()
    close_annoying_modal()
    download_orders_file(csv_url, workbook_path)
    robot_orders = csv_to_table(csv_columns)
    for order in robot_orders:
        preview_path = input_order(order)
        save_receipt_to_pdf(preview_path)
        press_new_order()
    create_zip()

def open_order_website():
    """Creates a browser session and navigates to the RSB webpage"""
    browser.configure(slowmo=500)
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_file(csv_url, workbook_path):
    """Downloads orders csv file"""
    http = HTTP()
    http.download(
        url=csv_url, 
        overwrite=True, 
        target_file=workbook_path
        )
    
def csv_to_table(csv_columns):
    """Assigns orders csv to a variable as a table"""
    Orders = Tables()
    orders_table = Orders.read_table_from_csv(
        path=workbook_path,
        header=True,
        columns = csv_columns
        )
    return orders_table

def close_annoying_modal():
    """Closes a popup on the website"""
    page = browser.page()
    page.click("text=OK")

def input_order(order):
    """Inputs and submits orders into the website, takes a screenshot and returns path to the screenshot"""
    #initialising variables
    retry_count = 0
    retry_limit = 3
    #Filling form
    page = browser.page()
    page.select_option('#head', str(order["Head"]))
    page.click("#id-body-" + str(order["Body"]))
    page.fill(
        "xpath=//html/body/div/div/div[1]/div/div[1]/form/div[3]/input",
        str(order["Legs"])
        )
    page.fill("#address", str(order["Address"]))
    #clicks 'preview' and screenshots the robot preview image
    page.click("#preview")
    preview_path = screenshot_preview(str(order["Order number"]))
    visible = False
    #This loop clicks the order button and checks if the next page appears, tries again if req'd
    while visible == False and retry_count < retry_limit:
        page.click("#order")
        time.sleep(1)
        visible = page.is_visible("#receipt")
        retry_count =+ 1
    return preview_path

def screenshot_preview(order):
    """Screenshots the preview and returns output path"""
    page = browser.page()
    preview_path = output_path + "screenshot" + order + ".jpeg"
    page.locator("#robot-preview-image").screenshot(
        type = 'jpeg',
        path = preview_path
        )
    return preview_path

def save_receipt_to_pdf(preview_path):
    """Adds receipt and preview to a pdf"""
    page = browser.page()
    receipt_number = page.text_content("xpath=//html/body/div/div/div[1]/div/div[1]/div/div/p[1]")
    receipt_path = output_path + receipt_number + ".jpeg"
    pdf_path = output_path + receipt_number + ".pdf"
    page.locator("#receipt").screenshot(
        type = 'jpeg',
        path = receipt_path
        )
    pdf = PDF()
    pdf.add_files_to_pdf(files = [preview_path, receipt_path], target_document = pdf_path)
    #removes screenshots
    FileSystem.remove_file(self = FileSystem ,path = preview_path)
    FileSystem.remove_file(self = FileSystem, path = receipt_path)

def press_new_order():
    """Opens a new order page and closes the popup"""
    page = browser.page()
    page.click("#order-another")
    close_annoying_modal()

def create_zip():
    """Packages pdf files into a zip and removes original directory"""
    lib = FileSystem()
    matches = lib.find_files("C:\\Users\\bot1\\Documents\\Robots\\level_two\\output\\*.pdf")
    
    if lib.does_directory_exist(path = output_path + "Archive") == False:
        lib.create_directory(path = output_path + "Archive")
    
    lib.move_files(sources = matches, destination = output_path + "Archive", overwrite = True)
    shutil.make_archive(output_path + "Archived", 'zip', output_path + "Archive")

    if lib.does_directory_exist(path = output_path + "Archive") == True:
        lib.remove_directory(path = output_path + "Archive", recursive =True)