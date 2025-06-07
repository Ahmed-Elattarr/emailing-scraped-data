from bs4 import BeautifulSoup
from openpyxl import Workbook
import smtplib
import requests
import csv
from openpyxl.utils import get_column_letter
from email.message import EmailMessage
import ssl

def scraping():
    main_link = "https://books.toscrape.com/"
    page_link = "catalogue/page-1.html"
    headers = {"User-Agent":"Mozilla/5.0"}

    with open("scraped.csv","w",newline='', encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["full title","price","book link"])
        
        while True:
            current_link = main_link + page_link
            print(page_link)
            response = requests.get(current_link,headers=headers)
            soup = BeautifulSoup(response.text,"html.parser")

            body = soup.find_all("li",class_ = "col-xs-6 col-sm-4 col-md-3 col-lg-3")
            if not body:
                    print("No products found. Stopping.")
                    break

            for book in body:
                title = book.find("h3")
                full_title = title.find("a")["title"]
                price = book.find("p", class_ = "price_color").get_text()
                book_link = main_link +"catalogue/"+ book.find("a")["href"]
                writer.writerow([full_title,price,book_link])

            next_button = soup.find("li",class_ = "next")
            if next_button:
                next_page_link = next_button.find("a")["href"]
                page_link = "catalogue/"+ next_page_link
            else:
                break
def csv2excel():
     wb = Workbook()
     ws = wb.active
     with open("scraped.csv",newline='',encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        s = list(reader)
        for row in s:
            ws.append(row)
        wb.save("scraped_data.xlsx")
def mailing():
    sender_email = "ahmedelattarpyauto@gmail.com"
    receiver_email = "3attar0@gmail.com"
    password = "hryboskqwulbwxmf"
    subject = "Automated scraped Report"
    body = "Hi,\n\nPlease find attached the latest scraped report.\n\nBest regards,\nAutomation Bot"

    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)

    with open("scraped_data.xlsx", "rb") as f:
        file_data = f.read()
        file_name = "scraped_data.xlsx"
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(sender_email, password)
        smtp.send_message(msg)

    print("Email sent successfully!")



scraping()
csv2excel()
mailing()