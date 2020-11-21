import os
import win32com.client as win32
from PyPDF2 import PdfFileReader, PdfFileWriter


def split_pdf(dest_path, file_update, start_page, end_page):
    """Split the PDF file"""

    pdf = PdfFileReader(file_update)

    if start_page:
        start_page = start_page
    else:
        start_page = 0

    if end_page:
        end_page = int(end_page) + 1
    else:
        end_page = pdf.getNumPages()

    clean_result_folder(dest_path)
    for page in range(int(start_page), int(end_page)):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page - 1))
        output_file = "{}\\{}.pdf".format(dest_path, page)

        with open(output_file, 'wb') as dest_f:
            pdf_writer.write(dest_f)

        send_mail(output_file, page)


def create_path():
    """Create the source path and Destination path"""

    sou_path = "C:\\temp\\pdf_split"
    dest_path = "C:\\temp\\pdf_split\\result"

    if "pdf_split" in os.listdir("C:\\temp"):
        info_1 = "Please copy the PDF file in to {}, the input the page number which you want to split.".format(sou_path)
        print(info_1)
    else:
        info_2 = "Folder create successfully in {}, please copy the pdf in to {}".format(
            sou_path, sou_path)
        print(info_2)
        try:
            os.makedirs(sou_path)
            os.makedirs(dest_path)
        except FileExistsError:
            pass

    return sou_path, dest_path


def check_file(path):
    """Check if the pdf file in the source folder"""

    if len(path) > 1:
        return
    else:
        print("No pdf file in {}, please check!".format(path))
        exit()


def clean_result_folder(path):
    """Clean the pdf result folder before split"""

    result_path = os.listdir(path)
    if result_path:
        for file in result_path:
            os.remove("{}\\{}".format(path, file))


def send_mail(path, i):
    """pop-up the outlook new mail window"""

    outlook = win32.Dispatch("outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Attachments.Add(path)
    mail_name = os.environ.get("username")
    mail.To = "{}@kuehne-nagel.com".format(mail_name)
    mail.subject = "split pdf{}".format(i)
    mail.Send()


def main():
    try:
        sou_path, dest_path = create_path()

        while True:
            check_file(sou_path)
            file_path = os.listdir(sou_path)
            start_page = input('Please input start page: ')
            if start_page == "exit":
                exit()
            end_page = input('Please input end page: ')
            for file in file_path:
                if file != "result":
                    file_update = "{}\\{}".format(sou_path, file)
                    split_pdf(dest_path, file_update, start_page, end_page)
                    print("Done, please check the floder: {}".format(dest_path))

    except Exception as error:
        print("{}, please contact the FUO IT".format(error))


if __name__ == '__main__':
    main()
