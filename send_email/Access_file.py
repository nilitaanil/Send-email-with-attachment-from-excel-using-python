import pandas as pd
email_list = pd.read_excel("/home/pachi/test_sheet.ods", engine="odf")
all_image = email_list['Image']


def is_accessible(path, mode='r'):
    try:
        f = open(path, mode)
        f.close()
    except IOError:
        return False
    return True


for image in all_image:
    file_location = "/home/pachi/" + image
    if not is_accessible(file_location):
        print(image, "not found in the given directory.")