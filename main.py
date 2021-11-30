from win32com.client import Dispatch
import pathlib
import qrcode
import cv2
from PIL import Image
from pyzbar import pyzbar


def start_barcode_video():
    # begin read qr code in video
    camera = cv2.VideoCapture(1)
    ret, frame = camera.read()

    while ret:
        ret, frame = camera.read()
        frame = read_barcodes(frame)
        cv2.imshow('Barcode/QR code reader', frame)
        if cv2.waitKey(1) & 0xFF == 27:
            break

    camera.release()
    cv2.destroyAllWindows()


def generate_qrcode(data, img_name):
    img = qrcode.make(data)
    img.save(img_name)
    return img_name


def read_barcodes(frame):

    barcodes = pyzbar.decode(frame)
    for barcode in barcodes:
        x, y, w, h = barcode.rect

        # 1
        barcode_info = barcode.data.decode('utf-8')
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

        # 2
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, barcode_info, (x + 6, y - 6), font, 1.0, (255, 255, 255), 1)

        # 3
        with open("barcode_result.txt", mode='w') as file:
            file.write("Recognized Barcode:" + barcode_info)
    return frame


def print_label(data):
    # Declare variables
    barcode_val = data
    barcode_left = barcode_val[:8]
    barcode_right = barcode_val[-8:]
    barcode_image_filename = generate_qrcode(barcode_val, 'qrcodes/' + barcode_val + '-QR.png')
    label_path = pathlib.Path('label_templates/double_QR.label')
    my_printer = 'DYMO LabelWriter 450'

    # Open printer and label template
    printer_com = Dispatch('Dymo.DymoAddIn')
    printer_com.SelectPrinter(my_printer)
    printer_com.Open(label_path)
    printer_label = Dispatch('Dymo.DymoLabels')

    # Set label variables
    printer_label.SetImageFile('QR1', barcode_image_filename)
    printer_label.SetImageFile('QR2', barcode_image_filename)
    printer_label.SetField('FIRST8_1', barcode_left)
    printer_label.SetField('FIRST8_2', barcode_left)
    printer_label.SetField('LAST8_1', barcode_right)
    printer_label.SetField('LAST8_2', barcode_right)

    # Print the label
    printer_com.StartPrintJob()
    printer_com.Print(1, False)
    printer_com.EndPrintJob()

    print("barcode printed")


def main():

    valid = False
    while not valid:
        data = input("enter a device id: ")
        if len(data) == 16:
            valid = True

    print_label(data)

    start_barcode_video()

    print('End')


if __name__ == '__main__':
    main()