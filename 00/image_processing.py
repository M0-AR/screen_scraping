# import pytesseract
# import pyautogui
# from PIL import Image
#
# # take a screenshot of the entire screen
# screenshot = pyautogui.screenshot()
#
# # save the screenshot as a file for pytesseract to read
# screenshot.save("number.png")
#
# # read the text from the screenshot using pytesseract
# text = pytesseract.image_to_string(Image.open("number.png"))
#
# # search for the word "example text" in the text extracted from the screenshot
# if "example text" in text:
#     print("Found the text!")
# else:
#     print("Text not found.")

# # Two solutions text or image in order to click it when find it
# import time
# import cv2
# import numpy as np
#
# text = "Improve ChatGPT"
# font = cv2.FONT_HERSHEY_SIMPLEX
# font_scale = 1
# thickness = 2
# size = cv2.getTextSize(text, font, font_scale, thickness)[0]
# width = size[0]
# height = size[1] * 2  # Double the height of the image
# image = np.zeros((height, width, 3), dtype=np.uint8)
# cv2.putText(image, text, (0, height - size[1]), font, font_scale, (255, 255, 255), thickness)
# cv2.imwrite("text.png", image)
#
# import cv2
# import numpy as np
# import pyautogui
#
# # take a screenshot of the entire screen
# screenshot = pyautogui.screenshot()
#
# # save the screenshot as a file for to read
# screenshot.save("number.png")
#
# # Load the screenshot
# screenshotNP = np.array(pyautogui.screenshot())
#
# # Convert to grayscale
# gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)
#
#
# # Load the template image # TODO: Ask why
# template = cv2.imread('text.png', cv2.IMREAD_GRAYSCALE)
#
# # Find the text on the screenshot using template matching
# result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
# y, x = np.unravel_index(result.argmax(), result.shape)
#
# # Click on the center of the text
# width, height = template.shape[::-1]
# # pyautogui.click(x + width/2, y + height/2)
# pyautogui.moveTo(x + width/2, y + height/2)
#
# # Delay the click for 1 second
# time.sleep(1)
#
# # Click on the center of the text
# pyautogui.click()



# # todo: is working but missing value
# from PIL import Image
# import pytesseract
# # im_file = "image.PNG"
# im_file = "image.jpg"
# # im_file = "template.png"
# img = Image.open(im_file)
# # rgb_im = im.convert('RGB')
# # rgb_im.save("image.jpg")
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
#
# ocr_result = pytesseract.image_to_string(img, lang='dan')
# print(ocr_result)


# import cv2
# import numpy as np
# import pytesseract
#
# # Load the image
# img = cv2.imread('image.jpg')
#
# # Convert to grayscale
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
#
# # Apply image preprocessing if necessary
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Extract the text using Pytesseract
# text = pytesseract.image_to_string(gray, lang='dan')
#
# # Parse the text to extract the data
# data = []
# lines = text.split('\n')
# for line in lines:
#     items = line.split()
#     if len(items) > 0:
#         if ':' in items[0]:
#             data.append(items[1:])
#         elif data:
#             data[-1] += items
#
# # Print the extracted data
# for row in data:
#     # print(row)
#     print([elem for elem in row if elem.isnumeric()])

# import cv2
# import numpy as np
# import pytesseract
#
# # Load the image
# img = cv2.imread('image.jpg')
#
# # Convert to grayscale
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
#
# # Apply image preprocessing if necessary
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Extract the text using Pytesseract
# text = pytesseract.image_to_string(gray, lang='eng')
#
# # Parse the text to extract the data
# data = []
# lines = text.split('\n')
# for line in lines:
#     items = line.split()
#     if len(items) > 0:
#         if ':' in items[0]:
#             data.append(items[1:])
#         elif data:
#             data[-1] += items
#
# # Print the extracted data
# for row in data:
#     print(row)

# import cv2
# import numpy as np
# import pytesseract
#
# # Load the image
# image = cv2.imread('table.jpg')
#
# # Convert the image to grayscale
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
#
# # Display the grayscale image
# # cv2.imshow('Grayscale Image', gray)
# # cv2.waitKey(0)
# # cv2.destroyAllWindows()
#
# # Apply image pre-processing techniques if necessary
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Extract the text using Pytesseract
# text = pytesseract.image_to_string(gray, lang='dan')
#
# print(text)
# # Split the text into rows
# rows = text.split('\n')
#
# # Extract the column headers
# header = rows[0].split()
#
# # Extract the row headers and data
# data = {}
# for row in rows[1:]:
#     items = row.split()
#     if len(items) > 1:
#         data[items[0]] = [float(i.replace(',', '.')) if i.replace(',', '.').isnumeric() else i for i in items[1:]]
#
# # Print the extracted data
# print(' '.join(header))
# for row in data:
#     print(row, ' '.join(map(str, data[row])))


# import cv2
# import pytesseract
# import re
# import numpy as np
#
# # Load the image
# img = cv2.imread('table.jpg')
#
# # Convert the image to grayscale
# gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
#
# # Apply thresholding to the image to make it brighter
# thresh_value = 150
# _, thresh = cv2.threshold(gray, thresh_value, 255, cv2.THRESH_BINARY)
#
# # Display the thresholded image
# cv2.imshow('Thresholded Image', thresh)
# cv2.waitKey(0)
# cv2.destroyAllWindows()
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Extract text using Pytesseract
# text = pytesseract.image_to_string(thresh, lang='dan')
#
# print(text)
#
# # Extract the rows of data
# rows = text.split('\n')
#
# # Extract the dates and times
# datetime_row = rows[0].split()
# dates = datetime_row[::2]
# times = datetime_row[1::2]
#
# # Extract the column names
# cols = [re.sub(r'[^a-zA-Z0-9 ]', '', row).split() for row in rows[1:3]]
# cols = [col[1:] for col in cols]
#
# # Extract the data
# data = []
# for row in rows[3:]:
#     row_data = []
#     items = row.split()
#     for item in items[1:]:
#         if ',' in item:
#             row_data.append(float(item.replace(',', '.')))
#         else:
#             row_data.append(item)
#     data.append(row_data)
#
# # Print the data
# print(' '.join(['{:<20}'.format('')] + [f'{d:<10}' for d in dates]))
# print(' '.join(['{:<20}'.format('')] + [f'{t:<10}' for t in times]))
# for i in range(len(cols)):
#     print(f'{cols[i][0]:<20}' + ' '.join([f'{d:<10}' for d in data[i]]))
#




# TODO: good so far

# import cv2
# import numpy as np
# import pytesseract
#
# # Load the image
# image = cv2.imread('table.jpg')
#
# # Convert the image to grayscale
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
#
# # Apply thresholding to the image
# thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
#
# # Apply different preprocessing techniques
# preprocess_options = [
#     ('Original', gray),
#     ('Thresholding', thresh),
#     ('Gaussian Blurring', cv2.GaussianBlur(gray, (3, 3), 0)),
#     ('Median Filtering', cv2.medianBlur(gray, 3)),
#     ('Bilateral Filtering', cv2.bilateralFilter(gray, 5, 75, 75)),
# ]
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Loop through each option and extract the text using Pytesseract
# for option in preprocess_options:
#     name, img = option
#     cv2.imshow(name, img)
#     cv2.waitKey(0)
#     cv2.destroyAllWindows()
#     text = pytesseract.image_to_string(img, lang='dan', config='--psm 6')
#     print(f'Preprocessing Option: {name}\n\n{text}\n\n')

# import cv2
# import numpy as np
# import pytesseract
#
# # Load the image
# image = cv2.imread('table.jpg')
#
# # Convert the image to grayscale
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
#
# # Apply thresholding to the image
# thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
#
# # Apply different preprocessing techniques
# preprocess_options = [
#     ('Gaussian Blurring', cv2.GaussianBlur(gray, (3, 3), 0)),
#     ('Median Filtering', cv2.medianBlur(gray, 3)),
#     ('Bilateral Filtering', cv2.bilateralFilter(gray, 5, 75, 75)),
# ]
#
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Loop through each option and extract the text using Pytesseract
# for option in preprocess_options:
#     name, img = option
#
#     # Apply preprocessing to the image
#     img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
#
#     # Extract the text using Pytesseract
#     text = pytesseract.image_to_string(img, lang='dan', config='--psm 6')
#     print(f'Preprocessing Option: {name}\n\n{text}\n\n')


# import cv2
# import numpy as np
# import pytesseract
# import re
#
# # Set the path to the Tesseract executable
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# # Load the image
# image = cv2.imread('02.jpg')
# # image = cv2.imread('table.jpg')
#
# # Convert the image to grayscale
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
#
# # Apply thresholding to the image
# thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
#
# # Apply different preprocessing techniques
# preprocess_options = [
#     ('Gaussian Blurring', cv2.GaussianBlur(gray, (3, 3), 0)),
#     ('Median Filtering', cv2.medianBlur(gray, 3)),
#     ('Bilateral Filtering', cv2.bilateralFilter(gray, 5, 75, 75)),
# ]
#
# # Loop through each option and extract the text using Pytesseract
# for option in preprocess_options:
#     name, img = option
#
#     # Apply preprocessing to the image
#     img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
#
#     # Extract the text using Pytesseract
#     text = pytesseract.image_to_string(img, lang='dan', config='--psm 6')
#
#     print(f'Preprocessing Option: {name}\n\n{text}\n\n')
