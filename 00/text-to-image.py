# # Two solutions text or image in order to click it when find it
# import time
# import cv2
# import numpy as np
#
# text = "Improve ChatGPT"
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

# from datetime import time
#
# import cv2  # import numpy as np
# import pyautogui
# import numpy as np
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
# # Load the template image # TODO: Ask why
# # template = cv2.imread('text.png', cv2.IMREAD_GRAYSCALE)
# template = cv2.imread('date.jpg', cv2.IMREAD_GRAYSCALE)
#
# # Find the text on the screenshot using template matching
# result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
# y, x = np.unravel_index(result.argmax(), result.shape)
#
# # Click on the center of the text
# width, height = template.shape[::-1]
# # pyautogui.click(x + width/2, y + height/2)
# pyautogui.moveTo(x + width / 2, y + height / 2)
#
# # Delay the click for 1 second
# time.sleep(1)

# # Click on the center of the text
# pyautogui.click()

# import pyautogui
#
# x = 100  # example starting x coordinate
# y = 100  # example starting y coordinate
# width = 200  # example selection width
# height = 150  # example selection height
#
# # Move the mouse to the starting position
# # pyautogui.moveTo(x, y)
#
# # Simulate dragging the mouse to select an area
# pyautogui.dragTo(x + width, y + height, button='left')
