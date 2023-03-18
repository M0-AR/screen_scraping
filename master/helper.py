import cv2
import numpy as np

# Load the image
img = cv2.imread('image.png', cv2.IMREAD_GRAYSCALE)

# Load the template image of the text
template = cv2.imread('text_template.png', cv2.IMREAD_GRAYSCALE)

# Get the width and height of the template image
w, h = template.shape[::-1]

# Perform template matching
res = cv2.matchTemplate(img, template, cv2.TM_CCOEFF_NORMED)

# Set a threshold for the matching result
threshold = 0.8

# Find the positions where the matching result is above the threshold
locs = np.where(res >= threshold)

# Loop through all the positions and draw a rectangle around the matching area
for pt in zip(*locs[::-1]):
    cv2.rectangle(img, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

# Display the result image with the matching areas highlighted
cv2.imshow('result', img)
cv2.waitKey(0)
cv2.destroyAllWindows()
