import cv2
import numpy as np

def preprocess_image(image):
    """
    Converts image to grayscale and applies thresholding to improve OCR accuracy.
    """
    img = np.array(image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(thresh)