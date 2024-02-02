import numpy as np
import cv2

np_array = np.ones((5, 5), np.float32)

blur = cv2.GaussianBlur(np_array,(5,5),0)