from array import array
import numpy as np
import cv2 as cv
from sys import argv

img_file = 'C:\TSP\png.png'
if (len(argv) > 1) :
    img_file = argv[1]

# 加载图片到灰度图像
img = cv.imread(img_file, cv.IMREAD_GRAYSCALE)
h,w = img.shape
print(h,w)
print(img.size)

f = open('C:\TSP\BITMAP', 'w')
line = '%d %d %d\n' % (h,w, img.size)
f.write(line)

lst = ['0'] * w
for m in range(h):
    for n in range(w):
        if img[m,n] < 127:
            lst[n] = '1'
        else:
            lst[n] = '0'
    line = ' '.join(lst)
    f.write(line+'\n')

# cv.imshow('image',img)
# cv.waitKey(5000)
# cv.destroyAllWindows()
