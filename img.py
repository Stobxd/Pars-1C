import openpyxl
from openpyxl.drawing.image import Image


book = openpyxl.open('Рузльтат Дату укажи (работает только в буратино).xlsx')
sheete = book.active

img = Image('dfgdf.jpg')

sheete.add_image(img, 'D41')
img.width = 350 
img.height = 400


book.save('huy.xlsx')