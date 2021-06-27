from openpyxl import load_workbook  

from openpyxl.drawing.image import Image  

  

# Let's use the hello_world spreadsheet since it has less data  

workbook = load_workbook(filename="student_chart1.xlsx")  

sheet = workbook.active  

  

logo = Image(r"C:\Users\Sai Reddy\Pictures\Screenshots\image.png")  

  

# A bit of resizing to not fill the whole spreadsheet with the logo  

logo.height = 150  

logo.width = 150  

  

sheet.add_image(logo, "E2")  

workbook.save(filename="hello_world_logo1.xlsx")
