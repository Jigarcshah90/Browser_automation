import openpyxl
import cv2
import numpy as np
import re
import urllib
from PIL import Image
from pytesseract import image_to_string
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# create a new chrome session
driver = webdriver.Chrome(executable_path = "D:\\Jigar\\tesseract-python\\chromedriver.exe")
driver.implicitly_wait(30)
driver.maximize_window()

# Navigate to the application home page
driver.get("http://www.esic.in/EmployeePortal/login.aspx#")

######################### Excel code Started #####################################
#Code to fetch the data from excel and save it in YourUsername
excel_document = openpyxl.load_workbook('D:\\Jigar\\tesseract-python\\employee_id.xlsx')
sheet = excel_document.get_sheet_by_name('Sheet1')

last_row_found = 'N'
number_of_round = 10
while number_of_round >= 1:
    print number_of_round
    number_of_round -= 1
    cell_row = 1
    last_row_found = 'N'
    while last_row_found == 'N':
        print "cell_row"+str(cell_row)
        cell_row = cell_row+1
        
        

        excel_data =  sheet.cell(row = cell_row, column = 1).value
        print excel_data
        if excel_data == '' or excel_data is None:
            last_row_found = 'Y'
            continue
        YourUsername = int(excel_data)
        # Navigate to the application home page
        driver.get("http://www.esic.in/EmployeePortal/login.aspx#")
        ######################### Excel code ended #####################################

        ######################### Croping the image after screenshot #####################
        driver.save_screenshot("D:\\Jigar\\tesseract-python\\screenshot.png")

        image = Image.open("D:\\Jigar\\tesseract-python\\screenshot.png")

        def crop(image_path, coords, saved_location):
            
            image_obj = Image.open("D:\\Jigar\\tesseract-python\\screenshot.png")
            cropped_image = image_obj.crop(coords)
            cropped_image.save("D:\\Jigar\\tesseract-python\\crop.png")
            #cropped_image.show()
         
        if __name__ == '__main__':
            crop(image, (1100, 340, 1236, 370), 'cropped.jpg')

        #########################Croping Done and saved at location ######################
        ###################################################################################

        ######################### Capcha code Started #####################################
        #code to fetch capcha and save it in txtCaptcha

        img = cv2.imread("D:\\Jigar\\tesseract-python\\crop.png")
        img=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)


        img = cv2.bitwise_not(img)
        th2 = cv2.adaptiveThreshold(img,255, cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,15,-2)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        horizontal = th2
        vertical = th2
        rows,cols = horizontal.shape

        #inverse the image, so that lines are black for masking
        horizontal_inv = cv2.bitwise_not(horizontal)
        #perform bitwise_and to mask the lines with provided mask
        masked_img = cv2.bitwise_and(img, img, mask=horizontal_inv)
        #reverse the image back to normal
        masked_img_inv = cv2.bitwise_not(masked_img)
        #cv2.imshow("masked img", masked_img_inv)
        #cv2.imwrite("result2.jpg", masked_img_inv)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        horizontalsize = int(cols / 30)
        horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (horizontalsize,1))
        horizontal = cv2.erode(horizontal, horizontalStructure, (-1, -1))
        horizontal = cv2.dilate(horizontal, horizontalStructure, (-1, -1))
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        verticalsize = int(rows / 30)
        verticalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, verticalsize))
        vertical = cv2.erode(vertical, verticalStructure, (-1, -1))
        vertical = cv2.dilate(vertical, verticalStructure, (-1, -1))
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        vertical = cv2.bitwise_not(vertical)
        cv2.waitKey(0)
        cv2.destroyAllWindows()


        # step3
        smooth = vertical.copy()

        #step 4
        smooth = cv2.blur(smooth, (4,4))
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        #step 5
        (rows, cols) = np.where(img == 0)
        vertical[rows, cols] = smooth[rows, cols]

        cv2.waitKey(0)
        cv2.destroyAllWindows()

        upper_img = vertical[1:13]
        below_img = vertical[15:32]

        Merge_img = np.vstack((upper_img,below_img)) 

        cv2.imwrite("D:\\Jigar\\tesseract-python\\vertical_final.jpg", Merge_img)

        im = Image.open(r'D:\\Jigar\\tesseract-python\\vertical_final.jpg')

        img1 = (image_to_string(im))

        img2 = img1.replace("i","1")
        img3 = img2.replace("t","1")
        img4 = img3.replace(".","")
        img5 = img4.replace("s","8")
        img6 = img5.replace("c", "e")
        img7 = img6.replace("C", "c")
        img8 = img7.replace("O", "0")
        img9 = img8.replace("Z", "2")
        img10 = img9.replace("q", "a")
        img11 = img10.replace("s", "5")
        img12 = img11.replace("Q", "9")
        img13 = img12.replace("U", "0")
        img14 = img13.replace("D", "0")
        img15 = img14.replace("o", "9")
        img16 = img15.replace("I", "1")
        img17 = img16.replace("l", "1")
        img18 = img17.replace("E", "a")
        img19 = img18.replace("r", "c")
        img20 = img19.replace("G", "e")
        img21 = img20.replace("g", "8")
        img22 = img21.replace("k", "c")
        img23 = img22.replace("K", "f")
        img24 = img23.replace("M", "4")

        #text_main = re.sub(r'[^\x00-\x7F]+','', img24)        
        text_main = re.sub('\W','',img24)
        txtCaptcha = text_main.lower()
        print txtCaptcha
        
        ######################### Captcha code ended #####################################

        username = driver.find_element_by_id("txtUserName")
        Captcha = driver.find_element_by_id("txtCaptcha")

        username.send_keys(YourUsername)
        Captcha.clear()
        Captcha.send_keys(txtCaptcha)

        driver.find_element_by_id('btnLogin').click()

#<input type="submit" name="btnLogin" value="Login" id="btnLogin" class="loginbtn">

driver.implicitly_wait(30)

#driver.find_element_by_xpath('//a[img/@src="/Includes/images/excel.png"]').click()





