import ClointFusion as cf
from tkinter.constants import TRUE
import time
from datetime import date


date_today = str(date.today())
cars24_link = "https://www.cars24.com"

CREDENTIALS_JSON = "C:\Cars24_car_details_download_automation\credentials.json"

DETAILS_JSON = "C:\Cars24_car_details_download_automation\details.json"
city_details = cf.file_get_json_details(path_of_json_file=DETAILS_JSON,section='city')
city_name=city_details.get('city_name')

DETAILS_JSON = "C:\Cars24_car_details_download_automation\details.json"
car_details = cf.file_get_json_details(path_of_json_file=DETAILS_JSON,section='car_details')
car_name=car_details.get('car_name')
sheet_name = car_name +"_"+ city_name + "_" + date_today

folder_location='C:\Cars24_car_details_download_automation\ '
sheet_location= folder_location[:-1] + sheet_name + ".xlsx"


def open_website():

    browser_state = False
    
    try:
        browser_state = cf.launch_website_h(cars24_link)
        
    except:
        print("Error in opening website")

    finally:
        return browser_state

def location_select():

    try:
        cf.browser_mouse_click_h("SELECT MANUALLY")
        time.sleep(0.5)
        cf.browser_write_h(city_name,User_Visible_Text_Element="Search City")
        time.sleep(0.5)
        cf.browser_mouse_click_h(city_name)

    except:
        print("location cannot be selected")


def car_select():
    
    try:
        cf.browser_mouse_click_h("VIEW ALL CARS")
        time.sleep(0.5)
        cf.browser_mouse_click_h("Find your dream car with us")
        time.sleep(0.5)
        cf.browser_write_h(car_name,User_Visible_Text_Element="Find your dream car with us")
        time.sleep(0.5)
        cf.key_write_enter(strMsg=" ")
        time.sleep(0.5)


    except:
        print("car cannot be selected")

def create_excel_sheet():
    
    try:
        cf.excel_create_excel_file_in_given_folder(folder_location[:-1],excelFileName=sheet_name )
    
    except:
        print("excel sheet cannot be created")


def store_excel_sheet():
    
    time.sleep(1)
    
    try:
        d =cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//h2[@itemprop='name']")
        i= 0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            print(h[0])
            cf.excel_set_single_cell(sheet_location,columnName="Name",cellNumber=i,setText=h[0])
            i=i+1
  
    except:
        print("error in collecting the car names") 
    time.sleep(1)
   
    try:
        c=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//h3")
        i=0
        for len in c:
            g=str(len).split(">")
            h=g[1].split("<")
            print(h[0])
            cf.excel_set_single_cell(sheet_location,columnName="Price",cellNumber=i,setText=h[0])
            i=i+1

    except:
        print("error in collecting the car price")

  
    time.sleep(1)


    try:
        c=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//p//span")

        i=0
        q=0
        for len in c:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
  
            if(i%4==0):
                print(t[0]) 
                cf.excel_set_single_cell(sheet_location,columnName="Kilometres used",cellNumber=q,setText=t[0])
                q=q+1 
            i=i+1
            
    except:
        print("error in collecting the kilometres used")

  
    time.sleep(1)


    try:
        g=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//p//span[@itemprop='name']")
        i=0
        for len in g:
          g=str(len).split(">")
          h=g[1].split("<")
          print(h[0])
          cf.excel_set_single_cell(sheet_location,columnName="Engine type",cellNumber=i,setText=h[0])
          i=i+1
   
    except:
        print("error in collecting the engine type ")


def send_outlook_email():
    try:
        
        outlook_details = cf.file_get_json_details(path_of_json_file=CREDENTIALS_JSON,section='Outlook')
        outlook_username = outlook_details.get('username')
        outlook_password = outlook_details.get('password')
        to = outlook_details.get('send_to')

        cf.browser_navigate_h('outlook.com')
        time.sleep(1)
        cf.browser_mouse_click_h('Sign in')
        time.sleep(0.5)

        cf.browser_write_h(outlook_username,User_Visible_Text_Element='Email, phone, or Skype')
        time.sleep(0.5)
        cf.browser_mouse_click_h('Next')

        time.sleep(0.5)

        cf.browser_write_h(outlook_password,User_Visible_Text_Element='Password')
        time.sleep(1)
        cf.browser_mouse_click_h('Sign in')
        time.sleep(1)

        cf.browser_mouse_click_h('New message')

        time.sleep(1)
        cf.browser_write_h(to,User_Visible_Text_Element='To')
        time.sleep(1)

        cf.browser_write_h('car details from cars 24',User_Visible_Text_Element='Add a subject')
        
        body_elem = cf.browser_locate_element_h("//*[@aria-label='Message body']")
        cf.browser_write_h('Please find the attached Report.\n\n\nThanks & Regards\nMohammad Yawar',User_Visible_Text_Element=body_elem)
        
        time.sleep(1)

        cf.browser_mouse_click_h(User_Visible_Text_Element='Attach')
        cf.browser_mouse_click_h(User_Visible_Text_Element='Browse this computer')
        cf.key_write_enter(strMsg=sheet_location)

        time.sleep(2)

        cf.browser_mouse_click_h('Send')


    except:
        print("Error in Sending Outlook Email")





if __name__ == '__main__':
    
    try:    

        browser_state= open_website()
        time.sleep(1)
     
     
        if browser_state==TRUE:
   
            #seting the location in cars24 website
            location_select()
            time.sleep(1)

            #selecting the car model
            car_select()
            time.sleep(1)
            
            #creating the excel sheet
            create_excel_sheet()
            time.sleep(1)
            
            #entering the details in excel sheet
            store_excel_sheet()
            time.sleep(1)
            
            #sending the outlook email
            send_outlook_email()
            time.sleep(3)
            cf.browser_quit_h()
        
        else:
            print("browser not opened")
    
    except:
        print("error")

