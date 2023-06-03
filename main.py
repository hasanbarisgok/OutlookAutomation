import streamlit as st  
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep


class auto_mail: 

    """These are xpath, and css_selector variables. We're using the variables on about Selenium Functions. As using the variables, computer does things by itself.
    The variables created as class variables.
    """
    
    email_area_xpath = '//*[@id="i0116"]'
    
    password_area_xpath = '//*[@id="i0118"]'
    
    new_mail_button_xpath = '//*[@id="innerRibbonContainer"]/div[1]/div/div/div/div[2]/div/div/span/button[1]'
    
    to_area_css_selector = 'div[aria-label="To"]'
    
    subject_area_css_selector = 'input[placeholder="Add a subject"]'
    
    message_area_css_selector = 'div[aria-label="Message body, press Alt+F10 to exit"]'
    
    send_message_button_css_selector = 'button[aria-label="Send"]'
    
    
    def __init__(self):
        
        """Constructor method. We're using variable $self.counter to count how much messages are sended, when finished the process.
        """
        
        self.main_area()
        
        #self.counter = 0
    
    def main_area(self):
        
        """These function creates the web-site architectural, so creates Streamlit pages with makes work to all design functions.
        """
        
        self.outlook_login()
        
        self.subject_area()
        
        self.message_area()
        
        self.users_area()  
        
        self.start_automation()
        
    
    
    def outlook_login(self):
        
        """Mail information areas. As using the function, we sign in the outlook system. So your account information must be entered these areas truly.
        """
        
        self.mail : str = st.text_input("Please enter your outlook mail: ")
        
        self.mail_passwd: str = st.text_input("Please enter your outlook mail password: ", type="password")
        
    def subject_area(self):
        
        """This is subject's area, so mail's subject.
        """
        
        self.subject_text = st.text_input("Type the Subject: ")
        
        
    def message_area(self):
        
        """This is messages' area, mail's draft.
        """
        
        self.message_text = st.text_area("Type your message/mail draft:", value = "", height= 300)
        
        if self.message_text:
            
            #Split the message into a list of messages using the unique character
            
            self.message = self.message_text.split("\n")
        
        
    def users_area(self):
        
        """This is e-mails' area, messages will be sended to which mails.
        """
        
        self.users_text = st.text_area("Type the users/mails: ", height=200)
        
        self.users_list = self.users_text.splitlines()
        
            
    def start_automation(self): 
        
        """The function makes work other functions step by step.
        """
        
        self.start_button = st.button("Start automation")
        
        if self.start_button:
            
            self.go_page()
            
            sleep(1)
            
            self.enter_mail()
            
            sleep(1)
            
            self.enter_password()
            
            sleep(1)
            
            self.click_new_mail()
            
            sleep(1)
            
            self.type_mail()
            
            sleep(1)
            
            self.num_of_sended_messages()
    
        
    def go_page(self):
        
        """The function openining the outlook's website.
        """
        
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        
        self.driver.maximize_window()
        
        self.driver.get("https://outlook.live.com/owa/?nlp=1")
        
        
    def enter_mail(self):
        
        """The function writes your outlook mail to mail area. (on sign-in page)
        """
        
        self.email_area = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, auto_mail.email_area_xpath)))
        
        self.email_area.send_keys(self.mail)
        
        self.email_area.send_keys(Keys.ENTER)
                
        
    def enter_password(self):
        
        """The function writes your outlook mail password to password area. (on sign-in page, password section)
        """
        
        self.password_area = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, auto_mail.password_area_xpath)))
        
        self.password_area.send_keys(self.mail_passwd)
        
        self.password_area.send_keys(Keys.ENTER)
        
        
    def click_new_mail(self):
        
        """The function clicks new e-mail button.
        """
        
        self.new_email_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, auto_mail.new_mail_button_xpath)))
        
        self.new_email_button.click()
        
    
    def type_mail(self):
        
        """The function combines all functions.
        """
        
        with st.spinner("Mails are sending now"):
            
            for user in self.users_list:
            
                self.person_mail_area = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, auto_mail.to_area_css_selector)))
            
                #writing person's mail to person mail area.
                self.person_mail_area.send_keys(user)
                
                sleep(1)
                
                #to select user's mail
                self.person_mail_area.send_keys(Keys.ENTER)
                
                self.add_subject()
                
                sleep(1)
                
                self.type_message()
                            
                self.send_mail()
                
                #self.counter += 1
                
                sleep(1)
                
                self.click_new_mail()
                
        st.success("All done.")

    
    def add_subject(self):
        
        """The function writes your subject to subject area.
        """
        
        self.subject_area = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, auto_mail.subject_area_css_selector)))
        
        self.subject_area.send_keys(self.subject_text)
        
    def type_message(self):
        
        """The function writes your message to draft area.
        """
        
        self.draft_area = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, auto_mail.message_area_css_selector)))
        
        for message in self.message:
            
            self.message_with_spaces = message + " "
            
            self.draft_area.send_keys(self.message_with_spaces)
            
            self.draft_area.send_keys(Keys.SHIFT + Keys.ENTER)
        
        
    def send_mail(self):
        
        """The function clicks `send mail` button
        """
        
        self.send_message_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, auto_mail.send_message_button_css_selector)))
        
        self.send_message_button.click()
        
    
    """def num_of_sended_messages(self):
        
       The function writes on the screen how much messages is sended
        
        
        self.info_num_sended_messages = st.write("Sended message num is: ", self.counter)"""
        
        
obj_1 = auto_mail()