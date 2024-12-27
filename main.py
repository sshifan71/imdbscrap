from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "IMDb Movies"
sheet.append(["Serial Number", "Movie Name", "Year of Release", "Rating"])


driver = webdriver.Chrome() 
driver.get("https://www.imdb.com/search/title/?title_type=feature&release_date=2015-01-01,2015-12-31")

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div[2]/div[3]')))


movies = driver.find_elements(By.XPATH, '/html/body/div[2]/main/div[2]/div[3]/section/section/div/section/section/div[2]/div/section/div[2]/div[2]/ul')
serial_number = 1

print(f"Found {len(movies)} movies")

for movie in movies:
    
    name = movie.find_element(By.CLASS_NAME, 'ipc-title__text').text
    
    year = movie.find_element(By.XPATH, '//*[@id="__next"]/main/div[2]/div[3]/section/section/div/section/section/div[2]/div/section/div[2]/div[2]/ul/li[1]/div/div/div/div[1]/div[2]/div[2]/span[1]').text.strip("()")
    
    rating = float(movie.find_element(By.CLASS_NAME, 'ipc-rating-star--rating').text)
    
    
    if rating >= 7:
        sheet.append([serial_number, name, year, rating])
        serial_number += 1


workbook.save("IMDb_Top_Movies.xlsx")


driver.quit()

print("Data has been scraped and saved to IMDb_Top_Movies.xlsx")
