from selenium import webdriver
from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException
import time
from webdriver_manager.chrome import ChromeDriverManager

# Load Data Excel
wb = load_workbook(filename=r"data.xlsx")

# Mengambil Sheet
sheetRange = wb['Sheet1']

# Mengambil Data Website
inputan = webdriver.Chrome(ChromeDriverManager().install())
inputan.get('https://aldev.my.id/testform/')

# Looping
i = 2

# Looping data yang tersedia pada excel
while i <= len(sheetRange['A']):
    Name = sheetRange['A'+str(i)].value
    Email = sheetRange['B'+str(i)].value
    Pekerjaan = sheetRange['C'+str(i)].value
    Alamat = sheetRange['D'+str(i)].value

# memulai inputan kedalam website
    try:
        inputan.find_element("id", "exampleFormControlInput1").send_keys(Name)
        inputan.find_element("id", "exampleFormControlInput2").send_keys(Email)
        inputan.find_element("name", "pekerjaan").send_keys(Pekerjaan)
        inputan.find_element("name", "alamat").send_keys(Alamat)
        inputan.find_element(
            "xpath", '/html/body/div/div/div[2]/form/div[2]/button[1]').click()

# Kalau lag atau apa akan gagal
    except TimeoutException:
        print("Gagal")
        pass

# Clear Kembali hasil inputan tadi
    inputan.find_element("id", "exampleFormControlInput1").clear()
    inputan.find_element("id", "exampleFormControlInput2").clear()
    inputan.find_element("name", "pekerjaan").clear()
    inputan.find_element("name", "alamat").clear()
    time.sleep(3)
    i = i + 1

print("Form Sudah Selesai Diinput !")
