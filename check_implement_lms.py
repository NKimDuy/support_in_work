from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
import time
import requests
from datetime import datetime
from openpyxl import Workbook
import json

def get_string(str, start = "IDNumber"):
      position = str.find(start)
      if position != -1:
            return str[position:]


def delete_string(str, delete_word = "IDNumber", semester = "242"):
      return str.replace("".join([delete_word, " ", semester, "_"]), "")


def get_lsa(semester, url):
      value = "Đăng nhập bằng HCMCOU-SSO"
      semester = " ".join(["[LIVE] LMS TX", semester])
      general = "http://lsa.ou.edu.vn/vi/admin/mm/report/usersiteoverviews"
      xpath = f"//button[text()='{value}']"
      chrome_options = Options()
      #chrome_options.add_argument("--ignore-certificate-errors")
      #chrome_options.add_argument("--disable-features=StrictTransportSecurity")
      chrome_options.add_argument("--headless")
      #chrome_options.add_argument("--allow-insecure-localhost")  # Nếu là localhost
      driver = webdriver.Chrome(options=chrome_options)
      #driver = webdriver.Chrome()
      #driver.get("http://lsa.ou.edu.vn")
      driver.get(url)

      try:
            button_semester = WebDriverWait(driver, 15).until(
                  EC.element_to_be_clickable((By.XPATH, xpath))
            )
            button_semester.click()
      except:
            print("Không tìm thấy nút đăng nhập bằng HCMCOU-SSO")

      try:
            dropdown = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "form-usertype"))
            )
            select_type_user = Select(dropdown)
            select_type_user.select_by_visible_text("Cán bộ-Nhân viên / Giảng viên")
      except:
            print("Không tìm thấy nút Cán bộ nhân viên/Giảng viên")

      try:
            username = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.ID, "form-username"))
            )
            username.send_keys("duy.nk")
      except:
            print("Không tìm thấy ô để nhập tài khoản")

      try:
            password = driver.find_element(By.ID, "form-password")
            password.send_keys("tonyTeo!998")
      except:
            print("Không tìm thấy ô để nhập mật khẩu")

      try:
            captcha = driver.find_element(By.ID, "form-captcha")
            captcha.send_keys("clcl")
      except:
            print("Không tìm thấy ô để nhập Capcha")

      try:
            button_login = WebDriverWait(driver, 15).until(
                  EC.element_to_be_clickable((By.XPATH, "//button[text()='Đăng nhập']"))
            )
            button_login.click()
      except:
            print("Không tìm thấy nút đăng nhập")

      try:
            has_found_button_allow = driver.find_elements(By.CSS_SELECTOR, ".btn.btn-success.btn-approve")
            if has_found_button_allow:
                  button_allow = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.CSS_SELECTOR, ".btn btn-success btn-approve"))
                  )
                  button_allow.click()
      except:
            print("Không tìm thấy nút để nhấn đồng ý")

      try:
            dropdown_semester = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "moodlesiteid"))
            )

            select_type_semester = Select(dropdown_semester)
            select_type_semester.select_by_value("54")
      except:
            print("Không tìm thấy dropdownlist thể hiện học kỳ")

      try:
            driver.execute_script("arguments[0].style.display='block';", driver.find_element(By.ID, "menu_1_sub"))
            overview_link = WebDriverWait(driver, 20).until(
                  EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='usersiteoverviews']"))
            )
            overview_link.click()
      except:
            print("Không tìm thấy nút report")

      try:
            table = WebDriverWait(driver, 40).until(
                  EC.presence_of_element_located((By.ID, "ourptlistcourse"))  # Thay "myTable" bằng ID thực tế
            )
      except:
            print("không tìm thấy bảng")

      try:
            rows = WebDriverWait(driver, 20).until(
                  EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
            )
      except: 
            print("không tìm thấy các dòng")

      get_subject = list()
      for row in rows:
            cells = row.find_elements(By.XPATH, ".//td")
            for cell in cells:
                  if "[242]" in cell.text and len(get_string(cell.text)) > len("IDNumber"):
                        get_subject.append(delete_string(str(get_string(cell.text))))

      return get_subject


# định dạng ngày có dạng yyyy-MM-DD
def get_subject_by_day(semester, from_day, to_day):
      url_link_unit = "https://api.ou.edu.vn/api/v1/hdmdp"
      url_list_subject_semester = "https://api.ou.edu.vn/api/v1/tkblopdp"
      headers = {
            "Authorization": "Bearer 52C4E470AF3AE6C56276FAE8666788291F7AEA1667FE67C9DF743FF49FD5C74B"
      }
      from_day = datetime.strptime(from_day, "%Y-%m-%d")
      to_day = datetime.strptime(to_day, "%Y-%m-%d")
      list_subject_in_range = []

      get_list_unit = requests.get(url_link_unit, headers=headers)
      list_unit = get_list_unit.json()
      for unit in list_unit.get("data", []):
            params_list_subject_semester = {
                  "nhhk": semester,
                  "madp": unit["MaDP"]
            }
            get_list_subject_semester = requests.get(url_list_subject_semester, headers=headers, params=params_list_subject_semester)
            list_subject_semester = get_list_subject_semester.json()
            for lst in list_subject_semester.get("data", []):
                  if lst["TUNGAYTKB"] is not None:
                        if from_day <= datetime.strptime(lst["TUNGAYTKB"], "%Y-%m-%d") <= to_day:

                              per_subject = {
                                    "NhomTo": lst["NhomTo"], 
                                    "MaMH": lst["MaMH"],
                                    "TenMH": lst["TenMH"],
                                    "TUNGAYTKB": lst["TUNGAYTKB"],
                                    "MaLop": lst["MaLop"],
                                    "TenLop": lst["TenLop"],
                                    "MaDP": lst["MaDP"],
                                    "TenDP": lst["TenDP"]
                              }
                              if per_subject not in list_subject_in_range:
                                    list_subject_in_range.append(per_subject)
     
      return list_subject_in_range


def main():
      semester = "242"
      from_day = "2025-04-21"
      to_day = "2025-04-27"
      url_lsa = "http://lsa.ou.edu.vn"
      report_final = []

      list_lsa = get_lsa(semester, url_lsa)
      list_subject_by_day = get_subject_by_day(semester, from_day, to_day)


      for subject in list_subject_by_day:
            sub_group = subject["MaMH"] + "_" + subject["NhomTo"]
            print(subject["TUNGAYTKB"])
            if any(sub_group in lsa for lsa in list_lsa):
                  report_final.append(subject)
      
      print(len(report_final))
if __name__ == "__main__":
      main()