import os
import playsound
from gtts import gTTS
from datetime import datetime
import speech_recognition
import re
import requests
from youtube_search import YoutubeSearch
import webbrowser
import json
import urllib.request as urllib2
import ctypes
from webdriver_manager.chrome import ChromeDriverManager
import wikipedia
from unidecode import unidecode
import smtplib
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys

path = ChromeDriverManager().install()

def noi(text):
    print("Robot: {}".format(text))
    tts = gTTS(text=text, lang='vi')
    date_string = datetime.now().strftime("%d%m%Y%H%M%S")
    tts.save("sound"+date_string+".mp3")
    playsound.playsound("sound"+date_string+".mp3")
    os.remove("sound"+date_string+".mp3")

def noi2(text):
    tts = gTTS(text=text, lang='vi')
    date_string = datetime.now().strftime("%d%m%Y%H%M%S")
    tts.save("sound"+date_string+".mp3")
    playsound.playsound("sound"+date_string+".mp3")
    os.remove("sound"+date_string+".mp3")


def nghe():
	robot_ear=speech_recognition.Recognizer()
	with speech_recognition.Microphone() as mic:
		#print("Robot: Tôi nghe đây")
		audio=robot_ear.listen(mic)
	try:
		text=robot_ear.recognize_google(audio, language="vi-VN")
		print(text)
		return(text)		
	except:
		text="Bạn có thể nói lại được không"
		return 0


def stop():
    noi("Hẹn gặp lại bạn sau!")

def get_text():
    for i in range(3):
        text = nghe()
        if text:
            return text.lower()
        elif i < 2:
            noi("Tôi không nghe rõ. Bạn nói lại được không!")
    #time.sleep(2)
    stop()
    return 0

def hello(name):
	now=datetime.now()
	day_time=int(now.hour)
	if day_time<12:
		noi("Chào buổi sáng bạn "+name+". Chúc bạn một ngày tốt lành")
	elif 12<=day_time<18:
		noi("Chào buổi chiều bạn "+name+". Chúc bạn một ngày tốt lành")
	else:
		noi("Chào buổi tối bạn "+name+". Chúc bạn một ngày tốt lành")  


def handle_score(text):
    s=text
    text=""
    i=0
    while i < len(s):
        try:
            if s[i]=='.':
                i+=2
            elif s[i+1]=='0' and s[i]!='.':
                text=text+s[i:i+2]+"|"
                i+=2
            elif s[i+1]=='.':
                text=text+s[i:i+3]+"|"
                i+=3
            else:
                text=text+s[i]+"|"
                i+=1
        except:
            text=text+s[i]
            i+=1
    return(text)


def score():

    browser = webdriver.Chrome(executable_path="D:\project4\chromedriver.exe")

    browser.get("https://tracuu.vnedu.vn/so-lien-lac/")

    sleep(3)

    choose_city = browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[1]/div/div/div[2]/form/div[1]/select")
    choose_city.click()

    sleep(3)

    choose_city = browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[1]/div/div/div[2]/form/div[1]/select/option[16]")
    choose_city.click()

    sleep(3)

    txtUser = browser.find_element_by_id("keyword")
    txtUser.send_keys("0766514695")


    txtUser.send_keys(Keys.ENTER)

    sleep(3)

    choose_name = browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[1]/div/div/div[3]/ul/li/a/h3")
    choose_name.click()
    sleep(3)

    txtPass = browser.find_element_by_id("password")
    txtPass.send_keys("0766514695")
    txtPass.send_keys(Keys.ENTER)

    sleep(3)

    rows = len(browser.find_elements_by_xpath("/html/body/div[3]/div[2]/div/div[2]/table[2]/tbody/tr"))
    cols = len(browser.find_elements_by_xpath("/html/body/div[3]/div[2]/div/div[2]/table[2]/tbody/tr[1]/th"))
    a=['','điểm đầu giờ thường xuyên: ','điểm giữa kỳ: ','điểm cuối kỳ: ','điểm trung bình: ']
    for i in range(2,rows+1):
        for j in range(1,cols+1):
            value = browser.find_element_by_xpath("/html/body/div[3]/div[2]/div/div[2]/table[2]/tbody/tr["+str(i)+"]/td["+str(j)+"]").text
            try:
                m=float(value)
                print(handle_score(value),end='      ')
                noi2(a[j-1]+handle_score(value))
            except:
                print(value,end='       ')
                noi2(a[j-1]+value)
        print("\n")

    browser.quit()


def time(text):
    now = datetime.now()
    if "giờ" in text:
        noi('Bây giờ là %d giờ %d phút' % (now.hour, now.minute))
    elif "ngày" in text:
        noi("Hôm nay là ngày %d tháng %d năm %d" % (now.day, now.month, now.year))
    else:
        noi("Tôi chưa hiểu ý của bạn. Bạn nói lại được không?")




def open_application(text):
    if "google" in text:
        print("Mở Google Chrome")
        os.startfile('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
    elif "word" in text:
        print("Mở Microsoft Word")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk')
    elif "excel" in text:
        print("Mở Microsoft Excel")
        os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
    else:
        print("Ứng dụng chưa được cài đặt. Bạn hãy thử lại!")



def open_website(text):
    reg_ex = re.search('mở  (.+)', text)
    if reg_ex:
        domain = reg_ex.group(1)
        url = 'https://www.' + domain
        webbrowser.open(url)
        print("Trang web bạn yêu cầu đã được mở.")
        return True
    else:
        return False

#
#
#
def open_google_and_search(text):
	search_for=text.split("kiếm",1)[1]
	url=f"https://www.google.com/search?q={search_for}"
	webbrowser.get().open(url)
	noi('Đây là thông tin bạn cần tìm kiếm')


def send_email():
	t='không'
	while "đúng" not in t:
		noi("Hãy cho tôi biết địa chỉ email người nhận")
		a=str(get_text())
		s=a.split(" ")
		s="".join(s)
		s=unidecode(s).lower()
		noi('Đây có phải là địa chỉ bạn cần muốn gửi đến. '+a+' @gmail.com')
		t=str(nghe()).lower()
		if "đúng" in t:
			t='đúng'
		else:
			t='không'
	addreas = s+'@gmail.com'
	print(addreas)
	noi('Nội dung bạn cần gửi là gì')
	msg = get_text()
	

	client = smtplib.SMTP("smtp.gmail.com",587)
	client.starttls()
	try:
		client.login('khailecoder2004@gmail.com','khaikhongquen')
		client.sendmail('khailecoder2004@gmail.com',addreas,msg.encode('utf-8'))
		noi("Đã gửi thành công đến "+addreas)
	except:
		noi("Gặp lỗi khi gửi đến "+addreas)

	client.quit()



def current_weather():
    noi("Bạn muốn xem thời tiết ở thành phố nào ạ.")
    ow_url = "http://api.openweathermap.org/data/2.5/weather?"
    city = get_text()
    if not city:
        pass
    api_key = "fe8d8c65cf345889139d8e545f57819a"
    call_url = ow_url + "appid=" + api_key + "&q=" + city + "&units=metric"
    response = requests.get(call_url)
    data = response.json()
    if data["cod"] != "404":
        city_res = data["main"]
        current_temperature = city_res["temp"]
        current_pressure = city_res["pressure"]
        current_humidity = city_res["humidity"]
        suntime = data["sys"]
        sunrise = datetime.fromtimestamp(suntime["sunrise"])
        sunset = datetime.fromtimestamp(suntime["sunset"])
        wthr = data["weather"]
        weather_description = wthr[0]["description"]
        now = datetime.now()
        content = """
        Hôm nay là ngày {day} tháng {month} năm {year}
        Mặt trời mọc vào {hourrise} giờ {minrise} phút
        Mặt trời lặn vào {hourset} giờ {minset} phút
        Nhiệt độ trung bình là {temp} độ C
        Áp suất không khí là {pressure} héc tơ Pascal
        Độ ẩm là {humidity}%.""".format(day = now.day,month = now.month, year= now.year, hourrise = sunrise.hour, minrise = sunrise.minute,
                                                                           hourset = sunset.hour, minset = sunset.minute, 
                                                                           temp = current_temperature, pressure = current_pressure, humidity = current_humidity)
        noi(content)
    else:
        noi("Không tìm thấy địa chỉ của bạn")



def play_song():
    noi('Xin mời bạn chọn tên bài hát')
    mysong = get_text()
    while True: 
        result = YoutubeSearch(mysong, max_results=10).to_dict()
        if result:
            break
    url = 'https://www.youtube.com' + result[0]['url_suffix']
    webbrowser.open(url)
    noi("Bài hát bạn yêu cầu đã được mở.")





def change_wallpaper():
    api_key = 'RF3LyUUIyogjCpQwlf-zjzCf1JdvRwb--SLV6iCzOxw'
    url = 'https://api.unsplash.com/photos/random?client_id=' + \
        api_key  
    f = urllib2.urlopen(url)
    json_string = f.read()
    f.close()
    parsed_json = json.loads(json_string)
    photo = parsed_json['urls']['full']
    
    date_string = datetime.now().strftime("%d%m%Y%H%M%S")
    urllib2.urlretrieve(photo, "C:/Users/PC/Downloads/a"+date_string+".png")
    image=os.path.join("C:/Users/PC/Downloads/a"+date_string+".png")
    ctypes.windll.user32.SystemParametersInfoW(20,0,image,3)
    noi('Hình nền máy tính vừa được thay đổi')

def read_news(text):
    #noi("Bạn muốn đọc báo về gì")
    #print("Bạn muốn đọc báo về gì")
    queue = text
    params = {
        'apiKey': '30d02d187f7140faacf9ccd27a1441ad',
        "q": queue,
    }
    api_result = requests.get('http://newsapi.org/v2/top-headlines?', params)
    api_response = api_result.json()
    print("Tin tức")

    for number, result in enumerate(api_response['articles'], start=1):
        print(f"""Tin {number}:\nTiêu đề: {result['title']}\nTrích dẫn: {result['description']}\nLink: {result['url']}
    """)
        if number <= 3:
            webbrowser.open(result['url'])



def tell_me_about():
    try:
        noi("Bạn muốn nghe về gì ạ")
        text = get_text()
        wikipedia.set_lang("vi")
        contents = wikipedia.summary(text).split('\n')
        noi(contents[0])
        
        for content in contents[1:]:
            noi("Bạn muốn nghe thêm không")
            ans = get_text()
            if "có" not in ans:
                break    
            noi(content)
            

        noi('Cảm ơn bạn đã lắng nghe!!!')
    except:
        noi("Tôi không định nghĩa được thuật ngữ của bạn. Xin mời bạn nói lại")


def number(text):
    if "0" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[4]/div[5]/span").click()
    if "1" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[3]/div[6]/span").click()
    if "2" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[3]/div[7]/span").click()
    if "3" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[3]/div[8]/span").click()
    if "4" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[6]/span").click()
    if "5" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[7]/span").click()
    if "6" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[8]/span").click()
    if "7" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[6]/span").click()
    if "8" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[7]/span").click()
    if "9" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[8]/span").click()

def click(text):
    if "x" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[1]/span").click()
    if "y" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[2]/span").click()
    if "mu222" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[3]/span").click()
    if "mu" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[4]/span").click()
    if "(" in text or "mo" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[1]/span").click()
    if ")" in text or "dong" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[2]/span").click()
    if "|" in text or "doi" in text or "đoi" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[3]/div[1]/span").click()
    if "can" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[4]/div[2]/span").click()
    if "pi" in text or "bi" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[4]/div[3]/span").click()
    if "cong" in text or "+" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[4]/div[8]/span").click()
    if "tru" in text or "-" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[3]/div[9]/span").click()
    if "nhan" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[9]/span").click()
    if "chia" in text or "/" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[9]/span").click()
    if "thoat" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[2]/div[12]/span").click()
    if text.isdigit():
        for i in text:
            number(i)
    if "sin" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[11]/span").click()
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/div[1]/span").click()
    if "cos" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[11]/span").click()
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/div[4]/span").click()
    if "tan" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[11]/span").click()
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/div[7]/span").click()
    if "cot" in text:
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[1]/div[1]/div[11]/span").click()
        browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/div[16]/span").click()


def help_me():
    noi("""Tôi có thể giúp bạn thực hiện các câu lệnh sau đây:
    1. Chào hỏi
    2. Hiển thị giờ
    3. Mở website, application
    4. Tìm kiếm trên Google
    5. Gửi email
    6. Dự báo thời tiết
    7. Mở video nhạc
    8. Thay đổi hình nền máy tính
    9. Đọc báo hôm nay
    10. Kể bạn biết về thế giới """)                

noi("Xin chào, bạn tên là gì nhỉ?")
name = get_text()
if name:
	noi("Chào bạn {}".format(name))
	noi("Tôi có thể giúp gì được cho bạn ạ?")
	while True:
		text = get_text()
		if not text:
			break
		elif "dừng" in text or "tạm biệt" in text or "chào robot" in text or "ngủ thôi" in text:
			stop()
			break
		elif "có thể làm gì" in text:
			help_me()
		elif "chào trợ lý ảo" in text:
			hello(name)
		elif "hiện tại" in text:
			time(text)
		elif "mở" in text:
			if 'mở google và tìm kiếm' in text:
				open_google_and_search(text)
			elif "." in text:
				open_website(text)
			else:
				open_application(text)
		elif "email" in text or "mail" in text or "gmail" in text:
			send_email()
		elif "thời tiết" in text:
			current_weather()
		elif "nghe nhạc" in text:
			play_song()
		elif "hình nền" in text:
			change_wallpaper()
		elif "đọc báo" in text:
			noi("Bạn muốn đọc báo về gì")
			read_news(get_text())
		elif "định nghĩa" in text:
			tell_me_about()
		elif "bảng điểm" in text:
			score()
		elif " vẽ đồ thị" in text:
			browser = webdriver.Chrome(executable_path="D:\project4\chromedriver.exe")
			browser.get("https://www.desmos.com/calculator?lang=vi")
			sleep(3)
			choose_keyboard = browser.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[7]/div[2]/div/div/i[2]")
			choose_keyboard.click()
			sleep(3)
			noi("phương trình bạn cần tôi vẽ là gì")
			for i in unidecode(get_text()).split():
				click(i)
			sleep(10)
			browser.quit()
		else:
			noi("Bạn cần tôi giúp gì ạ?")
