![WhatsApp Image 2023-01-20 at 16 23 38](https://user-images.githubusercontent.com/43014726/213787899-16afd6a8-2b72-42bd-9ba6-7cc86a1522f9.jpeg)
![WhatsApp Image 2023-01-20 at 16 23 51](https://user-images.githubusercontent.com/43014726/213787901-42b2ac7d-5ba0-4d59-8a91-6bf882d16e8a.jpeg)
![WhatsApp Image 2023-01-20 at 15 06 37](https://user-images.githubusercontent.com/43014726/213787903-84c0052a-e3ae-4834-a3a7-5519af3dec59.jpeg)
![WhatsApp Image 2023-01-20 at 16 20 49](https://user-images.githubusercontent.com/43014726/213787905-944d5c6b-37cc-4679-8b05-a1a21b1196fc.jpeg)

# AUTOMATION WEB TO GET THE VARIATION OF COINS AROUND THE WORLD

### About This Project:
This project will

### My Evolution On This Project:
This was the first project that i build using the selenium extension to automate proccess, i had used the extension 'Pyautogui' before, but it´s not a good practice simule one click, also that, the selenium can offer to us another benefits that the pyautogui don't oofer to us, like ; 
<ul>
  <li>We can run this application in second place, this is, i can do another things on my computer while the script is running, the code will not be affected with this</li>
  <li>I put this code on AWS server, so making this way, the code will run in one server and will continue sending me emails about the variations of the coins</li>
  <li>If i want, i can become this application one executable file, with this, i can run this code in one machine that have no python installed for example</li>
</ul>

Also that, i used the library 'pandas for the first time, the library pandas was very used to work with datas, in this case, with use the pandas to read our sheet, update the datas in our sheet and localizate the datas

### Technologies Used In This Project And Yours Functions:

| Technologies | Function |
| ----------- | ----------- |
| Selenium & ChromeDriver | Selenium was used to make the web scrapping, with selenium and chromedriver i search for the variations of the coins of USA, China, Europe and Gold between websites, after that, i save the information that i get on this websites in my coder |
| Python | The language that i choose to build this code |
| Sheets | I used the sheets on this project to armazenate the datas of products, puurchase prices, sales prices and margin of sales |
| Pandas | Pandas is one library very used on github to manipulate datas, he can read files with extensions in csv, pdf, html, xlsx (our case in this project), with her we update the datas in our sheets and generate another sheet with the values refresheds to send to email |
| Win32.com | The library Win32.com was used to send emails on outlook |
| DateTime | Used to get the atual date |
| Time | The library 'time' was used to set the interval of teen seconds between the interations in website |

### Observations Of This Project:
If you want reply this project on your computer, you have to install the libraries above: 

~~~ List of libraries that you have to install in your computer to run this application
pip install selenium
pip install pandas
pip install numpy
pip install openpyxl
~~~

~~~Also that, you will need install the 'chromedriver', verify your google chrome version and install the chromedriver through the link above, after downloaded, you have to copy and paste the project in the same folder that you create the file python that will run you application
https://chromedriver.chromium.org/downloads
~~~

If you use the Firefox, don't worry, you can make the same step searching on google for 'geckodriver'. 

### Verify Project On Your Computer:
You can clone this project and run in your computer.
