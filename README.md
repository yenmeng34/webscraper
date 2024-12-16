This is a web scraping program designed for personal use only. We will not take any responsibility for its use in commercial applications or any consequences thereof.
It allows users to extract data from websites efficiently and store the results in a structured format (Excel and JSON file) for further analysis or use. 
The program is flexible and can be customized to suit various web scraping needs.

Requirements
To run this program, you need the following:

- Python 3.11+

- Required Python libraries:
    Selenium
    Pandas (optional, for data manipulation and storage)

Install the dependencies using:
    - pip install selenium
    - pip install pandas



Additionally, you need to install a WebDriver for Selenium. Follow these steps below:
Selenium WebDriver Installation

1. Choose a Browser: Decide which browser you want to use for scraping (e.g., Chrome, Firefox).

2. Download WebDriver:
      - For Chrome: Download ChromeDriver
      - For Firefox: Download GeckoDriver

3. Install the WebDriver:
      - Place the downloaded WebDriver in a directory included in your system's PATH, or specify its location in your script. # Recommend download to C drive
      - Verify Installation: Test if the WebDriver works by running:

        from selenium import webdriver

        driver = webdriver.Chrome()  # or webdriver.Firefox()
        driver.get("https://example.com")
        print(driver.title)
        driver.quit()



LIMITATION and USAGE
1. This program is intended for personal use only. Please respect the terms and conditions of the websites you scrape.
2. It does not handle advanced anti-bot mechanisms like CAPTCHAs.
3. Please ensure follow Robots.txt webscraping rules. Web scraping may violate the terms of service of some websites. Use this program responsibly and ensure compliance with applicable laws and regulations.
