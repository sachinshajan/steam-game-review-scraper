from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.keys import Keys
import re
from time import sleep
from datetime import datetime
from openpyxl import Workbook
import csv


game_id = #enter your game ID here
#you can alter the URL as you like,here I have set it accordingly so that I can get all the positive reviews in all the language
template = 'https://steamcommunity.com/app/{}/positivereviews/?browsefilter=toprated&p=1&filterLanguage=all'
template_with_language = 'https://steamcommunity.com/app/{}/positivereviews/?browsefilter=toprated&p=1&filterLanguage=all'
url = template_with_language.format(game_id)

# setup driver
options = EdgeOptions()
options.use_chromium = True
driver = Edge(options=options)

driver.maximize_window()
driver.get(url)

# get current position of y scrollbar
last_position = driver.execute_script("return window.pageYOffset;")

reviews = []
review_ids = set()
running = True

while running:
    # get cards on the page
    cards = driver.find_elements_by_class_name('apphub_Card')

    for card in cards[-20:]:  # only the tail end are new cards

        # gamer profile url
        profile_url = card.find_element_by_xpath('.//div[@class="apphub_friend_block"]/div/a[2]').get_attribute('href')

        # steam id
        steam_id = profile_url.split('/')[-2]

        # check to see if I've already collected this review
        if steam_id in review_ids:
            continue
        else:
            review_ids.add(steam_id)

        # username
        user_name = card.find_element_by_xpath('.//div[@class="apphub_friend_block"]/div/a[2]').text

        # language of the review
        date_posted = card.find_element_by_xpath('.//div[@class="apphub_CardTextContent"]/div').text
        review_content = card.find_element_by_xpath('.//div[@class="apphub_CardTextContent"]').text.replace(date_posted,
                                                                                                            '').strip()

        # review length
        review_length = len(review_content.replace(' ', ''))

        # recommendation
        thumb_text = card.find_element_by_xpath('.//div[@class="reviewInfo"]/div[2]').text
        thumb_text

        # amount of play hours
        play_hours = card.find_element_by_xpath('.//div[@class="reviewInfo"]/div[3]').text
        play_hours

        # save review
        review = (steam_id, profile_url, review_content, thumb_text, review_length, play_hours, date_posted)
        reviews.append(review)

        # attempt to scroll down thrice.. then break
    scroll_attempt = 0
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(0.9)
        curr_position = driver.execute_script("return window.pageYOffset;")

        if curr_position == last_position:
            scroll_attempt += 1
            sleep(0.9)

            if curr_position >= 3:
                running = False
                break
        else:
            last_position = curr_position
            break  # continue scraping the results

# shutdown the web driver
driver.close()

# save the file to Excel Worksheet
wb = Workbook()
ws = wb.worksheets[0]
ws.append(['SteamId', 'ProfileURL', 'ReviewText', 'Review', 'ReviewLength(Chars)', 'PlayHours', 'DatePosted'])
for row in reviews:
    ws.append(row)

today = datetime.today().strftime('%Y%m%d')
wb.save(f'Steam_Reviews_{game_id}_{today}.xlsx')
wb.close()

# save the file to a CSV file
today = datetime.today().strftime('%Y%m%d')
with open(f'Steam_Reviews_{game_id}_{today}.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['SteamId', 'ProfileURL', 'ReviewText', 'Review', 'ReviewLength(Chars)', 'PlayHours', 'DatePosted'])
    writer.writerows(reviews)