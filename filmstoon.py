import requests
from bs4 import BeautifulSoup
from datetime import datetime
from xlwt import Workbook


def get_web_link(titles):
    titles_list = titles.split(',')
    while True:  # Loop runs until user inputs valid title match setting
        match = input('Select title match type: partial / exact ')
        if match.lower() in ['exact', 'e', 'partial', 'p']:
            break
    web_links = []
    for search_title in titles_list:
        page = requests.get("https://filmstoon.in/?s=" + search_title)
        soup = BeautifulSoup(page.content, 'html.parser')
        try:
            n_pages = soup.find(class_='pagination').find_all('li')[-1].a.get('href')  # Getting total number of pages
            n_pages = int(n_pages.split('/page/')[1].split('/?s')[0])
        except AttributeError:
            n_pages = 1
        for page_number in range(1, n_pages + 1):
            page = requests.get('https://filmstoon.in/page/' + str(page_number) + '/?s=' + search_title)
            soup = BeautifulSoup(page.content, 'html.parser')
            list_full = soup.find(class_='movies-list movies-list-full')
            list_items = list_full.find_all(class_='ml-item')  # List of all results on page
            for list_item in list_items:  # Loops through all results and appends the link if title matches
                if match.lower() == "partial" or match.lower() == 'p':  # Matches part of title
                    if search_title.lower() in list_item.a.get('oldtitle').lower():
                        web_links.append(list_item.a.get('href'))
                elif match.lower() == "exact" or match.lower() == 'e':  # Matches the whole title
                    if search_title.lower() == list_item.a.get('oldtitle').lower():
                        web_links.append(list_item.a.get('href'))

    return web_links  # Returns a list of web links


def get_direct_link(web_links):  # Gets the direct link for every web link in the list
    for link in web_links:
        page = requests.get(link)
        soup = BeautifulSoup(page.content, 'html.parser')
        seasons = soup.find_all(class_='tvseason')
        if seasons:  # Runs if matched media is a TV show
            for season in seasons:
                episodes = season.find_all('a')
                for episode in episodes:
                    episode_link = episode.get('href')
                    page = requests.get(episode_link)
                    soup = BeautifulSoup(page.content, 'html.parser')
                    title = soup.find(class_="mvic-desc").h3.text
                    direct_link = soup.find('iframe').get('src')
                    write_output(episode_link, direct_link, title)

        else:  # Runs if matched media is a movie
            title = soup.find(class_="mvic-desc").h3.text
            direct_link = soup.find('iframe').get('src')
            write_output(link, direct_link, title)


def write_output(web_link, direct_link, title):  # Writes content info to spreadsheet
    if '.jpg' not in direct_link and '.png' not in direct_link:
        now = datetime.now()
        i = len(sheet1._Worksheet__rows)
        dt_string = now.strftime("%m/%d/%Y %H:%M")
        sheet1.write(i, 0, dt_string)
        sheet1.write(i, 1, web_link)
        sheet1.write(i, 2, direct_link)
        sheet1.write(i, 3, title)
        wb.save('output.xls')


titles = input('Input titles to be searched separated by commas: ')
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'Time')
sheet1.write(0, 1, 'Web link')
sheet1.write(0, 2, 'Direct link')
sheet1.write(0, 3, 'Title')
wb.save('output.xls')

web_links = get_web_link(titles)
get_direct_link(web_links)

