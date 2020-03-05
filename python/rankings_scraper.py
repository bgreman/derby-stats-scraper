import csv
import requests
from bs4 import BeautifulSoup

stats_url = "https://stats.wftda.com/rankings"

resp = requests.get(stats_url)

soup = BeautifulSoup(resp.content, 'html.parser')

rankings_table = soup.find(class_="rankingsTable")
rankings_body = rankings_table.tbody

data = []

for row in rankings_body.contents[1:None:2]:
    rank = row.find(class_="rankingsTable--position").string.strip()
    league = row.find(class_="rankingsTable--leagueTitleColumn").a.string.strip()
    gpa_td = row.find(class_="rankingsTable--gpa")
    gpa = next(gpa_td.strings).strip()
    weight = gpa_td.span.string.strip()
    print(f"#{rank} {league} has gpa {gpa} and weight {weight}")
    data.append({"rank": rank, "league": league, "gpa": gpa, "weight": weight})
    
print(f"Total # of items: {len(data)}")
    
with open('rankings_data.csv', 'w', newline='') as csvfile:
    fieldnames = ['rank', 'league', 'gpa', 'weight']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for item in data:
        writer.writerow(item)