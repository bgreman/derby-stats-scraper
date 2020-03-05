import requests
from bs4 import BeautifulSoup

stats_url = "https://stats.wftda.com/rankings-live"
resp = requests.get(stats_url)

soup = BeautifulSoup(resp.content, 'html.parser')

rankings_table = soup.find(class_="rankingsTable")
rankings_body = rankings_table.tbody

data = []

for row in rankings_body.contents[1:None:2]:
    rank = row.find(class_="rankingsTable--position").string.strip()
    league = row.find(class_="rankingsTable--leagueTitleColumn").a.string.strip()
    region = row.find(class_="rankingsTable--region").a.string.strip()
    gpa_td = row.find(class_="rankingsTable--gpa")
    gpa = next(gpa_td.strings).strip()
    weight = gpa_td.span.string.strip()
    data.append({"rank": rank, "league": league, "region": region, "gpa": gpa})
    
direct_to_champs = [league for league in data if int(league["rank"]) < 5]
print(f"Champs Bye (Top {len(direct_to_champs)}):")
for league in direct_to_champs:
    print(f"\t#{league['rank']} {league['league']} {league['gpa']}")
    
print("")

d1 = [league for league in data if int(league["rank"]) > 4 and int(league["rank"]) < 29]
print(f"D1 (Next {len(d1)}):")
for league in d1:
    print(f"\t#{league['rank']} {league['league']} {league['gpa']}")
    
print("")
    
na_west = [league for league in data if int(league["rank"]) > 28 and "N. America West" in league["region"]]
na_west_cup = na_west[0:12]
print(f"NA West Cup (Next {len(na_west_cup)} in NA West):")
for league in na_west_cup:
    print(f"\t#{league['rank']} {league['league']} {league['gpa']}")
    
print("")
    
na_east = [league for league in data if int(league["rank"]) > 28 and "N. America East" in league["region"]]
na_east_cup = na_east[0:12]
print(f"NA East Cup (Next {len(na_east_cup)} in NA East):")
for league in na_east_cup:
    print(f"\t#{league['rank']} {league['league']} {league['gpa']}")
    
print("")
    
europe = [league for league in data if int(league["rank"]) > 28 and "Europe" in league["region"]]
europe_cup = europe[0:8]
print(f"Europe Cup (Next {len(europe_cup)} in Europe):")
for league in europe_cup:
    print(f"\t#{league['rank']} {league['league']} {league['gpa']}")