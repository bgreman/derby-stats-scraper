# derby-stats-scraper
Various scripts for collecting derby data

# Powershell
In the Powershell directory are a few Powershell scripts for parsing collections of statsbooks.  Just drop the script in the same directory as your statsbooks and run via Powershell.

# Python
## rankings_scraper.py
This is a script that scrapes the stats.wftda.com/rankings site and exports the critical rankings data to a csv file in the same directory as the script.

### Requirements
Requires you to have installed the `requests` and `BeautifulSoup4` Python modules.

### Use
Open a command prompt, navigate to the directory where `rankings_scraper.py` is and run it using `python rankings_scraper.py`

## post_season_predictor.py
This is a script that scrapes the stats.wftda.com/rankings-live and comes up with post-season placements as if the live rankings at the time of run were the ones used to seed post-season tournaments.

### Requirements
Requires you to have installed the `requests` and `BeautifulSoup4` Python modules.

### Use
Open a command prompt, navigate to the directory where `post_season_predictor.py` is and run it using `python post_season_predictor.py`
