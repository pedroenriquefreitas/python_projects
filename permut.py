import bs4
import requests

res = requests.get('https://www.linkedin.com/sales/gmail/profile/viewByEmail/TLopes@netflix.com')
res.raise_for_status()

soup = bs4.BeautifulSoup(res.text)
elems = soup.select('#li-profile-name')
