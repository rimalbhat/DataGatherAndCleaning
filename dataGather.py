import requests
from bs4 import BeautifulSoup as bs
baseUrl = "https://www.eia.gov/dnav/pet/"
url = baseUrl + "pet_pnp_inpt_a_epc0_yir_mbbl_m.htm"
urlResponse = requests.get(url)
soup = bs(urlResponse.content, 'html5lib')
select_tag = soup.find("select", attrs={"class":"C"})
options = select_tag.find_all("option")

urlWithNeededUnits = ""

for option in options:
    if option.text.strip() == "Monthly-Thousand Barrels per Day":
        urlWithNeededUnits = baseUrl + option['value']

rightUrlResponse = requests.get(urlWithNeededUnits)
mainSoup = bs(rightUrlResponse.content, 'html5lib')
downloadHref = mainSoup.find("a", attrs={"class":"crumb"})['href']
downloadUrl = baseUrl + downloadHref

downloadRequest = requests.get(downloadUrl, allow_redirects=True)
open('data.xls', 'wb').write(downloadRequest.content)