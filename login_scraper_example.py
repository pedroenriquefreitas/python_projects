import requests
from lxml import html

USERNAME = "pefandrade@hotmail.com"
PASSWORD = "pefa1997"

LOGIN_URL = "https://www.linkedin.com/uas/login"
URL = "https://www.linkedin.com/feed/"

def main():
    session_requests = requests.session()

    # Get login csrf token
    result = session_requests.get(LOGIN_URL)
    tree = html.fromstring(result.text)
    authenticity_token = list(set(tree.xpath("//*[@id='app__container']/div[1]/div/form/input[1]/@value")))[0]

    # Create payload
    payload = {
        "session_key": USERNAME,
        "session_password": PASSWORD,
        "csrfToken": authenticity_token
    }

    # Perform login
    result = session_requests.post(LOGIN_URL, data = payload, headers = dict(referer = LOGIN_URL))
    print(result.ok)
    print(result.status_code)

    # Scrape url
    result = session_requests.get(URL, headers = dict(referer = URL))
    tree = html.fromstring(result.content)
    bucket_names = tree.xpath("//*[@id='ember139']/span/text()")

    print(bucket_names)

if __name__ == '__main__':
    main()
