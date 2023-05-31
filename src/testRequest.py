import requests
url = 'https://scholar.google.com/citations?view_op=view_citation&hl=en&user=iR8mYs0AAAAJ&citation_for_view=iR8mYs0AAAAJ:qUcmZB5y_30C'
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
headers = {
    'User-Agent': user_agent
}
response = requests.get(url, headers=headers)
print(response.text)