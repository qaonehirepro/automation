from urllib.request import urlopen
from bs4 import BeautifulSoup

# url = "https://s3.ap-southeast-1.amazonaws.com/test-all-hirepro-files/AT/candidateFiles/9c3e3fa4-e868-4835-afaf-77b4f9c4e667cf_versant_domains.txt?AWSAccessKeyId=AKIAIPKT65FIXSXQGDOQ&Signature=BIA0OQd93O6%2BcKK2v0VbIsBXtn0%3D&Expires=1657027638"
url = 'https://s3.ap-southeast-1.amazonaws.com/test-all-hirepro-files/AT/candidateFiles/9c3e3fa4-e868-4835-afaf-77b4f9c4e667cf_versant_domains.txt?AWSAccessKeyId=AKIAIPKT65FIXSXQGDOQ&Signature=BIA0OQd93O6%2BcKK2v0VbIsBXtn0%3D&Expires=1657027638'
html = urlopen(url).read()
soup = BeautifulSoup(html, features="html.parser")

# kill all script and style elements
# for script in soup(["script", "style"]):
#     script.extract()    # rip it out

# get text
text = soup.get_text()

# break into lines and remove leading and trailing space on each
lines = (line.strip() for line in text.splitlines())
# break multi-headlines into a line each
chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
# drop blank lines
text = '\n'.join(chunk for chunk in chunks if chunk)

print(text)
