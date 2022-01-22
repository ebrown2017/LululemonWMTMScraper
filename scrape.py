import requests
import json

def scrape(nValue, category):
    # open JSON query data
    jsonFile = open('query.json')
    jsonData = json.load(jsonFile)

    # convert query to scrape specified category
    jsonData['variables']['nValue'] = nValue
    jsonData['variables']['category'] = category

    # convert JSON data to sting to format
    payload = json.dumps(jsonData)
    payload = payload.replace('\n', '\\n')

    # url to send the queries to
    graphQLUrl = 'https://shop.lululemon.com/snb/graphql'

    # url to get a cookie from
    cookieUrl = 'https://shop.lululemon.com/'

    # create session
    session = requests.Session()

    # get cookie url to get cookie to use for query
    response = session.get(cookieUrl)

    # get cookie and convert to string
    cookies = session.cookies.get_dict()
    userCookie = "; ".join([str(x)+"="+str(y) for x,y in cookies.items()])

    # set headers for post
    headers = {
        "Content-type": "application/json",
        "Authority": "shop.lululemon.com",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36",
        'cookie': userCookie
    }

    response = session.post(graphQLUrl, headers=headers, json=json.loads(payload))

    return [nValue, category, response.json()]