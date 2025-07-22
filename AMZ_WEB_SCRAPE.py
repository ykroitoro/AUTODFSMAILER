import requests
from bs4 import BeautifulSoup

def fetch_amazon_search_results(search_query, number_of_results=20):
    base_url = "https://www.amazon.com/s?k="
    url = base_url + search_query.replace(" ", "+")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.5",
    }
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "lxml")
        results = []
        items = soup.select("div[data-index]")

        for item in items[:number_of_results]:
            title = item.select_one("span.a-size-medium")
            link = item.select_one("a.a-link-normal")

            if title and link and link["href"].startswith("/dp/"):
                results.append({
                    "title": title.text.strip(),
                    "link": "https://www.amazon.com" + link["href"]
                })

        return results

if __name__ == "__main__":
    search_query = input("Enter the search query: ")
    results = fetch_amazon_search_results(search_query)

    if results:
        for idx, result in enumerate(results, start=1):
            print(f"{idx}. {result['title']}\n{result['link']}\n")
    else:
        print("No results found.")
