import requests
from bs4 import BeautifulSoup
import sqlite3
import json
import csv
import argparse
from datetime import datetime
from openpyxl import Workbook



SOURCES = {
    "bbc": {
        "url": "https://www.bbc.com/news",
        "selector": "h2"
    },

    "guardian": {
        "url": "https://www.theguardian.com/international",
        "selector": "h2"
    },

    "reuters": {
        "url": "https://www.reuters.com/world/",
        "selector": "h2"
    }
}



def create_database():
    conn = sqlite3.connect("news.db")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS news (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            source TEXT,
            url TEXT,
            date TEXT,
            UNIQUE(title, source)
        )
    """)

    conn.commit()
    conn.close()



def scrape_source(source_name, source_info):

    print(f"\nFetching news from {source_name.upper()}...")

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:
        response = requests.get(
            source_info["url"],
            headers=headers,
            timeout=10
        )

        response.raise_for_status()

    except requests.RequestException as e:
        print(f"Error fetching {source_name}: {e}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")

    headlines = soup.select(source_info["selector"])

    results = []

    for headline in headlines:

        title = headline.get_text(strip=True)

        if not title:
            continue

        link = headline.find("a")

        if link and link.get("href"):
            url = link["href"]

            if url.startswith("/"):
                url = "https://www.bbc.com" + url
        else:
            url = source_info["url"]

        results.append({
            "title": title,
            "source": source_name,
            "url": url,
            "date": datetime.now().strftime("%Y-%m-%d")
        })

    return results


def collect_news():

    all_news = []

    for source_name, source_info in SOURCES.items():

        news = scrape_source(source_name, source_info)

        all_news.extend(news)

    print(f"\nCollected {len(all_news)} headlines.")

    return all_news


def remove_duplicates(news):

    unique_news = []
    seen = set()

    for item in news:

        key = item["title"].lower().strip()

        if key not in seen:
            seen.add(key)
            unique_news.append(item)

    return unique_news




def save_to_database(news):

    conn = sqlite3.connect("news.db")
    cursor = conn.cursor()

    added = 0

    for item in news:

        try:

            cursor.execute("""
                INSERT INTO news
                (title, source, url, date)
                VALUES (?, ?, ?, ?)
            """, (
                item["title"],
                item["source"],
                item["url"],
                item["date"]
            ))

            added += 1

        except sqlite3.IntegrityError:
            # Duplicate article
            pass

    conn.commit()
    conn.close()

    print(f"Added {added} new articles to database.")



def export_json(news):

    with open("news.json", "w", encoding="utf-8") as file:

        json.dump(
            news,
            file,
            indent=4,
            ensure_ascii=False
        )

    print("Exported to news.json")



def export_csv(news):

    with open(
        "news.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=[
                "title",
                "source",
                "url",
                "date"
            ]
        )

        writer.writeheader()

        writer.writerows(news)

    print("Exported to news.csv")




def export_excel(news):

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "News"

    headers = [
        "Title",
        "Source",
        "URL",
        "Date"
    ]

    sheet.append(headers)

    for item in news:

        sheet.append([
            item["title"],
            item["source"],
            item["url"],
            item["date"]
        ])

    workbook.save("news.xlsx")

    print("Exported to news.xlsx")




def load_news():

    conn = sqlite3.connect("news.db")

    cursor = conn.cursor()

    cursor.execute("""
        SELECT title, source, url, date
        FROM news
        ORDER BY date DESC
    """)

    rows = cursor.fetchall()

    conn.close()

    news = []

    for row in rows:

        news.append({
            "title": row[0],
            "source": row[1],
            "url": row[2],
            "date": row[3]
        })

    return news




def filter_news(news, source=None, keyword=None, date=None):

    filtered = news

    if source:

        filtered = [
            item for item in filtered
            if item["source"].lower() == source.lower()
        ]

    if keyword:

        filtered = [
            item for item in filtered
            if keyword.lower() in item["title"].lower()
        ]

    if date:

        filtered = [
            item for item in filtered
            if item["date"] == date
        ]

    return filtered




def display_news(news):

    if not news:

        print("\nNo news found.")

        return

    print("\n" + "=" * 80)

    for index, item in enumerate(news, start=1):

        print(f"\n{index}. {item['title']}")
        print(f"   Source : {item['source']}")
        print(f"   Date   : {item['date']}")
        print(f"   URL    : {item['url']}")

    print("\n" + "=" * 80)




def main():

    parser = argparse.ArgumentParser(
        description="News Aggregator CLI"
    )

    parser.add_argument(
        "--collect",
        action="store_true",
        help="Collect latest news"
    )

    parser.add_argument(
        "--source",
        help="Filter by source"
    )

    parser.add_argument(
        "--keyword",
        help="Search headlines by keyword"
    )

    parser.add_argument(
        "--date",
        help="Filter by date YYYY-MM-DD"
    )

    parser.add_argument(
        "--export",
        choices=["json", "csv", "excel"],
        help="Export news"
    )

    args = parser.parse_args()

    create_database()

    # Collect fresh news
    if args.collect:

        news = collect_news()

        news = remove_duplicates(news)

        save_to_database(news)

    # Load stored news
    news = load_news()

    # Apply filters
    news = filter_news(
        news,
        source=args.source,
        keyword=args.keyword,
        date=args.date
    )


    display_news(news)

    
    if args.export == "json":
        export_json(news)

    elif args.export == "csv":
        export_csv(news)

    elif args.export == "excel":
        export_excel(news)


if __name__ == "__main__":
    main()
